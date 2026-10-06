/**
 * Landing.js — 年度の着地見込み（統計）と、見通しの空模様（天気）。計画ごとにサーバーで計算する（ここではシートを読まない。計算だけ）。
 *
 * 旧来の計算の月の予測（OUTPUT の 29〜40 行の P10/P50/P90）は、その年度の実績を使わない（予測の月と過去売上の月が重ならない）。
 * そこで、締まった月の実績から「今年の水準」θ を学び、残りの月に伸ばす（共役の正規分布の更新。モンテカルロは使わない）:
 *   実績 a_m = θ・f_m + e_m（e_m ~ N(0, σ_m²)、月ごとに独立）。f_m = その月の P50（計画値 = P50 の約束に合わせる）
 *   σ_m = w・max((P90 − P10) / 2.5631, 0.05・F / 12)（F = 12 か月の P50 の合計。幅の無い月の下限）
 *   事前分布 θ ~ N(1, τ²)（学んだ偏りの補正は f_m に入っているので 1 を中心にする）。新しい月ほど重い（半減期 4 か月。旧来の B-5 と同じ）
 *   着地 L = 締まった月の実績の合計 + θ̂・残りの月の P50 の合計
 *   ばらつき sd² = 残りの P50 の合計² / 精度（水準のぶれ。残りの月に共通）+ 残りの月の σ² の合計（月ごとのぶれ）
 * τ と w は、全計画の検証の表（EVAL_COMPARE_MONTHLY）から学ぶ（appLandingPrior_。学べなければ 0.15 と 1）。
 * 空模様は着地 ÷ 年間予算で分ける。急な変化（天変地異）は、ひと月の大きな外れ・2 か月続いた同じ向きの外れ・予測の前提の大きな変化で拾う。
 */
const APP_LANDING = {
  TAU0: 0.15,              // 今年の水準 θ の事前分布の標準偏差（全計画から学べるまで）
  TAU_MIN: 0.05, TAU_MAX: 0.35,
  HALF_LIFE: 4,            // 月。APP_BIAS_HALF_LIFE（Learning.js）・旧来の B-5 と同じ
  Z80: 1.2815515655446004, // 標準正規分布の 90% 点（P10〜P90 = ±1.2816σ）
  FLOOR_REL: 0.05,         // 月の σ の下限 = 月の平均の予測の 5%
  THETA_MAX: 5,
  STALE_MONTHS: 3,         // 締まった年度の月のうち、実績が取り込まれていない月がこれだけあれば霧（旧来の C-1 の最低 3 か月と同じ）
  HOT: 1.5, CLEAR: 1.1, FAIR: 0.9, CLOUD: 0.5,   // 着地 ÷ 予算の境目（猛暑・快晴・晴れのち曇り・曇り。それより下は雨）
  SHOCK_Z: 3, SHIFT_Z: 2, MATERIAL: 0.05,        // 天変地異: ひと月 |z| ≥ 3、2 か月続けて |z| ≥ 2（同じ向き）。着地 ÷ 予算が 0.05 以上動いたときだけ
  PREMISE: 0.3, PREMISE_DAYS: 31, PREMISE_MATERIAL: 0.1,   // 前提の変化: 直近 2 回の予測の年間 P50 が 30% 以上違う（31 日以内・予算の 10% 以上）
  PRIOR_MIN_PLANS: 3, PRIOR_MIN_MONTHS: 3, W_MIN_MONTHS: 12, W_MAX: 3
};

/** 標準正規分布の累積分布関数（Abramowitz–Stegun 7.1.26。誤差 1.5e-7 未満） */
function appNormCdf_(x) {
  const t = 1 / (1 + 0.3275911 * Math.abs(x) / Math.SQRT2);
  const y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x / 2);
  return x >= 0 ? (1 + y) / 2 : (1 - y) / 2;
}

/** 年度の 12 か月（'yyyy/MM'。4 月〜翌 3 月） */
function appLandingFyYms_(fy) {
  const y = Number(fy);
  const out = [];
  for (let i = 0; i < 12; i++) { const m = 4 + i; out.push((m > 12 ? y + 1 : y) + '/' + ('0' + (m > 12 ? m - 12 : m)).slice(-2)); }
  return out;
}

/** θ の共役の正規分布の更新（obs = [{ f, a, s2 }] を古い順に。最後の月の重みが 1）。返り値 { theta, P }（P = 事後の精度） */
function appLandingPost_(obs, tau2) {
  let P = 1 / tau2, num = 1 / tau2;   // 事前分布の中心は 1
  obs.forEach((o, i) => {
    const lam = Math.pow(0.5, (obs.length - 1 - i) / APP_LANDING.HALF_LIFE);
    P += lam * o.f * o.f / o.s2;
    num += lam * o.f * o.a / o.s2;
  });
  return { theta: num / P, P: P };
}

/**
 * 着地見込みと空模様（計画 1 つ。シートは読まない）。
 * inp: { fy, months: [{ ym, p10, p50, p90 }]（OUTPUT の 29〜40 行）, actual: { 'yyyy/MM': 実績 }（検証の表の actual_total）,
 *        cutoffYm（締まった月の境目。この月より前が締まった月。分からなければ ''）, todayYm, budget,
 *        tau・w（全計画から学んだ値。省けば 0.15 と 1）, runs: [{ p50, ageDays }]（直近の予測 2 回。新しい順） }
 * 返り値: { k（締まった月の数）, actualYtd, budget, landing, landingSd, landingP10, landingP90, pAbove（予算以上で着地する確率）, ratio（着地 ÷ 予算）,
 *          theta, credibility（実績の重み 0〜1）, sky, skyReason, skyDir, zLast, dRatio1, dRatio2 }
 * 空模様（先に当てはまったもの）: 霧 mikakunin（no_budget）→ 雪 sekka（zero_sales）→ 霧（no_forecast）→ 霧（stale_actuals）
 *   → 天変地異 tenpen（shock / shift / premise。skyDir = up / down）→ 着地 ÷ 予算（ratio）: 猛暑 mousho ≥ 1.5・快晴 kaisei ≥ 1.1・
 *   晴れのち曇り harenochi > 0.9・曇り kumori > 0.5・雨 ame ≤ 0.5。
 * 着地の数字は、予測が使えれば空模様が霧・雪のときも出す（予算が無ければ ratio と pAbove は出さない）
 */
function appLandingSky_(inp) {
  const C = APP_LANDING;
  const fin = v => typeof v === 'number' && isFinite(v);
  const yms = appLandingFyYms_(inp.fy);
  const byYm = {};
  (inp.months || []).forEach(m => { if (m && m.ym) byYm[m.ym] = m; });
  const fc = yms.map(ym => byYm[ym] || null);
  const F = fc.reduce((s, m) => s + (m && fin(m.p50) ? Math.max(0, m.p50) : 0), 0);
  const fcOk = F > 0 && fc.every(m => m && fin(m.p10) && fin(m.p50) && fin(m.p90));
  const cutoff = String(inp.cutoffYm || '');
  const obsYm = cutoff ? yms.filter(ym => ym < cutoff) : [];   // 年度の月は古い順なので、先頭から k か月
  const k = obsYm.length;
  const act = inp.actual || {};
  const aOf = ym => { const v = Number(act[ym]); return act[ym] === null || act[ym] === '' || !isFinite(v) ? 0 : v; };   // 実績の行が無い締まった月は 0 円
  const A = obsYm.reduce((s, ym) => s + aOf(ym), 0);
  const closedByToday = inp.todayYm ? yms.filter(ym => ym < inp.todayYm).length : k;
  const bn = inp.budget === null || inp.budget === undefined || inp.budget === '' ? NaN : Number(inp.budget);
  const B = isFinite(bn) ? bn : null;
  const out = { k: k, actualYtd: A, budget: B, landing: null, landingSd: null, landingP10: null, landingP90: null, pAbove: null, ratio: null,
    theta: null, credibility: null, sky: 'mikakunin', skyReason: '', skyDir: '', zLast: null, dRatio1: null, dRatio2: null };
  const usable = fcOk || k >= 12;
  const tau = inp.tau > 0 ? inp.tau : C.TAU0;
  const tau2 = tau * tau;
  const w = inp.w > 0 ? inp.w : 1;
  const floor = C.FLOOR_REL * F / 12;
  const sig2 = m => Math.pow(w * Math.max((m.p90 - m.p10) / (2 * C.Z80), floor), 2);
  const obs = fcOk ? obsYm.map((ym, i) => ({ f: Math.max(0, fc[i].p50), a: aOf(ym), s2: sig2(fc[i]) })) : [];
  const clampTheta = x => Math.min(C.THETA_MAX, Math.max(0, x));
  const rest = fcOk ? fc.slice(k) : [];
  const FR = rest.reduce((s, m) => s + Math.max(0, m.p50), 0);
  let post = null;
  if (usable) {
    post = appLandingPost_(obs, tau2);
    const theta = clampTheta(post.theta);
    const L = A + theta * FR;
    const sd = Math.sqrt(FR * FR / post.P + rest.reduce((s, m) => s + sig2(m), 0));
    Object.assign(out, { landing: L, landingSd: sd, landingP10: Math.max(A, L - C.Z80 * sd), landingP90: L + C.Z80 * sd,
      theta: theta, credibility: 1 - (1 / tau2) / post.P });
    if (B !== null && B > 0) Object.assign(out, { ratio: L / B, pAbove: sd > 0 ? appNormCdf_((L - B) / sd) : (L >= B ? 1 : 0) });
  }
  // 1〜4: 霧・雪（数字で分けられない）
  if (!(B !== null && B > 0)) { out.skyReason = 'no_budget'; return out; }
  if (k >= 1 && Math.abs(A) < 1) { out.sky = 'sekka'; out.skyReason = 'zero_sales'; return out; }
  if (!usable) { out.skyReason = 'no_forecast'; return out; }
  if (closedByToday - k >= C.STALE_MONTHS) { out.skyReason = 'stale_actuals'; return out; }
  // 5: 天変地異（月の外れは、その月の前までの実績で立てた見込みと比べる）
  let t1 = false, t2 = false, dir = 0;
  if (fcOk && k >= 1 && k < 12) {
    const before = n => {   // 先頭の n か月の実績だけで立てた見込み
      const p = appLandingPost_(obs.slice(0, n), tau2);
      const th = clampTheta(p.theta);
      const fMid = obs.slice(n).reduce((s, o) => s + o.f, 0);
      return { theta: th, P: p.P, L: obs.slice(0, n).reduce((s, o) => s + o.a, 0) + th * (fMid + FR) };
    };
    const zOf = (o, p) => (o.a - o.f * p.theta) / Math.sqrt(o.f * o.f / p.P + o.s2);
    const b1 = before(k - 1);
    const z1 = zOf(obs[k - 1], b1);
    out.zLast = z1;
    out.dRatio1 = (out.landing - b1.L) / B;
    t1 = Math.abs(z1) >= C.SHOCK_Z && Math.abs(out.dRatio1) >= C.MATERIAL;
    if (t1) dir = z1;
    if (k >= 2) {
      const b2 = before(k - 2);
      const za = zOf(obs[k - 2], b2), zb = zOf(obs[k - 1], b2);
      out.dRatio2 = (out.landing - b2.L) / B;
      t2 = Math.abs(za) >= C.SHIFT_Z && Math.abs(zb) >= C.SHIFT_Z && za * zb > 0 && Math.abs(out.dRatio2) >= C.MATERIAL;
      if (t2 && !dir) dir = zb;
    }
  }
  const r = inp.runs || [];
  const t3 = r.length >= 2 && fin(r[0].p50) && fin(r[1].p50) && r[1].p50 > 0 && fin(r[0].ageDays) && r[0].ageDays <= C.PREMISE_DAYS
    && Math.abs(r[0].p50 / r[1].p50 - 1) >= C.PREMISE && Math.abs(r[0].p50 - r[1].p50) >= C.PREMISE_MATERIAL * B;
  if (t3 && !dir) dir = r[0].p50 - r[1].p50;
  if (t1 || t2 || t3) { out.sky = 'tenpen'; out.skyReason = t1 ? 'shock' : t2 ? 'shift' : 'premise'; out.skyDir = dir > 0 ? 'up' : 'down'; return out; }
  // 6: 着地 ÷ 予算（境目ちょうどの値が、割り算の誤差で隣に落ちないよう丸めてから比べる）
  const q = Math.round(out.ratio * 1e9) / 1e9;
  out.sky = q >= C.HOT ? 'mousho' : q >= C.CLEAR ? 'kaisei' : q > C.FAIR ? 'harenochi' : q > C.CLOUD ? 'kumori' : 'ame';
  out.skyReason = 'ratio';
  return out;
}

/**
 * 全計画で学ぶ τ（今年の水準のばらつき）と w（P10〜P90 の幅の倍率）。plans = [[{ f: P50, a: 実績, p10, p90 }]]（計画ごとの締まった月）。
 * τ: 締まった月が 3 か月以上ある計画が 3 つ以上あれば、計画ごとの水準 y = Σ実績 / ΣP50 − 1 から DerSimonian–Laird で
 *    計画の間のばらつき τ_b² を出し、τ = √(τ_b² + 平均²)（0.05〜0.35）。足りなければ 0.15。
 * w: 計画ごとの水準を除いた外れ |a − θ_i・f| / ((P90 − P10) / 2.5631) の 80% 点 ÷ 1.2816（1〜3）。月が 12 未満なら 1
 */
function appLandingPrior_(plans) {
  const C = APP_LANDING;
  const per = [];
  const zs = [];
  let ss = 0, dof = 0;
  (plans || []).forEach(ms => {
    const use = (ms || []).filter(m => m && typeof m.f === 'number' && m.f > 0 && typeof m.a === 'number' && isFinite(m.a));
    if (use.length < C.PRIOR_MIN_MONTHS) return;
    const Sf = use.reduce((s, m) => s + m.f, 0), Sa = use.reduce((s, m) => s + m.a, 0), Sf2 = use.reduce((s, m) => s + m.f * m.f, 0);
    const th = Sa / Sf;
    use.forEach(m => { ss += m.f * m.f * Math.pow(m.a / m.f - th, 2) / Sf2 * use.length; });
    dof += use.length - 1;
    per.push({ y: th - 1, h: Sf2 / (Sf * Sf) });
    use.forEach(m => {
      if (typeof m.p10 !== 'number' || typeof m.p90 !== 'number') return;
      const s = (m.p90 - m.p10) / (2 * C.Z80);
      if (isFinite(s) && s > 0) zs.push(Math.abs(m.a - th * m.f) / s);
    });
  });
  let tau = C.TAU0, learned = false;
  if (per.length >= C.PRIOR_MIN_PLANS && dof > 0) {
    const s2 = ss / dof;   // 計画の中の、月ごとの 実績 / P50 のばらつき（全計画でまとめる）
    const wt = per.map(p => 1 / Math.max(1e-6, s2 * p.h));
    const W = wt.reduce((a, b) => a + b, 0);
    const mu = per.reduce((s, p, i) => s + wt[i] * p.y, 0) / W;
    const Q = per.reduce((s, p, i) => s + wt[i] * Math.pow(p.y - mu, 2), 0);
    const c = W - wt.reduce((a, b) => a + b * b, 0) / W;
    const tauB2 = Math.max(0, (Q - (per.length - 1)) / (c || 1));
    tau = Math.min(C.TAU_MAX, Math.max(C.TAU_MIN, Math.sqrt(tauB2 + mu * mu)));
    learned = true;
  }
  const w = zs.length >= C.W_MIN_MONTHS ? Math.min(C.W_MAX, Math.max(1, appQuantile_(zs, 0.8) / C.Z80)) : 1;
  return { tau: tau, w: w, plans: per.length, months: zs.length, learned: learned };
}

/** OUTPUT の 29〜40 行（cells_json の並び）から、月ごとの P10/P50/P90（数でないセルは null） */
function appLandingMonths_(outRows) {
  const n = x => (x && x.charAt(0) === 'n' && isFinite(Number(x.slice(1))) ? Number(x.slice(1)) : null);
  const out = [];
  for (let r = 29; r <= 40; r++) {
    const row = (outRows || {})[r] || [];
    const ym = row[0] ? appCellYm_(row[0].charAt(0), row[0].slice(1)) : '';
    if (ym) out.push({ ym: ym, p10: n(row[1]), p50: n(row[2]), p90: n(row[3]) });
  }
  return out;
}

/** セルの日時をミリ秒に（日時の型・ISO の文字列はそのまま。'yyyy/MM/dd HH:mm' の文字列は日本の時刻として。読めなければ null） */
function appLandingTime_(t, text) {
  if (text === '' || text === null || text === undefined) return null;
  const s = String(text);
  if (t === 'd' || /T\d{2}:\d{2}/.test(s)) { const d = new Date(s.replace(/([+-]\d\d)(\d\d)$/, '$1:$2')); return isNaN(d.getTime()) ? null : d.getTime(); }
  const m = /^(\d{4})\D(\d{1,2})\D(\d{1,2})(?:\D+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/.exec(s);
  return m ? Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0) - 9, Number(m[5] || 0), Number(m[6] || 0)) : null;
}

/**
 * 締まった月の境目（'yyyy/MM'。この月より前が締まった月）。ENG_PROCESS_STATUS の B-1（実績の取り込み = step2）の後に
 * B-2（検証 = step5）が成功していれば、検証の表はちょうどその取り込みを映している（B-1 は実績を丸ごと入れ替え、取り込んだ月より前が締まった月）。
 * どちらかが成功していない・B-2 の方が古いときは ''（締まった月は数えない。空模様は霧になる）
 */
function appLandingCutoff_(stRows) {
  const ok = key => (stRows || []).filter(s => s.step_key === key && String(s.status).toLowerCase() === 'success')[0] || null;
  const b1 = ok('step2_status'), b2 = ok('step5_status');
  if (!b1 || !b2) return '';
  const typeOf = s => String(s._types || '').charAt(1);   // last_run_date は見出しの 2 列目
  const t1 = appLandingTime_(typeOf(b1), b1.last_run_date), t2 = appLandingTime_(typeOf(b2), b2.last_run_date);
  if (t1 === null || t2 === null || t2 < t1) return '';
  return appCellYm_(typeOf(b1), b1.last_run_date);
}

/** 予測を実行した日から今日まで、日本の暦で何日か（today = 'yyyy-MM-dd'。読めなければ null） */
function appLandingAgeDays_(finishedAt, today) {
  const ms = appLandingTime_('', finishedAt);
  if (ms === null || !today) return null;
  const day = Utilities.formatDate(new Date(ms), APP_TZ, 'yyyy-MM-dd');
  return Math.round((Date.parse(today + 'T00:00:00Z') - Date.parse(day + 'T00:00:00Z')) / 864e5);
}
