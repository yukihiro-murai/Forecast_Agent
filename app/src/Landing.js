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
 * τ と w は、業務の設定（landing.tau・landing.w。所有者が承認した値を saveSetting で書く）を使い、書いていなければ 0.15 と 1（Settings.js の appLandingApproved_。設定を読むのは計画の一覧の側）。
 * 全計画の検証の表（EVAL_COMPARE_MONTHLY）から学んだ値（appLandingPrior_。雪の計画は使わない）は、承認されるまで使わず、学びの画面に「試し」で出すだけ
 * （2026-10-08 村井さん承認の判断 29・11。前は、計画が 3 つ以上そろうと学んだ値をそのまま使っていた）。
 * 年度の見込みの試し（appLandingAligned_。判断 10）と予算に届く見込み（appLandingReach_・appLandingReachTotal_。判断 24・25）も同じ分布（appLandingDist_）を使う。
 * どれも影で、保存している数字（予測の年度の P10/P50/P90・予算）は変えない。
 * 空模様は着地 ÷ 年間予算で分ける。急な変化（天変地異）は、ひと月の大きな外れ・2 か月続いた同じ向きの外れ・予測の前提の大きな変化で拾う。
 *
 * 締まった月（2026-10-07 村井さん承認 D5）: 月末から APP_ACTUAL_CLOSE_LAG_DAYS 日たってから実績を取り込んだ（B-1）月だけ。
 * 月末 + 5 日 <= 取り込んだ日（日本の暦）。例: 10/03 の取り込みでは 9 月は途中の月、10/06 の取り込みなら 9 月は締まった月。
 * 旧来の計算と同じ決まり（数は旧来の側の定数と同じにする。旧来の側に定数があれば app/tests が照合する）。精度・学び・振り返り（Learning.js・Portfolio.js・Insights.js）も、
 * この境目（appLandingCutoff_）より前の月だけを使う（D4。締まっていない月の検証の行は消さずに読み飛ばす）
 */
const APP_ACTUAL_CLOSE_LAG_DAYS = 5;   // 締まった月: 月末からこの日数がたってから取り込んだ月だけ（D5）
const APP_LANDING = {
  TAU0: 0.15,              // 今年の水準 θ の事前分布の標準偏差（所有者が承認した値を設定 landing.tau に書くまで。判断 11）
  TAU_MIN: 0.05, TAU_MAX: 0.35,
  HALF_LIFE: 4,            // 月。APP_BIAS_HALF_LIFE（Learning.js）・旧来の B-5 と同じ
  Z80: 1.2815515655446004, // 標準正規分布の 90% 点（P10〜P90 = ±1.2816σ）
  FLOOR_REL: 0.05,         // 月の σ の下限 = 月の平均の予測の 5%
  THETA_MAX: 5,
  STALE_MONTHS: 3,         // 締まった年度の月のうち、実績が取り込まれていない月がこれだけあれば霧（旧来の C-1 の最低 3 か月と同じ）
  HOT: 1.5, CLEAR: 1.1, FAIR: 0.9, CLOUD: 0.5,   // 着地 ÷ 予算の境目（猛暑・快晴・晴れのち曇り・曇り。それより下は雨）
  SHOCK_Z: 3, SHIFT_Z: 2, MATERIAL: 0.05,        // 天変地異: ひと月 |z| ≥ 3、2 か月続けて |z| ≥ 2（同じ向き）。着地 ÷ 予算が 0.05 以上動いたときだけ
  PREMISE: 0.3, PREMISE_DAYS: 31, PREMISE_MATERIAL: 0.1,   // 前提の変化: 直近 2 回の予測の年間 P50 が 30% 以上違う（31 日以内・予算の 10% 以上）
  PRIOR_MIN_PLANS: 3, PRIOR_MIN_MONTHS: 3, W_MIN_MONTHS: 12, W_MAX: 3,
  // 予算に届く見込み（試し）で、届く金額を出す確率（%）と標準正規分布の点（Φ(z) = 確率）。90% 以上は出さない（判断 24）
  REACH: [[50, 0], [60, 0.2533471031357997], [70, 0.5244005127080407], [80, 0.8416212335729143]]
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

/** 予算の値（数でなければ null。0 以下はそのまま返す。使う側が 0 より大きいかを見る） */
function appLandingBudget_(v) {
  const n = v === null || v === undefined || v === '' ? NaN : Number(v);
  return isFinite(n) ? n : null;
}

/** θ を 0〜THETA_MAX に収める */
function appLandingTheta_(x) { return Math.min(APP_LANDING.THETA_MAX, Math.max(0, x)); }

/** 予算 B 以上で着地する確率（着地 L・ばらつき sd。ばらつきが 0 なら届くか届かないか） */
function appLandingPAbove_(L, sd, B) { return sd > 0 ? appNormCdf_((L - B) / sd) : (L >= B ? 1 : 0); }

/**
 * 年度の 12 か月の予測（OUTPUT の 29〜40 行。appLandingMonths_）を年度の月の順に並べる。
 * 12 か月とも P10/P50/P90 がそろい、P50 の合計が 0 より大きいときだけ返す（そうでなければ null = 予測が使えない）
 */
function appLandingFc_(fy, months) {
  const fin = v => typeof v === 'number' && isFinite(v);
  const byYm = {};
  (months || []).forEach(m => { if (m && m.ym) byYm[m.ym] = m; });
  const fc = appLandingFyYms_(fy).map(ym => byYm[ym] || null);
  const F = fc.reduce((s, m) => s + (m && fin(m.p50) ? Math.max(0, m.p50) : 0), 0);
  return F > 0 && fc.every(m => m && fin(m.p10) && fin(m.p50) && fin(m.p90)) ? fc : null;
}

/**
 * 着地の分布（着地見込み appLandingSky_・年度の見込みの試し appLandingAligned_・予算に届く見込み appLandingReach_ が同じ式を使う。シートは読まない）。
 * fc: そろった 12 か月の予測（appLandingFc_。null なら予測を使わない = 12 か月締まった年度だけ）、
 * acts: 締まった月の実績（年度の先頭から古い順。k = acts.length。行の無い月は 0 円）、tau・w: 今年の水準のぶれと幅の倍率。
 * 返り値: { landing, sd, p10, p90, theta, P, credibility, actualYtd, obs, FR, tau2 }（obs・FR・tau2 は天変地異の判定に使う）
 */
function appLandingDist_(fc, acts, tau, w) {
  const C = APP_LANDING;
  const F = fc ? fc.reduce((s, m) => s + Math.max(0, m.p50), 0) : 0;
  const floor = C.FLOOR_REL * F / 12;
  const sig2 = m => Math.pow(w * Math.max((m.p90 - m.p10) / (2 * C.Z80), floor), 2);
  const k = acts.length;
  const A = acts.reduce((s, a) => s + a, 0);
  const obs = fc ? acts.map((a, i) => ({ f: Math.max(0, fc[i].p50), a: a, s2: sig2(fc[i]) })) : [];
  const rest = fc ? fc.slice(k) : [];
  const FR = rest.reduce((s, m) => s + Math.max(0, m.p50), 0);
  const tau2 = tau * tau;
  const post = appLandingPost_(obs, tau2);
  const theta = appLandingTheta_(post.theta);
  const L = A + theta * FR;
  const sd = Math.sqrt(FR * FR / post.P + rest.reduce((s, m) => s + sig2(m), 0));
  return { landing: L, sd: sd, p10: Math.max(A, L - C.Z80 * sd), p90: L + C.Z80 * sd, theta: theta, P: post.P, credibility: 1 - (1 / tau2) / post.P,
    actualYtd: A, obs: obs, FR: FR, tau2: tau2 };
}

/**
 * 着地見込みと空模様（計画 1 つ。シートは読まない）。
 * inp: { fy, months: [{ ym, p10, p50, p90 }]（OUTPUT の 29〜40 行）, actual: { 'yyyy/MM': 実績 }（検証の表の actual_total）,
 *        cutoffYm（締まった月の境目。この月より前が締まった月。分からなければ ''）,
 *        todayYm（今日取り込んだとしたときの境目 = appCloseCutoffYm_(今日)。この月より前は、今日までに締まっているはずの月）, budget,
 *        pendingCutoffYm（B-1 の後に B-2 がまだのとき、その B-1 の取り込みの境目 = appLandingPendingCutoff_。B-2 が済んでいれば ''）,
 *        tau・w（承認した値 appLandingApproved_。省けば 0.15 と 1）, runs: [{ p50, ageDays }]（直近の予測 2 回。新しい順） }
 * 返り値: { k（締まった月の数）, actualYtd, budget, landing, landingSd, landingP10, landingP90, pAbove（予算以上で着地する確率）, ratio（着地 ÷ 予算）,
 *          theta, credibility（実績の重み 0〜1）, sky, skyReason, skyDir, zLast, dRatio1, dRatio2 }
 * 空模様（先に当てはまったもの）: 霧 mikakunin（no_budget）→ 雪 sekka（zero_sales）→ 霧（no_forecast）→ 霧（eval_pending / stale_actuals）
 *   → 天変地異 tenpen（shock / shift / premise。skyDir = up / down）→ 着地 ÷ 予算（ratio）: 猛暑 mousho ≥ 1.5・快晴 kaisei ≥ 1.1・
 *   晴れのち曇り harenochi > 0.9・曇り kumori > 0.5・雨 ame ≤ 0.5。
 * 着地の数字は、予測が使えれば雪・霧（予算が無い）のときも出す（予算が無ければ ratio と pAbove は出さない）。
 * 実績の取り込みが遅れている（stale_actuals の条件に当たる）ときは、空模様が霧（予算が無い）・雪でも出さない（古い実績のままの数字を見せない。合計にも入らない）。
 * 実績は取り込んだ（B-1）が当たり具合の計算（B-2）がまだで、締まった月が数えられないだけのとき（B-1 の取り込みの境目で数えれば遅れていない）は、
 * 霧の理由を eval_pending（当たり具合の計算待ち）にする。数字を出さないのは stale_actuals と同じ（2026-10-08。前は「実績の遅れ」と出ていた）。
 * B-1 の取り込みの境目で数えても 3 か月以上遅れていれば、stale_actuals のまま
 * 前提の変化（premise）は、締まっていない月があるときだけ（12 か月締まった年度の着地は実績の合計で動かない）
 */
function appLandingSky_(inp) {
  const C = APP_LANDING;
  const fin = v => typeof v === 'number' && isFinite(v);
  const yms = appLandingFyYms_(inp.fy);
  const fc = appLandingFc_(inp.fy, inp.months);   // そろった 12 か月の予測（そろわなければ null）
  const fcOk = !!fc;
  const cutoff = String(inp.cutoffYm || '');
  const obsYm = cutoff ? yms.filter(ym => ym < cutoff) : [];   // 年度の月は古い順なので、先頭から k か月
  const k = obsYm.length;
  const act = inp.actual || {};
  const aOf = ym => { const v = Number(act[ym]); return act[ym] === null || act[ym] === '' || !isFinite(v) ? 0 : v; };   // 実績の行が無い締まった月は 0 円
  const A = obsYm.reduce((s, ym) => s + aOf(ym), 0);
  const closedByToday = inp.todayYm ? yms.filter(ym => ym < inp.todayYm).length : k;
  const stale = closedByToday - k >= C.STALE_MONTHS;   // 今日までに締まっているはずの月（月末から 5 日たった月）のうち、3 か月以上の実績が取り込まれていない
  // B-2 待ち: B-1 の取り込みの境目で数えれば遅れていない（遅れて見えるのは、B-2 がまだで締まった月を数えられないため）
  const pend = String(inp.pendingCutoffYm || '');
  const evalPending = stale && !!pend && closedByToday - Math.max(k, yms.filter(ym => ym < pend).length) < C.STALE_MONTHS;
  const B = appLandingBudget_(inp.budget);
  const out = { k: k, actualYtd: A, budget: B, landing: null, landingSd: null, landingP10: null, landingP90: null, pAbove: null, ratio: null,
    theta: null, credibility: null, sky: 'mikakunin', skyReason: '', skyDir: '', zLast: null, dRatio1: null, dRatio2: null };
  const usable = fcOk || k >= 12;
  const tau = inp.tau > 0 ? inp.tau : C.TAU0;
  const w = inp.w > 0 ? inp.w : 1;
  let d = null;
  if (usable && !stale) {   // 実績の取り込みが遅れていれば、どの空模様でも着地の数字は出さない
    d = appLandingDist_(fc, obsYm.map(aOf), tau, w);
    Object.assign(out, { landing: d.landing, landingSd: d.sd, landingP10: d.p10, landingP90: d.p90, theta: d.theta, credibility: d.credibility });
    if (B !== null && B > 0) Object.assign(out, { ratio: d.landing / B, pAbove: appLandingPAbove_(d.landing, d.sd, B) });
  }
  // 1〜4: 霧・雪（数字で分けられない）
  if (!(B !== null && B > 0)) { out.skyReason = 'no_budget'; return out; }
  if (k >= 1 && Math.abs(A) < 1) { out.sky = 'sekka'; out.skyReason = 'zero_sales'; return out; }
  if (!usable) { out.skyReason = 'no_forecast'; return out; }
  if (stale) { out.skyReason = evalPending ? 'eval_pending' : 'stale_actuals'; return out; }
  // 5: 天変地異（月の外れは、その月の前までの実績で立てた見込みと比べる）
  let t1 = false, t2 = false, dir = 0;
  if (fcOk && k >= 1 && k < 12) {
    const obs = d.obs, tau2 = d.tau2, FR = d.FR;
    const before = n => {   // 先頭の n か月の実績だけで立てた見込み
      const p = appLandingPost_(obs.slice(0, n), tau2);
      const th = appLandingTheta_(p.theta);
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
  const t3 = k < 12 && r.length >= 2 && fin(r[0].p50) && fin(r[1].p50) && r[1].p50 > 0 && fin(r[0].ageDays) && r[0].ageDays <= C.PREMISE_DAYS
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
 * 締まった月の実績の合計が 0 円の計画（雪）は、τ にも w にも使わない（水準 −100% として τ を上限に押し上げ、ほかの計画の着地と空模様を動かすため）。
 * 0 円の月がいくつかあるだけの計画は、その月を 0 円のまま使う
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
    if (Math.abs(Sa) < 1) return;   // 雪（空模様の雪と同じ: 締まった月の売上の合計が 1 円未満）
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

/** 確率を 5% 刻みの整数（0〜100）に丸める（予算に届く見込みの見せ方。0 は「5% 未満」・100 は「95% 超」と読む） */
function appLandingPct5_(p) { return typeof p === 'number' && isFinite(p) ? Math.round(p * 20) * 5 : null; }

/**
 * 年度の見込みの試し（判断 10。影: 保存している年度の P10/P50/P90 は変えない）。
 * 今の年度の P10/P50/P90 は旧来の計算の試行（学んだ補正を入れず、月を別々に足す）なので、月の P50 の合計と違い、幅も狭い。
 * そこで、中心 = 最新の予測の月の P50（OUTPUT の 29〜40 行。補正を入れた値）の 12 か月の合計、幅 = 着地見込みと同じ式で締まった月が無いとき（appLandingDist_）。
 * months: appLandingMonths_ の形、tau・w: appLandingApproved_、budget: 予算以上になる確率を出す予算（無い・0 以下なら pAbove は null）。
 * opts: { k（締まった月の数。着地見込み appLandingSky_ の k）, frozen（年度を締めた計画） }。省けば締まった月なし。
 * 締まった月がある（k > 0）ときは、幅を出さない（sd・p10・p90・pAbove は null）: 予測の締まった月は実績が入っているので、年度全体に τ の幅を
 * つけると、ぶれを大きく見せ、着地見込みの幅とも食い違う（年度の途中の幅は着地見込みの方。2026-10-08）。中心はそのまま出す。
 * 12 か月締まった・年度を締めた計画は done（年度は終わった）。done も幅を出さない。
 * 予測がそろわなければ null。返り値: { center, sd, p10, p90, budget, pAbove, tau, w, k, done }
 */
function appLandingAligned_(fy, months, tau, w, budget, opts) {
  const fc = appLandingFc_(fy, months);
  if (!fc) return null;
  const o = opts || {};
  const k = Math.max(0, Math.min(12, Math.floor(Number(o.k) || 0)));
  const done = k >= 12 || !!o.frozen;
  const d = appLandingDist_(fc, [], tau, w);
  const B = appLandingBudget_(budget);
  const band = k === 0 && !done;
  return { center: d.landing, sd: band ? d.sd : null, p10: band ? d.p10 : null, p90: band ? d.p90 : null, budget: B,
    pAbove: band && B !== null && B > 0 ? appLandingPAbove_(d.landing, d.sd, B) : null, tau: tau, w: w, k: k, done: done };
}

/**
 * 予算に届く見込み（判断 24・25 の見せ方。影: 予算も予測も変えない。確率を選んで予算を書く操作は、まだ無い）。
 * sky: その計画の着地見込み（appLandingSky_ の返り値。年度の途中は締まった月の実績を入れた分布、先の年度は締まった月なし）。
 * budgets: { draft（今の予算 = 採用予測 + 上乗せ）, official（承認済みの公式版の最終予算）, officialNo }、prior: appLandingApproved_、
 * opts: { frozen（年度を締めた計画） }。
 * 着地見込みの数字が無い（実績の遅れ・予測が無い）ときは null。
 * 12 か月締まった・年度を締めた計画は done（年度は終わった。2026-10-08）: actual = 締まった月の実績の合計、届く金額（amounts）は空。
 * 画面は確率ではなく、届いた／届かなかったを出す（予算ごとの reached = 実績の合計 ≥ 予算。p・pct は今までどおりの値）。
 * 返り値: { center, sd, p10, p90, k, actualYtd, tau, w, tauSet, wSet, done, actual（done のときだけ数。ほかは null）,
 *          amounts: [{ pct, amount }]（その確率で届く金額。50〜80%。done なら空）,
 *          draft: { budget, p, pct, reached? } | null, official: { budget, p, pct, versionNo, reached? } | null }（reached は done のときだけ）
 */
function appLandingReach_(sky, budgets, prior, opts) {
  const fin = v => typeof v === 'number' && isFinite(v);
  if (!sky || !fin(sky.landing) || !fin(sky.landingSd)) return null;
  const L = sky.landing, sd = sky.landingSd, A = fin(sky.actualYtd) ? sky.actualYtd : 0;
  const done = (fin(sky.k) && sky.k >= 12) || !!(opts && opts.frozen);
  const one = b => {
    const B = appLandingBudget_(b);
    if (B === null || !(B > 0)) return null;
    const p = appLandingPAbove_(L, sd, B);
    const out = { budget: B, p: p, pct: appLandingPct5_(p) };
    if (done) out.reached = A >= B;
    return out;
  };
  const bs = budgets || {}, pr = prior || {};
  const official = one(bs.official);
  if (official) official.versionNo = bs.officialNo || null;
  return { center: L, sd: sd, p10: sky.landingP10, p90: sky.landingP90, k: sky.k, actualYtd: A,
    tau: fin(pr.tau) ? pr.tau : APP_LANDING.TAU0, w: fin(pr.w) ? pr.w : 1, tauSet: !!pr.tauSet, wSet: !!pr.wSet,
    done: done, actual: done ? A : null,
    // 届く金額は、締まった月の実績の合計より下にはしない（80% の幅の下と同じ）。年度が終わっていれば出さない
    amounts: done ? [] : APP_LANDING.REACH.map(x => ({ pct: x[0], amount: Math.max(A, L - x[1] * sd) })),
    draft: one(bs.draft), official: official };
}

/**
 * 全部のメーカーの合計の予算に届く見込み（判断 24 の見せ方。影）。rows: [{ landing, sd, budget }]（着地見込みと予算の両方がある計画）。
 * メーカーどうしが独立に動くとき（ばらつき = √Σsd²）と、全部が同じ向きに動くとき（ばらつき = Σsd）の 2 つ。本当の値はその間のどこか。
 * 返り値: { plans, budget, landing, independent: { p, pct }, together: { p, pct } }。計画が無ければ null
 */
function appLandingReachTotal_(rows) {
  const fin = v => typeof v === 'number' && isFinite(v);
  const use = (rows || []).filter(r => r && fin(r.landing) && fin(r.budget) && r.budget > 0);
  if (!use.length) return null;
  const L = use.reduce((s, r) => s + r.landing, 0), B = use.reduce((s, r) => s + r.budget, 0);
  const sdOf = r => (fin(r.sd) && r.sd > 0 ? r.sd : 0);
  const ind = Math.sqrt(use.reduce((s, r) => s + sdOf(r) * sdOf(r), 0)), tog = use.reduce((s, r) => s + sdOf(r), 0);
  const pi = appLandingPAbove_(L, ind, B), pt = appLandingPAbove_(L, tog, B);
  return { plans: use.length, budget: B, landing: L, independent: { p: pi, pct: appLandingPct5_(pi) }, together: { p: pt, pct: appLandingPct5_(pt) } };
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
 * 月（'yyyy/MM'）が、importDay（実績を取り込んだ日。日本の暦の 'yyyy-MM-dd'）の取り込みで締まっているか（D5）:
 * 月末 + APP_ACTUAL_CLOSE_LAG_DAYS 日 <= 取り込んだ日。9 月なら 10/05 以降の取り込みで締まる（10/04 までは途中の月）。読めなければ false
 */
function appActualClosed_(ym, importDay) {
  const m = /^(\d{4})\/(\d{2})$/.exec(String(ym || ''));
  const d = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(importDay || ''));
  if (!m || !d) return false;
  return Date.UTC(Number(m[1]), Number(m[2]), 0) + APP_ACTUAL_CLOSE_LAG_DAYS * 864e5 <= Date.UTC(Number(d[1]), Number(d[2]) - 1, Number(d[3]));
}

/**
 * 取り込んだ日（日本の暦の 'yyyy-MM-dd'）の締まった月の境目（'yyyy/MM'。この月より前が締まった月 = appActualClosed_ が真の月）。
 * 取り込んだ日の月は締まっていない（月末がまだ来ていない）ので、そこから前の月へたどる。読めなければ ''
 */
function appCloseCutoffYm_(importDay) {
  const d = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(importDay || ''));
  if (!d) return '';
  let y = Number(d[1]), mo = Number(d[2]);
  for (;;) {
    const py = mo === 1 ? y - 1 : y, pm = mo === 1 ? 12 : mo - 1;
    if (appActualClosed_(py + '/' + ('0' + pm).slice(-2), importDay)) return y + '/' + ('0' + mo).slice(-2);
    y = py; mo = pm;
  }
}

/** 'yyyy/MM' の月の初め（日本の暦の 1 日 0 時）のミリ秒。読めなければ null */
function appMonthStartMs_(ym) {
  const m = /^(\d{4})\/(\d{2})$/.exec(String(ym || ''));
  return m ? Date.UTC(Number(m[1]), Number(m[2]) - 1, 1) - 9 * 3600e3 : null;
}

/**
 * 締まった月の境目（'yyyy/MM'。この月より前が締まった月）。ENG_PROCESS_STATUS の B-1（実績の取り込み = step2）の後に
 * B-2（検証 = step5）が成功していれば、検証の表はちょうどその取り込みを映している（B-1 は実績を丸ごと入れ替える）。
 * 締まった月は、B-1 を動かした日（日本の暦）から appCloseCutoffYm_ で決める（月末から 5 日たってから取り込んだ月だけ。D5）。
 * どちらかが成功していない・B-2 の方が古いときは ''（締まった月は数えない。空模様は霧になる）。
 * stRows は、データ本体の行（値は文字のまま・型の並び _types つき）でも、読み戻した行（日時は日時の型）でもよい
 */
function appLandingCutoff_(stRows) {
  const st = appLandingSteps_(stRows);
  if (st.t1 === null || st.t2 === null || st.t2 < st.t1) return '';
  return appCloseCutoffYm_(Utilities.formatDate(new Date(st.t1), APP_TZ, 'yyyy-MM-dd'));   // 取り込んだ時刻の、日本の暦の日（文字の ISO 時刻も時差を見る）
}

/**
 * B-1 の後に B-2 がまだのとき（B-2 の成功が無い・B-1 の成功より古い）、その B-1 の取り込みの境目（'yyyy/MM'。appCloseCutoffYm_）。
 * B-2 が済めば締まった月の境目（appLandingCutoff_）になる値。B-2 が済んでいる・B-1 の成功が無い・時刻が読めなければ ''。
 * 空模様の霧の理由を、実績の遅れ（stale_actuals）と当たり具合の計算待ち（eval_pending）に分けるために使う（appLandingSky_ の pendingCutoffYm）
 */
function appLandingPendingCutoff_(stRows) {
  const st = appLandingSteps_(stRows);
  if (st.t1 === null || (st.t2 !== null && st.t2 >= st.t1)) return '';
  return appCloseCutoffYm_(Utilities.formatDate(new Date(st.t1), APP_TZ, 'yyyy-MM-dd'));
}

/** PROCESS_STATUS の行から、B-1（step2）と B-2（step5）の成功の時刻（ミリ秒。成功が無い・読めなければ null） */
function appLandingSteps_(stRows) {
  const ok = key => (stRows || []).filter(s => s.step_key === key && String(s.status).toLowerCase() === 'success')[0] || null;
  const timeOf = s => {
    if (!s) return null;
    if (appIsDate_(s.last_run_date)) { const t = s.last_run_date.getTime(); return isFinite(t) ? t : null; }
    return appLandingTime_(String(s._types || '').charAt(1), s.last_run_date);   // last_run_date は見出しの 2 列目
  };
  return { t1: timeOf(ok('step2_status')), t2: timeOf(ok('step5_status')) };
}

/** 予測を実行した日から今日まで、日本の暦で何日か（today = 'yyyy-MM-dd'。読めなければ null） */
function appLandingAgeDays_(finishedAt, today) {
  const ms = appLandingTime_('', finishedAt);
  if (ms === null || !today) return null;
  const day = Utilities.formatDate(new Date(ms), APP_TZ, 'yyyy-MM-dd');
  return Math.round((Date.parse(today + 'T00:00:00Z') - Date.parse(day + 'T00:00:00Z')) / 864e5);
}
