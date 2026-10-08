/**
 * Portfolio.js — 計画（クライアント × 年度）の一覧と、新しい計画を作る。
 *
 * 新しい計画は、計算用ブックの上で旧来の A-1 初期セットアップ（setupForecastBook の中身）と
 * 設定の保存（saveInitialSetupSettings）を動かし、できたシートをデータ本体に計画として保存する。
 * その後は、A-2 売上データの取り込み → A-3 → 予測、と進める。
 */

/** A-1 初期セットアップで作るシート（旧来の setupForecastBook の order と同じ順） */
const APP_SETUP_ORDER = ['GUIDE', 'CONFIG', 'SALES_INPUT', 'SALES_MONTHLY', 'AI_RESEARCH', 'PRODUCT', 'CLIENT', 'OPINIONS', 'DEV_SPOT', 'OUTPUT',
  'DASHBOARD', 'ACTUAL_EVAL_MONTHLY', 'EVAL_COMPARE_MONTHLY', 'EVAL_LOG', 'EVAL_INSIGHTS', 'QUARTERLY_REVIEW', 'QUARTERLY_REVIEW_LOG',
  'AI_RESEARCH_STRUCTURED', 'RUN_LOG', 'FORECAST_SNAPSHOT', 'PROCESS_STATUS', 'AI_SCORE_HISTORY', 'AI_IMPACT_HISTORY',
  'SUBJECTIVE_IMPACT_HISTORY', 'CALIBRATION_STATE', 'CALIBRATION_HISTORY', 'SOURCE_RELIABILITY', 'RELIABILITY_EVIDENCE'];
const APP_CANDIDATES_CACHE_KEY = 'APP_ZAC_CLIENTS';
/** 新しい計画の地域（旧ブックの既定と同じ） */
function appNewPlanTz_() { return { time_zone: APP_TZ, locale: 'ja_JP' }; }

// ---- 新しい計画 ----

/** ZAC の実績に出てくるクライアント（旧来の A-1 の候補と同じ。直近 2 年）。6 時間控える */
function appPlanCandidates_(ctx, input) {
  const cache = CacheService.getScriptCache();
  let names = null;
  if (!(input && input.refresh)) { try { names = JSON.parse(cache.get(APP_CANDIDATES_CACHE_KEY) || 'null'); } catch (e) { names = null; } }
  if (!names) {
    names = appLegacyCall_({}, { asOfMs: new Date().getTime(), seed: 'candidates', actor: ctx.actor }, 'getClientCandidatesForSetup_', []).value || [];
    try { cache.put(APP_CANDIDATES_CACHE_KEY, JSON.stringify(names), 21600); } catch (e) { /* 大きすぎれば控えない */ }
  }
  // 画面には、ふつうの表記を出す（作るときは ZAC の名前を使う。売上の取り込みで照合するため）
  return { candidates: names.map(n => ({ zac: n, display: appClientDisplayName_(n) })), defaultFy: appFy_(new Date()) + (new Date().getMonth() >= 9 ? 1 : 0), existing: appListPlans_().map(p => ({ clientName: p.clientName, fy: p.fy })) };
}

function appPlanCreateCheck_(p) {
  const clientName = String(p && p.clientName || '').trim();
  const fy = Number(p && p.fy);
  const people = String(p && p.peopleCsv || '').split(/[,、，]/).map(s => s.trim()).filter(Boolean);
  if (!clientName) throw new Error('メーカーを選んでください。');
  if (clientName.length > 100) throw new Error('メーカーの名前が長すぎます。');
  if (!fy || fy < 2000 || fy > 2100 || Math.floor(fy) !== fy) throw new Error('年度（FY）を 4 桁の数で入れてください。');
  appRequireOpenYear_(fy);   // 締めた年度には、新しい計画を作れない（組み立て・保存の両方で確かめる）
  if (!people.length) throw new Error('担当者を 1 人以上入れてください。');
  if (people.some(x => x.length > 40)) throw new Error('担当者の名前が長すぎます。');
  appPlanCheckArgs_([clientName].concat(people));
  const normalized = appNormalizeName_(clientName);
  const client = appReadTable_('CLIENTS').filter(c => c.normalized_name === normalized)[0] || null;
  if (client && appReadTable_('PLANS').some(x => x.client_id === client.client_id && String(x.fy) === String(fy) && x.state !== 'ARCHIVED')) {
    throw new Error(appClientDisplayName_(client.client_name) + ' の FY' + fy + ' の計画は、すでにあります。');
  }
  return { clientName: clientName, fy: fy, peopleCsv: people.join(','), client: client, normalized: normalized };
}

/** 作る（PLAN.CREATE）: 計算用ブックの上で旧来の A-1 初期セットアップと設定の保存を動かす（データ本体には書かない） */
function appPlanCreateBuild_(ctx, p) {
  const t0 = new Date().getTime();
  return appWithLock_(() => {
    appJournalRecover_(ctx);
    const c = appPlanCreateCheck_(p);
    const scratch = appWorkScratch_(appNewPlanTz_());
    const token = appId_('SCR');
    appScratchReset_(scratch, token);
    const asOfMs = new Date().getTime();
    const engine = appLegacySetupBook_(scratch, { asOfMs: asOfMs, seed: token, actor: ctx.actor }, APP_SETUP_ORDER, c.clientName, c.fy, c.peopleCsv);
    // A-1 が作る検証の表は 26 列。旧来の B-2 が書く 36 列まで広げてから保存する（Engine.js の APP_ENGINE_MIN_COLUMNS）
    appEnsureMinColumns_(scratch);
    const payload = { clientName: c.clientName, fy: c.fy, peopleCsv: c.peopleCsv, token: token, engine: engine, buildMs: new Date().getTime() - t0 };
    return { __next: { kind: 'PLAN.CREATE_SAVE', payload: payload }, audit: { entityId: c.clientName + ' FY' + c.fy } };
  });
}

/** 保存（PLAN.CREATE_SAVE）: できたシートを、新しい計画としてデータ本体に書く */
function appPlanCreateSave_(ctx, p) {
  const t0 = new Date().getTime();
  return appWithLock_(() => {
    appJournalRecover_(ctx);
    if (!appScratchOwnedBy_(p.token)) throw new Error('計算用ブックがほかの処理で使われました。もう一度作ってください。');
    const c = appPlanCreateCheck_(p);   // 作っている間に、同じ計画がほかで作られていないか
    const now = appNowIso_();
    const planId = appId_('PL');
    const ops = [];
    let client = c.client;
    if (!client) {
      client = { client_id: appId_('CL'), client_name: c.clientName, zac_code: '', normalized_name: c.normalized, aliases_json: '[]',
        is_active: true, note: '', created_at: now, created_by: ctx.actor, updated_at: now, updated_by: ctx.actor, row_version: 1 };
      ops.push({ table: 'CLIENTS', mode: 'ensure', rows: [client] });
    }
    const cap = appCaptureChanged_(appScratchBook_(), planId, {});
    const missing = Object.keys(APP_ENGINE_SHEETS).filter(n => APP_SETUP_ORDER.indexOf(n) >= 0 && !cap.changed.some(e => e.sheetRow.sheet === n));
    if (missing.length) throw new Error('初期セットアップでできなかったシートがあります: ' + missing.join('、'));
    const batchId = appId_('NEW');
    appChangedOps_(ctx, planId, cap.changed, batchId).forEach(op => ops.push(op));
    ops.push({ table: 'PLANS', mode: 'ensure', rows: [{ plan_id: planId, client_id: client.client_id, fy: String(c.fy), client_label: c.clientName,
      people_csv: c.peopleCsv, source_book_id: '', locale: appNewPlanTz_().locale, time_zone: appNewPlanTz_().time_zone, state: 'ACTIVE', note: '',
      created_at: now, created_by: ctx.actor, updated_at: now, updated_by: ctx.actor, row_version: 1 }] });
    const written = appJournalRun_(ctx, '計画の作成（' + c.clientName + ' FY' + c.fy + '）', planId, ops);
    appScratchMarkSynced_(planId, p.token, null);   // 作った後の A-2 などは組み立て直さずに使える
    return { planId: planId, clientName: appClientDisplayName_(c.clientName), fy: c.fy, sheets: cap.changed.length, written: written, engine: p.engine,
      timing: { buildMs: p.buildMs, saveMs: new Date().getTime() - t0 }, audit: { entityId: planId, clientId: client.client_id } };
  });
}

// ---- 一覧 ----

/** 計画の一覧（予測の画面の計画を選ぶ欄） */
function appListPlans_() {
  const clients = appClientNameMap_();   // 画面に出す名前（半角カナ・株式会社などを除いた、ふつうの表記）
  return appReadTable_('PLANS').map(p => {
    return { planId: p.plan_id, clientName: clients[p.client_id] || p.client_label, fy: p.fy, state: p.state, frozen: appYearIsFrozen_(p.fy) };
  });
}


/**
 * 全部の計画の要点（最新の予測・前回からの変化・予算・精度・手順の進み・着地見込みと空模様）。
 * 着地見込みと空模様は Landing.js（締まった月の実績から今年の水準を学ぶ）。今日の日付で変わる（霧の判定・前提の変化の日数）ので、控えは日ごと。
 * 検証の表（EVAL_COMPARE_MONTHLY）の予測（P10/P50/P90）・外れ・幅の外の印は、今の決まり（D4〜D6）で B-2 が書き直した計画だけで使う
 * （appPortfolioScored_）。表には版の印が無く、旧来の B-2 が丸ごと書き直すまでは、前の決まり（月が始まった後の予測）の値のままのため。
 * 書き直していない計画（締めた年度の計画はずっと）は、外れ幅・予実の差・幅の外を出さず、着地の τ・w の学びにも入れない。実績はそのまま使う
 * （暫定実績・着地見込み。実績は B-1 の取り込みのままで、どの回の予測で測るかに関係しない）
 * 着地見込みの τ・w は、所有者が承認して業務の設定に書いた値（無ければ 0.15 と 1。appLandingApproved_）。全計画から学んだ値（appLandingPrior_）は
 * prior.learned として返すだけで、着地には使わない（2026-10-08 判断 29）。
 * 計画ごとに足す項目（2026-10-08。どれも影で、保存している数字は変えない）:
 *   aligned      … 年度の見込みの試し（月の P50 の合計を中心に、着地見込みと同じ式の幅。appLandingAligned_。判断 10）。
 *                  締まった月がある（aligned.k > 0）・締まった月を数え直している間（aligned.pending）は幅を出さない。12 か月を数えた計画は aligned.done
 *   reach        … 予算に届く見込み（今の予算と承認済みの公式版の予算。届く金額 50〜80%。appLandingReach_。判断 24・25）。
 *                  12 か月を数えた計画は reach.done と reach.actual（実績の合計。届いた／届かなかった。届く金額は出さない）。年度を締めても 12 か月に足りなければ見込みのまま
 *   scoredMonths … 今の決まりで B-2 が測った締まった月の数（appPortfolioScored_。実績 0 円の月も数える。数）
 * 返り値: { plans: [計画の要点], prior: { learned: appLandingPrior_ の返り値, used: appLandingApproved_ の返り値 } }
 */
function appPortfolioData_() {
  const clients = appClientNameMap_();   // 画面に出す名前（半角カナ・株式会社などを除いた、ふつうの表記）
  const today = appToday_();
  const todayYm = appCloseCutoffYm_(today);   // 今日までに締まっているはずの月の境目（月末から 5 日たった月まで。霧の判定）
  const runs = {};
  appReadTable_('FORECAST_RUNS').filter(r => r.status === 'DONE').forEach(r => { (runs[r.plan_id] = runs[r.plan_id] || []).push(r); });
  const outRows = {};
  appReadTable_('ENG_ROWS').filter(r => r.sheet === 'OUTPUT' && Number(r.col_from) === 1).forEach(s => {
    const n = Number(s.row_no);
    if (n === 1 || n === 26 || (n >= 29 && n <= 40)) { (outRows[s.plan_id] = outRows[s.plan_id] || {})[n] = JSON.parse(s.cells_json); }
  });
  const steps = {};
  appReadTable_('ENG_PROCESS_STATUS').forEach(r => { (steps[r.plan_id] = steps[r.plan_id] || []).push(r); });
  const plans = appReadTable_('PLANS').filter(p => p.state !== 'ARCHIVED');
  // 締まった月の境目（B-1 → B-2 の順に動いた取り込み）と、今の決まりで測った月（EVAL_LOG を全計画の分 1 回で読む）
  const cut = {};
  plans.forEach(p => { cut[p.plan_id] = appLandingCutoff_(steps[p.plan_id] || []); });
  const scored = appPortfolioScored_(plans.map(p => p.plan_id), cut);
  const ape = {};   // 計画ごとの [{ ym, v }]（外れ幅は、締まった月で今の決まりで測った月の分だけを後で平均する）
  const cmp = {};   // 月ごとの実績と P10/P50/P90（検証の表）。暫定実績・着地見込み・予実の差・全計画の学びに使う（読むだけ。計算は変えない）
  const fnum = x => { const n = x === '' || x === null || x === undefined ? null : Number(x); return n !== null && isFinite(n) ? n : null; };
  appReadTable_('ENG_EVAL_COMPARE_MONTHLY').forEach(r => {
    const now = !!(scored[r.plan_id] && scored[r.plan_id].any);   // 今の決まりで B-2 が書き直した計画か
    const v = Number(r.ape_p50);
    const ym = appCellYm_(String(r._types || '').charAt(0), r.target_month);
    // 実績が空でも P50 のある行は残す（actual: null。ZAC に記録の無い月 = 売上 0。締まった月なら 0 円として数える）
    const blank = r.actual_total === '' || r.actual_total === null || r.actual_total === undefined;
    const act = blank ? null : Number(r.actual_total);
    const p50 = fnum(r.forecast_total_p50);
    // 外れは、今の決まりで測った月だけ（精度と同じ月。締まった後の予測 = 予測と実績がまったく同じ月も、精度と同じく除く）
    if (now && scored[r.plan_id].months[ym] && !blank && isFinite(act) && isFinite(v) && r.ape_p50 !== '' && !(p50 !== null && Math.abs(p50 - act) < 1e-6)) {
      (ape[r.plan_id] = ape[r.plan_id] || []).push({ ym: ym, v: v });
    }
    if (ym && (blank ? p50 !== null : isFinite(act))) {
      // 書き直していない計画は、予測と幅の外の印を使わない（実績だけ）
      (cmp[r.plan_id] = cmp[r.plan_id] || {})[ym] = now ? { actual: act, p10: fnum(r.forecast_total_p10), p50: p50, p90: fnum(r.forecast_total_p90),
        out: String(r.range_outside_flag) === '1' ? 1 : String(r.range_outside_flag) === '0' ? 0 : null } : { actual: act, p10: null, p50: null, p90: null, out: null };
    }
  });
  const ver = appVersionSummary_();
  const num = x => { if (!x) return null; const t = x.charAt(0); if (t !== 'n') return null; const v = Number(x.slice(1)); return isFinite(v) ? v : null; };
  // 全計画の締まった月から学ぶ τ・w（1 回だけ。書き直していない計画は予測が無いので入らない）。承認されるまで着地には使わない（学びの画面に試しで出す）
  const learned = appLandingPrior_(plans.map(p => {
    const c = cmp[p.plan_id] || {};
    return Object.keys(c).filter(ym => cut[p.plan_id] && ym < cut[p.plan_id]).sort()
      .map(ym => ({ f: c[ym].p50, a: c[ym].actual === null ? 0 : c[ym].actual, p10: c[ym].p10, p90: c[ym].p90 }));   // 実績の空の締まった月は 0 円（着地の計算と同じ。全部 0 円の計画 = 雪は appLandingPrior_ が除く）
  }));
  const used = appLandingApproved_();   // 着地に使う τ・w（所有者が承認して書いた値。無ければ 0.15 と 1）
  // 年度を締めたか。年度ごとに 1 回だけ確かめる（YEAR_CLOSURES を計画の数だけ読まない）。確かめられなければ締めていないとして見せる
  // （見せ方だけに使う。書き込みの入口の確かめ appRequireOpenYear_ は別で、そちらは止める）
  const frozenFy = {};
  const frozenOf = fy => {
    const k = String(fy);
    if (!Object.prototype.hasOwnProperty.call(frozenFy, k)) {
      try { frozenFy[k] = appYearIsFrozen_(fy); } catch (e) { Logger.log('年度の締めを確かめられません: ' + k + ' ' + (e && e.message ? e.message : e)); frozenFy[k] = false; }
    }
    return frozenFy[k];
  };
  const rows = plans.map(p => {
    const rs = (runs[p.plan_id] || []).sort((a, b) => String(b.finished_at).localeCompare(String(a.finished_at)) || b._row - a._row);
    const o = outRows[p.plan_id] || {};
    let adopted = null, uplift = null;
    for (let r = 29; r <= 40; r++) {
      const a = num((o[r] || [])[7]), u = num((o[r] || [])[8]);
      if (a !== null) adopted = (adopted || 0) + a;
      if (u !== null) uplift = (uplift || 0) + u;
    }
    const stored = o[26] ? { p10: num(o[26][1]), p50: num(o[26][2]), p90: num(o[26][3]) } : {};
    const latest = rs[0] || null;
    const prev = rs[1] || null;
    const v = ver[p.plan_id] || {};
    const budget = adopted === null && uplift === null ? null : (adopted || 0) + (uplift || 0);
    // 空模様の予算: 承認済みの公式版の最終予算、無ければ今の予算（採用予測 + 上乗せ。予測し直すと採用予測は P50 に戻る）
    const official = v.officialNo ? appNum_(v.officialFinal) : null;
    const budgetUsed = official !== null ? official : budget;
    const yt = appPlanYtd_(p.fy, cmp[p.plan_id] || {}, cut[p.plan_id]);
    const months = appLandingMonths_(o);
    const frozen = frozenOf(p.fy);   // 年度を締めた計画（画面の印。年度が終わったかは、数えた締まった月が 12 か月かで決める）
    const pendingCut = appLandingPendingCutoff_(steps[p.plan_id] || []);   // 実績を取り込んだ後、まだ当たり具合を計算していない（締まった月を数え直している間）
    const sky = appLandingSky_({ fy: p.fy, months: months, actual: yt.actual, cutoffYm: cut[p.plan_id], todayYm: todayYm, budget: budgetUsed,
      pendingCutoffYm: pendingCut, tau: used.tau, w: used.w, runs: rs.slice(0, 2).map(r => ({ p50: r.annual_p50, ageDays: appLandingAgeDays_(r.finished_at, today) })) });
    const st = steps[p.plan_id] || [];
    const errors = st.filter(s => String(s.status).toLowerCase() === 'error').map(s => s.step_key);
    // 外れ幅は締まった月で、今の決まりで測った月だけ（D4〜D6。精度 appAccuracyOf_ と同じ月。ほかの検証の行は消さずに読み飛ばす。境目が分からなければ数えない）
    const a = (ape[p.plan_id] || []).filter(x => cut[p.plan_id] && x.ym && x.ym < cut[p.plan_id]).map(x => x.v);
    return {
      planId: p.plan_id, clientName: clients[p.client_id] || p.client_label, fy: p.fy,
      p10: latest ? latest.annual_p10 : stored.p10 === undefined ? null : stored.p10,
      p50: latest ? latest.annual_p50 : stored.p50 === undefined ? null : stored.p50,
      p90: latest ? latest.annual_p90 : stored.p90 === undefined ? null : stored.p90,
      prevP50: prev ? prev.annual_p50 : null,
      lastRunAt: latest ? latest.finished_at : '', runs: rs.length,
      budget: budget,
      mape: a.length ? a.reduce((x, y) => x + y, 0) / a.length : null, mapeMonths: a.length,
      stepsDone: st.filter(s => String(s.status).toLowerCase() === 'success').length, stepsTotal: st.length, stepErrors: errors,
      createdAt: p.created_at,
      officialNo: v.officialNo || null, officialFinal: v.officialFinal === undefined ? null : v.officialFinal,
      pendingNo: v.pendingNo || null,
      actualYtd: yt.actualYtd, actualMonths: yt.actualMonths, forecastYtd: yt.forecastYtd, rangeOut: yt.rangeOut, rangeN: yt.rangeN,
      landing: sky.landing, landingSd: sky.landingSd, landingP10: sky.landingP10, landingP90: sky.landingP90, pAbove: sky.pAbove, ratio: sky.ratio,
      sky: sky.sky, skyReason: sky.skyReason, skyDir: sky.skyDir, theta: sky.theta, credibility: sky.credibility, k: sky.k,
      budgetUsed: budgetUsed, budgetSource: official !== null ? 'official' : budget !== null ? 'draft' : '',
      aligned: appLandingAligned_(p.fy, months, used.tau, used.w, budgetUsed, { k: sky.k, pending: pendingCut !== '' || sky.skyReason === 'eval_pending' || sky.skyReason === 'stale_actuals' }),
      frozen: !!frozen,
      reach: appLandingReach_(sky, { draft: budget, official: official, officialNo: v.officialNo || null }, used),
      scoredMonths: scored[p.plan_id] ? Object.keys(scored[p.plan_id].months).length : 0
    };
  }).sort((x, y) => String(y.fy).localeCompare(String(x.fy)) || String(x.clientName).localeCompare(String(y.clientName), 'ja'));
  return { plans: rows, prior: { learned: learned, used: used } };
}

/**
 * 予測の画面に出す、計画 1 つの試しの数（年度の見込みの試し・予算に届く見込み。計画の一覧の控え appPortfolioAll_ から。表は読み直さない）。
 * 読めないときは null（予測の画面はそのまま出す）。返り値: { planId, aligned, reach, scoredMonths } | null
 */
function appPlanShadow_(planId) {
  try {
    const p = appPortfolioAll_().plans.filter(x => x.planId === String(planId || ''))[0];
    return p ? { planId: p.planId, aligned: p.aligned || null, reach: p.reach || null, scoredMonths: p.scoredMonths } : null;
  } catch (e) {
    Logger.log('試しの数: ' + (e && e.message ? e.message : e));
    return null;
  }
}

/**
 * 計画ごとの、今の決まりで測った月（D4〜D6）: { 計画の ID: { any, months: { 'yyyy/MM': true } } }。
 * 月 = EVAL_LOG の neutral の行が appEvalRowCurrent_ の月（締まった月で、今の検証の版で B-2 が書いた行。精度・振り返りと同じ）。
 * any = その月が 1 つでもある = 今の決まりで B-2 が検証の表を書き直した計画。EVAL_LOG は全計画の分を 1 回で読む（appEngAll_）。
 * cut = 計画ごとの締まった月の境目（appLandingCutoff_）
 */
function appPortfolioScored_(ids, cut) {
  const rows = appEngAll_('EVAL_LOG', ids);
  const out = {};
  ids.forEach(id => {
    const o = out[id] = { any: false, months: {} };
    (rows[id] || []).forEach(r => {
      if (String(r.scenario) !== 'neutral' || !appEvalRowCurrent_(r, cut[id])) return;
      o.months[appYm_(r.target_month)] = true;
      o.any = true;
    });
  });
  return out;
}

/** セルの型と中身から 'yyyy/MM'（日時は日本の暦で。時刻つきの ISO の文字も時差を見て日本の暦で。読めなければ ''） */
function appCellYm_(t, text) {
  if (text === '' || text === null || text === undefined) return '';
  if (t === 'd') { const d = new Date(text); return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, APP_TZ, 'yyyy/MM'); }
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(String(text))) { const ms = appLandingTime_('', text); return ms === null ? '' : Utilities.formatDate(new Date(ms), APP_TZ, 'yyyy/MM'); }
  const m = /^(\d{4})\D(\d{1,2})/.exec(String(text));
  return m ? m[1] + '/' + ('0' + m[2]).slice(-2) : '';
}

/**
 * 年度の締まった月（cutoffYm より前）の暫定実績（2026-10-06）。締まっていない・途中の月（月末から 5 日たたずに取り込んだ月とその後の月。
 * appLandingCutoff_）は数えない。
 * 実績の行が無い・実績が空の締まった月は 0 円として数える（ZAC に記録が無い月）。actualMonths は締まった月の数。
 * forecastYtd は締まった月の予測（検証の表の P50）の合計（予実の差を見る）。予測の無い締まった月があれば出さない（比べられない）。
 * 今の決まりで B-2 が検証の表を書き直していない計画は、どの月も予測が無い（appPortfolioData_ が渡さない）ので、予実の差も幅の外も出ない。
 * 実績が空でも予測のある月は、0 円と予測で比べる。actual は締まった月ごとの実績（着地見込みの計算に渡す）
 */
function appPlanYtd_(fy, cmpByYm, cutoffYm) {
  const yms = cutoffYm ? appLandingFyYms_(fy).filter(ym => ym < cutoffYm) : [];
  const res = { actualYtd: null, actualMonths: yms.length, forecastYtd: null, rangeOut: 0, rangeN: 0, actual: {} };
  let fcN = 0;
  yms.forEach(ym => {
    const c = cmpByYm[ym];
    const a = c && typeof c.actual === 'number' ? c.actual : 0;
    res.actualYtd = (res.actualYtd || 0) + a;
    if (!c) return;
    res.actual[ym] = a;
    if (c.p50 !== null) { res.forecastYtd = (res.forecastYtd || 0) + c.p50; fcN++; }
    if (c.out !== null) { res.rangeN++; res.rangeOut += c.out; }
  });
  if (fcN !== yms.length) res.forecastYtd = null;
  return res;
}

