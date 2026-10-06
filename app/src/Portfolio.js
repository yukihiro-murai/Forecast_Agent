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
 * 着地見込みと空模様は Landing.js（締まった月の実績から今年の水準を学ぶ）。今日の日付で変わる（霧の判定・前提の変化の日数）ので、控えは日ごと
 */
function appPortfolio_() {
  const clients = appClientNameMap_();   // 画面に出す名前（半角カナ・株式会社などを除いた、ふつうの表記）
  const today = appToday_();
  const todayYm = today.slice(0, 7).replace('-', '/');
  const runs = {};
  appReadTable_('FORECAST_RUNS').filter(r => r.status === 'DONE').forEach(r => { (runs[r.plan_id] = runs[r.plan_id] || []).push(r); });
  const outRows = {};
  appReadTable_('ENG_ROWS').filter(r => r.sheet === 'OUTPUT' && Number(r.col_from) === 1).forEach(s => {
    const n = Number(s.row_no);
    if (n === 1 || n === 26 || (n >= 29 && n <= 40)) { (outRows[s.plan_id] = outRows[s.plan_id] || {})[n] = JSON.parse(s.cells_json); }
  });
  const ape = {};
  const cmp = {};   // 月ごとの実績と P10/P50/P90（検証の表）。暫定実績・着地見込み・予実の差・全計画の学びに使う（読むだけ。計算は変えない）
  const fnum = x => { const n = x === '' || x === null || x === undefined ? null : Number(x); return n !== null && isFinite(n) ? n : null; };
  appReadTable_('ENG_EVAL_COMPARE_MONTHLY').forEach(r => {
    const v = Number(r.ape_p50);
    if (r.actual_total !== '' && r.actual_total !== null && isFinite(v) && r.ape_p50 !== '') (ape[r.plan_id] = ape[r.plan_id] || []).push(v);
    const ym = appCellYm_(String(r._types || '').charAt(0), r.target_month);
    // 実績が空でも P50 のある行は残す（actual: null。ZAC に記録の無い月 = 売上 0。締まった月なら 0 円として数える）
    const blank = r.actual_total === '' || r.actual_total === null || r.actual_total === undefined;
    const act = blank ? null : Number(r.actual_total);
    const p50 = fnum(r.forecast_total_p50);
    if (ym && (blank ? p50 !== null : isFinite(act))) {
      (cmp[r.plan_id] = cmp[r.plan_id] || {})[ym] = { actual: act, p10: fnum(r.forecast_total_p10), p50: p50, p90: fnum(r.forecast_total_p90),
        out: String(r.range_outside_flag) === '1' ? 1 : String(r.range_outside_flag) === '0' ? 0 : null };
    }
  });
  const steps = {};
  appReadTable_('ENG_PROCESS_STATUS').forEach(r => { (steps[r.plan_id] = steps[r.plan_id] || []).push(r); });
  const ver = appVersionSummary_();
  const num = x => { if (!x) return null; const t = x.charAt(0); if (t !== 'n') return null; const v = Number(x.slice(1)); return isFinite(v) ? v : null; };
  const plans = appReadTable_('PLANS').filter(p => p.state !== 'ARCHIVED');
  // 締まった月の境目（B-1 → B-2 の順に動いた取り込み）と、全計画の締まった月から学ぶ τ・w（1 回だけ）
  const cut = {};
  plans.forEach(p => { cut[p.plan_id] = appLandingCutoff_(steps[p.plan_id] || []); });
  const prior = appLandingPrior_(plans.map(p => {
    const c = cmp[p.plan_id] || {};
    return Object.keys(c).filter(ym => cut[p.plan_id] && ym < cut[p.plan_id]).sort()
      .map(ym => ({ f: c[ym].p50, a: c[ym].actual === null ? 0 : c[ym].actual, p10: c[ym].p10, p90: c[ym].p90 }));   // 実績の空の締まった月は 0 円（着地の計算と同じ）
  }));
  return plans.map(p => {
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
    const sky = appLandingSky_({ fy: p.fy, months: appLandingMonths_(o), actual: yt.actual, cutoffYm: cut[p.plan_id], todayYm: todayYm, budget: budgetUsed,
      tau: prior.tau, w: prior.w, runs: rs.slice(0, 2).map(r => ({ p50: r.annual_p50, ageDays: appLandingAgeDays_(r.finished_at, today) })) });
    const st = steps[p.plan_id] || [];
    const errors = st.filter(s => String(s.status).toLowerCase() === 'error').map(s => s.step_key);
    const a = ape[p.plan_id] || [];
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
      budgetUsed: budgetUsed, budgetSource: official !== null ? 'official' : budget !== null ? 'draft' : ''
    };
  }).sort((x, y) => String(y.fy).localeCompare(String(x.fy)) || String(x.clientName).localeCompare(String(y.clientName), 'ja'));
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
 * 年度の締まった月（cutoffYm より前）の暫定実績（2026-10-06）。取り込んだ月とその後の月（締まっていない・途中の月）は数えない。
 * 実績の行が無い・実績が空の締まった月は 0 円として数える（ZAC に記録が無い月）。actualMonths は締まった月の数。
 * forecastYtd は締まった月の予測（検証の表の P50）の合計（予実の差を見る）。予測の無い締まった月があれば出さない（比べられない）。
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

