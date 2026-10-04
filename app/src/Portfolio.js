/**
 * Portfolio.js — 計画（クライアント × 年度）の一覧と、新しい計画を新アプリで作る。
 *
 * 新しい計画は、旧ブックを経ずに作る: 計算用ブックの上で旧来の A-1 初期セットアップ（setupForecastBook の中身）と
 * 設定の保存（saveInitialSetupSettings）を動かし、できたシートをデータ本体に計画として保存する。
 * その後は、A-2 売上データの取り込み → A-3 → 予測、と旧ブックから移した計画と同じに使える。
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
  if (!clientName) throw new Error('クライアントを選んでください。');
  if (clientName.length > 100) throw new Error('クライアントの名前が長すぎます。');
  if (!fy || fy < 2000 || fy > 2100 || Math.floor(fy) !== fy) throw new Error('年度（FY）を 4 桁の数で入れてください。');
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
    const scratch = appParityScratch_(appNewPlanTz_());
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
        is_active: true, note: '新アプリで作成', created_at: now, created_by: ctx.actor, updated_at: now, updated_by: ctx.actor, row_version: 1 };
      ops.push({ table: 'CLIENTS', mode: 'ensure', rows: [client] });
    }
    const cap = appCaptureChanged_(appScratchBook_(), planId, {});
    const missing = Object.keys(APP_ENGINE_SHEETS).filter(n => APP_SETUP_ORDER.indexOf(n) >= 0 && !cap.changed.some(e => e.sheetRow.sheet === n));
    if (missing.length) throw new Error('初期セットアップでできなかったシートがあります: ' + missing.join('、'));
    const batchId = appId_('NEW');
    appChangedOps_(ctx, planId, cap.changed, batchId).forEach(op => ops.push(op));
    ops.push({ table: 'PLANS', mode: 'ensure', rows: [{ plan_id: planId, client_id: client.client_id, fy: String(c.fy), client_label: c.clientName,
      people_csv: c.peopleCsv, source_book_id: '', locale: appNewPlanTz_().locale, time_zone: appNewPlanTz_().time_zone, state: 'ACTIVE', note: '新アプリで作成',
      created_at: now, created_by: ctx.actor, updated_at: now, updated_by: ctx.actor, row_version: 1 }] });
    const written = appJournalRun_(ctx, '計画の作成（' + c.clientName + ' FY' + c.fy + '）', planId, ops);
    return { planId: planId, clientName: appClientDisplayName_(c.clientName), fy: c.fy, sheets: cap.changed.length, written: written, engine: p.engine,
      timing: { buildMs: p.buildMs, saveMs: new Date().getTime() - t0 }, audit: { entityId: planId, clientId: client.client_id } };
  });
}

// ---- 一覧 ----

/** 全部の計画の要点（最新の予測・前回からの変化・予算・精度・手順の進み） */
function appPortfolio_() {
  const clients = appClientNameMap_();   // 画面に出す名前（半角カナ・株式会社などを除いた、ふつうの表記）
  const runs = {};
  appReadTable_('FORECAST_RUNS').filter(r => r.status === 'DONE').forEach(r => { (runs[r.plan_id] = runs[r.plan_id] || []).push(r); });
  const outRows = {};
  appReadTable_('ENG_ROWS').filter(r => r.sheet === 'OUTPUT' && Number(r.col_from) === 1).forEach(s => {
    const n = Number(s.row_no);
    if (n === 1 || n === 26 || (n >= 29 && n <= 40)) { (outRows[s.plan_id] = outRows[s.plan_id] || {})[n] = JSON.parse(s.cells_json); }
  });
  const ape = {};
  appReadTable_('ENG_EVAL_COMPARE_MONTHLY').forEach(r => {
    const v = Number(r.ape_p50);
    if (r.actual_total !== '' && r.actual_total !== null && isFinite(v) && r.ape_p50 !== '') (ape[r.plan_id] = ape[r.plan_id] || []).push(v);
  });
  const steps = {};
  appReadTable_('ENG_PROCESS_STATUS').forEach(r => { (steps[r.plan_id] = steps[r.plan_id] || []).push(r); });
  const ver = appVersionSummary_();
  const num = x => { if (!x) return null; const t = x.charAt(0); if (t !== 'n') return null; const v = Number(x.slice(1)); return isFinite(v) ? v : null; };
  return appReadTable_('PLANS').filter(p => p.state !== 'ARCHIVED').map(p => {
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
    const st = steps[p.plan_id] || [];
    const errors = st.filter(s => String(s.status).toLowerCase() === 'error').map(s => s.step_key);
    const a = ape[p.plan_id] || [];
    return {
      planId: p.plan_id, clientName: clients[p.client_id] || p.client_label, fy: p.fy, source: p.source_book_id ? 'book' : 'app',
      p10: latest ? latest.annual_p10 : stored.p10 === undefined ? null : stored.p10,
      p50: latest ? latest.annual_p50 : stored.p50 === undefined ? null : stored.p50,
      p90: latest ? latest.annual_p90 : stored.p90 === undefined ? null : stored.p90,
      prevP50: prev ? prev.annual_p50 : null,
      lastRunAt: latest ? latest.finished_at : '', runs: rs.length,
      budget: adopted === null && uplift === null ? null : (adopted || 0) + (uplift || 0),
      mape: a.length ? a.reduce((x, y) => x + y, 0) / a.length : null, mapeMonths: a.length,
      stepsDone: st.filter(s => String(s.status).toLowerCase() === 'success').length, stepsTotal: st.length, stepErrors: errors,
      createdAt: p.created_at,
      officialNo: (ver[p.plan_id] || {}).officialNo || null, officialFinal: (ver[p.plan_id] || {}).officialFinal === undefined ? null : ver[p.plan_id].officialFinal,
      pendingNo: (ver[p.plan_id] || {}).pendingNo || null
    };
  }).sort((x, y) => String(y.fy).localeCompare(String(x.fy)) || String(x.clientName).localeCompare(String(y.clientName), 'ja'));
}
