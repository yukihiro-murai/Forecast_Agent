/**
 * Forecast.js — 新アプリでの予測の実行（段階2-2b）。
 * 旧来の予測（A-9）を、データ本体から組み立てた計算用ブックで動かし、書き換わったシートだけをデータ本体へ戻す。
 * 結果は FORECAST_RUNS（1 回 1 行）と FORECAST_MONTHLY（月ごと）に残す。種 = 実行の ID、「今」= 実行を始めた時刻。
 * 1 回の実行の上限（6 分）に収まるよう、3 つの処理に分ける（2026-10-03 から。組み立ては何回かに分かれることがある）。
 *   組み立て（FORECAST.RUN）: データ本体から計算用ブックを組み立てる
 *   計算（FORECAST.RUN_CALC）: 旧来の予測を動かす（データ本体には書かない）
 *   保存（FORECAST.RUN_SAVE）: 書き換わったシートを控えの形にし、計算の間にデータ本体が変わっていないことを確かめてから書く
 * 計算用ブックは処理の印（APP_SCRATCH_OWNER）で、組み立てたままかを確かめる。書くときは保存の控え（Journal.js）を置く。
 */

function appPlanOf_(planId) {
  const plan = appReadTable_('PLANS').filter(x => x.plan_id === String(planId || ''))[0];
  if (!plan) throw new Error('計画が見つかりません。');
  return plan;
}

/** 計画の計算用の表の中身（シートごとのハッシュ）をまとめたハッシュ。計算の間に変わっていないかを確かめる。読むのはその計画の行だけ */
function appPlanInputHash_(planId) {
  const rows = appReadPlanTable_('ENG_SHEETS', planId).map(r => r.sheet + ':' + r.content_hash).sort();
  return appSha256Hex_(rows.join('|'));
}

/**
 * 組み立て（FORECAST.RUN）: データ本体から計算用ブックを組み立てる。1 回の上限に収まらなければ、同じ処理を続けて動かす。
 * 「今」（asOfMs）・種（runId）・入力のハッシュは最初の回に決め、続きの処理に渡す
 */
function appForecastRunBuild_(ctx, p, job) {
  const plan = appPlanOf_(p.planId);
  const jobStart = new Date().getTime();
  return appWithLock_(() => {
    if (!p.build) appJournalRecover_(ctx);   // 前の保存が途中で止まっていれば、先に書き終える
    const inputHash = appPlanInputHash_(plan.plan_id);
    if (p.build && inputHash !== p.inputHash) throw new Error('組み立てている間にデータ本体が変わりました。もう一度実行してください。');
    const runId = p.runId || appId_('RUN');
    const asOfMs = p.asOfMs || jobStart;
    const step = appScratchBuildStep_(appWorkScratch_(plan), plan.plan_id, null, p.build || null, appBuildDeadline_(jobStart));
    const payload = { planId: plan.plan_id, confirms: (p.confirms || []).map(String), runId: runId, seed: runId, asOfMs: asOfMs, inputHash: inputHash,
      startedAt: Utilities.formatDate(new Date(asOfMs), APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ"), build: step.state,
      buildMs: Number(p.buildMs || 0) + (new Date().getTime() - jobStart) };
    return { __next: { kind: step.complete ? 'FORECAST.RUN_CALC' : 'FORECAST.RUN', payload: payload },
      audit: { entityId: plan.plan_id, clientId: plan.client_id } };
  });
}

/** 計算（FORECAST.RUN_CALC）: 組み立てた計算用ブックで旧来の予測を動かす。確認が要るときは何も保存せずに返す */
function appForecastRunCalc_(ctx, p) {
  const plan = appPlanOf_(p.planId);
  const t0 = new Date().getTime();
  return appWithLock_(() => {
    if (!p.build || !appScratchOwnedBy_(p.build.token)) throw new Error('計算用ブックがほかの処理で使われました。もう一度実行してください。');
    if (appPlanInputHash_(plan.plan_id) !== p.inputHash) throw new Error('予測を計算している間にデータ本体が変わりました。もう一度実行してください。');
    const scratch = appWorkScratch_(plan);
    const run = appRunLegacyForecast_(scratch, { asOfMs: p.asOfMs, seed: p.seed, confirms: p.confirms || [], actor: ctx.actor });
    if (!run.ok) return { needConfirm: run.needConfirm, planId: plan.plan_id, audit: { entityId: plan.plan_id, clientId: plan.client_id } };
    const payload = Object.assign({}, p, { engine: { version: run.version, sourceSha256: run.sourceSha256 }, headline: appForecastHeadline_(scratch),
      runMs: new Date().getTime() - t0 });
    return { __next: { kind: 'FORECAST.RUN_SAVE', payload: payload }, audit: { entityId: plan.plan_id, clientId: plan.client_id } };
  });
}

/** 保存（FORECAST.RUN_SAVE）: 計算の間にデータ本体が変わっていないことを確かめ、書き換わったシートと予測の記録を、控えを置いてから書く */
function appForecastRunSave_(ctx, p) {
  const plan = appPlanOf_(p.planId);
  const t0 = new Date().getTime();
  return appWithLock_(() => {
    appJournalRecover_(ctx);   // 書きかけの控えを先に書き終える（途中の表から控えを作らない）
    if (!p.build || !appScratchOwnedBy_(p.build.token)) throw new Error('計算用ブックがほかの処理で使われました。もう一度実行してください。');
    if (appPlanInputHash_(plan.plan_id) !== p.inputHash) throw new Error('予測を計算している間にデータ本体が変わりました。もう一度実行してください。');
    const cap = appCaptureChanged_(appWorkScratch_(plan), plan.plan_id, appStoredHashes_(plan.plan_id));
    const t1 = new Date().getTime();
    const names = cap.changed.map(e => e.sheetRow.sheet);
    const h = p.headline || { annual: {}, objective: {}, monthly: [], objectiveMonthly: [] };
    const num = v => (typeof v === 'number' && isFinite(v) ? v : null);
    const now = appNowIso_();
    const obj = {};
    (h.objectiveMonthly || []).forEach(m => { obj[m.month] = m; });
    const ops = appChangedOps_(ctx, plan.plan_id, cap.changed, p.runId).concat([
      { table: 'FORECAST_RUNS', mode: 'ensure', rows: [{
        run_id: p.runId, plan_id: plan.plan_id, status: 'DONE', engine_version: p.engine.version, engine_sha256: p.engine.sourceSha256,
        seed: p.seed, as_of: Utilities.formatDate(new Date(p.asOfMs), APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ"), input_hash: p.inputHash,
        annual_p10: num(h.annual.p10), annual_p50: num(h.annual.p50), annual_p90: num(h.annual.p90),
        objective_p10: num(h.objective.p10), objective_p50: num(h.objective.p50), objective_p90: num(h.objective.p90),
        changed_sheets_json: names, confirms_json: p.confirms, started_at: p.startedAt, finished_at: now, actor_email: ctx.actor }] },
      { table: 'FORECAST_MONTHLY', mode: 'ensure', rows: (h.monthly || []).map(m => ({
        run_id: p.runId, plan_id: plan.plan_id, ym: m.month, p10: num(m.p10), p50: num(m.p50), p90: num(m.p90),
        obj_p10: obj[m.month] ? num(obj[m.month].p10) : null, obj_p50: obj[m.month] ? num(obj[m.month].p50) : null, obj_p90: obj[m.month] ? num(obj[m.month].p90) : null })) }
    ]);
    const written = appJournalRun_(ctx, '予測の保存（' + p.runId + '）', plan.plan_id, ops);
    appScratchMarkSynced_(plan.plan_id, p.build.token, null, p.build.problems);   // 次の予測は組み立て直さずに使える
    return { runId: p.runId, planId: plan.plan_id, changed: names, written: written, headline: h, unknown: cap.unknown,
      build: (p.build && p.build.problems) || [],
      timing: { buildMs: p.buildMs || 0, runMs: p.runMs || 0, captureMs: t1 - t0, saveMs: new Date().getTime() - t1 },
      audit: { entityId: p.runId, clientId: plan.client_id } };
  });
}

/**
 * 計算用ブックの各シートを控えの形にし、データ本体と中身（ハッシュ）が違うものだけを返す。
 * only を渡すと、そのシートだけを見る（一部だけを組み立てた計算用ブックのとき）
 */
function appCaptureChanged_(scratch, planId, stored, only) {
  const changed = [];
  const unknown = [];
  scratch.getSheets().forEach(sh => {
    const name = sh.getName();
    if (only && only.indexOf(name) < 0) return;
    const reg = APP_ENGINE_SHEETS[name];
    if (!reg) { unknown.push(name); return; }
    const enc = appEngEncodeSheet_(planId, appSheetSnapshot_(sh, reg.header ? reg.header.length : 0));
    if (stored[name] !== enc.sheetRow.content_hash) changed.push(enc);
  });
  return { changed: changed, unknown: unknown };
}

/** 計画のシートごとの中身のハッシュ（{ シート名: content_hash }）。読むのはその計画の行だけ */
function appStoredHashes_(planId) {
  const stored = {};
  appReadPlanTable_('ENG_SHEETS', planId).forEach(r => { stored[r.sheet] = r.content_hash; });
  return stored;
}

/**
 * 控えの形にしたシートで、データ本体の計画 planId の行を入れ替える書き方（ops。appJournalRun_ に渡す）。
 * 履歴の表は、前の行がそのままなら足された行だけを書く
 */
function appChangedOps_(ctx, planId, changed, batchId) {
  const now = appNowIso_();
  const names = changed.map(e => e.sheetRow.sheet);
  const ops = [];
  changed.forEach(enc => {
    const name = enc.sheetRow.sheet;
    if (APP_ENGINE_SHEETS[name].mode !== 'table') return;
    const rows = enc.sheetRow.mode === 'table' ? enc.tableRows : [];   // 見出しが違えば ENG_ROWS 側に持つ
    ops.push(appOpReplaceOrAppend_('ENG_' + name, planId, rows));
  });
  if (!changed.length) return ops;
  const segs = [].concat.apply([], changed.map(e => e.rowSegs));
  const fmts = [].concat.apply([], changed.map(e => e.formatRows));
  ops.push(appOpReplacePlan_('ENG_ROWS', planId, names, segs));
  ops.push(appOpReplacePlan_('ENG_FORMATS', planId, names, fmts));
  // シートの大きさとハッシュは最後に書く（途中で止まっても、控えから書き直すまで「入力のハッシュ」は前のまま）
  ops.push(appOpReplacePlan_('ENG_SHEETS', planId, names,
    changed.map(e => Object.assign({}, e.sheetRow, { import_batch_id: batchId, updated_at: now, updated_by: ctx.actor }))));
  return ops;
}

/** データ本体の OUTPUT（計画を作ったとき、または最後の予測）から、主な結果を読む。読むのはその計画の行だけ */
function appStoredHeadline_(planId) {
  const segs = appReadPlanTable_('ENG_ROWS', planId).filter(r => r.sheet === 'OUTPUT');
  const rows = {};
  segs.forEach(s => { if (Number(s.col_from) === 1) rows[Number(s.row_no)] = JSON.parse(s.cells_json); });
  const cell = (r, c) => { const x = (rows[r] || [])[c - 1]; return x ? appCellDecode_(x.charAt(0) === 'f' ? 'e' : x.charAt(0), x.slice(1)) : ''; };
  const num = v => (typeof v === 'number' && isFinite(v) ? v : null);
  if (!rows[26]) return null;
  const month = v => (appIsDate_(v) ? Utilities.formatDate(v, APP_TZ, 'yyyy/MM') : String(v));
  const monthly = [];
  for (let r = 29; r <= 40; r++) monthly.push({ month: month(cell(r, 1)), p10: num(cell(r, 2)), p50: num(cell(r, 3)), p90: num(cell(r, 4)) });
  return { title: String(cell(1, 1) || ''), annual: { p10: num(cell(26, 2)), p50: num(cell(26, 3)), p90: num(cell(26, 4)) },
    objective: { p10: num(cell(65, 2)), p50: num(cell(65, 3)), p90: num(cell(65, 4)) }, monthly: monthly };
}

/** 画面: 計画の最新の予測（新アプリで動かした記録）と、データ本体の OUTPUT の結果 */
function appForecastLatest_(planId) {
  const plan = appPlanOf_(planId);
  const clients = appClientNameMap_();   // 画面に出す名前（半角カナ・株式会社などを除いた、ふつうの表記）
  const runs = appReadTable_('FORECAST_RUNS').filter(r => r.plan_id === plan.plan_id).map(appStripRow_)
    .sort((a, b) => (a.finished_at < b.finished_at ? 1 : -1));
  const latest = runs[0] || null;
  const monthly = latest ? appReadTable_('FORECAST_MONTHLY').filter(m => m.run_id === latest.run_id).map(appStripRow_) : [];
  return {
    plan: { planId: plan.plan_id, clientName: clients[plan.client_id] || plan.client_label, fy: plan.fy },
    latest: latest, monthly: monthly,
    runs: runs.slice(0, 10).map(r => ({ runId: r.run_id, finishedAt: r.finished_at, actor: r.actor_email, p50: r.annual_p50 })),
    stored: appStoredHeadline_(plan.plan_id)
  };
}
