/**
 * Forecast.js — 新アプリでの予測の実行（段階2-2b）。
 * 旧来の予測（A-9）を、データ本体から組み立てた計算用ブックで動かし、書き換わったシートだけをデータ本体へ戻す。
 * 結果は FORECAST_RUNS（1 回 1 行）と FORECAST_MONTHLY（月ごと）に残す。種 = 実行の ID、「今」= 実行を始めた時刻。
 * 1 回の実行の上限（6 分）に収まるよう、2 つの処理に分ける。
 *   計算（FORECAST.RUN）: 組み立て → 実行 → 書き換わったシートの控えを取る（データ本体には書かない）
 *   保存（FORECAST.RUN_SAVE）: 計算の間にデータ本体が変わっていないことを確かめてから、控えを書き戻す
 * 旧ブックは変えない。並行運用の間に旧ブックを取り込み直すと、計算用の表は旧ブックの内容に戻る（予測の記録は残る）。
 */

function appPlanOf_(planId) {
  const plan = appReadTable_('PLANS').filter(x => x.plan_id === String(planId || ''))[0];
  if (!plan) throw new Error('計画が見つかりません。');
  return plan;
}

/** 計画の計算用の表の中身（シートごとのハッシュ）をまとめたハッシュ。計算の間に変わっていないかを確かめる */
function appPlanInputHash_(planId) {
  const rows = appReadTable_('ENG_SHEETS').filter(r => r.plan_id === planId).map(r => r.sheet + ':' + r.content_hash).sort();
  return appSha256Hex_(rows.join('|'));
}

/** 計算: データ本体から組み立てて旧来の予測を動かし、書き換わったシートを控える。確認が要るときは何も控えずに返す */
function appForecastRunCalc_(ctx, p, job) {
  const plan = appPlanOf_(p.planId);
  const t0 = new Date().getTime();
  const runId = appId_('RUN');
  const seed = runId;
  const confirms = (p.confirms || []).map(String);
  const inputHash = appPlanInputHash_(plan.plan_id);
  const stored = appStoredHashes_(plan.plan_id);
  return appWithLock_(() => {
    const scratch = appParityScratch_(plan);
    const build = appScratchFromStore_(scratch, plan.plan_id);
    const t1 = new Date().getTime();
    const run = appRunLegacyForecast_(scratch, { asOfMs: t0, seed: seed, confirms: confirms, actor: ctx.actor });
    if (!run.ok) return { needConfirm: run.needConfirm, planId: plan.plan_id, audit: { entityId: plan.plan_id, clientId: plan.client_id } };
    const t2 = new Date().getTime();
    const headline = appForecastHeadline_(scratch);
    const cap = appCaptureChanged_(scratch, plan.plan_id, stored);
    const changed = cap.changed;
    const unknown = cap.unknown;
    appJobPutResult_(job.id + '_SAVE', { changed: changed });
    const t3 = new Date().getTime();
    return {
      __next: { kind: 'FORECAST.RUN_SAVE', payload: {
        planId: plan.plan_id, runId: runId, seed: seed, asOfMs: t0, inputHash: inputHash, confirms: confirms, parentJobId: job.id,
        engine: { version: run.version, sourceSha256: run.sourceSha256 }, headline: headline,
        changed: changed.map(e => e.sheetRow.sheet), unknown: unknown, startedAt: Utilities.formatDate(new Date(t0), APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ"),
        build: build.filter(x => x.mismatch || x.forcedText || x.formatMismatches).map(x => x.sheet),
        timing: { buildMs: t1 - t0, runMs: t2 - t1, captureMs: t3 - t2 } } },
      audit: { entityId: plan.plan_id, clientId: plan.client_id }
    };
  });
}

/** 保存: 控えをデータ本体へ書き戻し、予測の記録を足す */
function appForecastRunSave_(ctx, p) {
  const plan = appPlanOf_(p.planId);
  const saved = appJobGetResult_(p.parentJobId + '_SAVE');
  if (!saved.found) throw new Error('計算した結果の控えが見つかりません（保存期間が過ぎた）。もう一度実行してください。');
  const changed = saved.value.changed;
  return appWithLock_(() => {
    if (appPlanInputHash_(plan.plan_id) !== p.inputHash) throw new Error('予測を計算している間にデータ本体が変わりました。もう一度実行してください。');
    const now = appNowIso_();
    const names = changed.map(e => e.sheetRow.sheet);
    const written = appWriteChanged_(ctx, plan.plan_id, changed, p.runId);
    const h = p.headline || { annual: {}, objective: {}, monthly: [], objectiveMonthly: [] };
    const num = v => (typeof v === 'number' && isFinite(v) ? v : null);
    appInsertRows_('FORECAST_RUNS', [{
      run_id: p.runId, plan_id: plan.plan_id, status: 'DONE', engine_version: p.engine.version, engine_sha256: p.engine.sourceSha256,
      seed: p.seed, as_of: Utilities.formatDate(new Date(p.asOfMs), APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ"), input_hash: p.inputHash,
      annual_p10: num(h.annual.p10), annual_p50: num(h.annual.p50), annual_p90: num(h.annual.p90),
      objective_p10: num(h.objective.p10), objective_p50: num(h.objective.p50), objective_p90: num(h.objective.p90),
      changed_sheets_json: names, confirms_json: p.confirms, started_at: p.startedAt, finished_at: now, actor_email: ctx.actor
    }]);
    const obj = {};
    (h.objectiveMonthly || []).forEach(m => { obj[m.month] = m; });
    appInsertRows_('FORECAST_MONTHLY', (h.monthly || []).map(m => ({
      run_id: p.runId, plan_id: plan.plan_id, ym: m.month, p10: num(m.p10), p50: num(m.p50), p90: num(m.p90),
      obj_p10: obj[m.month] ? num(obj[m.month].p10) : null, obj_p50: obj[m.month] ? num(obj[m.month].p50) : null, obj_p90: obj[m.month] ? num(obj[m.month].p90) : null
    })));
    return { runId: p.runId, planId: plan.plan_id, changed: names, written: written, headline: h, unknown: p.unknown || [],
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

/** 計画のシートごとの中身のハッシュ（{ シート名: content_hash }） */
function appStoredHashes_(planId) {
  const stored = {};
  appReadTable_('ENG_SHEETS').filter(r => r.plan_id === planId).forEach(r => { stored[r.sheet] = r.content_hash; });
  return stored;
}

/** 控えの形にしたシートを、データ本体の計画 planId の行と入れ替える（ロックの中で呼ぶ）。履歴の表は足された行だけを書く */
function appWriteChanged_(ctx, planId, changed, batchId) {
  const now = appNowIso_();
  const names = changed.map(e => e.sheetRow.sheet);
  const inChanged = r => r.plan_id === planId && names.indexOf(r.sheet) >= 0;
  const written = {};
  changed.forEach(enc => {
    const name = enc.sheetRow.sheet;
    if (APP_ENGINE_SHEETS[name].mode !== 'table') return;
    const rows = enc.sheetRow.mode === 'table' ? enc.tableRows : [];   // 見出しが違えば ENG_ROWS 側に持つ
    written['ENG_' + name] = appReplaceOrAppend_('ENG_' + name, planId, rows);
  });
  const segs = [].concat.apply([], changed.map(e => e.rowSegs));
  const fmts = [].concat.apply([], changed.map(e => e.formatRows));
  written.ENG_ROWS = appReplaceRows_('ENG_ROWS', r => !inChanged(r), segs);
  written.ENG_FORMATS = appReplaceRows_('ENG_FORMATS', r => !inChanged(r), fmts);
  written.ENG_SHEETS = appReplaceRows_('ENG_SHEETS', r => !inChanged(r),
    changed.map(e => Object.assign({}, e.sheetRow, { import_batch_id: batchId, updated_at: now, updated_by: ctx.actor })));
  return written;
}

/** データ本体の OUTPUT（取り込んだ旧ブック、または最後の予測）から、主な結果を読む */
function appStoredHeadline_(planId) {
  const segs = appReadTable_('ENG_ROWS').filter(r => r.plan_id === planId && r.sheet === 'OUTPUT');
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
  const clients = {};
  appReadTable_('CLIENTS').forEach(c => { clients[c.client_id] = c.client_name; });
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
