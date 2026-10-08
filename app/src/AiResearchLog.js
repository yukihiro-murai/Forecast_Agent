/**
 * AiResearchLog.js — AI 調査の記録（AI_RESEARCH_LOG。SCHEMA_PLAN_v10-12_JA.md の 3-2。2026-10-08 村井さん承認「版10はおすすめで」）。
 * A-4 の保存（Plan.js の appPlanRunSave_。操作が AI.RESEARCH で AI_RESEARCH_STRUCTURED が変わったとき）に、その回に書いた行を同じ控えで足す。
 * 旧来の A-4 は調べるたびに AI_RESEARCH_STRUCTURED を消して書き直すので、書き直した後の行がその回の中身。失敗した回は旧来の計算が前の行を残し
 * （保存まで進まないか、表が変わらない）、記録も足さない。
 * started_by: 週 1 回の自動の調査（AutoResearch.js）が始めた回は AUTO（自動の調査の控えの「始めた処理の番号」と、この処理の最初の処理の番号が同じ）。
 * ほかは始めた人のメール。row_json は元の行の 22 列そのまま（見出しの順の配列で、列ごとに「型 1 文字 + 値」。appAiResearchLogCells_ で元の値に戻る。
 * 列の名前を行ごとに持たないのは、年度の控え（8MB まで）を小さくするため）。記録は計算に使わない。
 * 見る画面（分析の「市場」・学びの「AI の学び」）は、それぞれの画面を作るときに足す。
 */

/** AI_RESEARCH_STRUCTURED の列（旧来の見出しと同じ 22 列） */
function appAiResearchHeader_() {
  return APP_ENGINE_SHEETS.AI_RESEARCH_STRUCTURED.header;
}

/**
 * A-4 の保存に足す記録の控え（ops）。A-4 でない・AI_RESEARCH_STRUCTURED が変わっていない・表の形でないときは []。
 * 作れないときはエラーのログに残して []（A-4 の保存は止めない。調べ直すと費用がかかる）
 */
function appAiResearchLogOps_(ctx, plan, p, changed, job) {
  if (!p || p.action !== 'AI.RESEARCH') return [];
  try {
    const enc = (changed || []).filter(e => e.sheetRow.sheet === 'AI_RESEARCH_STRUCTURED')[0];
    if (!enc) return [];
    if (enc.sheetRow.mode !== 'table') throw new Error('AI_RESEARCH_STRUCTURED の見出しが旧来の定義と違うので、AI 調査の記録を足しませんでした。');
    const startedBy = appAiResearchStartedBy_(ctx, plan.plan_id, job);
    const now = appNowIso_();
    return appLogOps_('AI_RESEARCH_LOG', enc.tableRows.map(r => appAiResearchLogRow_(plan.plan_id, r, {
      research_id: appStableLogId_(APP_LOG_PREFIX.AI_RESEARCH_LOG, [p.actionId, r.seq]), action_id: p.actionId, started_by: startedBy, recorded_at: now })));
  } catch (e) {
    appLogError_('LOG.AI_RESEARCH', e, ctx);
    return [];
  }
}

/** 誰が始めたか: 週 1 回の自動の調査が始めた処理の続きなら AUTO、ほかは始めた人（控えが読めなければ人が始めたとみなす） */
function appAiResearchStartedBy_(ctx, planId, job) {
  try {
    const chain = appAutoResearchState_().chain;
    if (chain && job && chain.planId === planId && appAutoResearchRootOf_(job) === chain.jobId) return 'AUTO';
  } catch (e) { /* 人が始めたとみなす */ }
  return String((ctx && ctx.actor) || '');
}

/**
 * データ本体の形の 1 行（ENG_AI_RESEARCH_STRUCTURED: 見出しの列 = 値の文字・_types = 型の並び）を、記録の 1 行にする。
 * 話題・種類・向き・効く時期は文字、点数は数（数でなければ空）、調べた日は yyyy-MM-dd（日付でなければ文字のまま）
 */
function appAiResearchLogRow_(planId, r, base) {
  const head = appAiResearchHeader_();
  const types = String(r._types || '');
  const cell = [];
  const val = {};
  head.forEach((h, j) => {
    const t = types.charAt(j) || 'e';
    const text = r[h] === undefined || r[h] === null ? '' : String(r[h]);
    cell.push(t + text);
    val[h] = t === 'f' ? text : appCellDecode_(t, text);
  });
  const str = v => (appIsDate_(v) ? Utilities.formatDate(v, APP_TZ, 'yyyy-MM-dd') : String(v === null || v === undefined ? '' : v));
  const num = v => { if (v === '' || v === null || typeof v === 'boolean' || appIsDate_(v)) return null; const n = typeof v === 'number' ? v : Number(v); return isFinite(n) ? n : null; };
  return Object.assign({ plan_id: planId, as_of_date: str(val.as_of_date), topic: str(val.topic), row_type: str(val.row_type), direction: str(val.direction),
    impact_score: num(val.impact_score), confidence: num(val.confidence), event_score: num(val.event_score), benchmark_score: num(val.benchmark_score),
    blended_score: num(val.blended_score), time_horizon: str(val.time_horizon), row_json: cell }, base);
}

/** 記録の row_json（見出しの順の 22 個の「型 1 文字 + 値」）を元の 22 列の値に戻す（{ 列: 値 }。テストと、記録を読む画面で使う） */
function appAiResearchLogCells_(rowJson) {
  const cell = typeof rowJson === 'string' ? JSON.parse(rowJson || '[]') : (rowJson || []);
  const out = {};
  appAiResearchHeader_().forEach((h, j) => {
    const s = String(cell[j] === undefined || cell[j] === null ? 'e' : cell[j]);
    out[h] = s.charAt(0) === 'f' ? s.slice(1) : appCellDecode_(s.charAt(0), s.slice(1));
  });
  return out;
}

/**
 * 一度だけの写し（V10.js の appV10BackfillAiResearchLog_）: 今の AI_RESEARCH_STRUCTURED の 1 回分を、計画ごとに BASELINE として写す。
 * 締めた年度の計画は飛ばす。キーは計画と行の番号から決める（2 つの操作が同時に動かしても、控えの書き直しでも二重にならない）。
 * 計画ごとにロックの中で、書きかけの保存を先に書き終えてから（appJournalRecover_）読む。BASELINE がもうある計画は飛ばす
 * （写しの控えが消えて写しがもう一度動いても、後の A-4 の中身を BASELINE として足さない）。1 回は約 20 秒までで、残りは次の操作で（{ more: true }）。
 * 前の回の根拠の文は残っていない（話題ごとの点数は AI_SCORE_HISTORY にある）
 */
function appAiResearchBackfill_(ctx) {
  if (!appLogReady_()) return { more: true, rows: 0 };   // 表の版がそろう前（次の操作でまた）
  const t0 = new Date().getTime();
  const ids = appReadTable_('PLANS').filter(pl => !appYearIsFrozen_(pl.fy)).map(pl => String(pl.plan_id)).sort();
  let rows = 0, plans = 0;
  for (let i = 0; i < ids.length; i++) {
    if (new Date().getTime() - t0 > APP_INPUT_LOG_BACKFILL_MS) return { more: true, rows: rows };
    const planId = ids[i];
    rows += appWithLock_(() => {
      appJournalRecover_(ctx);   // 書きかけの保存（A-4 の保存を含む）を先に書き終える
      const plan = appPlanOf_(planId);
      if (appYearIsFrozen_(plan.fy)) return 0;
      if (appReadPlanTable_('AI_RESEARCH_LOG', planId).some(r => r.action_id === 'BASELINE')) return 0;   // ほかの操作が先に写した
      const eng = appReadPlanTable_('ENG_AI_RESEARCH_STRUCTURED', planId);
      if (!eng.length) return 0;
      const now = appNowIso_();
      const ops = appLogOps_('AI_RESEARCH_LOG', eng.map(r => appAiResearchLogRow_(planId, r, {
        research_id: appStableLogId_(APP_LOG_PREFIX.AI_RESEARCH_LOG, ['BASELINE', planId, r.seq]), action_id: 'BASELINE', started_by: 'BASELINE', recorded_at: now })));
      if (!ops.length) return 0;
      const written = appJournalRun_(ctx, 'AI 調査の記録（版 10 の写し・' + planId + '）', planId, ops);
      const n = (written.AI_RESEARCH_LOG && written.AI_RESEARCH_LOG.appended) || 0;
      if (n) plans++;
      return n;
    });
  }
  return plans ? { rows: rows, plans: plans } : { rows: rows };
}
