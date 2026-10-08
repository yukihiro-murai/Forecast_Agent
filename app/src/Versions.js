/**
 * Versions.js — 公式版（確定した予測と予算）と承認。
 *
 * 予算策定担当が、今の計画の予測と予算を「公式版」として出す（承認待ち）。承認者が承認すると公式版になり、前の公式版は置き換わる。
 * 出した版の数字（年度と月ごとの P10/P50/P90・採用予測・上乗せ・最終予算）は、出した時点のまま変えない。
 * 計画はその後も直せる（次の版を出す）。版どうしの差は画面で比べる。
 */
const APP_VERSION_STATES = ['SUBMITTED', 'APPROVED', 'REJECTED', 'SUPERSEDED', 'WITHDRAWN'];

/** データ本体の OUTPUT から、今の予測と予算の数字（年度と月ごと）。読むのはその計画の行だけ */
function appPlanNumbers_(planId) {
  const rows = {};
  appReadPlanTable_('ENG_ROWS', planId).filter(r => r.sheet === 'OUTPUT' && Number(r.col_from) === 1)
    .forEach(s => { rows[Number(s.row_no)] = JSON.parse(s.cells_json); });
  const cell = (r, c) => { const x = (rows[r] || [])[c - 1]; return x ? (x.charAt(0) === 'f' ? '' : appCellDecode_(x.charAt(0), x.slice(1))) : ''; };
  const num = v => (typeof v === 'number' && isFinite(v) ? v : null);
  if (!rows[26]) return null;
  const monthly = [];
  for (let r = 29; r <= 40; r++) {
    const v = cell(r, 1);
    const adopted = num(cell(r, 8)), uplift = num(cell(r, 9));
    monthly.push({ month: appIsDate_(v) ? Utilities.formatDate(v, APP_TZ, 'yyyy/MM') : String(v), p10: num(cell(r, 2)), p50: num(cell(r, 3)), p90: num(cell(r, 4)),
      adopted: adopted, uplift: uplift, final: adopted === null && uplift === null ? null : (adopted || 0) + (uplift || 0) });
  }
  const sum = k => monthly.reduce((a, m) => (m[k] === null ? a : (a || 0) + m[k]), null);
  return { annual: { p10: num(cell(26, 2)), p50: num(cell(26, 3)), p90: num(cell(26, 4)) },
    budget: { adopted: sum('adopted'), uplift: sum('uplift'), final: sum('final') }, monthly: monthly };
}

/** 公式版の表（版を上げて足した表。まだ作られていなければ空） */
function appVersionTable_() {
  try { return appReadTable_('PLAN_VERSIONS'); } catch (e) { if (/表がありません/.test(String(e && e.message))) return []; throw e; }
}

function appVersionRows_(planId) {
  return appVersionTable_().filter(v => v.plan_id === planId).map(appStripRow_).sort((a, b) => b.version_no - a.version_no);
}

function appVersionOut_(v) {
  return { versionId: v.version_id, no: v.version_no, state: v.state, inputHash: v.input_hash, runId: v.forecast_run_id,
    annual: { p10: v.annual_p10, p50: v.annual_p50, p90: v.annual_p90 }, budget: { adopted: v.budget_adopted, uplift: v.budget_uplift, final: v.budget_final },
    monthly: appParseJsonList_(v.monthly_json), note: v.note, submittedAt: v.submitted_at, submittedBy: v.submitted_by,
    decidedAt: v.decided_at, decidedBy: v.decided_by, decisionNote: v.decision_note, rowVersion: v.row_version };
}

/** 画面: 計画の版の一覧と、今の数字（出す前に見比べる）。測る専用の計画（measure）は出せない（can.submit が false） */
function appVersionList_(ctx, input) {
  const plan = appPlanOf_(input && input.planId);
  const versions = appVersionRows_(plan.plan_id).map(appVersionOut_);
  const official = versions.filter(v => v.state === 'APPROVED')[0] || null;
  const pending = versions.filter(v => v.state === 'SUBMITTED')[0] || null;
  const inputHash = appPlanInputHash_(plan.plan_id);
  const frozen = appYearIsFrozen_(plan.fy);
  const measure = appPlanIsMeasure_(plan);
  return { planId: plan.plan_id, frozen: frozen, measure: measure, current: appPlanNumbers_(plan.plan_id), inputHash: inputHash, versions: versions, official: official, pending: pending,
    pendingChanged: !!pending && pending.inputHash !== inputHash,
    can: { submit: appHasRole_(ctx.roles, 'PLANNER', plan.client_id) && !frozen && !measure, approve: appHasRole_(ctx.roles, 'APPROVER', plan.client_id) && !frozen }, me: ctx.actor };
}

/** 今の予測と予算を、公式版として出す（承認待ち）。前の承認待ちは取り下げる。測る専用の計画は出さない（PlanPurpose.js） */
function appVersionSubmit_(ctx, input) {
  const note = String(input && input.note || '').trim().slice(0, 500);
  appPlanCheckArgs_([note]);
  return appWithLock_(() => {
    if (appJournalPending_()) throw new Error('データ本体の保存が途中で止まっています。予測の画面の「保存の続きを書く」を先に行ってください。');
    const plan = appRequireOpenPlan_(input && input.planId);   // 締めた年度には版を出せない（所有者でも同じ）
    appMeasureRequireBudgetPlan_(plan);
    const nums = appPlanNumbers_(plan.plan_id);
    if (!nums || nums.annual.p50 === null) throw new Error('予測がまだありません。予測を実行してから出してください。');
    const inputHash = appPlanInputHash_(plan.plan_id);
    if (input && input.inputHash && input.inputHash !== inputHash) throw new Error('画面を開いた後に計画が変わりました。読み直してから出してください。');
    const versions = appVersionRows_(plan.plan_id);
    const now = appNowIso_();
    const withdrawn = versions.filter(v => v.state === 'SUBMITTED');
    withdrawn.forEach(v => appUpdateByKey_('PLAN_VERSIONS', { version_id: v.version_id }, { state: 'WITHDRAWN', decided_at: now, decided_by: ctx.actor,
      decision_note: '新しい版を出したため取り下げ' }, v.row_version, ctx.actor));
    const run = appReadTable_('FORECAST_RUNS').filter(r => r.plan_id === plan.plan_id && r.status === 'DONE')
      .sort((a, b) => String(b.finished_at).localeCompare(String(a.finished_at)) || b._row - a._row)[0];
    const row = { version_id: appId_('VER'), plan_id: plan.plan_id, version_no: (versions.length ? versions[0].version_no : 0) + 1, state: 'SUBMITTED',
      input_hash: inputHash, forecast_run_id: run ? run.run_id : '', annual_p10: nums.annual.p10, annual_p50: nums.annual.p50, annual_p90: nums.annual.p90,
      budget_adopted: nums.budget.adopted, budget_uplift: nums.budget.uplift, budget_final: nums.budget.final, monthly_json: nums.monthly, note: note,
      submitted_at: now, submitted_by: ctx.actor, decided_at: '', decided_by: '', decision_note: '', updated_at: now, updated_by: ctx.actor, row_version: 1 };
    appInsertRows_('PLAN_VERSIONS', [row]);
    return { version: appVersionOut_(row), withdrawn: withdrawn.map(v => v.version_no), audit: { entityId: row.version_id, clientId: plan.client_id } };
  });
}

/**
 * 承認待ちの版を、承認・却下する。出した本人は承認できない（所有者だけは、ほかに承認者がいない運用のため承認できる。記録に残る）
 * 承認すると、前の公式版は SUPERSEDED になる
 */
function appVersionDecide_(ctx, input) {
  const decision = String(input && input.decision || '');
  if (['APPROVED', 'REJECTED'].indexOf(decision) < 0) throw new Error('承認か却下を選んでください。');
  const note = String(input && input.note || '').trim().slice(0, 500);
  if (decision === 'REJECTED' && !note) throw new Error('却下の理由を書いてください。');
  appPlanCheckArgs_([note]);
  return appWithLock_(() => {
    const v = appReadTable_('PLAN_VERSIONS').filter(x => x.version_id === String(input && input.versionId || ''))[0];
    if (!v) throw new Error('版が見つかりません。');
    if (v.state !== 'SUBMITTED') throw new Error('この版は承認待ちではありません（' + v.state + '）。画面を読み直してください。');
    if (v.submitted_by === ctx.actor && !ctx.user.isOwner) throw new Error('自分で出した版は承認・却下できません。ほかの承認者に頼んでください。');
    // 前の公式版を置き換える前に確かめる（置き換えた後に止まると、公式版が無くなる）
    if (input && input.rowVersion !== undefined && input.rowVersion !== null && input.rowVersion !== '' && Number(input.rowVersion) !== Number(v.row_version)) {
      throw new Error('ほかの人が先に更新しました。画面を読み直してから、もう一度選んでください。');
    }
    const plan = appRequireOpenPlan_(v.plan_id);   // 締めた年度の版は承認・却下できない
    const now = appNowIso_();
    const superseded = [];
    if (decision === 'APPROVED') {
      appReadTable_('PLAN_VERSIONS').filter(x => x.plan_id === v.plan_id && x.state === 'APPROVED').forEach(x => {
        appUpdateByKey_('PLAN_VERSIONS', { version_id: x.version_id }, { state: 'SUPERSEDED' }, x.row_version, ctx.actor);
        superseded.push(x.version_no);
      });
    }
    const res = appUpdateByKey_('PLAN_VERSIONS', { version_id: v.version_id }, { state: decision, decided_at: now, decided_by: ctx.actor, decision_note: note },
      input && input.rowVersion, ctx.actor);
    return { version: appVersionOut_(res.after), superseded: superseded, selfApproved: v.submitted_by === ctx.actor,
      before: { state: v.state }, audit: { entityId: v.version_id, clientId: plan.client_id } };
  });
}

/** 版の計画のクライアント（承認の権限を、その計画のクライアントの範囲で確かめるため） */
function appVersionClient_(input) {
  const v = appReadTable_('PLAN_VERSIONS').filter(x => x.version_id === String(input && input.versionId || ''))[0];
  return v ? appJobPlanClient_({ planId: v.plan_id }) : undefined;
}

/** 計画ごとの公式版と承認待ち（一覧に出す） */
function appVersionSummary_() {
  const out = {};
  appVersionTable_().forEach(v => {
    const o = out[v.plan_id] = out[v.plan_id] || { officialNo: null, officialFinal: null, pendingNo: null };
    if (v.state === 'APPROVED') { o.officialNo = v.version_no; o.officialFinal = v.budget_final; }
    if (v.state === 'SUBMITTED') o.pendingNo = v.version_no;
  });
  return out;
}
