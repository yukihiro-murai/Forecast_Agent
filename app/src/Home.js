/**
 * Home.js — ホームのダッシュボード（最重要のことだけ）。よみのセリフは画面がこの材料から選ぶ。
 * 数字と計画の状態は、データ本体が変わらない間は控えから返す（appCachedRead_）。バックアップの状態は、バックアップしたときに残した要点を読む。
 * 空模様（霧の判定・前提の変化の日数）は今日の日付で変わるので、計画の要点とホームの控えは日ごとに分ける。
 */

/** 全部の計画の要点（計画の一覧と同じもの）。データ本体が変わらない間（その日のうち）は控えから */
function appPortfolioCached_() {
  return appCachedRead_('PORTFOLIO\u0001' + appToday_(), () => appPortfolio_());
}

/** その人のホーム（数字・計画の状態・着地見込みと空模様・承認待ち）＋ 管理者には仕組みの状態 */
function appHome_(ctx) {
  const out = appCachedRead_('HOME\u0001' + ctx.actor + '\u0001' + appToday_(), () => appHomeData_(ctx));
  if (appHasRole_(ctx.roles, 'ADMIN')) out.system = appHomeSystem_();
  out.at = appNowIso_();
  return out;
}

function appHomeData_(ctx) {
  const plans = appPortfolioCached_();
  // 見せる年度: 今の年度の計画があればそれ、無ければ一番新しい年度
  const thisFy = String(appFy_(new Date()));
  const fy = plans.some(p => String(p.fy) === thisFy) ? thisFy : plans.reduce((m, p) => (m === null || String(p.fy) > m ? String(p.fy) : m), null) || thisFy;
  const planRows = appReadTable_('PLANS');
  const clientOf = {};
  planRows.forEach(p => { clientOf[p.plan_id] = p.client_id; });
  const names = {};
  plans.forEach(p => { names[p.planId] = { clientName: p.clientName, fy: p.fy }; });
  // 承認待ち（承認できる人には「確認の依頼」、出した人には「待っている」）
  const approvals = [];
  const mine = [];
  appVersionTable_().filter(v => v.state === 'SUBMITTED' && names[v.plan_id]).forEach(v => {
    const o = { planId: v.plan_id, clientName: names[v.plan_id].clientName, fy: names[v.plan_id].fy, versionNo: v.version_no, budgetFinal: v.budget_final,
      submittedBy: String(v.submitted_by || '').split('@')[0], submittedAt: v.submitted_at };
    if (v.submitted_by === ctx.actor) mine.push(o);
    if ((v.submitted_by !== ctx.actor || ctx.user.isOwner) && appHasRole_(ctx.roles, 'APPROVER', clientOf[v.plan_id])) approvals.push(o);
  });
  return {
    fy: fy,
    plans: plans.map(p => ({ planId: p.planId, clientName: p.clientName, fy: String(p.fy), p10: p.p10, p50: p.p50, p90: p.p90, prevP50: p.prevP50, budget: p.budget, mape: p.mape,
      mapeMonths: p.mapeMonths, runs: p.runs, lastRunAt: p.lastRunAt, stepErrors: p.stepErrors, stepsDone: p.stepsDone, stepsTotal: p.stepsTotal,
      officialNo: p.officialNo, officialFinal: p.officialFinal, pendingNo: p.pendingNo,
      actualYtd: p.actualYtd, actualMonths: p.actualMonths, forecastYtd: p.forecastYtd, rangeOut: p.rangeOut, rangeN: p.rangeN,
      landing: p.landing, landingSd: p.landingSd, landingP10: p.landingP10, landingP90: p.landingP90, pAbove: p.pAbove, ratio: p.ratio,
      sky: p.sky, skyReason: p.skyReason, skyDir: p.skyDir, theta: p.theta, credibility: p.credibility, k: p.k, budgetUsed: p.budgetUsed, budgetSource: p.budgetSource })),
    approvals: approvals, mine: mine
  };
}

/** 仕組みの状態（管理者）。バックアップは、バックアップしたときに残した要点を読む（ドライブを数えない。Backup.js の appBackupStatusFast_） */
function appHomeSystem_() {
  let backup = null;
  try { backup = appBackupStatusFast_(); } catch (e) { backup = { error: String(e && e.message || e) }; }
  let journal = null;
  try { journal = appJournalPending_(); } catch (e) { journal = null; }
  let autoResearch = null;   // 自動の AI 調査（AutoResearch.js）
  try { autoResearch = appAutoResearchStatus_(); } catch (e) { autoResearch = { error: String(e && e.message || e) }; }
  return { backup: backup, housekeeping: appHousekeepingLast_(), journal: journal, autoResearch: autoResearch };
}
