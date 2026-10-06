/**
 * Home.js — ホームのダッシュボード（最重要のことだけ）。よみのセリフは画面がこの材料から選ぶ。
 * 数字と計画の状態は、データ本体が変わらない間は控えから返す（appCachedRead_）。バックアップの状態は、バックアップしたときに残した要点を読む。
 */
const APP_HOME_RECENT = 6;

/** 全部の計画の要点（計画の一覧と同じもの）。データ本体が変わらない間は控えから */
function appPortfolioCached_() {
  return appCachedRead_('PORTFOLIO', () => appPortfolio_());
}

/** その人のホーム（数字・計画の状態・承認待ち・最近の動き）＋ 管理者には仕組みの状態 */
function appHome_(ctx) {
  const out = appCachedRead_('HOME\u0001' + ctx.actor, () => appHomeData_(ctx));
  if (appHasRole_(ctx.roles, 'ADMIN')) out.system = appHomeSystem_();
  out.at = appNowIso_();
  return out;
}

function appHomeData_(ctx) {
  const plans = appPortfolioCached_();
  const thisFy = String(appFy_(new Date()));
  const fys = plans.map(p => String(p.fy)).filter((x, i, a) => a.indexOf(x) === i).sort();
  const fy = fys.indexOf(thisFy) >= 0 ? thisFy : (fys.length ? fys[fys.length - 1] : thisFy);
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
  // 最近の動き（予測の実行と、計画への保存・実行）
  const recent = [];
  appReadTable_('FORECAST_RUNS').filter(r => r.status === 'DONE' && names[r.plan_id])
    .forEach(r => recent.push({ at: r.finished_at, label: '予測の実行', planId: r.plan_id, actor: String(r.actor_email || '').split('@')[0] }));
  (() => { try { return appReadTable_('PLAN_ACTIONS'); } catch (e) { return []; } })().filter(r => r.status === 'OK' && names[r.plan_id])
    .forEach(r => {
      let label = r.action;
      try { label = appPlanAction_(r.action).label; } catch (e) { /* 前の名前の操作はそのまま */ }
      recent.push({ at: r.finished_at, label: label, planId: r.plan_id, actor: String(r.actor_email || '').split('@')[0] });
    });
  recent.sort((a, b) => String(b.at).localeCompare(String(a.at)));
  return {
    fy: fy, fys: fys,
    plans: plans.map(p => ({ planId: p.planId, clientName: p.clientName, fy: String(p.fy), p50: p.p50, prevP50: p.prevP50, budget: p.budget, mape: p.mape,
      mapeMonths: p.mapeMonths, runs: p.runs, lastRunAt: p.lastRunAt, stepErrors: p.stepErrors, stepsDone: p.stepsDone, stepsTotal: p.stepsTotal,
      officialNo: p.officialNo, officialFinal: p.officialFinal, pendingNo: p.pendingNo,
      actualYtd: p.actualYtd, actualMonths: p.actualMonths, forecastYtd: p.forecastYtd, landing: p.landing, rangeOut: p.rangeOut, rangeN: p.rangeN })),
    approvals: approvals, mine: mine,
    recent: recent.slice(0, APP_HOME_RECENT).map(x => Object.assign(x, { clientName: names[x.planId].clientName, fy: names[x.planId].fy }))
  };
}

/** 仕組みの状態（管理者）。バックアップは、バックアップしたときに残した要点を読む（ドライブを数えない。Backup.js の appBackupStatusFast_） */
function appHomeSystem_() {
  let backup = null;
  try { backup = appBackupStatusFast_(); } catch (e) { backup = { error: String(e && e.message || e) }; }
  let journal = null;
  try { journal = appJournalPending_(); } catch (e) { journal = null; }
  return { backup: backup, housekeeping: appHousekeepingLast_(), journal: journal };
}
