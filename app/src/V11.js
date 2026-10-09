/**
 * V11.js — 表の版 11 の共通の道具（SCHEMA_PLAN_v10-12_JA.md の 4 章。2026-10-09 村井さん承認「推奨する対応を継続してください」）。
 * - 足した表: 予算の下書き BUDGET_DRAFTS（追記だけ。書くのは BudgetDrafts.js。足すのは版 10 と同じ appLogOps_ → 控えの書き方 append）
 * - 足した列: 公式版 PLAN_VERSIONS の後ろの 9 列（出したときの届く見込みと前提。Versions.js が出すときに書き、後から変えない）。
 *   列を足す移行は版 10 と同じ仕組み（Setup.js の appMigrateColumns_。その日のバックアップを確かめ、続きから終わる）
 * - 一度だけの写し（APP_V11_BACKFILLS）: 表の版がそろった後の操作の初めに、版 10 の写しの後で動く（Setup.js の appRunBackfills_）。
 *   今のデータ本体の予算（OUTPUT の採用予測と上乗せ）に手で直した跡がある計画だけ、BASELINE の下書きとして 1 回写す（予測し直しても戻るように）。
 * 新しい自動の処理とメールは足さない（今ある操作の中で書く）。
 */

/**
 * 表の版がそろうまで足さない記録の表（V10.js の appLogReady_(table)）。版 11 の予算の下書きは、移行の後の一度だけの写し（BASELINE）と
 * 予測し直したときの戻しが、版 11 の決まりで動くため。ほかの記録の表（版 10）は、表があって見出しが今の列なら、移行を待っている間も足す
 */
const APP_LOG_WAIT_TABLES = ['BUDGET_DRAFTS'];

/** 版 11 の一度だけの写し（名前だけ。関数は動かすときに探す） */
const APP_V11_BACKFILLS = ['appV11BackfillBudgetDrafts_'];
/** 一度だけの写し（と、写す前に予測し直したときの同じ写し）の行の「した人」。人ではない（画面には人として出さない: appLogSystemActor_） */
const APP_V11_BACKFILL_ACTOR = 'SYSTEM:V11_BACKFILL';
/** 1 回の写しで使ってよい時間の目安（超えたら続きは次の操作で） */
const APP_V11_BACKFILL_MS = 30 * 1000;

/**
 * 予算の下書きの BASELINE（4-1 の移行）: 版 11 の移行の後に 1 回。計画ごとにロックの中で、書きかけの保存を先に書き終えてから読む。
 * 写すのは、データ本体の OUTPUT の採用予測が月の真ん中と違う・上乗せが空でない月が 1 つでもある計画だけ（12 か月分を写す）。
 * 下書きがもうある計画（写しの前に予算を保存した・予測し直して同じ写しを足した）・締めた年度の計画・測る専用の計画は飛ばす。
 * 返り値は Setup.js の appRunBackfills_ の決まり（{ rows }・{ more, rows }）
 */
function appV11BackfillBudgetDrafts_(ctx) {
  if (!appLogReady_()) return { more: true, rows: 0 };
  const t0 = new Date().getTime();
  let drafts;
  try { drafts = appReadTable_('BUDGET_DRAFTS'); } catch (e) { return { more: true, rows: 0 }; }   // 表がまだ無い（移行の前）: 次の操作でまた
  const had = {};
  drafts.forEach(r => { had[r.plan_id] = true; });
  const plans = appReadTable_('PLANS').filter(p => !had[p.plan_id] && !appPlanIsMeasure_(p) && !appYearIsFrozen_(p.fy));
  let rows = 0;
  for (let i = 0; i < plans.length; i++) {
    if (new Date().getTime() - t0 > APP_V11_BACKFILL_MS) return { more: true, rows: rows };
    const planId = plans[i].plan_id;
    rows += appWithLock_(() => {
      appJournalRecover_(ctx);   // 書きかけの保存を先に書き終える（その保存の予算が入ってから読む）
      const plan = appPlanOf_(planId);
      if (appYearIsFrozen_(plan.fy) || appPlanIsMeasure_(plan)) return 0;
      if (appReadPlanTable_('BUDGET_DRAFTS', plan.plan_id).length) return 0;   // ほかの操作が先に下書きを足した
      const list = appBudgetBaselineRows_(plan, APP_V11_BACKFILL_ACTOR);
      if (!list.length) return 0;
      const ops = appLogOps_('BUDGET_DRAFTS', list);
      if (!ops.length) return 0;
      appJournalRun_(ctx, '予算の下書きの初めの姿（' + plan.plan_id + '）', plan.plan_id, ops);
      return list.length;
    });
  }
  return { rows: rows };
}
