/**
 * Parity.js — 計算の一致の確認（段階2-2）。所有者だけ。データ本体にも旧ブックにも書かない（計算用ブックだけを使う）。
 * A: 旧ブックの計算に使うシートをそのまま計算用ブックに写し、旧来の予測（A-9）を動かして、計算後の全シートの控えを取る。
 * B: データ本体の計画から計算用ブックを組み立て、同じ種・同じ「今」で予測を動かし、計算後の全シートを A と比べる。
 * 1 回の実行の上限（6 分）に収まるよう、A と B は別の処理（続きの処理）に分ける。
 */

/** 'yyyy-MM-ddTHH:mm:ss+0900' を日時に（コロンのない時差も読む） */
function appParseIso_(s) {
  const t = String(s || '').replace(/([+-]\d{2})(\d{2})$/, '$1:$2');
  const d = new Date(t);
  return isNaN(d.getTime()) ? null : d;
}

function appParityPlan_(planId) {
  const plan = appReadTable_('PLANS').filter(x => x.plan_id === String(planId || ''))[0];
  if (!plan) throw new Error('計画が見つかりません。先に旧ブックを取り込んでください。');
  const last = appLastImport_(plan.plan_id);
  if (!last) throw new Error('この計画はまだ取り込まれていません。');
  return { plan: plan, last: last };
}

function appParityScratch_(plan) {
  const scratch = appScratchBook_();
  try { scratch.setSpreadsheetTimeZone(plan.time_zone || APP_TZ); if (plan.locale) scratch.setSpreadsheetLocale(plan.locale); } catch (e) {
    Logger.log('計算用ブックの地域設定: ' + (e && e.message ? e.message : e));
  }
  return scratch;
}

/** A: 旧ブックの写しで予測を動かし、控えを取る。続きの処理（B）を返す */
function appParityA_(ctx, p, job) {
  const pl = appParityPlan_(p.planId);
  const t0 = new Date().getTime();
  const legacy = SpreadsheetApp.openById(pl.plan.source_book_id);
  const updated = DriveApp.getFileById(pl.plan.source_book_id).getLastUpdated();
  const imported = appParseIso_(pl.last.finished_at);
  const asOfMs = t0;
  const seed = 'PARITY:' + pl.plan.plan_id + ':' + asOfMs;
  return appWithLock_(() => {
    const scratch = appParityScratch_(pl.plan);
    const copied = appScratchCopyLegacy_(scratch, legacy);
    const previous = appForecastHeadline_(scratch);   // 旧ブックで前回に実行した結果（写した OUTPUT）
    const pre = appScratchDigest_(scratch);            // 計算の前の控え（組み立ての違いか、計算での違いかを分けるため）
    const t1 = new Date().getTime();
    const run = appRunLegacyForecast_(scratch, { asOfMs: asOfMs, seed: seed, confirms: ['extreme'] });
    if (!run.ok) throw new Error('旧来の計算が確認を求めて止まりました（' + JSON.stringify(run.needConfirm) + '）。');
    const t2 = new Date().getTime();
    const headline = appForecastHeadline_(scratch);
    appJobPutResult_(job.id + '_A', { pre: pre, digest: appScratchDigest_(scratch) });
    const t3 = new Date().getTime();
    return {
      __next: { kind: 'FORECAST.PARITY_B', payload: {
        planId: pl.plan.plan_id, asOfMs: asOfMs, seed: seed, parentJobId: job.id, copied: copied.length,
        legacyChangedAfterImport: !!(imported && updated && updated.getTime() > imported.getTime()),
        engine: { version: run.version, sourceSha256: run.sourceSha256 },
        previous: previous, legacy: headline, timingA: { copyMs: t1 - t0, runMs: t2 - t1, digestMs: t3 - t2 } } },
      audit: { entityId: pl.plan.plan_id, clientId: pl.plan.client_id }
    };
  });
}

/** B: データ本体から組み立てて予測を動かし、A の控えと比べる */
function appParityB_(ctx, p) {
  const pl = appParityPlan_(p.planId);
  const t0 = new Date().getTime();
  return appWithLock_(() => {
    appJournalRecover_(ctx);   // 保存が途中で止まっていれば、先に書き終える（食い違ったまま組み立てない）
    const scratch = appParityScratch_(pl.plan);
    const build = appScratchFromStore_(scratch, pl.plan.plan_id);
    const pre = appScratchDigest_(scratch);
    const t1 = new Date().getTime();
    const run = appRunLegacyForecast_(scratch, { asOfMs: p.asOfMs, seed: p.seed, confirms: ['extreme'] });
    if (!run.ok) throw new Error('旧来の計算が確認を求めて止まりました（' + JSON.stringify(run.needConfirm) + '）。');
    const t2 = new Date().getTime();
    const headline = appForecastHeadline_(scratch);
    const a = appJobGetResult_(p.parentJobId + '_A');
    if (!a.found) throw new Error('旧ブックの写しでの結果の控えが見つかりません（保存期間が過ぎた）。もう一度始めてください。');
    const diff = appDigestDiff_(a.value.digest, appScratchDigest_(scratch));
    const preDiff = a.value.pre ? appDigestDiff_(a.value.pre, pre) : null;
    const t3 = new Date().getTime();
    return {
      planId: pl.plan.plan_id, asOf: Utilities.formatDate(new Date(p.asOfMs), APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ"), seed: p.seed,
      same: diff.length === 0, diff: diff, preSame: preDiff ? preDiff.length === 0 : null, preDiff: preDiff,
      sheets: p.copied, engine: p.engine, legacyChangedAfterImport: p.legacyChangedAfterImport,
      legacy: p.legacy, store: headline, previous: p.previous,
      build: build.filter(x => x.mismatch || x.forcedText || x.formatMismatches || x.formatFixed || x.blankMethod),
      timing: { a: p.timingA, b: { buildMs: t1 - t0, runMs: t2 - t1, compareMs: t3 - t2 } },
      audit: { entityId: pl.plan.plan_id, clientId: pl.plan.client_id }
    };
  });
}

/** 計画の一覧（管理画面: 取り込み・一致の確認の対象） */
function appListPlans_() {
  const clients = {};
  appReadTable_('CLIENTS').forEach(c => { clients[c.client_id] = c.client_name; });
  return appReadTable_('PLANS').map(p => {
    const last = appLastImport_(p.plan_id);
    return { planId: p.plan_id, clientName: clients[p.client_id] || p.client_label, fy: p.fy, state: p.state,
      lastImportedAt: last ? last.finished_at : '', source: p.source_book_id ? 'book' : 'app' };
  });
}
