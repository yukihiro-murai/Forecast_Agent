#!/usr/bin/env node
/*
 * app-pipeline.test.mjs — 裏の処理の壊れにくさ（2026-10-03 夜間の見直し）の契約テスト。
 * - 組み立ては 1 回の上限に収まらなければ何回かに分かれ、分かれても結果は同じ
 * - 組み立て → 計算 → 保存の間に計算用ブックがほかの処理で使われたら、保存しない
 * - 保存が途中で止まっても（Google 側のエラーなど）、保存の控えのとおりに書き直せる（続きの処理・次の操作のどちらでも）
 * - 数式として動いてしまう入力は受け付けない（計算用ブックは所有者の権限で動くため）
 *
 *   node app/tests/app-pipeline.test.mjs
 */

import assert from 'node:assert/strict';
import { J, OWNER, setUpEnv } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const SNAP_HEADER = ['snapshot_id', 'run_date', 'client', 'target_month', 'scenario', 'base_pred', 'subjective_adj', 'ai_adj', 'deterministic_adj', 'final_pred',
  'confidence_interval_lower', 'confidence_interval_upper', 'key_factors_json', 'subjective_input_date', 'calibration_applied_json'];
function legacyBook(env) {
  const output = [['FY2026 売上予測（テスト製薬）']];
  for (let r = 2; r <= 25; r++) output.push([]);
  output.push(['年度合計（予測）', 900, 1000, 1100]);
  output.push([], []);
  for (let i = 0; i < 12; i++) output.push([D(2026, 4 + i), 70, 80, 90]);
  return env.makeBook('クライアント別売上予測', {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', 'テスト製薬'], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
    OUTPUT: { values: output },
    FORECAST_SNAPSHOT: { values: [SNAP_HEADER, ['S1', D(2026, 9, 1), 'テスト製薬', '2026/04', 'neutral', 80, 0, 0, 0, 80, 70, 90, '{}', '', '{}']], formats: { D: '@' } },
    PROCESS_STATUS: { values: [['step_key', 'last_run_date', 'last_run_by', 'status', 'target_client', 'record_count', 'error_summary'],
      ['step4_status', D(2026, 9, 1), 'owner', 'success', 'テスト製薬', 12, '']] },
    PRODUCT: { values: [['Person', 'ProductName', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Reason'], ['鷹野', '製品A', D(2026, 5), 5, '新規']] },
    CLIENT: { values: [['Person', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Reason']] },
    OPINIONS: { values: [['Person', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Confidence(0..1)', 'Note']] },
    DEV_SPOT: { values: [['Person', 'Month(yyyy/mm/dd)', 'Project', 'Amount(JPY)', 'Confidence(0..1)']] },
  });
}
/** 旧来の予測の代わり（決まった形で OUTPUT・FORECAST_SNAPSHOT・PROCESS_STATUS を書く） */
const STUB = `appLegacyEngine_ = function (svc) {
  return {
    WEB_UI_CONFIRMS_: {}, VERSION: 'stub-1', SOURCE_SHA256: 'stub', WEB_SOURCE_SHA256: 'stub-web',
    webGetBootstrap_() { return { setup: { client: 'テスト製薬' } }; },
    runPhase1Forecast() {
      const ss = svc.SpreadsheetApp.getActiveSpreadsheet();
      const r = () => Math.round(Math.random() * 1000);
      const out = ss.getSheetByName('OUTPUT');
      out.getRange(26, 1, 1, 4).setValues([['年度合計（予測）', 100000 + r(), 200000 + r(), 300000 + r()]]);
      for (let i = 0; i < 12; i++) out.getRange(29 + i, 1, 1, 4).setValues([[new svc.Date(2026, 3 + i, 1), r(), 1000 + r(), 2000 + r()]]);
      const snap = ss.getSheetByName('FORECAST_SNAPSHOT');
      snap.getRange(snap.getLastRow() + 1, 1, 1, 15).setValues([[svc.Utilities.getUuid(), new svc.Date(), 'テスト製薬', '2026/05', 'neutral', 1, 0, 0, 0, 1, 0, 2, '{}', '', '{}']]);
      ss.getSheetByName('PROCESS_STATUS').getRange(2, 2, 1, 3).setValues([[new svc.Date(), 'owner', 'success']]);
    }
  };
};`;
function imported(stub = true) {
  const env = setUpEnv();
  const book = legacyBook(env);
  const url = 'https://docs.google.com/spreadsheets/d/' + book.getId() + '/edit';
  const dry = env.runJob('MIGRATION.DRYRUN', { bookUrl: url });
  assert.equal(dry.status, 'DONE', dry.error);
  const imp = env.runJob('MIGRATION.IMPORT', { bookUrl: url, contentHash: dry.result.contentHash });
  assert.equal(imp.status, 'DONE', imp.error);
  if (stub) env.run(STUB);
  return { env, planId: imp.result.planId };
}
const engRows = (env, sheet, planId) => env.table('ENG_' + sheet).filter((r) => r.plan_id === planId);
/** データ本体の計画を組み立て直して、計算用ブックと同じか（値・数式・表示形式・大きさ） */
function storeMatchesScratch(env, planId, names) {
  const scratch = env.scratch();
  const loaded = env.run('appEngLoadPlanSheets_(__p)', { __p: planId });
  for (const name of names) {
    const snap = env.run('appSheetSnapshot_(__s, __n)', { __s: scratch.getSheetByName(name), __n: Number(loaded[name].fmtCols) });
    assert.deepEqual(J(env.run('appEngCompare_(__a, __b)', { __a: snap, __b: loaded[name] })), [], name + ' がデータ本体と計算用ブックで同じ');
  }
}
const ends = (env, re) => env.audit().filter((a) => re.test(a.action) && a.phase === 'END').map((a) => [a.action, a.result]);

// ==== 1. 組み立ては何回かに分かれても、結果は同じ ====
{
  const { env, planId } = imported();
  const sheets = env.table('ENG_SHEETS').filter((r) => r.plan_id === planId).length;
  assert.ok(sheets >= 5, 'テストの計画には何枚かのシートがある');
  env.run('appBuildDeadline_ = function () { return 0; };');   // 1 回に 1 枚ずつ
  const st = env.runJob('FORECAST.RUN', { planId });
  assert.equal(st.status, 'DONE', st.error);
  const builds = ends(env, /^FORECAST\.RUN\.BUILD$/);
  assert.equal(builds.length, Object.keys(J(env.run('APP_ENGINE_SHEETS'))).length, '登録したシートの数だけ（無いシートも 1 回に 1 枚ずつ数える）');
  assert.ok(builds.every((b) => b[1] === 'OK'));
  assert.deepEqual(ends(env, /^FORECAST\.RUN\.(CALC|SAVE)$/), [['FORECAST.RUN.CALC', 'OK'], ['FORECAST.RUN.SAVE', 'OK']]);
  assert.deepEqual([...st.result.changed].sort(), ['FORECAST_SNAPSHOT', 'OUTPUT', 'PROCESS_STATUS']);
  assert.equal(env.table('FORECAST_RUNS').length, 1);
  assert.ok(st.result.timing.buildMs >= 0 && st.result.timing.saveMs >= 0, 'かかった時間を残す');
  storeMatchesScratch(env, planId, ['OUTPUT', 'FORECAST_SNAPSHOT', 'PROCESS_STATUS', 'CONFIG']);
  assert.equal(env.scratch().getSheets().filter((s) => /^_EMPTY_/.test(s.getName())).length, 0, '組み立て終えたら空のシートを消す');
}

// ==== 2. 組み立て・計算・保存の間に、計算用ブックがほかの処理で使われたら保存しない ====
{
  const { env, planId } = imported();
  const before = env.table('ENG_SHEETS').map((r) => r.content_hash);
  const started = env.call('apiStartJob(__in)', { __in: { kind: 'FORECAST.RUN', payload: { planId } } });
  env.fireTriggers('triggerRunJob');   // 組み立て
  const st = env.call('apiJobStatus(__in)', { __in: { jobId: started.jobId } });
  assert.equal(st.status, 'CONTINUED');
  assert.equal(st.nextKind, 'FORECAST.RUN_CALC');
  env.run('appScratchReset_(appScratchBook_())');   // ほかの処理が計算用ブックを空にした
  env.fireTriggers('triggerRunJob');
  const st2 = env.call('apiJobStatus(__in)', { __in: { jobId: st.nextJobId } });
  assert.equal(st2.status, 'FAILED');
  assert.match(st2.error, /計算用ブックがほかの処理で使われました/);
  assert.deepEqual(env.table('ENG_SHEETS').map((r) => r.content_hash), before, 'データ本体は変えない');
  assert.equal(env.table('FORECAST_RUNS').length, 0);
}

// ==== 3. 保存が途中で止まっても、控えのとおりに書き直せる（「保存の続き」の処理） ====
{
  const { env, planId } = imported();
  env.run(`(() => { const orig = appReplaceWhole_; let n = 0;
    appReplaceWhole_ = function (a, b) { if (a === 'ENG_FORMATS' && ++n === 1) throw new Error('Service Spreadsheets timed out'); return orig(a, b); };
    globalThis.__restore = () => { appReplaceWhole_ = orig; }; })()`);
  const st = env.runJob('FORECAST.RUN', { planId });
  assert.equal(st.status, 'FAILED');
  assert.match(st.error, /timed out/);
  const pending = J(env.run('appJournalPending_()'));
  assert.ok(pending && pending.planId === planId && /予測の保存/.test(pending.label), '控えが残る');
  // 書きかけ: ENG_ROWS は新しく、ENG_SHEETS は前のまま（表どうしが食い違っている）
  const v = env.call('apiPlanView(__in)', { __in: { planId } });
  assert.equal(v.boot, null, '食い違っている間は画面の中身を組み立てない');
  assert.ok(v.pendingWrite && /予測の保存/.test(v.pendingWrite.label));
  env.run('__restore()');
  const rec = env.runJob('SYSTEM.RECOVER', {});
  assert.equal(rec.status, 'DONE', rec.error);
  assert.equal(rec.result.planId, planId);
  assert.equal(J(env.run('appJournalPending_()')), null, '書き終えたら控えを消す');
  assert.equal(env.table('FORECAST_RUNS').length, 1, '予測の記録も控えのとおりに足す');
  assert.equal(env.table('FORECAST_MONTHLY').length, 12);
  storeMatchesScratch(env, planId, ['OUTPUT', 'FORECAST_SNAPSHOT', 'PROCESS_STATUS']);
  assert.ok(env.runLog().some((r) => r.kind === 'JOURNAL.RECOVER'), '書き直したことを実行ログに残す');
  const v2 = env.call('apiPlanView(__in)', { __in: { planId } });
  assert.ok(v2.boot && !v2.pendingWrite, '書き直した後は画面が出る');
  // もう一度書き直しても同じ（何もしない）
  const again = env.runJob('SYSTEM.RECOVER', {});
  assert.equal(again.status, 'DONE');
  assert.equal(again.result.nothing, true);
  assert.equal(env.table('FORECAST_RUNS').length, 1);
}

// ==== 4. 止まった後の次の操作でも、先に控えのとおりに書き終えてから進む ====
{
  const { env, planId } = imported();
  env.run(`(() => { const orig = appReplaceWhole_; let n = 0;
    appReplaceWhole_ = function (a, b) { if (a === 'ENG_SHEETS' && ++n === 1) throw new Error('Exception: Service error'); return orig(a, b); }; })()`);
  const st = env.runJob('FORECAST.RUN', { planId });
  assert.equal(st.status, 'FAILED');
  const st2 = env.runJob('FORECAST.RUN', { planId });
  assert.equal(st2.status, 'DONE', st2.error);
  assert.equal(env.table('FORECAST_RUNS').length, 2, '止まった回の記録も、次の回の記録も残る');
  assert.equal(engRows(env, 'FORECAST_SNAPSHOT', planId).length, 3, '履歴は 1 行ずつ足される（止まった回の分も）');
  assert.equal(J(env.run('appJournalPending_()')), null);
  storeMatchesScratch(env, planId, ['OUTPUT', 'FORECAST_SNAPSHOT', 'PROCESS_STATUS']);
  // 控えのファイルが読めないときは、はっきり止める
  env.run(`(() => { appProps_().setProperty(APP_JOURNAL_PROP, JSON.stringify({ id: 'JNL-x', fileId: 'missing', label: 'テスト', planId: '${planId}', at: '' })); })()`);
  const st3 = env.runJob('FORECAST.RUN', { planId });
  assert.equal(st3.status, 'FAILED');
  assert.match(st3.error, /控え（テスト）が読めません/);
}

// ==== 5. 数式として動いてしまう入力は受け付けない（数・増減率・ふつうの文は通す） ====
{
  const { env, planId } = imported(false);   // 旧来の webSaveInputs を本物のまま
  const view = env.call('apiPlanView(__in)', { __in: { planId } });
  const bad = env.runJob('PLAN.EDIT', { planId, action: 'INPUT.SAVE', inputHash: view.inputHash,
    args: { kind: 'product', rows: [{ person: '鷹野', product: '製品A', ym: '2026-06', step: '+5%', reason: '=IMPORTXML("https://example.com","//a")' }] } });
  assert.equal(bad.status, 'FAILED');
  assert.match(bad.error, /数式として動いてしまう文字/);
  const bad2 = env.runJob('PLAN.EDIT', { planId, action: 'INPUT.SAVE', inputHash: view.inputHash,
    args: { kind: 'product', rows: [{ person: '鷹野', product: '製品A', ym: '2026-06', step: '+5%', reason: '+SUM(A1:A9)' }] } });
  assert.equal(bad2.status, 'FAILED');
  assert.equal(engRows(env, 'PRODUCT', planId)[0].Reason, '新規', '止めたときは何も書かない');
  const ok = env.runJob('PLAN.EDIT', { planId, action: 'INPUT.SAVE', inputHash: view.inputHash,
    args: { kind: 'product', rows: [{ person: '鷹野', product: '製品A', ym: '2026-06', step: '-5%', reason: '- 前年の反動（=見込み）' }] } });
  assert.equal(ok.status, 'DONE', ok.error);
  assert.equal(engRows(env, 'PRODUCT', planId)[0].Reason, '- 前年の反動（=見込み）');
  assert.equal(env.scratch().getSheetByName('PRODUCT').fmls.flat().filter(Boolean).length, 0, '計算用ブックに数式を作らない');
}

console.log('app-pipeline: all tests passed');
