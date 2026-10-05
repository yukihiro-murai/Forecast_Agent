#!/usr/bin/env node
/**
 * app-home.test.mjs — ホーム（最重要のこと）・読んだ結果の控え（書いたら古くする）・短い保存をその場で動かす。
 *
 *   node app/tests/app-home.test.mjs
 */
import assert from 'node:assert/strict';
import { OWNER, MEMBER, setUpEnv, STATS } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const env = setUpEnv();
const output = [['FY2026 売上予測（テスト製薬）']];
for (let r = 2; r <= 25; r++) output.push([]);
output.push(['年度合計（予測）', 900, 1000, 1100]);
const book = env.makeBook('見本', {
  CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', 'テスト製薬'], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
  OUTPUT: { values: output },
  PROCESS_STATUS: { values: [['step_key', 'last_run_date', 'last_run_by', 'status', 'target_client', 'record_count', 'error_summary'],
    ['step1_status', D(2026, 9, 30), 'owner', 'error', 'テスト製薬', 0, '取り込めませんでした']] },
});
const planId = env.seedPlan(book);

// ==== 1. ホーム: 計画の状態・承認待ち・最近の動き。管理者には仕組みの状態。最初の画面にも入れる ====
{
  const h = env.call('apiHome()');
  assert.equal(h.fy, '2026');
  assert.equal(h.plans.length, 1);
  assert.deepEqual([h.plans[0].planId, h.plans[0].clientName, h.plans[0].p50], [planId, 'テスト製薬', 1000]);
  assert.deepEqual(h.plans[0].stepErrors, ['step1_status'], '手順のエラーを拾う（よみが「進み」を見るよう頼む）');
  assert.deepEqual([h.approvals, h.mine, h.recent], [[], [], []]);
  assert.ok(h.system && 'backup' in h.system && 'journal' in h.system, '管理者には仕組みの状態');
  assert.equal(env.call('apiBootstrap()').home.plans.length, 1, '最初の画面にホームの中身を入れる');
  env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'M' })`);
  env.as(MEMBER);
  const hm = env.call('apiHome()');
  assert.equal(hm.system, undefined, '管理者でなければ仕組みの状態は出さない');
  assert.equal(hm.plans.length, 1, '閲覧は社内全員');
  env.as(OWNER);
}

// ==== 2. 読んだ結果の控え: データ本体が変わらない間は表を読まない。書いたらすぐ読み直す ====
{
  env.call('apiListPlans()');
  env.call('apiPortfolio()');
  env.call('apiHome()');
  const r0 = STATS.reads;
  assert.equal(env.call('apiListPlans()').plans.length, 1);
  env.call('apiPortfolio()');
  env.call('apiHome()');
  assert.equal(STATS.reads - r0, 0, '控えから返す（表を読まない）');
  env.call(`apiSaveClient({ clientName: '別の製薬' })`);
  const r1 = STATS.reads;
  env.call('apiListPlans()');
  assert.ok(STATS.reads > r1, '書いた後は読み直す');
  // 版の印が消えても（控えが追い出されても）、古い結果は返さない
  env.run(`CacheService.getScriptCache().remove('APP_DATA_GEN')`);
  const r2 = STATS.reads;
  env.call('apiListPlans()');
  assert.ok(STATS.reads > r2, '印が無ければ読み直す');
}

// ==== 3. 短い保存は、トリガーを待たずに頼んだ通信の中で動かす（最近かかった時間から見て収まるとき） ====
{
  const v1 = env.call('apiPlanView(__in)', { __in: { planId } });
  const first = env.runJob('PLAN.EDIT', { planId, action: 'SETUP.PEOPLE', args: { peopleCsv: '鷹野, 佐藤' }, inputHash: v1.inputHash });
  assert.equal(first.status, 'DONE', first.error);
  const v2 = env.call('apiPlanView(__in)', { __in: { planId } });
  assert.notEqual(v2.inputHash, v1.inputHash, '保存の後は画面の中身も新しい（控えを使わない）');
  const started = env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.EDIT', payload: { planId, action: 'SETUP.PEOPLE', args: { peopleCsv: '鷹野' }, inputHash: v2.inputHash } } });
  assert.deepEqual([started.status, started.direct], ['DONE', true], '2 回目からはその場で動く');
  assert.equal(env.triggers.filter((t) => t.handler === 'triggerRunJob').length, 0, '使わなかったトリガーは消す');
  const st = env.call('apiJobStatus(__in)', { __in: { jobId: started.jobId } });
  assert.equal(st.status, 'DONE');
  assert.equal(env.table('PLANS')[0].people_csv, '鷹野');
  // 時間がかかる処理（予測の実行など）は、これまでどおりトリガーで動かす
  const run = env.call('apiStartJob(__in)', { __in: { kind: 'SYSTEM.RECOVER', payload: {} } });
  assert.equal(run.status, 'QUEUED');
  env.fireTriggers('triggerRunJob');
  assert.equal(env.state.lockHeld, false);
}

console.log('app-home: all tests passed');
