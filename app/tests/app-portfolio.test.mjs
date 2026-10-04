/**
 * app-portfolio.test.mjs — 新しい計画を新アプリで作る（本物の旧来の A-1 初期セットアップで）・全部の計画の一覧。
 */
import assert from 'node:assert/strict';
import { OWNER, setUpEnv } from './gas-mock.mjs';

const env = setUpEnv();

// ==== 1. 計画を作る: 旧来の A-1 と同じシートができ、データ本体に計画として残る ====
const st = env.runJob('PLAN.CREATE', { clientName: 'テスト製薬', fy: 2027, peopleCsv: '鷹野, 佐藤' });
assert.equal(st.status, 'DONE', st.error);
const planId = st.result.planId;
const plan = env.table('PLANS').filter((p) => p.plan_id === planId)[0];
assert.equal(plan.fy, '2027');
assert.equal(plan.source_book_id, '', '旧ブックは無い');
assert.equal(plan.people_csv, '鷹野,佐藤');
assert.equal(env.table('CLIENTS').filter((c) => c.client_name === 'テスト製薬').length, 1);
const sheets = env.table('ENG_SHEETS').filter((r) => r.plan_id === planId).map((r) => r.sheet);
for (const name of ['CONFIG', 'SALES_INPUT', 'PRODUCT', 'OUTPUT', 'PROCESS_STATUS', 'EVAL_LOG', 'CALIBRATION_STATE']) assert.ok(sheets.includes(name), name + ' ができる');

// 画面の中身（旧来の webGetBootstrap_）がそのまま組み立てられる
const view = env.call('apiPlanView(__in)', { __in: { planId } });
assert.equal(view.boot.setup.client, 'テスト製薬');
assert.equal(Number(view.boot.setup.fy), 2027);
assert.deepEqual(view.boot.setup.people, ['鷹野', '佐藤']);

// ==== 2. 同じクライアント × 年度は 2 つ作らない。数式や空の担当者は受け付けない ====
const dup = env.runJob('PLAN.CREATE', { clientName: 'テスト製薬', fy: 2027, peopleCsv: '鷹野' });
assert.equal(dup.status, 'FAILED');
assert.match(dup.error, /すでにあります/);
assert.match(env.runJob('PLAN.CREATE', { clientName: '=IMPORTXML("x","y")', fy: 2027, peopleCsv: '鷹野' }).error, /数式として動いてしまう文字/);
assert.match(env.runJob('PLAN.CREATE', { clientName: '別製薬', fy: 2027, peopleCsv: ' , ' }).error, /担当者を 1 人以上/);
assert.match(env.runJob('PLAN.CREATE', { clientName: '別製薬', fy: 27, peopleCsv: '鷹野' }).error, /4 桁/);

// 次の年度は作れる（同じクライアントを使う）
const next = env.runJob('PLAN.CREATE', { clientName: 'テスト製薬', fy: 2028, peopleCsv: '鷹野' });
assert.equal(next.status, 'DONE', next.error);
assert.equal(env.table('CLIENTS').filter((c) => c.client_name === 'テスト製薬').length, 1, 'クライアントは 1 つのまま');

// ==== 3. 一覧: 作った計画が並ぶ（新しい年度が上） ====
const pf = env.call('apiPortfolio()');
assert.deepEqual(pf.plans.map((p) => p.fy), ['2028', '2027']);
assert.equal(pf.plans[1].source, 'app');
assert.equal(pf.plans[1].p50, null, 'まだ予測していない');
assert.ok(pf.plans[1].stepsTotal > 0, '手順の進みの行がある');

// ==== 4. 管理者でない社内の人は、作れない（一覧は見られる） ====
env.as('member@bigm2y.com');
assert.throws(() => env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.CREATE', payload: { clientName: '別製薬', fy: 2027, peopleCsv: '鷹野' } } }), /権限/);
assert.throws(() => env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.CREATE_SAVE', payload: { clientName: '別製薬', fy: 2027, peopleCsv: '鷹野' } } }), /画面から始められません|権限/);
assert.equal(env.call('apiPortfolio()').plans.length, 2);
env.as(OWNER);
console.log('app-portfolio: all tests passed');
