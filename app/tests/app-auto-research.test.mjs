#!/usr/bin/env node
/**
 * app-auto-research.test.mjs — 市場の動き（A-4 AI 調査）の週 1 回の自動収集（AutoResearch.js）。
 * トリガーの作成（1 日 1 回だけ確かめる・二重にしない・上限・止めるスイッチ）、入口の本人確認、
 * いちばん古い計画を選ぶ・ほかの処理が動いていれば何もしない・5 時を過ぎたら続けない・締めた年度には書かない。
 * Vertex AI の答えはテストの決まった形（app-ai.test.mjs と同じ）。
 *
 *   node app/tests/app-auto-research.test.mjs
 */
import assert from 'node:assert/strict';
import { OWNER, MEMBER, setUpEnv, jstDay, J } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const DAY = 864e5;
/** 今日から offset 日の、日本時間の hh:mm（ミリ秒） */
const jst = (offset, hh, mm = 0) => new Date(`${jstDay(offset)}T${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}:00+09:00`).getTime();

function stubVertex(env, code = 200) {
  env.run(`(() => {
    ScriptApp.getOAuthToken = () => 'test-token';
    const gemini = (text) => JSON.stringify({ candidates: [{ content: { parts: [{ text }] }, groundingMetadata: { groundingChunks: [{ web: { uri: 'https://example.com/a', title: 'a' } }] } }],
      usageMetadata: { promptTokenCount: 10, candidatesTokenCount: 20, totalTokenCount: 30 } });
    globalThis.UrlFetchApp = { fetch: (u, o) => {
      if (${code} !== 200) return { getResponseCode: () => ${code}, getContentText: () => JSON.stringify({ error: { code: ${code}, status: 'PERMISSION_DENIED', message: 'denied' } }) };
      const body = JSON.parse(o.payload);
      const structured = !!(body.generationConfig && body.generationConfig.responseMimeType);
      return { getResponseCode: () => 200, getContentText: () => gemini(structured ? JSON.stringify({ rows: [] }) : '市場は横ばい。') };
    } };
    appAiDeadline_ = () => Date.now() + 60000;
  })()`);
}

/** A-4 を動かせる計画の見本（ai = 0 なら AI 調査を無効にした計画） */
function seed(env, client, fy, ai = 1) {
  return env.seedPlan(env.makeBook(client, {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', client], ['[必須] 予測年度FY（YYYY）', fy], ['[必須] 担当者（カンマ区切り）', '鷹野'],
      ['VERTEX_PROJECT_ID', 'test-project'], ['VERTEX_LOCATION', 'asia-northeast1'], ['VERTEX_GEMINI_MODEL', 'gemini-test'], ['AI_RESEARCH_ENABLED（0/1）', ai]] },
    PROCESS_STATUS: { values: [['step_key', 'last_run_date', 'last_run_by', 'status', 'target_client', 'record_count', 'error_summary'],
      ['step1_status', D(2026, 9, 1), 'owner', 'success', client, 10, '']] },
    PRODUCT: { values: [['Person', 'ProductName', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Reason']] },
    CLIENT: { values: [['Person', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Reason']] },
    OPINIONS: { values: [['Person', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Confidence(0..1)', 'Note']] },
    DEV_SPOT: { values: [['Person', 'Month(yyyy/mm/dd)', 'Project', 'Amount(JPY)', 'Confidence(0..1)']] },
  }));
}

let actSeq = 0;
/** 過去に A-4 が終わった記録を足す（finished_at = ms） */
function researchedAt(env, planId, ms) {
  env.run(`appWithLock_(() => appInsertRows_('PLAN_ACTIONS', [{ action_id: __id, plan_id: __p, action: 'AI.RESEARCH', status: 'DONE',
    started_at: appAutoResearchIso_(new Date(__ms)), finished_at: appAutoResearchIso_(new Date(__ms)), actor_email: 'owner@bigm2y.com' }]))`,
  { __id: 'ACT-TEST-' + (++actSeq), __p: planId, __ms: ms });
}

const setNow = (env, ms) => env.run(`appAutoResearchNow_ = () => new Date(${ms})`);
const autoTriggers = (env) => env.triggers.filter((t) => t.handler === 'triggerAutoResearch');
const state = (env) => J(env.run('appAutoResearchState_()'));
const jobs = (env) => J(env.run('appJobList_()'));
const ensure = (env) => J(env.run(`appEnsureAutoResearchTrigger_({ requestId: 'TEST', actor: '${OWNER}' })`));
/** 毎日のトリガー（triggerAutoResearch）を動かす（トリガーの中では操作者のメールが空） */
function tick(env) {
  const uid = autoTriggers(env)[0].uid;
  const prev = env.state.active;
  env.as('');
  try { return env.call(`triggerAutoResearch({ triggerUid: '${uid}' })`); } finally { env.as(prev); }
}
/** 裏の処理を、until() が真になるか、待っている処理が無くなるまで動かす */
function drain(env, until) {
  for (let i = 0; i < 300; i++) {
    if (until && until()) return;
    if (!env.triggers.some((t) => t.handler === 'triggerRunJob')) return;
    env.fireTriggers('triggerRunJob');
  }
  throw new Error('裏の処理が終わらない');
}
/** テストの中で動かした A-4 の数（足した過去の記録は数えない） */
const researched = (env, planId) => env.table('PLAN_ACTIONS').filter((r) => r.plan_id === planId && r.action === 'AI.RESEARCH' && r.status === 'DONE' && !r.action_id.startsWith('ACT-TEST-')).length;
const autoRuns = (env) => env.runLog().filter((r) => r.kind === 'AUTO.RESEARCH').map((r) => Object.assign(JSON.parse(r.detail_json || '{}'), { status: r.status }));

// ==== 1. トリガー: 所有者が開いたとき・毎日の手入れで、無ければ作る（1 日 1 回だけ確かめる・二重にしない・上限・止めるスイッチ） ====
{
  const env = setUpEnv();
  const audits = env.audit().length;
  env.run('doGet({})');
  assert.equal(autoTriggers(env).length, 1, '所有者が開くとトリガーを作る');
  assert.deepEqual([autoTriggers(env)[0].everyDays, autoTriggers(env)[0].atHour, autoTriggers(env)[0].nearMinute], [1, 1, 0], '毎日 1 時ごろ');
  assert.equal(env.audit().length, audits, '画面を開くだけでは監査の記録を増やさない');
  assert.deepEqual(env.runLog().filter((r) => r.kind === 'AUTO.RESEARCH.TRIGGER').map((r) => r.status), ['CREATED'], '作ったことは実行ログに残す');
  // 同じ日は確かめ直さない（開くたびに重くしない）
  env.triggers.splice(env.triggers.indexOf(autoTriggers(env)[0]), 1);
  env.run('doGet({})');
  assert.equal(autoTriggers(env).length, 0, '同じ日のうちは確かめ直さない');
  assert.equal(ensure(env).reason, 'CHECKED_TODAY');
  // 次の日（控えの日付が違う）: 所有者以外が開いても作らない。所有者が開けば作る
  env.props.APP_AUTO_RESEARCH_CHECKED = JSON.stringify({ date: jstDay(-1), installed: true });
  env.as(MEMBER);
  env.run('doGet({})');
  assert.equal(autoTriggers(env).length, 0, '所有者でなければ確かめない');
  env.as(OWNER);
  env.run('doGet({})');
  assert.equal(autoTriggers(env).length, 1);
  // 二重にできていたら 1 つにする
  env.run(`ScriptApp.newTrigger('triggerAutoResearch').timeBased().everyDays(1).atHour(1).create()`);
  delete env.props.APP_AUTO_RESEARCH_CHECKED;
  assert.deepEqual([ensure(env).installed, autoTriggers(env).length], [true, 1], '二重のトリガーは 1 つにする');
  // 止めるスイッチ: 作らない
  env.triggers.splice(env.triggers.indexOf(autoTriggers(env)[0]), 1);
  delete env.props.APP_AUTO_RESEARCH_CHECKED;
  env.props.AUTO_RESEARCH_ENABLED = 'false';
  assert.deepEqual([ensure(env).reason, autoTriggers(env).length], ['DISABLED', 0]);
  delete env.props.AUTO_RESEARCH_ENABLED;
  // トリガーの上限（20）: 裏の処理の分（3）を空けて、足りなければ作らない
  for (let i = 0; i < 17; i++) env.run(`ScriptApp.newTrigger('triggerDummy').timeBased().everyDays(1).atHour(5).create()`);
  delete env.props.APP_AUTO_RESEARCH_CHECKED;
  assert.deepEqual([ensure(env).reason, autoTriggers(env).length], ['TRIGGER_LIMIT', 0], '上限に余裕が無ければ作らない');
  assert.ok(env.runLog().some((r) => r.kind === 'AUTO.RESEARCH.TRIGGER' && r.status === 'SKIPPED'));
  env.triggers.splice(env.triggers.findIndex((t) => t.handler === 'triggerDummy'), 1);
  delete env.props.APP_AUTO_RESEARCH_CHECKED;
  assert.deepEqual([ensure(env).created, autoTriggers(env).length], [true, 1], '16 個なら作れる（20 - 3 の余裕）');
  for (let i = env.triggers.length - 1; i >= 0; i--) if (env.triggers[i].handler === 'triggerDummy') env.triggers.splice(i, 1);
  // 毎日の手入れ（triggerDailyBackup）でも作る
  env.triggers.splice(env.triggers.indexOf(autoTriggers(env)[0]), 1);
  delete env.props.APP_AUTO_RESEARCH_CHECKED;
  env.call('apiEnableBackup()');
  const uid = env.triggers.find((t) => t.handler === 'triggerDailyBackup').uid;
  env.as('');
  const daily = env.call(`triggerDailyBackup({ triggerUid: '${uid}' })`);
  env.as(OWNER);
  assert.equal(daily.autoResearch.created, true, '毎日の手入れでも作る');
  assert.equal(autoTriggers(env).length, 1);
  // 入口の本人確認（triggerDailyBackup と同じ）: このプロジェクトのトリガーの UID でなければ所有者として扱わない
  const auid = autoTriggers(env)[0].uid;
  env.as('');
  assert.throws(() => env.call(`triggerAutoResearch({ triggerUid: 'forged' })`), /権限がありません/);
  assert.throws(() => env.call('triggerAutoResearch()'), /権限がありません/);
  env.as(MEMBER);
  assert.throws(() => env.call(`triggerAutoResearch({ triggerUid: '${auid}' })`), /権限がありません/, 'UID を知っていても所有者にならない');
  env.as(OWNER);
  const r = tick(env);
  assert.deepEqual([r.status, r.reason], ['SKIPPED', 'NO_STALE_PLAN'], '計画が無ければ何もしない');
  assert.equal(env.audit().filter((x) => x.action === 'AUTO.RESEARCH' && x.phase === 'DENIED').length, 3, '断ったことも記録に残る');
  const a = env.audit().filter((x) => x.action === 'AUTO.RESEARCH' && x.phase !== 'DENIED');
  assert.deepEqual(a.map((x) => [x.phase, x.actor_email]), [['START', OWNER], ['END', OWNER]], '所有者として記録つきで動く');
}

// ==== 2. いちばん古い計画を選ぶ・ほかの処理が動いていれば何もしない・終わったら 5 時まで次へ ====
const env = setUpEnv();
stubVertex(env);
setNow(env, jst(0, 1, 5));
const fy = env.run('appFy_(appAutoResearchNow_())');
const pPast = seed(env, '己製薬', fy - 1);          // 前の年度（対象外）
const pNever = seed(env, '甲製薬', fy);             // 一度も無い
const pOff = seed(env, '乙製薬', fy, 0);            // AI 調査を無効にしてある（30 日前）
const p10 = seed(env, '丙製薬', fy);                // 10 日前
const pNext = seed(env, '丁製薬', fy + 1);          // 来年度・8 日前
const p2 = seed(env, '戊製薬', fy);                 // 2 日前（まだ新しい）
researchedAt(env, pOff, jst(-30, 12));
researchedAt(env, p10, jst(-10, 12));
researchedAt(env, pNext, jst(-8, 12));
researchedAt(env, p2, jst(-2, 12));
env.run('doGet({})');
assert.equal(autoTriggers(env).length, 1);
{
  const plans = J(env.run(`appAutoResearchPlans_(${fy})`)).map((p) => p.planId);
  assert.deepEqual(plans, [pNever, pOff, p10, pNext, p2], '進行中・今年度か来年度。古い順（一度も無いものが先）');
  const r = tick(env);
  assert.deepEqual([r.status, r.planId, r.via], ['STARTED', pNever, 'TICK'], r.reason);
  const q = jobs(env).filter((j) => j.status === 'QUEUED');
  assert.equal(q.length, 1);
  assert.deepEqual([q[0].kind, q[0].payload.planId, q[0].payload.action, q[0].requestedBy], ['PLAN.RUN', pNever, 'AI.RESEARCH', OWNER], '画面から始めるのと同じ処理を所有者として');
  assert.equal(state(env).chain.jobId, q[0].id, '自動で始めた印は状態の控えに置く');
  // ほかの処理が待っている間は何もしない
  const busy = tick(env);
  assert.deepEqual([busy.status, busy.reason], ['SKIPPED', 'BUSY']);
  assert.equal(jobs(env).filter((j) => j.status === 'QUEUED').length, 1, '二重に始めない');
  // 終わったら（1 時台）次に古い計画へ。準備ができていない計画（AI 調査が無効）は飛ばす
  drain(env, () => (state(env).chain || {}).planId === p10);
  assert.equal(researched(env, pNever), 1, '一度も無かった計画の A-4 が終わった');
  assert.equal(state(env).chain.planId, p10, '続けて次に古い計画（10 日前）を始める');
  assert.equal(state(env).tried[pOff].status, 'NOT_READY');
  assert.equal(state(env).tried[pOff].reason, 'AI_DISABLED');
  assert.equal(state(env).lastResult.status, 'DONE');
  assert.equal(state(env).lastResult.planId, pNever);
  const save = env.audit().filter((x) => x.action === 'PLAN.AI.RESEARCH.SAVE' && x.phase === 'END');
  assert.ok(save.length >= 1 && save.every((x) => x.actor_email === OWNER && x.result === 'OK'), '段ごとに監査の記録が残る');
  // 5 時を過ぎて終わったら、次は始めない
  setNow(env, jst(0, 5, 10));
  drain(env);
  assert.equal(researched(env, p10), 1);
  assert.equal(researched(env, pNext), 0, '5 時を過ぎたら続けない');
  assert.equal(state(env).chain, null);
  const last = autoRuns(env).slice(-1)[0];
  assert.deepEqual([last.status, last.reason, last.via], ['SKIPPED', 'AFTER_HOURS', 'CHAIN']);
  assert.ok(autoRuns(env).some((x) => x.status === 'DONE' && x.event === 'FINISHED' && x.planId === pNever), '結果を実行ログに残す');
  // 次の日の 1 時: 残りの古い計画（来年度・9 日前）。準備ができていない計画は 3 日あける
  setNow(env, jst(1, 1, 5));
  const r2 = tick(env);
  assert.deepEqual([r2.status, r2.planId], ['STARTED', pNext], r2.reason);
  assert.equal(r2.notReady, undefined, '準備ができていない計画は 3 日たつまで確かめ直さない');
  drain(env);
  assert.equal(researched(env, pNext), 1);
  const done = autoRuns(env).slice(-1)[0];
  assert.deepEqual([done.status, done.reason], ['SKIPPED', 'NO_STALE_PLAN'], 'もう古い計画は無い');
  assert.equal(researched(env, p2), 0, '7 日たっていない計画はそのまま');
  assert.equal(researched(env, pPast), 0, '前の年度の計画は集めない');
  assert.equal(researched(env, pOff), 0);
}

// ==== 3. 失敗も控える（次の計画へは進む・3 日あける）。状態をホームに出す ====
{
  stubVertex(env, 403);
  setNow(env, jst(6, 1, 5));
  const r = tick(env);
  assert.deepEqual([r.status, r.planId], ['STARTED', p2], '2 日前だった計画が 8 日前になった');
  assert.deepEqual(r.notReady, [{ planId: pOff, reason: 'AI_DISABLED' }], '3 日たったので確かめ直す（まだ無効）');
  drain(env);
  const st = state(env);
  assert.equal(st.lastResult.status, 'FAILED');
  assert.match(st.lastResult.error, /問い合わせの失敗/);
  assert.equal(st.tried[p2].status, 'FAILED');
  assert.equal(autoRuns(env).slice(-1)[0].reason, 'NO_STALE_PLAN', '失敗した計画は 3 日あける（すぐには繰り返さない）');
  const s = J(env.run('appAutoResearchStatus_()'));
  assert.equal(s.enabled, true);
  assert.equal(s.failures7d, 1);
  assert.equal(s.lastStartedPlan.planId, p2);
  assert.equal(s.lastResult.status, 'FAILED');
  assert.ok(s.lastTickAt && s.trigger.installed);
  assert.deepEqual(s.stalePlans.map((p) => p.planId), [pOff, p2], '古いままの計画（古い順）と、自動で最後に試した結果');
  assert.deepEqual([s.stalePlans[0].lastTry.status, s.stalePlans[0].lastTry.reason], ['NOT_READY', 'AI_DISABLED']);
  assert.equal(env.call('apiHome()').system.autoResearch.failures7d, 1, '管理者のホームの仕組みの状態に出す');
  // 終わりの知らせなしに止まった自動の A-4 は、次の回に「止まった」として控える
  env.run(`appAutoResearchUpdate_(s => { s.chain = { jobId: 'JOB-GONE', planId: __p, requestedBy: '${OWNER}', startedAt: appAutoResearchIso_(appAutoResearchNow_()) }; })`, { __p: p10 });
  setNow(env, jst(6, 1, 30));
  const r3 = tick(env);
  assert.equal(r3.stalled, p10);
  assert.equal(state(env).chain, null);
  assert.equal(J(env.run('appAutoResearchStatus_()')).failures7d, 2);
  // 知らせだけが届かなかった（始めた後に A-4 の記録がある）ものは、終わったとして控える
  env.run(`appAutoResearchUpdate_(s => { s.chain = { jobId: 'JOB-QUIET', planId: __p, requestedBy: '${OWNER}', startedAt: appAutoResearchIso_(new Date(${jst(-1, 1, 5)})) }; })`, { __p: pNever });
  const r4 = tick(env);
  assert.equal(r4.stalled, undefined);
  assert.equal(state(env).tried[pNever].status, 'DONE');
  assert.equal(J(env.run('appAutoResearchStatus_()')).failures7d, 2, '失敗には数えない');
}

// ==== 4. 止めるスイッチ（毎日の回も、続きも）・1 晩の上限 ====
{
  stubVertex(env);
  setNow(env, jst(20, 1, 5));
  const r = tick(env);
  assert.equal(r.status, 'STARTED', r.reason);
  env.props.AUTO_RESEARCH_ENABLED = 'false';
  drain(env);
  const last = autoRuns(env).slice(-1)[0];
  assert.deepEqual([last.status, last.reason, last.via], ['SKIPPED', 'DISABLED', 'CHAIN'], '止めたら続きも始めない');
  assert.equal(jobs(env).filter((j) => j.status === 'QUEUED').length, 0);
  const off = tick(env);
  assert.deepEqual([off.status, off.reason], ['SKIPPED', 'DISABLED']);
  assert.equal(jobs(env).filter((j) => j.status === 'QUEUED').length, 0);
  assert.equal(J(env.run('appAutoResearchStatus_()')).enabled, false);
  delete env.props.AUTO_RESEARCH_ENABLED;
  env.run(`appAutoResearchUpdate_(s => { s.night = { date: __d, count: 12 }; })`, { __d: jstDay(20) });
  const lim = tick(env);
  assert.deepEqual([lim.status, lim.reason], ['SKIPPED', 'NIGHT_LIMIT'], '1 晩に始める数の上限');
  // 画面から始めた A-4 は、同じ人・同じ計画でも、自動で始めた処理（の続き）でなければ自動の続きにしない
  setNow(env, jst(21, 1, 5));
  env.run(`appAutoResearchUpdate_(s => { s.night = null; s.chain = { jobId: 'JOB-OLD', planId: __p, requestedBy: '${OWNER}', startedAt: appAutoResearchIso_(appAutoResearchNow_()) }; })`, { __p: p10 });
  env.as(OWNER);
  const before = autoRuns(env).length;
  const manual = env.runJob('PLAN.RUN', { planId: p10, action: 'AI.RESEARCH' });
  assert.equal(manual.status, 'DONE', manual.error);
  assert.equal(autoRuns(env).length, before, '画面から始めた処理は自動の続きにしない');
  assert.equal(state(env).chain.jobId, 'JOB-OLD', '自動の印はそのまま（次の回に止まったとして控える）');
  assert.equal(env.state.lockHeld, false);
}

// ==== 5. 締めた年度の計画には書かない（選ばない・続きでも選ばない） ====
{
  const fz = setUpEnv();
  stubVertex(fz);
  setNow(fz, jst(0, 1, 5));
  const y = fz.run('appFy_(appAutoResearchNow_())');
  const pNextFy = seed(fz, '甲製薬', y + 1);   // 先に作る（ID の順では先に選ばれるはず）
  const pThisFy = seed(fz, '乙製薬', y);
  const close = (year) => fz.run(`appWithLock_(() => appInsertRows_('YEAR_CLOSURES', [{ fy: __fy, state: 'CLOSED', file_id: 'FILE-TEST', snapshot_sha256: __h,
    plan_count: 1, row_count: 1, bytes: 1, closed_at: appNowIso_(), closed_by: '${OWNER}' }]))`, { __fy: String(year), __h: 'a'.repeat(64) });
  close(y + 1);
  assert.deepEqual(J(fz.run(`appAutoResearchPlans_(${y})`)).map((p) => p.planId), [pThisFy], '締めた年度の計画は候補にしない');
  fz.run('doGet({})');
  const r = tick(fz);
  assert.deepEqual([r.status, r.planId], ['STARTED', pThisFy]);
  close(y);   // 始めた後で今年度も締まった
  drain(fz);
  assert.equal(researched(fz, pThisFy), 0, '裏の処理が年度の締めを確かめて止める');
  assert.equal(researched(fz, pNextFy), 0);
  const runs = autoRuns(fz);
  assert.deepEqual(runs.slice(-1).map((x) => [x.status, x.reason]), [['SKIPPED', 'NO_STALE_PLAN']], '続きでも締めた年度は選ばない');
  assert.ok(!J(fz.run('appJobList_()')).some((j) => j.payload && j.payload.planId === pNextFy), '締めた年度の計画の処理は一度も作らない');
  assert.equal(J(fz.run('appAutoResearchStatus_()')).stalePlans.length, 0);
}

console.log('app-auto-research: all tests passed');
