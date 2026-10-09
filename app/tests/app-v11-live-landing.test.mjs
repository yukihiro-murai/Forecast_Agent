#!/usr/bin/env node
/**
 * app-v11-live-landing.test.mjs — 実績を取り込んで（B-1）から当たり具合を計算する（B-2）までの間の着地見込み（2026-10-09 村井さん承認）。
 *   1. 締まった月の境目（appLandingCutoff_ の hist）: B-1 の後に B-2 がまだの間は、最後に成功した B-2 の前の、最後の B-1 の境目で数え続ける
 *      （PROCESS_STATUS は最後の回しか残さないので、消さずに足す記録から探す）。組が無ければ ''（前と同じ）
 *   2. 空模様（appLandingSky_）: 年度の初めの間（締まった 2 か月）は k = 2 のまま数字を出す。初めての取り込み（前の組が無い）だけ計算待ち（eval_pending）。
 *      ふつうの状態（B-2 が済んでいる・前の組がある）は 045e7db の計算とまったく同じ（新しい項目 wait のほか）
 *   3. 計画の一覧を通して: PLAN_ACTIONS の記録・旧来の RUN_LOG の記録から前の組を見つける。初めての取り込みは計算待ちで、年度の見込みの試しも幅を出さない。
 *      ふつうの計画は今までどおり
 * モックの上の確かめで、本物の Apps Script の上では動かしていない。
 *
 *   node app/tests/app-v11-live-landing.test.mjs
 */
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { execFileSync } from 'node:child_process';
import { OWNER, makeEnv, J, repoRoot } from './gas-mock.mjs';

const fyYms = (fy) => Array.from({ length: 12 }, (_, i) => { const m = 4 + i; return (m > 12 ? fy + 1 : fy) + '/' + String(m > 12 ? m - 12 : m).padStart(2, '0'); });
const yms = fyYms(2026);
const acts = (v) => Object.fromEntries(v.map((x, i) => [yms[i], x]));
const pure = makeEnv();
const run = (code, vars) => J(pure.run(code, vars));
const flat = yms.map((ym) => ({ ym, p10: 80, p50: 100, p90: 120 }));
const base = { fy: 2026, months: flat, budget: 1200 };
const sky = (over) => run('appLandingSky_(__in)', { __in: Object.assign({}, base, over) });
const nums = (x) => [x.landing, x.landingSd, x.landingP10, x.landingP90, x.pAbove, x.ratio, x.theta, x.credibility];

// ==== 1. 締まった月の境目: B-1 の後に B-2 がまだの間は、前の B-1 → B-2 の組の B-1 の境目 ====
{
  const jst = (s) => Date.parse(s + '+09:00');
  const stp = (b1, b2) => [['step2_status', b1], ['step5_status', b2]].filter((x) => x[1]).map(([k, t]) => ({ step_key: k, status: 'success', last_run_date: t }));
  const cut = (rows, hist) => pure.run('appLandingCutoff_(__r, __h)', { __r: rows, __h: hist });
  const b1 = (s) => ({ step: 'b1', t: jst(s) }), b2 = (s) => ({ step: 'b2', t: jst(s) });
  const now = stp('2026-07-06T10:00:00+0900', '2026-06-06T11:00:00+0900');   // B-1（7/06）の後に B-2 がまだ（最後の B-2 は 6/06）
  const hist = [b1('2026-05-06T10:00:00'), b2('2026-05-06T11:00:00'), b1('2026-06-06T10:00:00'), b2('2026-06-06T11:00:00'), b1('2026-07-06T10:00:00')];
  assert.equal(cut(now, hist), '2026/06', '最後の B-2（6/06）の前の B-1（6/06）の境目: 4・5 月が締まった月');
  assert.equal(cut(now), '', '記録を渡さなければ前と同じ（数えない）');
  assert.equal(cut(now, []), '', '組が無い（初めての取り込み）');
  assert.equal(cut(now, [b1('2026-07-06T10:00:00')]), '', '今の B-1 しか無い');
  assert.equal(cut(now, [b1('2026-06-03T10:00:00'), b2('2026-06-03T11:00:00'), b1('2026-07-06T10:00:00')]), '2026/05', '前の B-1 が 6/03（5 月はまだ途中）なら 4 月まで');
  assert.equal(cut(now, [b1('2026-06-06T10:00:00'), b1('2026-06-07T09:00:00'), b1('2026-07-06T10:00:00')]), '2026/06', 'B-2 の時刻（PROCESS_STATUS）より後の B-1 は使わない');
  // PROCESS_STATUS に B-2 の成功が無い: 記録の B-2 のうち、今の B-1 より前で一番新しいもの
  const noB2 = stp('2026-07-06T10:00:00+0900', null);
  assert.equal(cut(noB2, hist), '2026/06');
  assert.equal(cut(noB2, [b1('2026-06-06T10:00:00'), b1('2026-07-06T10:00:00'), b2('2026-07-06T11:00:00')]), '', '今の B-1 より後の B-2 の記録は使わない（PROCESS_STATUS が正）');
  // ふつうの状態（B-2 が済んでいる）は記録を渡しても同じ
  const done = stp('2026-07-06T10:00:00+0900', '2026-07-06T11:00:00+0900');
  assert.deepEqual([cut(done), cut(done, hist), cut(done, [])], ['2026/07', '2026/07', '2026/07']);
  assert.deepEqual([cut([]), cut([], hist)], ['', ''], 'B-1 が無ければ空');
  // 文字の ISO 時刻（PLAN_ACTIONS の as_of の形）・日時の型（RUN_LOG）の記録も読める（読むのは計画の一覧の側。ここでは時刻の数）
  assert.equal(pure.run(`appLandingTime_('', '2026-06-06T10:00:00+0900')`), jst('2026-06-06T10:00:00'));
}

// ==== 2. 空模様: 年度の初めの間は前の組の締まった月で数える。初めての取り込みだけ計算待ち ====
{
  // 年度の初めの間（7/06 に B-1、最後の B-2 は 6/06 → 締まった 4・5 月）: k = 2 のまま着地を出す（前は境目が '' で、締まった月 0 として数えていた）
  const w = sky({ actual: acts([100, 90]), cutoffYm: '2026/06', pendingCutoffYm: '2026/07', todayYm: '2026/07' });
  assert.deepEqual([w.k, w.actualYtd, w.sky, w.skyReason, w.wait], [2, 190, 'harenochi', 'ratio', '']);
  assert.ok(w.landing > 0 && w.landingSd > 0 && w.credibility > 0, '締まった 2 か月の実績を入れた着地');
  const w0 = sky({ actual: acts([100, 90]), cutoffYm: '2026/06', todayYm: '2026/07' });
  assert.deepEqual(nums(w), nums(w0), 'B-2 待ちでも、前の組の境目で数えた着地は B-2 が済んだときの前の値と同じ');
  // 初めての取り込み（前の組が無い）: 遅れていなくても数字を出さず、計算待ち
  const f = sky({ actual: {}, cutoffYm: '', pendingCutoffYm: '2026/06', todayYm: '2026/06' });
  assert.deepEqual([f.k, f.sky, f.skyReason, f.wait, ...nums(f)], [0, 'mikakunin', 'eval_pending', 'eval_pending', null, null, null, null, null, null, null, null], '初めての取り込みは計算待ち');
  assert.deepEqual([sky({ budget: null, actual: {}, cutoffYm: '', pendingCutoffYm: '2026/06', todayYm: '2026/06' }).skyReason,
    sky({ budget: null, actual: {}, cutoffYm: '', pendingCutoffYm: '2026/06', todayYm: '2026/06' }).wait], ['no_budget', 'eval_pending'], '予算が無いのが先（数字を出さないわけは wait に）');
  assert.equal(sky({ budget: null, actual: {}, cutoffYm: '', pendingCutoffYm: '2026/06', todayYm: '2026/06' }).landing, null, '予算が無くても数字は出さない');
  // 初めての取り込みでも、その B-1 がこの年度の締まった月を足さない（先の年度の計画）なら、そのまま数える
  const nx = sky({ fy: 2027, months: fyYms(2027).map((ym) => ({ ym, p10: 80, p50: 100, p90: 120 })), actual: {}, cutoffYm: '', pendingCutoffYm: '2026/10', todayYm: '2026/10' });
  assert.deepEqual([nx.k, nx.skyReason, nx.wait, nx.landing], [0, 'ratio', '', 1200], '先の年度は締まった月が無いので計算待ちにしない');
  // 初めての取り込みの B-1 そのものが 3 か月遅れていれば、実績の遅れ
  assert.deepEqual([sky({ actual: {}, cutoffYm: '', pendingCutoffYm: '2026/05', todayYm: '2026/08' }).skyReason, sky({ actual: {}, cutoffYm: '', pendingCutoffYm: '2026/05', todayYm: '2026/08' }).wait],
    ['stale_actuals', 'stale_actuals']);
  assert.equal(sky({ actual: {}, cutoffYm: '', pendingCutoffYm: '2026/05', todayYm: '2026/07' }).skyReason, 'eval_pending', 'B-1 の境目で遅れが 2 か月なら計算待ち');
  // 前の組があっても、その組が 3 か月以上前（B-1 の境目では遅れていない）なら、これまでどおり計算待ち
  assert.equal(sky({ actual: acts([100]), cutoffYm: '2026/05', pendingCutoffYm: '2026/10', todayYm: '2026/10' }).skyReason, 'eval_pending');
  assert.equal(sky({ actual: acts([100, 100, 100, 100]), cutoffYm: '2026/08', todayYm: '2026/10' }).wait, '', 'ふつうの状態の wait は空');
  assert.equal(sky({ actual: acts([100]), cutoffYm: '2026/05', todayYm: '2026/10' }).wait, 'stale_actuals', '遅れのわけは空模様によらず wait に');

  // ふつうの状態（B-2 が済んでいる・前の組がある・B-1 がこの年度の締まった月を足さない）は、045e7db の計算とまったく同じ（新しい項目 wait のほか）
  let oldSrc = null;
  try { oldSrc = execFileSync('git', ['show', '045e7db:app/src/Landing.js'], { cwd: repoRoot, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }); } catch (e) { oldSrc = null; }
  if (!oldSrc) console.log('app-v11-live-landing: 045e7db の Landing.js を読めないので、前の計算との照合を飛ばしました');
  else {
    const old = vm.createContext({});
    vm.runInContext(oldSrc, old);
    const variants = [flat, yms.map((ym, i) => ({ ym, p10: 60 + i, p50: 90 + 2 * i, p90: 130 + 3 * i }))];
    const actSets = [{}, acts([100, 90, 110]), acts([80, 80, 80, 80, 80, 80]), acts([100, 100, 100, 100, 100, 30]), acts(Array(12).fill(100)), acts([0, 0])];
    let n = 0;
    for (const months of variants) for (const actual of actSets) for (const cutoffYm of ['', '2026/04', '2026/06', '2026/08', '2026/10', '2027/04']) {
      for (const todayYm of ['', '2026/04', '2026/07', '2026/10', '2027/04']) for (const pendingCutoffYm of ['', '2026/04', '2026/08', '2026/10', '2027/04']) for (const budget of [1200, null, 600]) {
        const pendK = pendingCutoffYm ? yms.filter((ym) => ym < pendingCutoffYm).length : 0;
        if (!cutoffYm && pendK > 0) continue;   // 初めての取り込みの間だけが変わった
        const inp = { fy: 2026, months, actual, cutoffYm, todayYm, pendingCutoffYm, budget, runs: [{ p50: 1500, ageDays: 3 }, { p50: 1000, ageDays: 40 }] };
        const a = run('appLandingSky_(__in)', { __in: inp });
        old.__in = JSON.parse(JSON.stringify(inp));
        const b = JSON.parse(JSON.stringify(vm.runInContext('appLandingSky_(__in)', old)));
        delete a.wait;
        assert.deepEqual(a, b, JSON.stringify(inp));
        n++;
      }
    }
    assert.ok(n > 1000, '照合した数: ' + n);
  }
}

// ==== 3. 計画の一覧を通して（PLAN_ACTIONS・RUN_LOG の記録から前の組を見つける） ====
{
  const env = makeEnv();
  env.run(`appToday_ = function () { return '2026-07-08'; }`);
  env.as(OWNER);
  env.call('apiSetup()');
  const HC = J(env.run('APP_ENGINE_SHEETS.EVAL_COMPARE_MONTHLY.header'));
  const HS = J(env.run('APP_ENGINE_SHEETS.PROCESS_STATUS.header'));
  const HE = J(env.run('APP_ENGINE_SHEETS.EVAL_LOG.header'));
  const HR = J(env.run('APP_ENGINE_SHEETS.RUN_LOG.header'));
  const obj = (H, o) => H.map((h) => (o[h] === undefined ? '' : o[h]));
  const output = [['FY2026 売上予測']];
  for (let r = 2; r <= 25; r++) output.push([]);
  output.push(['年度合計（予測）', 1100, 1200, 1300]);
  output.push([], ['月', 'P10', 'P50', 'P90', '', '', '', '採用予測', '上乗せ']);
  yms.forEach((ym) => output.push([ym, 80, 100, 120, '', '', '', 100, '']));
  // 検証の表: 前の B-2（6/06）が書いた 4・5 月（6 月の途中の実績も）。EVAL_LOG も 4・5 月（今の検証の版）
  const cmp = [['2026/04', 105], ['2026/05', 90], ['2026/06', 40]].map(([ym, a]) => obj(HC, { target_month: ym, actual_total: a, forecast_total_p10: 80, forecast_total_p50: 100,
    forecast_total_p90: 120, ape_p50: Math.abs(100 - a) / a, range_outside_flag: 0 }));
  const evalRows = [['2026/04', 105], ['2026/05', 90]].map(([ym, a]) => obj(HE, { eval_id: 'E-' + ym, evaluated_at: new Date(2026, 5, 6, 11), client: 'x', target_month: ym, scenario: 'neutral',
    pred: 100, actual: a, evaluation_policy_version: 'policy-2026H1-v3', constraint_relevant_flag: 1 }));
  const st = (b1, b2) => [HS, ['step2_status', b1, 'owner', 'success', 'x', 10, ''], ...(b2 ? [['step5_status', b2, 'owner', 'success', '', 2, '']] : [])];
  const book = (client, b1, b2, runLog) => env.makeBook(client, {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', client], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
    OUTPUT: { values: output },
    EVAL_COMPARE_MONTHLY: { values: [HC, ...cmp] },
    EVAL_LOG: { values: [HE, ...evalRows], formats: { D: '@' } },
    PROCESS_STATUS: { values: st(b1, b2) },
    ...(runLog ? { RUN_LOG: { values: [HR, ...runLog.map(([fn, at]) => obj(HR, { run_id: 'R-' + fn + at.getTime(), run_at: at, run_by: 'owner', function_name: fn, client, status: 'success', count: 1 }))] } } : {}),
  });
  const B1 = new Date(2026, 6, 6, 10), B2old = new Date(2026, 5, 6, 11);
  // 前の組を PLAN_ACTIONS（新アプリで動かした回）から見つける計画
  const idA = env.seedPlan(book('記録製薬', B1, B2old));
  env.run(`appWithLock_(() => appInsertRows_('PLAN_ACTIONS', __r))`, { __r: [
    ['ACT-1', 'IMPORT.ACTUALS', '2026-06-06T10:00:00+0900'], ['ACT-2', 'EVAL.REPORT', '2026-06-06T11:00:00+0900'], ['ACT-3', 'IMPORT.ACTUALS', '2026-07-06T10:00:00+0900'],
    ['ACT-4', 'INPUT.SAVE', '2026-07-01T10:00:00+0900']].map(([id, action, asOf]) => ({ action_id: id, plan_id: idA, action, status: 'DONE', as_of: asOf, started_at: asOf, finished_at: asOf, actor_email: OWNER })) });
  // 前の組を旧来の RUN_LOG（旧来のブックから写した回）から見つける計画
  const idR = env.seedPlan(book('旧ログ製薬', B1, B2old, [['importActualEvalMonthly', new Date(2026, 5, 6, 10)], ['updatePhase1EvaluationReport', B2old],
    ['runForecastA9', new Date(2026, 5, 20, 10)], ['importActualEvalMonthly', B1]]));
  // 初めての取り込み（B-1 だけ。前の組が無い）
  const idF = env.seedPlan(book('初回製薬', B1, null));
  // ふつう（B-1 の後に B-2 が済んでいる）
  const idN = env.seedPlan(book('ふつう製薬', B1, new Date(2026, 6, 6, 11)));
  const ps = Object.fromEntries(env.call('apiPortfolio()').plans.map((p) => [p.planId, p]));
  for (const [id, what] of [[idA, 'PLAN_ACTIONS'], [idR, 'RUN_LOG']]) {
    const p = ps[id];
    assert.deepEqual([p.k, p.actualMonths, p.actualYtd, p.skyReason, p.scoredMonths, p.mapeMonths], [2, 2, 195, 'ratio', 2, 2], what + ' の記録から前の組（6/06）: 締まった 4・5 月のまま');
    assert.ok(p.landing > 0 && p.landingSd > 0, what + ': 着地の数字を出す');
    assert.deepEqual([p.aligned.k, p.aligned.pending, p.aligned.p10], [2, false, null], what + ': 年度の見込みの試しは締まった月がある形（数え直し中ではない）');
    assert.ok(p.reach && p.reach.k === 2, what + ': 予算に届く見込みも締まった 2 か月で');
  }
  const f = ps[idF];
  assert.deepEqual([f.k, f.actualMonths, f.actualYtd, f.sky, f.skyReason, f.landing, f.reach], [0, 0, null, 'mikakunin', 'eval_pending', null, null], '初めての取り込みは計算待ち');
  assert.deepEqual([f.aligned.pending, f.aligned.p10, f.aligned.sd], [true, null, null], '試しの年度も幅を出さない');
  const n = ps[idN];
  assert.deepEqual([n.k, n.actualMonths, n.actualYtd, n.skyReason, n.aligned.pending], [3, 3, 235, 'shock', false], 'ふつうの計画は B-1（7/06）の境目で 4〜6 月（6 月の 40 も締まった月。急な落ち込み）');
  // ふつうの計画は、記録を探さない（B-1 の後に B-2 がまだの計画だけ）
  assert.deepEqual(J(env.run('appPortfolioStepHist_([])')), {});
  const h = J(env.run('appPortfolioStepHist_(__ids)', { __ids: [idA, idR, idN] }));
  assert.deepEqual(h[idA].map((x) => x.step), ['b1', 'b2', 'b1'], 'PLAN_ACTIONS の B-1・B-2 だけ（ほかの操作は入れない）');
  assert.deepEqual(h[idR].map((x) => x.step), ['b1', 'b2', 'b1'], 'RUN_LOG の B-1・B-2 の success だけ');
  assert.equal(h[idN], undefined, '記録の無い計画は無い');
  // ホームの合計: 前の組で数えた計画は着地の合計に入り、計算待ちの計画は入らない
  const ht = env.call('apiHome()').totals;
  assert.deepEqual([ht.plans, ht.landingPlans], [4, 3]);
}

console.log('app-v11-live-landing: all tests passed');
