#!/usr/bin/env node
/*
 * app-legacy-closed-months.test.mjs — 検証の決まり（2026-10-07 村井さん承認 D4〜D6）を、本物の旧来の計算をモックの上で動かして確かめる。
 *   0. 締まりの日数（月末から 5 日）は旧来の計算と新アプリで同じ数（新アプリに *_CLOSE_LAG_DAYS の数があれば、同じ値であること）
 *   1. B-1（実績の取り込み）: 月末から 3 日後の取り込みでは 9 月は open、6 日後なら closed（取り込んだ日時を source_updated_at に残す）
 *   2. B-2（検証）: 締まった月だけを、その月が始まる前の最後の予測（6/19 の回）で測る。後から作った回（10/02）は使わない。
 *      月が始まる前の予測が無い月（4〜6 月）は測らずに数える。前の版で書いた EVAL_LOG の行は消さない
 *   3. B-3・B-4・B-5・C-1 は、前の版で書いた行（締まっていない月の途中の売上・後から作った予測で測った行）を使わない
 *   4. B-5 は締まった月だけで学ぶ（3 日後の取り込みでは 2 か月で足りず学ばない。6 日後なら 7〜9 月の 3 か月で、6/19 の予測から学ぶ）
 *   5. C-1 は締まった 3 か月（7〜9 月）を、6/19 の回の AI と人の押しで数える
 *   6. 計画の画面の「人の記入」（旧来の webParseEval_）も、今の決まりで測った月の行だけ（前からある行は消さず、出さない）
 * 数字はテスト用の作りもの。本物の Apps Script での確認の代わりではない。
 *
 *   node app/tests/app-legacy-closed-months.test.mjs
 */
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { setUpEnv, sources, repoRoot } from './gas-mock.mjs';

/** 日本の時刻（月は 1 始まり） */
const jst = (y, m, d, H = 0, M = 0) => new Date(Date.UTC(y, m - 1, d, H - 9, M));
const CLIENT = 'テスト製薬';
const POLICY = 'policy-2026H1-v3';
const isDate = (v) => Object.prototype.toString.call(v) === '[object Date]';   // 計算用ブックの日付は別の realm で作られる
const ymOf = (v, t) => {
  if (isDate(v) || t === 'd') { const d = new Date(v); return d.getFullYear() + '/' + String(d.getMonth() + 1).padStart(2, '0'); }
  return String(v);
};

// ==== 0. 締まりの日数は 1 つの数（旧来の計算）。新アプリに同じ意味の数があれば同じ値 ====
{
  const legacySrc = await readFile(path.join(repoRoot, 'Forecast_Agent.js'), 'utf8');
  const lagOf = (text) => [...text.matchAll(/^\s*const ([A-Z_]*CLOSE_LAG_DAYS)\s*=\s*(\d+)\s*;/gm)].map((m) => [m[1], Number(m[2])]);
  assert.deepEqual(lagOf(legacySrc), [['ACTUAL_CLOSE_LAG_DAYS', 5]], '旧来の計算: 月末から 5 日（1 か所だけ）');
  assert.deepEqual(lagOf(sources['LegacyEngine.js']), [['ACTUAL_CLOSE_LAG_DAYS', 5]], '新アプリに取り込んだ旧来の計算も同じ');
  const appSide = Object.keys(sources).filter((n) => n !== 'LegacyEngine.js').flatMap((n) => lagOf(sources[n]).map(([k, v]) => [n, k, v]));
  appSide.forEach(([n, k, v]) => assert.equal(v, 5, `${n} の ${k} は旧来の計算の ACTUAL_CLOSE_LAG_DAYS と同じ日数`));
  assert.match(legacySrc, /^const EVAL_CALENDAR_TZ = 'Asia\/Tokyo';/m, '日本の暦で数える');
  assert.match(sources['Config.js'], /const APP_TZ = 'Asia\/Tokyo';/, '新アプリの暦も日本');
}

const env0 = setUpEnv();
const H = env0.run('(() => { const o = {}; Object.keys(APP_ENGINE_SHEETS).forEach(k => { o[k] = APP_ENGINE_SHEETS[k].header || null; }); return o; })()');
const row = (sheet, o) => H[sheet].map((h) => (o[h] === undefined ? '' : o[h]));
const CONFIG = { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', CLIENT], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] };

// ==== 1. B-1: 取り込んだ日で締まりが決まる ====
{
  const env = env0;
  const ext = Array.from({ length: 70 }, () => '');
  const rec = (month, amount) => { const r = ext.slice(); r[40] = CLIENT; r[45] = 'ベース'; r[49] = '製品A'; r[56] = month; r[65] = amount; return r; };
  const zac = env.makeBook('Veeva 売上分析ツール', { '*2026_actual_value': { cols: 70, values: [ext.map((_, i) => 'c' + (i + 1)),
    rec(new Date(2026, 7, 20), 1000), rec(new Date(2026, 8, 20), 1000), rec(new Date(2026, 9, 2), 300)] } });
  env.call('apiSaveSetting(__in)', { __in: { key: 'source.zac_spreadsheet', value: 'https://docs.google.com/spreadsheets/d/' + zac.getId() + '/edit' } });
  const importAt = (at) => {
    const book = env.makeBook('クライアント別売上予測', {
      CONFIG,
      PROCESS_STATUS: { values: [H.PROCESS_STATUS, ['step1_status', new Date(2026, 3, 1), 'owner', 'success', CLIENT, 48, '']] },
      RUN_LOG: { values: [H.RUN_LOG] },
      ACTUAL_EVAL_MONTHLY: { values: [H.ACTUAL_EVAL_MONTHLY] },
    });
    env.run('appLegacyCall_(__book, { asOfMs: __at, seed: "b1", actor: "owner@bigm2y.com" }, "webRunImportActuals", [])', { __book: book, __at: at.getTime() });
    return book.getSheetByName('ACTUAL_EVAL_MONTHLY').getDataRange().getValues().slice(1)
      .map((r) => [ymOf(r[3]), r[5], isDate(r[6]) ? r[6].getTime() : r[6]]);
  };
  const at3 = jst(2026, 10, 3, 12);
  assert.deepEqual(importAt(at3), [['2026/08', 'closed', at3.getTime()], ['2026/09', 'open', at3.getTime()], ['2026/10', 'open', at3.getTime()]],
    '月末から 3 日後の取り込み: 9 月は締まっていない（前は取り込んだ日の暦だけで closed にしていた）');
  const at6 = jst(2026, 10, 6, 9);
  assert.deepEqual(importAt(at6), [['2026/08', 'closed', at6.getTime()], ['2026/09', 'closed', at6.getTime()], ['2026/10', 'open', at6.getTime()]],
    '6 日後の取り込み: 9 月は締まった');
  assert.deepEqual(importAt(jst(2026, 10, 5, 0, 0)).map((r) => r[1]), ['closed', 'closed', 'open'], '境目: 月末 + 5 日（10/05 0:00）は締まった');
  assert.deepEqual(importAt(jst(2026, 10, 4, 23, 59)).map((r) => r[1]), ['closed', 'open', 'open'], '境目: 10/04 23:59 は締まっていない');
}

// ==== 2〜5. B-2〜C-1（PLAN.RUN で本物の旧来の計算を動かす） ====
const FY_MONTHS = ['2026/04', '2026/05', '2026/06', '2026/07', '2026/08', '2026/09', '2026/10', '2026/11', '2026/12', '2027/01', '2027/02', '2027/03'];
/** 実績（取り込んだ値）。10 月以降は途中の売上（締まっていない） */
const ACT = { '2026/04': 1000, '2026/05': 1000, '2026/06': 1000, '2026/07': 1000, '2026/08': 1000, '2026/09': 1000, '2026/10': 300, '2026/11': 50, '2026/12': 40, '2027/03': 30 };
const PRE = 1100;   // 6/19 の回（R1）の P50: 7 月以降の月が始まる前の予測
const LATE = 900;   // 10/02 の回（R2）の P50: 後から作った予測（7〜10 月には使わない）
const R1_AT = jst(2026, 6, 19, 10), R2_AT = jst(2026, 10, 2, 10);
const QUANT = 1050; // 統計だけの予測（7〜9 月の実績 1000 は下 = 驚きは下向き）

function planBook(env, importedAt) {
  const snap = [H.FORECAST_SNAPSHOT];
  const run = (sid, at, p50, factor) => FY_MONTHS.forEach((ym) => ['nega', 'neutral', 'posi'].forEach((sc, i) => {
    const v = p50 * [0.9, 1, 1.1][i];
    snap.push([sid, at, CLIENT, ym, sc, p50, 0, 0, 0, v, p50 * 0.9, p50 * 1.1, JSON.stringify({ opinion: '', forecast_source: 'forecast_open' }), '',
      JSON.stringify({ version: 'test', bias_correction_factor: factor, residual_month_bias_json: '' })]);
  }));
  run('R1', R1_AT, PRE, 1);
  run('R2', R2_AT, LATE, 0.9);
  // 前の決まりの B-1 で付いた印（取り込んだ日の暦: 9 月までは closed）
  const actual = [H.ACTUAL_EVAL_MONTHLY].concat(Object.keys(ACT).map((ym) => [CLIENT, 'BASE', '製品A', ym, ACT[ym], ym <= '2026/09' ? 'closed' : 'open', importedAt]));
  // 前の版の B-2（10/02 の回で、締まっていない月も測った）が書いた EVAL_LOG。この 30 行は消さない
  const evalLog = [H.EVAL_LOG];
  Object.keys(ACT).forEach((ym) => ['nega', 'neutral', 'posi'].forEach((sc, i) => {
    const p = LATE * [0.9, 1, 1.1][i];
    evalLog.push(row('EVAL_LOG', { eval_id: 'OLD-' + ym + sc, evaluated_at: jst(2026, 10, 2, 11), client: CLIENT, target_month: ym, scenario: sc, pred: p, actual: ACT[ym],
      ape: Math.abs(p - ACT[ym]) / ACT[ym], signed_error: p - ACT[ym], abs_error: Math.abs(p - ACT[ym]), bias_direction: p > ACT[ym] ? 'over' : 'under',
      model_version: '2.4.0-dev', evaluation_policy_version: 'policy-2026H1-v2', constraint_relevant_flag: sc === 'neutral' ? 1 : 0 }));
  }));
  // 前の版の B-2 が書いた検証の表（途中の売上の月も誤差あり）
  const cmp = [H.EVAL_COMPARE_MONTHLY].concat(FY_MONTHS.map((ym) => {
    const a = ACT[ym];
    return row('EVAL_COMPARE_MONTHLY', { target_month: ym, forecast_total: LATE, actual_total: a === undefined ? '' : a, forecast_total_p10: LATE * 0.9, forecast_total_p50: LATE,
      forecast_total_p90: LATE * 1.1, signed_error_p50: a === undefined ? '' : LATE - a, abs_error_p50: a === undefined ? '' : Math.abs(LATE - a),
      ape_p50: a === undefined ? '' : Math.abs(LATE - a) / a, half_label: ym < '2026/10' ? 'FY2026-H1' : 'FY2026-H2', over_flag: a !== undefined && LATE > a ? 1 : 0,
      range_outside_flag: a === undefined ? '' : (a < LATE * 0.9 || a > LATE * 1.1 ? 1 : 0) });
  }));
  // C-1 の材料: 7〜9 月。6/19 の回（R1）は当たり、10/02 の回（R2）は外れ
  const impact = (id, at, ym, dir, k) => row('AI_IMPACT_HISTORY', { run_id: id, run_at: at, client: CLIENT, target_month: ym, k_ai: k, ai_direction: dir,
    pred_p50: id === 'R1' ? PRE : LATE, pred_p50_quant_only: QUANT, forecast_source: 'forecast_open' });
  const push = (id, at, ym, dir) => row('SUBJECTIVE_IMPACT_HISTORY', { run_id: id, run_at: at, client: CLIENT, target_month: ym, source_type: 'opinion', source_key: '鷹野',
    push_step: dir * 0.05, push_direction: dir, applied_reliability_r: 1, forecast_source: 'forecast_open' });
  const Q2 = ['2026/07', '2026/08', '2026/09'];
  return env.makeBook('クライアント別売上予測', {
    CONFIG,
    // 前の版の B-2 は 10/02 に動いた（B-1 の前。B-4 は B-2 が成功していれば動く）
    PROCESS_STATUS: { values: [H.PROCESS_STATUS, ['step2_status', importedAt, 'owner', 'success', CLIENT, 10, ''], ['step4_status', R2_AT, 'owner', 'success', CLIENT, 12, ''],
      ['step5_status', jst(2026, 10, 2, 11), 'owner', 'success', '', 30, '']] },
    RUN_LOG: { values: [H.RUN_LOG] },
    ACTUAL_EVAL_MONTHLY: { values: actual, formats: { D: '@' } },
    FORECAST_SNAPSHOT: { values: snap, formats: { D: '@' } },
    EVAL_LOG: { values: evalLog, formats: { D: '@' } },
    EVAL_COMPARE_MONTHLY: { values: cmp, cols: 40 },
    // 前の版の B-4 が書いた、締まっていない 11 月の行（消さない・書き換えない）
    EVAL_INSIGHTS: { values: [H.EVAL_INSIGHTS, row('EVAL_INSIGHTS', { evaluated_at: jst(2026, 10, 2, 12), client: CLIENT, target_month: '2026/11', actual_total: 50, pred_p50: LATE,
      diff: 50 - LATE, error_rate: (50 - LATE) / 50, diagnostic_type: 'range_breach', cause_bucket: 'range_outside', action_type: 'update', status: 'open', review_cycle: 'monthly_light' })],
    formats: { C: '@' } },
    DASHBOARD: { values: [H.DASHBOARD] },
    AI_IMPACT_HISTORY: { values: [H.AI_IMPACT_HISTORY].concat(Q2.map((ym) => impact('R1', R1_AT, ym, 'down', 0.997)), Q2.map((ym) => impact('R2', R2_AT, ym, 'up', 1.05))) },
    SUBJECTIVE_IMPACT_HISTORY: { values: [H.SUBJECTIVE_IMPACT_HISTORY].concat(Q2.map((ym) => push('R1', R1_AT, ym, -1)), Q2.map((ym) => push('R2', R2_AT, ym, 1))) },
    AI_SCORE_HISTORY: { values: [H.AI_SCORE_HISTORY] },
    CALIBRATION_STATE: { values: [H.CALIBRATION_STATE, [CLIENT, jst(2026, 10, 2), 'auto', '', '', '[]', 1, '', '', '', '', 1, '']] },
    CALIBRATION_HISTORY: { values: [H.CALIBRATION_HISTORY] },
    SOURCE_RELIABILITY: { values: [H.SOURCE_RELIABILITY] },
    POOL_PRIOR: { values: [H.POOL_PRIOR] },
    QUARTERLY_REVIEW: { values: [['']] },
    QUARTERLY_REVIEW_LOG: { values: [H.QUARTERLY_REVIEW_LOG] },
    RELIABILITY_EVIDENCE: { values: [H.RELIABILITY_EVIDENCE] },
  });
}

function setUpPlan(importedAt) {
  const env = setUpEnv();
  const planId = env.seedPlan(planBook(env, importedAt));
  const rows = (sheet) => env.table('ENG_' + sheet).filter((r) => r.plan_id === planId);
  const runAction = (action) => env.runJob('PLAN.RUN', { planId, action });
  const ok = (action) => { const st = runAction(action); assert.equal(st.status, 'DONE', action + ': ' + st.error); return st; };
  const metric = (name) => { const m = rows('DASHBOARD').find((r) => r.metric === name); return m ? m.value : undefined; };
  const cal = () => { const c = rows('CALIBRATION_STATE'); assert.equal(c.length, 1); return c[0]; };
  const shownInsights = () => env.call('apiPlanView(__in)', { __in: { planId } }).boot.eval.insights.map((r) => [r.month, r.row]);
  return { env, planId, rows, runAction, ok, metric, cal, shownInsights };
}
const evalByYm = (rows) => {
  const o = {};
  rows('EVAL_LOG').forEach((r) => { const ym = ymOf(r.target_month, r._types.charAt(3)); (o[ym] = o[ym] || []).push(r); });
  return o;
};
const cmpByYm = (rows) => { const o = {}; rows('EVAL_COMPARE_MONTHLY').forEach((r) => { o[ymOf(r.target_month, r._types.charAt(0))] = r; }); return o; };
const near = (a, b, msg) => assert.ok(Math.abs(Number(a) - b) < 1e-9, `${msg}: ${a} / ${b}`);

// ---- 3 日後に取り込んだ計画（9 月は締まっていない。締まった月は 4〜8 月で、月が始まる前の予測があるのは 7・8 月だけ） ----
{
  const P = setUpPlan(jst(2026, 10, 3, 9));
  const insightsBefore = JSON.stringify(P.rows('EVAL_INSIGHTS'));

  // 3. B-2 の前（前の版の行だけ）: B-3・B-4・B-5・C-1 は前の版の行を使わない
  P.ok('EVAL.DASHBOARD');
  assert.equal(P.metric('annual_abs_error_rate_latest'), '', 'B-3: 前の版の検証の表（途中の売上の月）は使わない');
  assert.equal(P.metric('annual_constraint_pass'), 'N/A');
  assert.equal(P.metric('half2_wape_latest'), '', '下期（途中の月だけ）の外れ幅も出さない');
  assert.equal(P.metric('range_outside_count_latest'), '');
  P.ok('EVAL.INSIGHTS');
  assert.equal(JSON.stringify(P.rows('EVAL_INSIGHTS')), insightsBefore, 'B-4: 前の版の行からは書かない（前からある 11 月の行も消さない）');
  assert.deepEqual(P.shownInsights(), [], '画面の「人の記入」にも、締まっていない 11 月の古い行は出さない');
  let st = P.ok('LEARN.MONTHLY');
  assert.equal(st.result.result.result.skipped, 'insufficient_eval_months', 'B-5: 前の版の行では学ばない');
  assert.equal(st.result.result.result.n, 0);
  assert.equal(Number(P.cal().bias_correction_factor), 1, '補正は変わらない');
  st = P.runAction('REVIEW.GENERATE');
  assert.equal(st.status, 'FAILED');
  assert.match(st.error, /実績が3月分不足/, 'C-1: 前の版の行は数えない');

  // 2. B-2: 締まった 7・8 月だけを 6/19 の回で測る
  P.ok('EVAL.REPORT');
  const ev = evalByYm(P.rows);
  assert.equal(P.rows('EVAL_LOG').length, 30, 'EVAL_LOG の行は消さない（7・8 月の行は同じ行で書き直す）');
  const cur = P.rows('EVAL_LOG').filter((r) => r.evaluation_policy_version === POLICY);
  assert.deepEqual([...new Set(cur.map((r) => ymOf(r.target_month, r._types.charAt(3))))].sort(), ['2026/07', '2026/08'], '測ったのは締まった 7・8 月だけ');
  cur.forEach((r) => assert.equal(Number(r.pred), PRE * { nega: 0.9, neutral: 1, posi: 1.1 }[r.scenario], '月が始まる前の最後の回（6/19）。10/02 の回は新しくても使わない'));
  ['2026/04', '2026/09', '2026/10', '2026/11', '2027/03'].forEach((ym) => assert.ok(ev[ym].every((r) => r.evaluation_policy_version === 'policy-2026H1-v2'),
    ym + ': 前の版の行はそのまま残る（読む側が使わない）'));
  const cmp = cmpByYm(P.rows);
  near(cmp['2026/07'].ape_p50, 0.1, '7 月の外れ'); near(cmp['2026/07'].forecast_total_p50, PRE, '7 月の予測は 6/19 の回');
  ['2026/09', '2026/10', '2026/11', '2027/03'].forEach((ym) => {
    assert.equal(cmp[ym].ape_p50, '', ym + ': 締まっていない月は誤差を出さない');
    assert.equal(cmp[ym].range_outside_flag, '', ym);
    assert.equal(cmp[ym].note_for_investigation, '実績が締まっていない月（検証しない）', ym);
    assert.equal(Number(cmp[ym].actual_total), ACT[ym], ym + ': 実績の列は取り込んだ値のまま');
  });
  ['2026/04', '2026/05', '2026/06'].forEach((ym) => {
    assert.equal(cmp[ym].forecast_total_p50, '', ym + ': 月が始まる前の予測が無い（6/19・10/02 とも月が始まった後）');
    assert.equal(cmp[ym].note_for_investigation, '月が始まる前の予測が無い月（検証しない）', ym);
  });
  near(cmp['2026/11'].forecast_total_p50, LATE, '11 月の予測は 10/02 の回（11 月が始まる前）');
  const log = P.rows('RUN_LOG').filter((r) => r.function_name === 'updatePhase1EvaluationReport');
  assert.equal(log.length, 1);
  assert.equal(log[0].error_summary, POLICY + ': scored_months=2; open_months=5; no_pre_month_forecast=3', '測った月・締まっていない月・前の予測が無い月の数');
  // B-2 の後の自動の B-5: 2 か月では学ばない
  const learn = P.rows('PROCESS_STATUS').find((r) => r.step_key === 'learn_status');
  assert.equal(learn.error_summary, 'skipped:insufficient_eval_months');
  assert.equal(Number(P.cal().bias_correction_factor), 1);
  assert.equal(P.rows('CALIBRATION_HISTORY').length, 0);

  // 3. B-2 の後: B-3・B-4 は 7・8 月だけ
  P.ok('EVAL.DASHBOARD');
  near(P.metric('annual_abs_error_rate_latest'), 0.1, 'B-3: 7・8 月だけ（予測 2200 / 実績 2000）');
  near(P.metric('half1_wape_latest'), 0.1, 'B-3 上期');
  assert.equal(P.metric('half2_wape_latest'), '', 'B-3 下期は締まった月が無い');
  near(P.metric('monthly_secondary_metric_latest'), 0.1, 'B-3 月の外れの平均');
  assert.equal(Number(P.metric('range_outside_count_latest')), 0);
  P.ok('EVAL.INSIGHTS');
  const ins = P.rows('EVAL_INSIGHTS');
  assert.deepEqual(ins.map((r) => ymOf(r.target_month, r._types.charAt(2))), ['2026/11', '2026/07', '2026/08'], 'B-4: 7・8 月の行を足す。前からある 11 月の行はそのまま');
  assert.equal(JSON.stringify(ins[0]), JSON.stringify(JSON.parse(insightsBefore)[0]), '11 月の行は書き換えない');
  ins.slice(1).forEach((r) => { near(r.pred_p50, PRE, 'B-4 の予測は 6/19 の回'); near(r.actual_total, 1000, 'B-4 の実績'); });
  assert.deepEqual(P.shownInsights(), [['2026/07', 3], ['2026/08', 4]], '画面には測った 7・8 月だけ（行の番号はシートのまま）');
  st = P.runAction('REVIEW.GENERATE');
  assert.equal(st.status, 'FAILED');
  assert.match(st.error, /実績が1月分不足/, 'C-1: 締まった月で測った月は 2 か月だけ');
}

// ---- 6 日後に取り込んだ計画（9 月も締まった。7〜9 月を 6/19 の回で測る） ----
{
  const P = setUpPlan(jst(2026, 10, 6, 9));
  P.ok('EVAL.REPORT');
  const cur = P.rows('EVAL_LOG').filter((r) => r.evaluation_policy_version === POLICY);
  assert.deepEqual([...new Set(cur.map((r) => ymOf(r.target_month, r._types.charAt(3))))].sort(), ['2026/07', '2026/08', '2026/09'], '9 月も測る');
  assert.equal(P.rows('EVAL_LOG').length, 30, '行は消さない');
  assert.equal(P.rows('RUN_LOG').find((r) => r.function_name === 'updatePhase1EvaluationReport').error_summary,
    POLICY + ': scored_months=3; open_months=4; no_pre_month_forecast=3');

  // 4. B-2 の後の自動の B-5: 締まった 7〜9 月を、6/19 の予測（補正なし・10% 多い）で学ぶ。期待値は設計の式から（エンジンとは別に計算）
  const w = [0, 1, 2].map((i) => Math.pow(0.5, i / 4));   // 新しい月（9 月）から
  const S = w.reduce((a, b) => a + b, 0);
  const e = (PRE - 1000) / 1000;
  const factor = 1 - (e * S) / (S + 3);
  const c = P.cal();
  near(c.bias_correction_factor, factor, 'B-5 の係数（7〜9 月・6/19 の予測）');
  assert.ok(Number(c.bias_correction_factor) < 1, '6/19 の予測は多すぎたので下げる（10/02 の回 900 なら上げていた）');
  assert.match(c.note, /\(n=3, raw, corrected=0\)$/, '3 か月だけで学んだ（前の版の行・締まっていない月は入らない）');
  const mb = JSON.parse(c.residual_month_bias_json);
  assert.deepEqual(Object.keys(mb).sort(), ['7', '8', '9'], '暦月の偏りも 7〜9 月だけ');
  ['9', '8', '7'].forEach((m, i) => near(mb[m], Number((-(w[i] * e) / (w[i] + 2)).toFixed(4)), m + ' 月の偏り'));
  assert.deepEqual(P.rows('CALIBRATION_HISTORY').map((r) => [r.review_id, r.factor_name, r.old_value]),
    [['AUTO-MONTHLY', 'bias_correction_factor', '1'], ['AUTO-MONTHLY', 'residual_month_bias_json', '{}']], '履歴に前の値を残す');

  // 5. C-1: 締まった 7〜9 月を、6/19 の回の押しで数える（10/02 の回は外れの向きなので、使えば当たりは 0）
  P.ok('REVIEW.GENERATE');
  const evd = P.rows('RELIABILITY_EVIDENCE');
  assert.deepEqual(evd.map((r) => [r.source_type, r.source_key, r.quarter_label, ymOf(r.quarter_end_month, r._types.charAt(4)), r.n, r.hit]),
    [['opinion', '鷹野', 'FY2026-Q2', '2026/09', '3', '3']], '人の押し: 6/19 の回で 3 か月とも当たり');
  const ai = P.rows('QUARTERLY_REVIEW_LOG').filter((r) => r.target_field === 'ai_weight_override');
  assert.equal(ai.length, 1);
  near(ai[0].proposed_value, 0.0008 * 1.5, 'AI: 6/19 の回で一致 100%・効き 0.3% → 1.5 倍の案（10/02 の回を混ぜると 0.7 倍・10/02 だけなら 0.5 倍）');
  assert.match(ai[0].rationale, /AI方向一致率=100\.0%/);
  assert.ok(P.rows('QUARTERLY_REVIEW_LOG').every((r) => r.quarter_label === 'FY2026-Q2'), '提案の四半期は締まった 7〜9 月');

  // B-3・B-4 は 7〜9 月
  P.ok('EVAL.DASHBOARD');
  near(P.metric('annual_abs_error_rate_latest'), 0.1, 'B-3: 7〜9 月（予測 3300 / 実績 3000）');
  P.ok('EVAL.INSIGHTS');
  assert.deepEqual(P.rows('EVAL_INSIGHTS').map((r) => ymOf(r.target_month, r._types.charAt(2))), ['2026/11', '2026/07', '2026/08', '2026/09']);
}

console.log('app-legacy-closed-months: all tests passed');
