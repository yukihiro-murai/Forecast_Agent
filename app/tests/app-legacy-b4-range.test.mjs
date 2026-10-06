#!/usr/bin/env node
/*
 * app-legacy-b4-range.test.mjs — B-4（学習インサイトの更新）の記録の直し（2026-10-07 村井さん承認）を、本物の旧来の計算をモックの上で動かして確かめる。
 * 準備は app-legacy-records.test.mjs と同じ（5 月だけ実績が予測の幅の外。計算用ブックは Sheets と同じく 'yyyy/MM' を日付に変える）。
 *   1. 検証の表（EVAL_COMPARE_MONTHLY）の月が日付でも、幅の外の印を拾う（前は String() で照合して拾えず、幅の外の月が range_breach にならなかった）
 *   2. 前からある行のうち、人の記入の跡が無い行（B-4 が既定値を入れただけの行）は、今回の判定で対応・次回への反映・状態を書き直す
 *      （状態が B-4 の入れる組 update・open / keep・monitoring と違えば、人が選んだものとして残す）。
 *      人の記入の跡がある行は残す。原因の区分（cause_bucket）は画面から書かない機械の列なので、どちらの行も毎回書き直す（逆向きの過去の値も直る）
 *   3. 予測の数字に関わる表は変わらない
 *
 *   node app/tests/app-legacy-b4-range.test.mjs
 */
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { setUpEnv, uiHtml, OWNER } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const env = setUpEnv();
const H = env.run('(() => { const o = {}; Object.keys(APP_ENGINE_SHEETS).forEach(k => { o[k] = APP_ENGINE_SHEETS[k].header || null; }); return o; })()');
const row = (sheet, o) => H[sheet].map((h) => (o[h] === undefined ? '' : o[h]));
const CLIENT = 'テスト製薬';
/** 検証の月: [月, 実績, P50]。4 月は予測が低すぎた・5 月は高すぎた・6 月は少し低すぎた */
const MONTHS = [['2026/04', 1200, 1000], ['2026/05', 800, 1000], ['2026/06', 1050, 1000]];
/** 検証の記録（P10・P90 は ±5%）。月は旧来の B-2 と同じく書式なしテキストの列に書いた形 */
const evalLog = [H.EVAL_LOG].concat(...MONTHS.map(([ym, act, p50], i) => [['nega', p50 * 0.95], ['neutral', p50], ['posi', p50 * 1.05]]
  .map(([sc, p]) => row('EVAL_LOG', { eval_id: 'E' + i + sc, evaluated_at: D(2026, 9, 1), client: CLIENT, target_month: ym, scenario: sc, pred: p, actual: act,
    signed_error: p - act, abs_error: Math.abs(p - act), bias_direction: p > act ? 'over' : 'under', constraint_relevant_flag: sc === 'neutral' ? 1 : 0 }))));
/** 予測の記録の月は、旧来が書式を付けずに書くので、計算用ブックでは日付になっている（その形で置く） */
const run1 = D(2026, 3, 20);
const impact = (m, quant) => row('AI_IMPACT_HISTORY', { run_id: 'R1', run_at: run1, client: CLIENT, target_month: D(2026, m), k_ai: 1, ai_direction: 'flat',
  pred_p50: 1000, pred_p50_quant_only: quant, forecast_source: 'forecast_open' });
const push = (m, type, key, dir) => row('SUBJECTIVE_IMPACT_HISTORY', { run_id: 'R1', run_at: run1, client: CLIENT, target_month: D(2026, m), source_type: type,
  source_key: key, push_step: dir * 0.05, push_direction: dir, applied_reliability_r: 1, forecast_source: 'forecast_open' });
const book = env.makeBook('クライアント別売上予測', {
  CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', CLIENT], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野,佐藤']] },
  PROCESS_STATUS: { values: [H.PROCESS_STATUS, ['step4_status', D(2026, 9, 1), 'owner', 'success', CLIENT, 12, ''], ['step5_status', D(2026, 9, 1), 'owner', 'success', CLIENT, 9, '']] },
  RUN_LOG: { values: [H.RUN_LOG] },
  EVAL_LOG: { values: evalLog, formats: { D: '@' } },
  // 旧来の B-2 は月を書式なしで書くので日付になっている（5 月だけ幅の外の印を付ける）
  EVAL_COMPARE_MONTHLY: { values: [H.EVAL_COMPARE_MONTHLY].concat(MONTHS.map(([ym, act, p50]) => row('EVAL_COMPARE_MONTHLY', {
    target_month: D(Number(ym.slice(0, 4)), Number(ym.slice(5))), forecast_total: p50, actual_total: act, forecast_total_p50: p50, signed_error_p50: p50 - act,
    abs_error_p50: Math.abs(p50 - act), half_label: 'FY2026-H1', over_flag: p50 > act ? 1 : 0, under_flag: p50 < act ? 1 : 0, range_outside_flag: ym === '2026/05' ? 1 : 0 }))) },
  // 前からある行（直す前の B-4 が書いた形。月は日付）: 4 月は人が原因の仮説と状態を書いた行、5 月は既定値だけの行（向きも逆）
  EVAL_INSIGHTS: { values: [H.EVAL_INSIGHTS,
    row('EVAL_INSIGHTS', { evaluated_at: D(2026, 8, 1), client: CLIENT, target_month: D(2026, 4), actual_total: 1200, pred_p50: 1000, diff: 200, error_rate: 0.1667,
      diagnostic_type: 'monthly_diagnostic', cause_hypothesis: '大型案件の前倒し', cause_bucket: 'over_forecast', impacted_assumption: 'CONFIG:環境前提',
      action_type: '入力を修正', next_cycle_reflection: '見解を見直す', owner: '鷹野', status: 'in_progress', review_cycle: 'monthly_light' }),
    row('EVAL_INSIGHTS', { evaluated_at: D(2026, 8, 1), client: CLIENT, target_month: D(2026, 5), actual_total: 800, pred_p50: 1000, diff: -200, error_rate: -0.25,
      diagnostic_type: 'monthly_diagnostic', cause_bucket: 'under_forecast', impacted_assumption: 'CONFIG:環境前提',
      action_type: 'keep', next_cycle_reflection: '現行運用を継続', status: 'monitoring', review_cycle: 'monthly_light' }),
    // 6 月: 見守りの月に、人が状態だけ「未着手」（open）を選んだ行（B-4 が入れる組 keep・monitoring と違うので人の記入として残す）
    row('EVAL_INSIGHTS', { evaluated_at: D(2026, 8, 1), client: CLIENT, target_month: D(2026, 6), actual_total: 1050, pred_p50: 1000, diff: 50, error_rate: 0.0476,
      diagnostic_type: 'monthly_diagnostic', cause_bucket: 'over_forecast', impacted_assumption: 'CONFIG:環境前提',
      action_type: 'keep', next_cycle_reflection: '現行運用を継続', status: 'open', review_cycle: 'quarterly_full' })] },
  // 統計だけの予測はどの月も 1000。実績の向きは 4 月 上・5 月 下・6 月 上
  AI_IMPACT_HISTORY: { values: [H.AI_IMPACT_HISTORY, impact(4, 1000), impact(5, 1000), impact(6, 1000)] },
  // 鷹野（見解）: 3 か月とも実績の向きに押した。佐藤（製品）: 3 か月とも上に押した（当たり 2・外れ 1）
  SUBJECTIVE_IMPACT_HISTORY: { values: [H.SUBJECTIVE_IMPACT_HISTORY,
    push(4, 'opinion', '鷹野', 1), push(5, 'opinion', '鷹野', -1), push(6, 'opinion', '鷹野', 1),
    push(4, 'factor_product', '佐藤', 1), push(5, 'factor_product', '佐藤', 1), push(6, 'factor_product', '佐藤', 1)] },
  AI_SCORE_HISTORY: { values: [H.AI_SCORE_HISTORY] },
  CALIBRATION_STATE: { values: [H.CALIBRATION_STATE, [CLIENT, D(2026, 9, 1), 'owner', '', '', '[]', 1, '', '{}', '', '', 1, '']] },
  SOURCE_RELIABILITY: { values: [H.SOURCE_RELIABILITY] },
  POOL_PRIOR: { values: [H.POOL_PRIOR] },
  QUARTERLY_REVIEW: { values: [['']] },
  QUARTERLY_REVIEW_LOG: { values: [H.QUARTERLY_REVIEW_LOG] },
  RELIABILITY_EVIDENCE: { values: [H.RELIABILITY_EVIDENCE] },
});
const planId = env.seedPlan(book);
const engRows = (sheet) => env.table('ENG_' + sheet).filter((r) => r.plan_id === planId);
const ym = (iso) => { const d = new Date(iso); return d.getFullYear() + '/' + String(d.getMonth() + 1).padStart(2, '0'); };
const runAction = (action) => { const st = env.runJob('PLAN.RUN', { planId, action }); assert.equal(st.status, 'DONE', action + ': ' + st.error); };

// ==== B-4 を動かす ====
const FC_TABLES = ['FORECAST_SNAPSHOT', 'CALIBRATION_STATE', 'EVAL_LOG', 'EVAL_COMPARE_MONTHLY', 'SOURCE_RELIABILITY', 'POOL_PRIOR', 'AI_IMPACT_HISTORY'];
const snap = () => JSON.stringify(FC_TABLES.map((t) => env.table('ENG_' + t).filter((r) => r.plan_id === planId).map((r) => { const o = Object.assign({}, r); delete o.row_version; delete o.updated_at; return o; })));
const before = snap();
runAction('EVAL.INSIGHTS');
const cmp = engRows('EVAL_COMPARE_MONTHLY');
assert.deepEqual(cmp.map((r) => r._types.charAt(0)), ['d', 'd', 'd'], '前提: 計算用ブックでは検証の表の月が日付になる（Sheets と同じ）');
const ins = engRows('EVAL_INSIGHTS');
assert.equal(ins.length, 3, '前からある 4・5・6 月の行を上書きする（行は増えない）');
const by = {}; ins.forEach((r) => { by[ym(r.target_month)] = r; });
assert.deepEqual(Object.keys(by).sort(), ['2026/04', '2026/05', '2026/06']);

// 1. 幅の外の印
assert.equal(by['2026/05'].diagnostic_type, 'range_breach', '幅の外の 5 月は range_breach（直す前は monthly_diagnostic）');
assert.equal(Number(by['2026/05'].range_breach), 1, '5 月の range_breach の印');
assert.equal(Number(by['2026/05'].overforecast_breach), 1, '5 月は予測が実績を 25% 上回った（過大予測の制約を超える）');
assert.equal(by['2026/04'].diagnostic_type, 'monthly_diagnostic', '幅の中の月はそのまま');
assert.equal(Number(by['2026/06'].range_breach), 0);

// 2. 既定値だけの 5 月の行は、今回の判定で書き直す。人が書いた 4 月の行は残す。cause_bucket はどちらも書き直す
assert.deepEqual([by['2026/05'].cause_bucket, by['2026/05'].action_type, by['2026/05'].next_cycle_reflection, by['2026/05'].status],
  ['range_outside', 'update', '次回サイクルで前提更新を反映', 'open'], '既定値だけの行は書き直す（直す前は keep・現行運用を継続・monitoring のまま）');
assert.deepEqual([by['2026/04'].cause_hypothesis, by['2026/04'].action_type, by['2026/04'].next_cycle_reflection, by['2026/04'].owner, by['2026/04'].status],
  ['大型案件の前倒し', '入力を修正', '見解を見直す', '鷹野', 'in_progress'], '人が書いた行は残す');
assert.equal(by['2026/04'].cause_bucket, 'under_forecast', '4 月は実績 > 予測なので under_forecast（逆向きの過去の値 over_forecast を書き直す）');
assert.equal(by['2026/06'].cause_bucket, 'under_forecast', '6 月も cause_bucket は書き直す');
assert.deepEqual([by['2026/06'].action_type, by['2026/06'].status], ['keep', 'open'], '人が状態だけを選んだ行も残す（B-4 の組 keep・monitoring に戻さない）');

// 3. 予測の数字に関わる表は変わらない
assert.equal(snap(), before, '予測の表・補正・検証の表は B-4 で変わらない');
console.log('app-legacy-b4-range: all tests passed');
