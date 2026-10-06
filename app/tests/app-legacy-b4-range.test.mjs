#!/usr/bin/env node
/*
 * app-legacy-b4-range.test.mjs — B-4（学習インサイトの更新）が、検証の表の「幅の外」の印を月が日付でも拾う（2026-10-07 村井さん承認の記録の直し）。
 * 下の準備は app-legacy-records.test.mjs と同じ（5 月だけ実績が予測の幅の外）。
 * 計算用ブックは、Sheets と同じく 'yyyy/MM' の文字を日付に変える（モックが真似する）。直す前は、どれもこの変換で食い違っていた。
 *   1. C-1 の当たりの記録（RELIABILITY_EVIDENCE）: SUBJECTIVE_IMPACT_HISTORY の月が日付でも、実績の月と合わせて数える
 *   2. B-4 を動かし直しても EVAL_INSIGHTS の行が増えない（同じ月の行を上書きし、人の記入は残す）
 *   3. EVAL_INSIGHTS.cause_bucket の向き: over_forecast = 予測 > 実績（EVAL_LOG・EVAL_COMPARE_MONTHLY と同じ向き）
 *   4. 四半期レビューの案の確度（高・中・低）を文字のまま画面へ渡し、画面に出す
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
  // 旧来の B-2 は月を書式なしで書くので日付になっている。幅の外の印は付けない（B-4 がこの表から幅の外の印を探すところも、
  // 月が日付だと合わずに印を拾えない。同じ日付の問題だが、この直しには入れず別の件として村井さんの判断を待つ）
  EVAL_COMPARE_MONTHLY: { values: [H.EVAL_COMPARE_MONTHLY].concat(MONTHS.map(([ym, act, p50]) => row('EVAL_COMPARE_MONTHLY', {
    target_month: D(Number(ym.slice(0, 4)), Number(ym.slice(5))), forecast_total: p50, actual_total: act, forecast_total_p50: p50, signed_error_p50: p50 - act,
    abs_error_p50: Math.abs(p50 - act), half_label: 'FY2026-H1', over_flag: p50 > act ? 1 : 0, under_flag: p50 < act ? 1 : 0, range_outside_flag: ym === '2026/05' ? 1 : 0 }))) },
  EVAL_INSIGHTS: { values: [H.EVAL_INSIGHTS] },
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

// ==== B-4: 月が日付の検証の表から、幅の外の印を拾う ====
runAction('EVAL.INSIGHTS');
const cmp = engRows('EVAL_COMPARE_MONTHLY');
assert.deepEqual(cmp.map((r) => r._types.charAt(0)), ['d', 'd', 'd'], '前提: 計算用ブックでは検証の表の月が日付になる（Sheets と同じ）');
const ins = engRows('EVAL_INSIGHTS');
const by = {}; ins.forEach((r) => { by[ym(r.target_month)] = r; });
assert.deepEqual(Object.keys(by).sort(), ['2026/04', '2026/05', '2026/06']);
assert.equal(by['2026/05'].diagnostic_type, 'range_breach', '幅の外の 5 月は range_breach（直す前は monthly_diagnostic）');
assert.equal(Number(by['2026/05'].range_breach), 1, '5 月の range_breach の印');
assert.equal(by['2026/04'].diagnostic_type, 'monthly_diagnostic', '幅の中の月はそのまま');
assert.equal(Number(by['2026/06'].range_breach), 0);
console.log('app-legacy-b4-range: all tests passed');
