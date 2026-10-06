#!/usr/bin/env node
/**
 * app-insights.test.mjs — 責任者の 3 つの画面の材料（分析・人の学び・AI の学び）。読むだけ。
 * 計算用の表をまとめて読む（appEngAll_）・月の日付と文字列の食い違い・振り返りの重複・ベータ分布の区間・Wilson の区間・空のとき。
 *
 *   node app/tests/app-insights.test.mjs
 */
import assert from 'node:assert/strict';
import { setUpEnv, STATS, OWNER, MEMBER, OUTSIDER, J } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const close = (a, b, eps = 1e-9, msg = '') => assert.ok(Math.abs(a - b) < eps, `${msg} ${a} ≈ ${b}`);
const env = setUpEnv();
const H = env.run('(() => { const o = {}; Object.keys(APP_ENGINE_SHEETS).forEach(k => { o[k] = APP_ENGINE_SHEETS[k].header || null; }); return o; })()');
const row = (sheet, o) => H[sheet].map((h) => (o[h] === undefined ? '' : o[h]));
const config = (client, fy) => ({ values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', client], ['[必須] 予測年度FY（YYYY）', fy], ['[必須] 担当者（カンマ区切り）', '鷹野']] });
/** 検証の記録: [月, 実績, P50, 幅]（P10 = P50 × (1 − 幅)・P90 = P50 × (1 + 幅)） */
const evalLog = (client, months) => [H.EVAL_LOG].concat(...months.map(([ym, act, p50, w], i) => [['nega', p50 * (1 - w)], ['neutral', p50], ['posi', p50 * (1 + w)]]
  .map(([sc, p]) => row('EVAL_LOG', { eval_id: 'E' + i + sc, evaluated_at: D(2026, 9, 1), client, target_month: ym, scenario: sc, pred: p, actual: act, constraint_relevant_flag: 1 }))));
const impact = (client, runAt, ym, quant, src = 'forecast_open') => row('AI_IMPACT_HISTORY', { run_id: 'R' + runAt.getTime(), run_at: runAt, client, target_month: ym, k_ai: 1,
  pred_p50_quant_only: quant, forecast_source: src });
const push = (client, runAt, ym, type, key, dir, src = 'forecast_open') => row('SUBJECTIVE_IMPACT_HISTORY', { run_id: 'R' + runAt.getTime(), run_at: runAt, client, target_month: ym,
  source_type: type, source_key: key, push_step: dir * 0.05, push_direction: dir, applied_reliability_r: 1, forecast_source: src });
const research = (client, asOf, topic, rowType, dir, score) => row('AI_RESEARCH_STRUCTURED', { client, as_of_date: asOf, topic, row_type: rowType, direction: dir, impact_score: 3,
  confidence: 0.7, blended_score: score, relative_position_label: '中位' });

// ---- 甲製薬 FY2026: SUBJECTIVE_IMPACT_HISTORY の月が日付（旧来の書き方で、計算用ブックが日付に変えたもの）----
const r1 = D(2026, 3, 1), r2 = D(2026, 3, 15), r3 = D(2026, 3, 20);
const output = [['FY2026 売上予測（甲製薬）']];
for (let r = 2; r <= 25; r++) output.push([]);
output.push(['年度合計（予測）', 900, 1000, 1100], [], []);
for (let i = 0; i < 12; i++) { const m = 4 + i; output.push([(m > 12 ? 2027 : 2026) + '/' + String(m > 12 ? m - 12 : m).padStart(2, '0'), 90, 100, 110, '', '', '', 100, 10]); }
const bookA = env.makeBook('甲', {
  CONFIG: config('甲製薬', 2026),
  OUTPUT: { values: output },
  EVAL_LOG: { values: evalLog('甲製薬', [['2026/04', 1250, 1000, 0.05], ['2026/05', 900, 1000, 0.05], ['2026/06', 1000, 1000, 0.05], ['2026/07', 1100, 1050, 0.05]]), formats: { D: '@' } },
  EVAL_COMPARE_MONTHLY: { values: [H.EVAL_COMPARE_MONTHLY].concat([['2026/04', 1250], ['2026/05', 900], ['2026/06', 1000], ['2026/07', 1100]]
    .map(([ym, a]) => row('EVAL_COMPARE_MONTHLY', { target_month: ym, actual_total: a, forecast_total_p50: 1000, ape_p50: Math.abs(1000 - a) / a, range_outside_flag: 1 }))), formats: { A: '@' } },
  AI_IMPACT_HISTORY: { values: [H.AI_IMPACT_HISTORY,
    impact('甲製薬', r1, '2026/04', 1300), impact('甲製薬', r1, '2026/05', 950),
    impact('甲製薬', r2, '2026/04', 1000), impact('甲製薬', r2, '2026/05', 950), impact('甲製薬', r2, '2026/06', 1000),
    impact('甲製薬', r3, '2026/05', 800, 'actual_locked')] },
  SUBJECTIVE_IMPACT_HISTORY: { values: [H.SUBJECTIVE_IMPACT_HISTORY,
    push('甲製薬', r1, D(2026, 4), 'opinion', '鷹野', -1),
    push('甲製薬', r2, D(2026, 4), 'opinion', '鷹野', 1), push('甲製薬', r2, D(2026, 5), 'opinion', '鷹野', -1), push('甲製薬', r2, D(2026, 6), 'opinion', '鷹野', 1),
    push('甲製薬', r2, D(2026, 4), 'factor_product', '佐藤', -1), push('甲製薬', r2, D(2026, 5), 'factor_product', '佐藤', 1),
    push('甲製薬', r2, D(2026, 4), 'ai_topic', 'Market', 1), push('甲製薬', r2, D(2026, 5), 'ai_topic', 'Market', 1),
    push('甲製薬', r3, D(2026, 5), 'opinion', '鷹野', 1, 'actual_locked'),
    push('甲製薬', r2, D(2026, 10), 'opinion', '鷹野', 1)] },
  POOL_PRIOR: { values: [H.POOL_PRIOR, ['reliability:opinion', 'reliability_r', 1.2, 10, 3, D(2026, 9, 1), 'owner', '']] },
  SOURCE_RELIABILITY: { values: [H.SOURCE_RELIABILITY, ['甲製薬', 'opinion', '鷹野', 1.1, '', 'FY2025-Q3', D(2026, 1, 12), 'owner', 'applied via R1']] },
  EVAL_INSIGHTS: { values: [H.EVAL_INSIGHTS].concat([
    { evaluated_at: D(2026, 5, 10), target_month: D(2026, 4), actual_total: 1250, pred_p50: 1000, cause_hypothesis: '大型案件の前倒し', cause_bucket: 'over_forecast',
      action_type: '入力を修正', next_cycle_reflection: '見解を見直す', owner: '鷹野', status: 'in_progress', range_breach: 0 },
    { evaluated_at: D(2026, 6, 10), target_month: D(2026, 4), actual_total: 1250, pred_p50: 1000, cause_bucket: 'over_forecast', action_type: 'update',
      next_cycle_reflection: '次回サイクルで前提更新を反映', status: 'open', range_breach: 0 },
    { evaluated_at: D(2026, 6, 10), target_month: D(2026, 5), actual_total: 900, pred_p50: 1000, cause_bucket: 'under_forecast', action_type: 'update', status: 'open', range_breach: 1 },
    { evaluated_at: D(2026, 6, 10), target_month: D(2026, 6), actual_total: 1000, pred_p50: 1000, cause_bucket: 'under_forecast', action_type: 'keep', status: 'monitoring', range_breach: 0 },
    { evaluated_at: D(2026, 6, 10), target_month: '2025/05', actual_total: 800, pred_p50: 1000, cause_hypothesis: '失注', cause_bucket: 'under_forecast', action_type: 'update',
      owner: '佐藤', status: 'done', range_breach: 0 },
  ].map((o) => row('EVAL_INSIGHTS', Object.assign({ client: '甲製薬' }, o)))) },
  QUARTERLY_REVIEW_LOG: { values: [H.QUARTERLY_REVIEW_LOG].concat([
    { review_id: 'R1', proposal_id: 'P-B-1', reviewed_at: D(2026, 1, 10), quarter_label: 'FY2025-Q3', phase: 'B', target_field: 'reliability:opinion:鷹野', current_value: '1',
      proposed_value: '1.1', confidence: '中', rationale: '的中率=67% / n=3', approval_status: '承認', approval_decided_at: D(2026, 1, 12), approval_decided_by: OWNER, applied: 1, applied_at: D(2026, 1, 12) },
    { review_id: 'R1', proposal_id: 'P-A-1', reviewed_at: D(2026, 1, 10), quarter_label: 'FY2025-Q3', phase: 'A', target_field: 'ai_weight_override', current_value: '0.5',
      proposed_value: '0.35', confidence: '高', approval_status: '却下', approval_decided_at: D(2026, 1, 12), approval_decided_by: OWNER, applied: 0 },
    { review_id: 'R2', proposal_id: 'P-B-2', reviewed_at: D(2026, 7, 10), quarter_label: 'FY2026-Q1', phase: 'B', target_field: 'reliability:factor_product:佐藤', current_value: '1',
      proposed_value: '0.8', confidence: '中', approval_status: '保留', applied: 0 },
  ].map((o) => row('QUARTERLY_REVIEW_LOG', Object.assign({ client: '甲製薬' }, o)))) },
  QUARTERLY_REVIEW: { values: [['【四半期レビュー: FY2026-Q1】'], ['検証期間: 2026/04 〜 2026/06'], [], [], ['承認列を入力後、適用してください。'], [],
    ['提案ID', '対象', '現在値', '提案値', '自信度', '根拠', '影響見積もり', '承認列', 'ロールバック'],
    ['P-B-2', 'reliability:factor_product:佐藤', '1', '0.8', '中', '的中率=0% / n=2', '…', '承認', '…', 'R2']] },
  CALIBRATION_HISTORY: { values: [H.CALIBRATION_HISTORY].concat([
    { change_id: 'C1', changed_at: D(2026, 5, 10), quarter_label: 'FY2026-Q1', review_id: 'AUTO-MONTHLY', factor_name: 'bias_correction_factor', old_value: '1', new_value: '0.95' },
    { change_id: 'C2', changed_at: D(2026, 1, 12), quarter_label: 'FY2025-Q3', review_id: 'R1', factor_name: 'reliability:opinion:鷹野', old_value: '1', new_value: '1.1' },
    { change_id: 'C3', changed_at: D(2026, 6, 10), quarter_label: 'FY2026-Q1', review_id: 'AUTO-MONTHLY', factor_name: 'residual_month_bias_json', old_value: '', new_value: '{"04":-0.05}' },
  ].map((o) => row('CALIBRATION_HISTORY', Object.assign({ client: '甲製薬', changed_by: 'owner' }, o)))) },
  CALIBRATION_STATE: { values: [H.CALIBRATION_STATE, ['甲製薬', D(2026, 6, 10), 'owner', '', '', '', 0.75, '', '{}', '', '', 1, '']] },
  AI_RESEARCH_STRUCTURED: { values: [H.AI_RESEARCH_STRUCTURED, research('甲製薬', D(2026, 8, 1), 'Market', 'event', 'down', -30),
    research('甲製薬', D(2026, 9, 1), 'Market', 'event', 'up', 10), research('甲製薬', D(2026, 9, 1), 'Market', 'benchmark', 'positive', 14),
    research('甲製薬', D(2026, 9, 1), 'Competitor', 'event', '低下', -5)] },
});

// ---- 乙製薬 FY2026: 月は文字列（'@'）----
const bookB = env.makeBook('乙', {
  CONFIG: config('乙製薬', 2026),
  EVAL_LOG: { values: evalLog('乙製薬', [['2026/04', 500, 520, 0.01], ['2026/05', 600, 560, 0.01], ['2026/06', 650, 600, 0.01], ['2026/07', 700, 680, 0.01],
    ['2026/08', 720, 700, 0.01], ['2026/09', 740, 760, 0.01]]), formats: { D: '@' } },
  AI_IMPACT_HISTORY: { values: [H.AI_IMPACT_HISTORY, impact('乙製薬', r2, '2026/04', 450), impact('乙製薬', r2, '2026/05', 650), impact('乙製薬', r2, '2026/07', 600)] },
  SUBJECTIVE_IMPACT_HISTORY: { values: [H.SUBJECTIVE_IMPACT_HISTORY,
    push('乙製薬', r2, '2026/04', 'opinion', '鷹野', 1), push('乙製薬', r2, '2026/05', 'opinion', '鷹野', 1), push('乙製薬', r2, '2026/07', 'opinion', '鷹野', 1),
    push('乙製薬', r2, '2026/04', 'opinion', '田中', -1)], formats: { D: '@' } },
  POOL_PRIOR: { values: [H.POOL_PRIOR, ['reliability:opinion', 'reliability_r', 0.8, 20, 3, D(2026, 3, 1), 'owner', '']] },   // 古い方（使わない）
  SOURCE_RELIABILITY: { values: [H.SOURCE_RELIABILITY, ['乙製薬', 'opinion', '鷹野', 0.9, '', 'FY2025-Q3', D(2026, 1, 12), 'owner', '']] },
  QUARTERLY_REVIEW_LOG: { values: [H.QUARTERLY_REVIEW_LOG, row('QUARTERLY_REVIEW_LOG', { review_id: 'R3', proposal_id: 'P-A-3', reviewed_at: D(2026, 8, 1), client: '乙製薬',
    quarter_label: 'FY2026-Q1', phase: 'A', target_field: 'ai_weight_override', current_value: '0.5', proposed_value: '0.75', confidence: '中', approval_status: '保留', applied: 0 })] },
  CALIBRATION_STATE: { values: [H.CALIBRATION_STATE, ['乙製薬', D(2026, 6, 10), 'owner', '', '', '', 1, '', '{}', '', '', 1, '']] },
  AI_RESEARCH_STRUCTURED: { values: [H.AI_RESEARCH_STRUCTURED, research('乙製薬', D(2026, 9, 15), 'Market', 'event', 'up', 20)] },
});
// ---- 丙製薬 FY2025: 記録が何も無い計画 ----
const bookC = env.makeBook('丙', { CONFIG: config('丙製薬', 2025) });
const A = env.seedPlan(bookA), B = env.seedPlan(bookB), C = env.seedPlan(bookC);

// 新アプリの予測の実行（14 回の完了と 1 回の失敗）と、公式版
env.run(`appInsertRows_('FORECAST_RUNS', __runs)`, { __runs: Array.from({ length: 15 }, (_, i) => ({ run_id: 'FR-' + String(i).padStart(2, '0'), plan_id: A,
  status: i === 14 ? 'ERROR' : 'DONE', annual_p10: 900 + i, annual_p50: 1000 + i, annual_p90: 1100 + i,
  started_at: `2026-09-${String(i + 1).padStart(2, '0')}T09:00:00+0900`, finished_at: `2026-09-${String(i + 1).padStart(2, '0')}T10:00:00+0900` })) });
env.run(`appInsertRows_('PLAN_VERSIONS', __v)`, { __v: [{ version_id: 'V1', plan_id: A, version_no: 1, state: 'APPROVED', annual_p10: 900, annual_p50: 1000, annual_p90: 1100,
  budget_adopted: 1200, budget_uplift: 120, budget_final: 1320, row_version: 1 }] });

// 日付に変わった月（旧来の不具合の元）と、文字列の月の両方がある
const types = (id) => env.table('ENG_SUBJECTIVE_IMPACT_HISTORY').filter((r) => r.plan_id === id).map((r) => r._types.charAt(3));
assert.ok(types(A).every((t) => t === 'd') && types(B).every((t) => t === 's'), '甲は日付・乙は文字列の月');

// ==== 1. 分析: 年度のメーカーを横に並べる ====
{
  const x = env.call('apiCrossMaker(__in)', { __in: {} });
  assert.equal(x.fy, '2026', '既定は今の年度');
  assert.deepEqual(x.fys, ['2025', '2026']);
  assert.deepEqual(x.plans.map((p) => p.clientName).sort(), ['乙製薬', '甲製薬']);
  const a = x.plans.find((p) => p.planId === A), b = x.plans.find((p) => p.planId === B);
  // 計画の一覧の項目はそのまま渡す（着地の見込み・暫定実績など。別の作業で足す項目も）
  const port = env.call('apiPortfolio()').plans.find((p) => p.planId === A);
  for (const k of Object.keys(port)) assert.deepEqual(a[k], port[k], '一覧の項目をそのまま: ' + k);
  assert.equal(a.p50, 1013, '一番新しい予測');
  assert.equal(a.budget, 1320);
  assert.equal(a.actualYtd, port.actualYtd, '暫定実績は計画の一覧と同じ（締まった月だけ）');
  // 予測の改訂: 完了した回だけ、古い順に最近の 12 回
  assert.equal(a.revisions.length, 12);
  assert.deepEqual([a.revisions[0].p50, a.revisions[11].p50], [1002, 1013]);
  assert.ok(a.revisions.every((r, i) => i === 0 || r.at > a.revisions[i - 1].at), '古い順');
  assert.deepEqual(b.revisions, []);
  // 精度（締まった後の予測の月は除く）
  assert.equal(a.accuracy.n, 3);
  assert.equal(a.accuracy.leaks, 1);
  close(a.accuracy.mape, (0.2 + 100 / 900 + 50 / 1100) / 3, 1e-9, 'MAPE');
  close(a.accuracy.coverage, 1 / 3, 1e-9, 'P10〜P90 に入った割合');
  assert.equal(b.accuracy.coverage, 0);
  // AI の話題: 一番新しい回だけ。向きはそろえる
  assert.equal(a.topics.length, 3);
  assert.deepEqual(a.topics.map((t) => [t.topic, t.direction, t.score]), [['Market', 'up', 10], ['Market', 'up', 14], ['Competitor', 'down', -5]]);
  assert.equal(a.topics[0].asOf, '2026/09');
  // 市場のまとめ: 話題ごとに、上向き・下向きのメーカーの数と点数の平均（メーカーごとの平均の平均）
  assert.deepEqual(x.market.map((m) => [m.topic, m.makers, m.up, m.down, m.meanScore]), [['Market', 2, 2, 0, 16], ['Competitor', 1, 0, 1, -5]]);
  // 合計（比は予算のある計画だけ）
  assert.equal(x.totals.plans, 2);
  assert.equal(x.totals.budget, 1320);
  assert.equal(x.totals.budgetPlans, 1);
  assert.equal(x.totals.p50, 1013 + (b.p50 || 0));
  close(x.totals.ratioP50, 1013 / 1320);
  const landed = [a, b].filter((p) => typeof p.landing === 'number');
  assert.equal(x.totals.landing, landed.length ? landed.reduce((s, p) => s + p.landing, 0) : null, '着地の合計は着地のある計画だけ（霧で数字の無い計画は入れない）');
  // 年度を選ぶ: 記録の無い計画でも形はそろう
  const y = env.call('apiCrossMaker(__in)', { __in: { fy: 2025 } });
  assert.equal(y.fy, '2025');
  assert.equal(y.plans.length, 1);
  const c = y.plans[0];
  assert.deepEqual([c.planId, c.revisions, c.topics, c.accuracy.n, c.accuracy.mape], [C, [], [], 0, null]);
  assert.deepEqual(y.market, []);
  assert.equal(y.totals.budget, null);
  assert.equal(y.totals.ratioP50, null);
  assert.equal(env.call('apiCrossMaker(__in)', { __in: { fy: 'x' } }).fy, '2026', '年度でない値は既定にする');
}

// ==== 2. 人の学び ====
const people = env.call('apiPeopleLearning()');
{
  // 振り返り: 計画 × 月で 1 つ（B-4 の再実行で増えた行を除く）。人の書いた欄は残し、数字は新しい行から。向きは予測 − 実績から
  const L = people.lessons;
  assert.deepEqual(L.map((x) => x.ym), ['2025/05', '2026/04', '2026/05', '2026/06'], '誤差の大きい順');
  const apr = L[1];
  assert.equal(apr.duplicates, 1);
  assert.deepEqual([apr.hypothesis, apr.owner, apr.status, apr.statusLabel, apr.actionType, apr.human], ['大型案件の前倒し', '鷹野', 'in_progress', '対応中', '入力を修正', true]);
  assert.deepEqual([apr.direction, apr.directionLabel], ['under', '予測が低すぎた'], 'cause_bucket（over_forecast・逆向き）は使わない');
  close(apr.err, (1000 - 1250) / 1250);
  assert.equal(apr.miss, true);
  const may = L[2];
  assert.deepEqual([may.direction, may.range, may.miss, may.statusLabel, may.actionLabel, may.human], ['over', true, true, '未着手', '前提を更新', false]);
  assert.deepEqual([L[3].miss, L[3].direction, L[3].statusLabel], [false, 'exact', '見守り']);
  assert.deepEqual(people.summary, { months: 4, misses: 3, withNotes: 2, duplicatesRemoved: 1 });
  assert.deepEqual(people.openActions.map((x) => [x.ym, x.statusLabel]), [['2026/04', '対応中'], ['2026/05', '未着手']], '済み・見守りは出さない');
  assert.equal(people.causes.reduce((s, c) => s + c.n, 0), 3, '外れた月だけ数える');
  assert.ok(people.causes.some((c) => c.direction === 'under' && c.actionType === '入力を修正' && !c.range));
  assert.ok(people.causes.some((c) => c.direction === 'over' && c.range && c.actionType === 'update'));
  assert.deepEqual(people.repeats, [{ clientName: '甲製薬', month: '05', fys: ['2025', '2026'], n: 2 }], '違う年度の同じ月にまた外れた');
}
{
  // 当たり: 日付の月（甲）も文字列の月（乙）も数える。予測の対象だった月（forecast_open）の一番新しい回だけ。全部の評価できた月で
  const S = people.scoreboard;
  const get = (type, key) => S.find((s) => s.type === type && s.key === key);
  const tk = get('opinion', '鷹野');
  assert.deepEqual([tk.n, tk.hit], [5, 4], '甲 2/2 + 乙 2/3');
  assert.deepEqual(tk.plans.map((p) => [p.clientName, p.n, p.hit, p.appliedR]).sort(), [['乙製薬', 3, 2, 0.9], ['甲製薬', 2, 2, 1.1]]);
  close(tk.appliedR, 1.0, 1e-9, '今の信頼度（計画の平均）');
  // 事前分布: 一番新しい POOL_PRIOR（μ = 1.2 / 2 = 0.6・強さ 10 → Beta(6, 4)）。事後 Beta(10, 5)
  assert.deepEqual([tk.prior.alpha0, tk.prior.beta0, tk.prior.from], [6, 4, 'pool']);
  assert.deepEqual([tk.alpha, tk.beta], [10, 5]);
  close(tk.postMean, 10 / 15);
  assert.ok(tk.ci80[0] < tk.postMean && tk.postMean < tk.ci80[1]);
  assert.deepEqual([get('opinion', '田中').n, get('opinion', '田中').hit], [1, 0]);
  const sato = get('factor_product', '佐藤');
  assert.deepEqual([sato.n, sato.hit, sato.alpha, sato.beta, sato.prior.from, sato.appliedR], [2, 0, 2, 4, 'default', null], '事前分布が無ければ Beta(2, 2)');
  assert.deepEqual([get('ai_topic', 'Market').n, get('ai_topic', 'Market').hit], [2, 1]);
  assert.equal(S.length, 4);
  assert.ok(S.every((s, i) => i === 0 || s.ci80[0] <= S[i - 1].ci80[0]), '80% の区間の下の端が高い順');
  assert.equal(S[0].key, '鷹野');
}
{
  // 判断の記録（新しい順）
  const Dd = people.decisions;
  assert.deepEqual(Dd.map((d) => d.proposalId), ['P-A-3', 'P-B-2', 'P-B-1', 'P-A-1']);
  const p1 = Dd.find((d) => d.proposalId === 'P-B-1');
  assert.deepEqual([p1.status, p1.applied, p1.decidedBy, p1.current, p1.proposed, p1.phaseLabel, p1.targetLabel],
    ['承認', true, 'owner', 1, 1.1, '情報源の信頼度', '見解「鷹野」の信頼度']);
  assert.deepEqual([Dd[1].targetLabel, Dd[1].status, Dd[1].applied, Dd[1].decidedAt], ['製品の入力「佐藤」の信頼度', '保留', false, '']);
  assert.equal(Dd.find((d) => d.proposalId === 'P-A-1').targetLabel, 'AI の重み');
  // 予算の上乗せの実績（公式版の予算と、年度の実績の合計）
  assert.equal(people.uplift.length, 1);
  const u = people.uplift[0];
  assert.deepEqual([u.planId, u.budgetFinal, u.complete], [A, 1320, false]);
}

// ==== 3. AI の学び ====
const ai = env.call('apiAiLearning()');
{
  // 信頼度の学び: 種類ごとに、四半期ごとに積み上げた事後（Beta(α0 + 当たり, β0 + 外れ)）
  assert.deepEqual(ai.curves.map((c) => c.type), ['factor_product', 'factor_client', 'opinion', 'ai_topic', 'vertex_forecast']);
  const op = ai.curves.find((c) => c.type === 'opinion');
  assert.deepEqual([op.prior.alpha0, op.prior.beta0, op.prior.from], [6, 4, 'pool']);
  assert.deepEqual(op.quarters.map((q) => [q.quarter, q.n, q.hit, q.cumN, q.alpha, q.beta]), [['FY2026-Q1', 5, 3, 5, 9, 6], ['FY2026-Q2', 1, 1, 6, 10, 6]]);
  assert.deepEqual([op.n, op.hit], [6, 4]);
  const width = (ci) => ci[1] - ci[0];
  assert.ok(width(op.quarters[1].ci80) < width(op.quarters[0].ci80) && width(op.quarters[0].ci80) < width(op.prior.ci80), '数えるほど区間が狭くなる');
  const fp = ai.curves.find((c) => c.type === 'factor_product');
  assert.deepEqual([fp.prior.alpha0, fp.prior.beta0, fp.prior.from, fp.quarters.length], [2, 2, 'default', 1]);
  assert.deepEqual(ai.curves.find((c) => c.type === 'vertex_forecast').quarters, []);
  // 全計画で縮めた偏り（3 か月以上ある計画）
  const sh = ai.shrinkage;
  assert.equal(sh.pooled, true);
  assert.deepEqual(sh.plans.map((p) => p.planId).sort(), [A, B].sort());
  const accA = J(env.run(`appAccuracyOf_('${A}')`));
  const pa = sh.plans.find((p) => p.planId === A);
  close(pa.se, Math.sqrt(Math.max(accA.errVar, 1e-4) / Math.max(1, accA.biasW)), 1e-12, '標準誤差');
  close(pa.bias, accA.bias);
  for (const p of sh.plans) assert.ok(p.shrunk >= Math.min(p.bias, sh.mu) - 1e-12 && p.shrunk <= Math.max(p.bias, sh.mu) + 1e-12, '自分の偏りと平均の間');
  close(sh.tau, Math.sqrt(sh.tau2));
  // 精度の推移（全計画の月ごと。締まった後の予測の月は除く）
  const T = ai.timeline;
  assert.deepEqual(T.map((t) => t.ym), ['2026/04', '2026/05', '2026/06', '2026/07', '2026/08', '2026/09']);
  const apes = { '2026/04': [0.2, 20 / 500], '2026/05': [100 / 900, 40 / 600], '2026/06': [50 / 650], '2026/07': [50 / 1100, 20 / 700] };
  close(T[0].mape, (0.2 + 0.04) / 2);
  assert.deepEqual([T[2].n, T[2].coverageN], [1, 1], '甲の 6 月は締まった後の予測');
  close(T[2].rolling3, [].concat(apes['2026/04'], apes['2026/05'], apes['2026/06']).reduce((a, b) => a + b) / 5, 1e-12, '3 か月の平均');
  assert.equal(T[2].rolling3N, 5);
  assert.deepEqual([T[3].coverage, T[3].coverageN, T[3].target], [0.5, 2, 0.8]);
  assert.ok(T[3].coverageCi80[0] < 0.5 && T[3].coverageCi80[1] > 0.5);
  assert.equal(T[0].coverageCi80[0], 0);
  // 補正の変化（古い順。月次の自動学習と四半期レビュー）
  const ca = ai.calibration.find((c) => c.planId === A);
  assert.equal(ca.factorNow, 0.75);
  assert.deepEqual(ca.path.map((p) => [p.factor, p.old, p.new, p.source, p.sourceLabel]), [
    ['reliability:opinion:鷹野', 1, 1.1, 'R1', '四半期レビュー'],
    ['bias_correction_factor', 1, 0.95, 'AUTO-MONTHLY', '月次の自動学習'],
    ['residual_month_bias_json', '', '{"04":-0.05}', 'AUTO-MONTHLY', '月次の自動学習']]);
  assert.equal(ca.path[1].factorLabel, '偏りの補正');
  assert.equal(ai.calibration.find((c) => c.planId === B).path.length, 0);
  assert.equal(ai.calibration.find((c) => c.planId === C), undefined, '補正の記録が無い計画は出さない');
  // 承認待ちの提案: 一番新しいレビューで、まだ適用していないもの。保存した判断も出す
  assert.deepEqual(ai.pending.map((p) => [p.planId, p.reviewId, p.proposals.map((x) => [x.proposalId, x.targetLabel, x.current, x.proposed, x.decision])]).sort(), [
    [A, 'R2', [['P-B-2', '製品の入力「佐藤」の信頼度', 1, 0.8, '承認']]],
    [B, 'R3', [['P-A-3', 'AI の重み', 0.5, 0.75, '']]]].sort());
  // 気になる点: 補正が下限に張りつく（甲）・P10〜P90 に入る割合が低い（乙: 6 か月で 0）
  const keys = ai.health.map((h) => [h.planId, h.key]);
  assert.ok(keys.some(([p, k]) => p === A && k === 'factor_clamp'));
  assert.ok(keys.some(([p, k]) => p === B && k === 'coverage_low'));
  assert.ok(!keys.some(([p, k]) => p === A && k.startsWith('coverage')), '月が少ない計画は見ない');
  assert.ok(ai.health.every((h) => ['shadow_worse', 'factor_clamp', 'coverage_low', 'coverage_high'].includes(h.key) && h.label));
  const hx = J(env.run(`appInsightHealth_([{ planId: 'P', clientName: 'x', fy: '2026' }], { P: { coverage: 1, coverageN: 12 } },
    { plans: { P: { mapeNow: 0.1, mapeShadow: 0.12 } } }, { P: 1.25 })`));
  assert.deepEqual(hx.map((h) => h.key), ['shadow_worse', 'factor_clamp', 'coverage_high']);
}

// ==== 4. 読み方: 計算用の表は 1 回ずつ。計画ごとに組み立て直さない。控えから返す ====
{
  for (const k of Object.keys(env.cache)) delete env.cache[k];
  for (const k of Object.keys(STATS)) STATS[k] = k === 'bySheet' ? {} : 0;
  const n = env.run(`(() => { APP_STORE_CACHE_ = {}; let n = 0; const orig = appEngTableObjects_;
    appEngTableObjects_ = function () { n++; return orig.apply(null, arguments); };
    try { appAiLearning_(); appPeopleLearning_(); appCrossMaker_(); } finally { appEngTableObjects_ = orig; }
    return n; })()`);
  assert.equal(n, 0, '表の形の計画は appEngTableObjects_ を使わない');
  for (const s of ['ENG_EVAL_LOG', 'ENG_SUBJECTIVE_IMPACT_HISTORY', 'ENG_AI_IMPACT_HISTORY', 'ENG_EVAL_INSIGHTS', 'ENG_QUARTERLY_REVIEW_LOG', 'ENG_POOL_PRIOR']) {
    assert.ok(STATS.bySheet[s].reads <= 2, s + ' は 1 回だけ読む（見出しの確認と中身）');
  }
  assert.equal(STATS.bySheet.ENG_FORMATS, undefined, '表示形式は読まない');
  // 画面からの呼び出しは控えから返す（データ本体を読まない）
  env.call('apiPeopleLearning()'); env.call('apiAiLearning()'); env.call('apiCrossMaker(__in)', { __in: {} });
  const r0 = STATS.reads;
  assert.deepEqual(env.call('apiPeopleLearning()'), people);
  env.call('apiAiLearning()'); env.call('apiCrossMaker(__in)', { __in: {} });
  assert.equal(STATS.reads, r0, '控えから返す');
  // 閲覧は社内全員（外の人は読めない）
  env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'M' })`);
  env.as(MEMBER);
  assert.equal(env.call('apiCrossMaker(__in)', { __in: {} }).plans.length, 2);
  assert.equal(env.call('apiPeopleLearning()').scoreboard.length, 4);
  assert.equal(env.call('apiAiLearning()').curves.length, 5);
  env.as(OUTSIDER);
  assert.throws(() => env.call('apiPeopleLearning()'), /権限がありません/);
  env.as(OWNER);
}

// ==== 5. まとめて読む表（appEngAll_）は、計画ごとに組み立てた表（appEngTableObjects_）と同じ。見出しの違う計画は行ごとに持った方から読む ====
{
  const e2 = setUpEnv();
  const H2 = H;
  const r = (sheet, o) => H2[sheet].map((h) => (o[h] === undefined ? '' : o[h]));
  const p1 = e2.seedPlan(e2.makeBook('一', {
    CONFIG: config('一製薬', 2026),
    EVAL_LOG: { values: [H2.EVAL_LOG, r('EVAL_LOG', { eval_id: 'E1', evaluated_at: D(2026, 9, 1), client: '一製薬', target_month: D(2026, 4), scenario: 'neutral', pred: 100, actual: 110, constraint_relevant_flag: 1 }),
      [], r('EVAL_LOG', { eval_id: 'E2', evaluated_at: D(2026, 9, 1), client: '一製薬', target_month: '2026/05', scenario: 'neutral', pred: -0, actual: 90, was_overridden: true })],
    formulas: { H2: '=1+1' } },
    SOURCE_RELIABILITY: { values: [H2.SOURCE_RELIABILITY, ['一製薬', 'opinion', '鷹野', 1.2, '', '', '', '', '']] },
  }));
  const renamed = H2.SOURCE_RELIABILITY.map((h) => (h === 'note' ? 'memo' : h));
  const p2 = e2.seedPlan(e2.makeBook('二', {
    CONFIG: config('二製薬', 2026),
    EVAL_LOG: { values: [H2.EVAL_LOG, r('EVAL_LOG', { eval_id: 'E1', evaluated_at: D(2026, 9, 1), client: '二製薬', target_month: '2026/04', scenario: 'neutral', pred: 200, actual: 210 })] },
    SOURCE_RELIABILITY: { values: [renamed, ['二製薬', 'opinion', '佐藤', 0.8, '', '', '', '', 'メモ']] },
  }));
  assert.equal(e2.table('ENG_SHEETS').find((s) => s.plan_id === p2 && s.sheet === 'SOURCE_RELIABILITY').mode, 'rows', '見出しが違う → 行ごと');
  const same = e2.run(`(() => { APP_STORE_CACHE_ = {}; const bad = [];
    ['EVAL_LOG', 'SOURCE_RELIABILITY', 'EVAL_INSIGHTS'].forEach(s => { const all = appEngAll_(s);
      ['${p1}', '${p2}'].forEach(id => { const one = appEngTableObjects_(id, [s])[s];
        if (JSON.stringify(all[id] || []) !== JSON.stringify(one)) bad.push(s + ' ' + id + ' ' + JSON.stringify(all[id]) + ' / ' + JSON.stringify(one)); }); });
    return bad; })()`);
  assert.deepEqual(J(same), []);
  const ev = J(e2.run(`appEngAll_('EVAL_LOG')['${p1}']`));
  assert.equal(ev.length, 2, '空の行は除く');
  assert.equal(ev[0].ape, '', '数式のセルは空');
  assert.equal(ev[1].was_overridden, true);
  assert.equal(J(e2.run(`appEngAll_('SOURCE_RELIABILITY')['${p2}']`))[0].memo, 'メモ', '行ごとに持った計画は、その見出しのまま');
  assert.deepEqual(J(e2.run(`appEngAll_('EVAL_LOG', [])`)), {});
  assert.deepEqual(Object.keys(J(e2.run(`appEngAll_('EVAL_LOG', ['${p2}'])`))), [p2], '計画を選べる');
  // 精度は、まとめて読んだ行から計算しても同じ
  assert.equal(e2.run(`JSON.stringify(appAccuracyOf_('${p1}'))`), e2.run(`JSON.stringify(appAccuracyOf_('${p1}', appEngAll_('EVAL_LOG')['${p1}']))`));
  // 読むのは表 1 回（表示形式や行ごとの表は読まない）
  for (const k of Object.keys(STATS)) STATS[k] = k === 'bySheet' ? {} : 0;
  e2.run(`(() => { APP_STORE_CACHE_ = {}; return appEngAll_('EVAL_LOG'); })()`);
  assert.ok(STATS.bySheet.ENG_EVAL_LOG.reads <= 2, '見出しの確認と中身');
  assert.equal(STATS.bySheet.ENG_FORMATS, undefined);
  assert.equal(STATS.bySheet.ENG_ROWS, undefined);
}

// ==== 6. ベータ分布の分位点・Wilson の区間・四半期 ====
{
  const q = (p, a, b) => env.run(`appBetaQuantile_(${p}, ${a}, ${b})`);
  close(q(0.1, 1, 1), 0.1, 1e-12, '一様分布');
  close(q(0.5, 2, 2), 0.5, 1e-12, '対称');
  close(q(0.1, 0.5, 0.5), Math.sin(0.05 * Math.PI) ** 2, 1e-10, '逆正弦分布（a, b が整数でない）');
  close(env.run('appBetaInc_(0.3, 1, 4)'), 1 - 0.7 ** 4, 1e-12, 'I_x(1, b) = 1 − (1 − x)^b');
  close(q(0.1, 2.5, 7.3), 1 - q(0.9, 7.3, 2.5), 1e-12, 'a と b を入れ替えると裏返し');
  // Beta(3, 7) の 10% 点を、密度の数値積分（シンプソン）で確かめる
  const x10 = q(0.1, 3, 7);
  const dens = (x) => x ** 2 * (1 - x) ** 6 / (2 * 720 / 362880);
  const N = 20000; let s = dens(0) + dens(x10);
  for (let i = 1; i < N; i++) s += (i % 2 ? 4 : 2) * dens((x10 * i) / N);
  close((s * x10) / N / 3, 0.1, 1e-8, '数値積分');
  // n が大きいときは正規分布に近い
  close(q(0.9, 300, 700), 0.3 + 1.2815515655446004 * Math.sqrt(300 * 700 / (1000 ** 2 * 1001)), 1e-3, '正規近似');
  const s1 = J(env.run('appBetaSummary_(10, 5)'));
  close(s1.mean, 2 / 3);
  assert.ok(s1.ci80[0] < s1.mean && s1.mean < s1.ci80[1]);
  // Wilson の区間（80%）
  const w = (k, n) => J(env.run(`appWilson_(${k}, ${n}, APP_INSIGHT_Z80)`));
  const z2 = 1.2815515655446004 ** 2;
  assert.deepEqual(w(0, 10)[0], 0);
  close(w(0, 10)[1], z2 / (10 + z2), 1e-12, 'k = 0 の上の端');
  close(w(8, 10)[0], 1 - w(2, 10)[1], 1e-12, '対称');
  close(w(8, 10)[1], 1 - w(2, 10)[0], 1e-12, '対称');
  assert.ok(w(8, 10)[0] < 0.8 && w(8, 10)[1] > 0.8);
  assert.ok(w(80, 100)[1] - w(80, 100)[0] < w(8, 10)[1] - w(8, 10)[0], '数が多いほど狭い');
  assert.equal(w(0, 0), null);
  assert.deepEqual(['2026/04', '2026/06', '2026/07', '2026/12', '2027/01', '2027/03'].map((m) => env.run(`appInsightQuarter_('${m}')`)),
    ['FY2026-Q1', 'FY2026-Q1', 'FY2026-Q2', 'FY2026-Q3', 'FY2026-Q4', 'FY2026-Q4'], '旧来の quarterLabelFromYm_ と同じ');
  assert.equal(env.run(`appInsightPrevYm_('2026/02', 2)`), '2025/12');
}

// ==== 7. 計画が無いとき: どの画面も空の形で返す ====
{
  const e3 = setUpEnv();
  const x = e3.call('apiCrossMaker(__in)', { __in: {} });
  assert.deepEqual([x.plans, x.fys, x.market, x.totals.plans, x.totals.budget, x.totals.p50, x.totals.ratioP50], [[], [], [], 0, null, null, null]);
  const p = e3.call('apiPeopleLearning()');
  assert.deepEqual([p.lessons, p.causes, p.openActions, p.repeats, p.scoreboard, p.decisions, p.uplift], [[], [], [], [], [], [], []]);
  assert.deepEqual(p.summary, { months: 0, misses: 0, withNotes: 0, duplicatesRemoved: 0 });
  const a = e3.call('apiAiLearning()');
  assert.equal(a.curves.length, 5);
  assert.ok(a.curves.every((c) => c.prior.alpha0 === 2 && c.prior.beta0 === 2 && c.prior.from === 'default' && c.quarters.length === 0 && c.n === 0));
  close(a.curves[0].prior.mean, 0.5);
  assert.deepEqual([a.shrinkage.plans, a.shrinkage.pooled, a.shrinkage.mu, a.shrinkage.tau, a.timeline, a.calibration, a.pending, a.health],
    [[], false, 0, 0, [], [], [], []]);
}

console.log('app-insights: all tests passed');
