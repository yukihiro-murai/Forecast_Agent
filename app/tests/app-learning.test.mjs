/**
 * app-learning.test.mjs — 精度の推移・全計画で縮めた偏りの補正（影）・幅の較正・情報源の信頼度の事前分布（ベータ二項）。
 */
import assert from 'node:assert/strict';
import { setUpEnv, STATS } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const env = setUpEnv();
const H = env.run('(() => { const o = {}; Object.keys(APP_ENGINE_SHEETS).forEach(k => { o[k] = APP_ENGINE_SHEETS[k].header || null; }); return o; })()');
/** 12 か月の検証の記録: 予測は実績の over 倍。P10〜P90 は ±width。最後の月は締まった後の予測（予測 = 実績） */
function book(client, over, width, hit) {
  const ev = [H.EVAL_LOG];
  for (let i = 0; i < 12; i++) {
    const ym = '2025/' + String(i + 1).padStart(2, '0');
    const act = 1000 + i * 10 * (i % 3 === 0 ? -1 : 1);
    const p50 = i === 11 ? act : act * over;
    for (const [sc, p] of [['nega', p50 * (1 - width)], ['neutral', p50], ['posi', p50 * (1 + width)]]) {
      ev.push(['E' + i + sc, D(2026, 1, 5), client, ym, sc, p, act, Math.abs(p - act) / act, 0, '', '', sc === 'neutral' ? 1 : 0, p - act, Math.abs(p - act), '', '', '', '', '', '', '', 1]);
    }
  }
  const evid = [H.RELIABILITY_EVIDENCE];
  for (const [type, n, h] of [['opinion', 10, Math.round(10 * hit)], ['factor_product', 8, 4], ['ai_topic', 2, 1]]) evid.push([client, type, 'k', '2025Q4', '2025/12', n, h, h / n, D(2026, 1, 5), 'R', '']);
  return env.makeBook(client, {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', client], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
    EVAL_LOG: { values: ev, formats: { D: '@' } },
    RELIABILITY_EVIDENCE: { values: evid },
    POOL_PRIOR: { values: [H.POOL_PRIOR] },
    CALIBRATION_STATE: { values: [H.CALIBRATION_STATE, [client, D(2026, 1, 5), 'owner', '', '', '', 0.98, '', '{}', '', '', 1, '']] },
  });
}
const ids = {};
for (const [c, over, width, hit] of [['甲製薬', 1.10, 0.05, 0.7], ['乙製薬', 1.04, 0.15, 0.5], ['丙製薬', 0.97, 0.10, 0.6]]) {
  const b = book(c, over, width, hit);
  ids[c] = env.seedPlan(b);
}

// ==== 1. 精度の推移（影）: 締まった後の予測（誤差 0）は除く ====
const a = env.call('apiLearningView(__in)', { __in: { planId: ids['甲製薬'] } });
assert.equal(a.accuracy.n, 11);
assert.equal(a.accuracy.leaks, 1, '予測 = 実績の月は除く');
assert.ok(Math.abs(a.accuracy.mape - 0.10) < 1e-9, '10% 高い予測');
assert.ok(Math.abs(a.accuracy.bias - 0.10) < 1e-9, '偏りは +10%');
assert.equal(a.accuracy.coverage, 0, 'P10〜P90（±5%）に実績は入らない');
assert.ok(a.accuracy.widthScale > 1.5, '幅は今の 2 倍ほど要る: ' + a.accuracy.widthScale);
assert.equal(a.accuracy.months.length, 12);
assert.equal(a.factorNow, 0.98);

// ==== 2. 全計画で縮めた偏りの補正（影）: 3 つの計画の平均へ寄る。過去の月に当てると誤差が減る ====
assert.equal(a.pooled, true);
assert.equal(a.pooledPlans, 3);
assert.ok(a.mu > 0 && a.mu < 0.10, '全体の平均の偏り: ' + a.mu);
assert.ok(a.shadow.shrunk <= 0.10 && a.shadow.shrunk >= a.mu - 1e-9, '自分の偏りと平均の間: ' + a.shadow.shrunk);
assert.ok(a.shadow.factorShadow < 1, '下げる補正');
assert.ok(a.shadow.mapeShadow < a.shadow.mapeNow, '過去の月に当てはめると誤差が減る');
const c = env.call('apiLearningView(__in)', { __in: { planId: ids['丙製薬'] } });
assert.ok(c.accuracy.bias < 0 && c.accuracy.coverage === 1, '低めの予測・幅の中');

// ==== 3. 情報源の信頼度の事前分布（ベータ二項の経験ベイズ） ====
const pv = env.call('apiPoolPreview()');
const op = pv.types.filter((t) => t.type === 'opinion')[0];
assert.equal(op.ok, true);
assert.equal(op.plans, 3);
assert.ok(Math.abs(op.hitRate - 18 / 30) < 1e-9, '当たった割合: 全計画の合計');
assert.ok(op.precision >= 2 && op.precision <= 50, '事前分布の強さ: ' + op.precision);
assert.ok(Math.abs(op.pooledR - 1.2) < 1e-9, '旧来と同じ r の尺度（2 × 当たる割合）');
assert.ok(op.perPlan.every((x) => x.posterior > Math.min(x.hitRate, op.hitRate) - 1e-9 && x.posterior < Math.max(x.hitRate, op.hitRate) + 1e-9), '事後の平均は自分と全体の間');
assert.equal(pv.types.filter((t) => t.type === 'ai_topic')[0].ok, false, '数が少ない情報源は作らない');

// 書く: 各計画の POOL_PRIOR に入る（旧来の C-1 が読む形）
const st = env.runJob('LEARN.POOL', {});
assert.equal(st.status, 'DONE', st.error);
assert.equal(st.result.plans.length, 3);
for (const id of Object.values(ids)) {
  const rows = env.table('ENG_POOL_PRIOR').filter((r) => r.plan_id === id);
  const o = rows.filter((r) => r.pool_scope === 'reliability:opinion')[0];
  assert.ok(o, 'reliability:opinion の行');
  assert.equal(o.param_key, 'reliability_r');
  assert.equal(o.pooled_value, '1.2');
}
const pv2 = env.call('apiPoolPreview()');
assert.equal(pv2.current[ids['乙製薬']]['reliability:opinion'].value, 1.2);
assert.ok(env.audit().some((x) => x.action === 'LEARN.POOL' && x.phase === 'END' && x.result === 'OK'));
// 計画を組み立て直しても同じ（データ本体と計算用ブックが合う）
assert.equal(env.call('apiPlanView(__in)', { __in: { planId: ids['甲製薬'] } }).pendingWrite || null, null);
// ==== 4. 表が大きいときは、計画の行だけを探して読む（結果は全部読んだときと同じ・読むセルは少ない） ====
const load = (id) => env.run(`(() => { APP_STORE_CACHE_ = {}; return JSON.stringify(appEngLoadPlanSheets_('${id}')); })()`);
const reset = () => { for (const k of Object.keys(STATS)) STATS[k] = k === 'bySheet' ? {} : 0; };
reset(); const full = load(ids['乙製薬']); const fullCells = STATS.readCells;
env.run('appPlanReadMinRows_ = () => 0');
reset(); const part = load(ids['乙製薬']); const partCells = STATS.readCells;
assert.equal(part, full, '計画の行だけ読んでも同じ');
assert.ok(partCells < fullCells * 0.7, '読むセルが減る: ' + fullCells + ' → ' + partCells);
assert.equal(env.call('apiLearningView(__in)', { __in: { planId: ids['甲製薬'] } }).accuracy.n, 11, '計画ごとに読んでも同じ結果');
const again = env.runJob('LEARN.POOL', {});
assert.equal(again.status, 'DONE', again.error);
env.run('appPlanReadMinRows_ = () => 3000');
console.log('app-learning: all tests passed');
