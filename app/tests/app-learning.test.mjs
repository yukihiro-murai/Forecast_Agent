/**
 * app-learning.test.mjs — 精度の推移・全計画で縮めた偏りの補正（影）・幅の較正・情報源の信頼度の事前分布（ベータ二項）。
 */
import assert from 'node:assert/strict';
import { setUpEnv, STATS, J } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const env = setUpEnv();
const H = env.run('(() => { const o = {}; Object.keys(APP_ENGINE_SHEETS).forEach(k => { o[k] = APP_ENGINE_SHEETS[k].header || null; }); return o; })()');
/**
 * 実績の取り込み（B-1）と検証（B-2）の記録。締まった月は、B-1 の日に月末から 5 日たった月だけ（2026-10-07 D5）なので、
 * 精度を測るには B-1 → B-2 の記録が要る（無ければ締まった月は無い）。2026-01-05 の取り込みなら 2025/12 まで締まっている
 */
const status = (client, b1 = new Date(2026, 0, 5, 10)) => ({ values: [H.PROCESS_STATUS, ['step2_status', b1, 'owner', 'success', client, 36, ''],
  ['step5_status', new Date(b1.getTime() + 3600e3), 'owner', 'success', client, 36, '']] });
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
    PROCESS_STATUS: status(client),
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

// ==== 5. 締まった月だけで測る（2026-10-07 村井さん承認 D4・D5）。締まっていない月の EVAL_LOG の行は消さずに読み飛ばす ====
{
  const e2 = setUpEnv();
  /** 甲製薬と同じ 12 か月（2025/01〜12。最後の月は予測 = 実績）に、extra の月（途中の実績 10 に予測 1000。誤差がとても大きい）を足した記録 */
  const evalRows = (client, extra) => {
    const ev = [H.EVAL_LOG];
    const add = (i, ym, act, p50) => { for (const [sc, p] of [['nega', p50 * 0.95], ['neutral', p50], ['posi', p50 * 1.05]]) ev.push(['E' + i + sc, D(2026, 2, 3), client, ym, sc, p, act, Math.abs(p - act) / act, 0, '', '', sc === 'neutral' ? 1 : 0, p - act, Math.abs(p - act), '', '', '', '', '', '', '', 1]); };
    for (let i = 0; i < 12; i++) { const act = 1000 + i * 10 * (i % 3 === 0 ? -1 : 1); add(i, '2025/' + String(i + 1).padStart(2, '0'), act, i === 11 ? act : act * 1.1); }
    extra.forEach((ym, j) => add(12 + j, ym, 10, 1000));
    return ev;
  };
  const plan = (client, extra, b1) => e2.seedPlan(e2.makeBook(client, Object.assign({
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', client], ['[必須] 予測年度FY（YYYY）', 2025], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
    EVAL_LOG: { values: evalRows(client, extra), formats: { D: '@' } },
  }, b1 ? { PROCESS_STATUS: status(client, b1) } : {})));
  const base = plan('基準製薬', [], new Date(2026, 0, 5, 10));                                   // 締まっていない月の行が無い
  const day3 = plan('三日製薬', ['2026/01', '2026/02'], new Date(2026, 1, 3, 10));               // 2/03 の取り込み: 1 月はまだ途中
  const day6 = plan('六日製薬', ['2026/01', '2026/02'], new Date(2026, 1, 6, 10));               // 2/06 の取り込み: 1 月は締まった（2 月は途中）
  const none = plan('記録無し製薬', [], null);                                                   // B-1 → B-2 の記録が無い
  const acc = (id) => J(e2.run(`appAccuracyOf_('${id}')`));
  const strip = (a) => Object.assign({}, a, { cutoffYm: undefined });
  const a0 = acc(base), a3 = acc(day3), a6 = acc(day6), aN = acc(none);
  assert.deepEqual([a0.cutoffYm, a3.cutoffYm, a6.cutoffYm, aN.cutoffYm], ['2026/01', '2026/01', '2026/02', '']);
  assert.deepEqual(strip(a3), strip(a0), '締まっていない月の行があっても、精度は無いときと同じ（読み飛ばす）');
  assert.deepEqual([a0.n, a0.leaks, a0.months.length], [11, 1, 12]);
  assert.deepEqual([a6.n, a6.months.length, a6.months[12].month], [12, 13, '2026/01'], '6 日の取り込みなら 1 月を数える（2 月は数えない）');
  assert.ok(a6.mape > a0.mape + 0.05, '締まった 1 月の大きな外れが入る: ' + a6.mape);
  assert.deepEqual([aN.n, aN.months, aN.mape, aN.coverage, aN.widthScale], [0, [], null, null, null], 'B-1 → B-2 の記録が無ければ締まった月は無い');
  // 境目を渡しても同じ（全計画をまとめて読む学びの画面が使う形）
  assert.deepEqual(J(e2.run(`appAccuracyOf_('${day3}', appEngTableObjects_('${day3}', ['EVAL_LOG']).EVAL_LOG, '2026/01')`)), a3);
  // 検証の画面: 全計画で縮めた偏りの影も、締まった月だけ（day3 の今の誤差は基準と同じ）
  const v3 = e2.call('apiLearningView(__in)', { __in: { planId: day3 } }), v0 = e2.call('apiLearningView(__in)', { __in: { planId: base } });
  assert.equal(v3.accuracy.n, 11);
  assert.ok(Math.abs(v3.shadow.mapeNow - v0.shadow.mapeNow) < 1e-12 && Math.abs(v3.shadow.bias - v0.shadow.bias) < 1e-12, '縮めた偏りの材料も同じ');
  assert.ok(!v3.accuracy.months.some((m) => m.month >= '2026/01'), '画面に締まっていない月を出さない');
  // 行は消さない
  assert.equal(e2.table('ENG_EVAL_LOG').filter((r) => r.plan_id === day3).length, 14 * 3, '締まっていない月の行も残る');
}
console.log('app-learning: all tests passed');
