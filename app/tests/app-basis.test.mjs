/**
 * app-basis.test.mjs — 予測の根拠（月ごとの内訳・入力と AI の押し・補正・AI 調査の根拠・前回からの変化とそのわけ changeCause）。
 * 旧来の予測の代わりに、旧来と同じ見出しで記録のシートを書く（本物の予測はモックでは動かさない）。
 */
import assert from 'node:assert/strict';
import { setUpEnv } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const env = setUpEnv();
const H = env.run('(() => { const o = {}; Object.keys(APP_ENGINE_SHEETS).forEach(k => { o[k] = APP_ENGINE_SHEETS[k].header || null; }); return o; })()');
const output = [['FY2026 売上予測（テスト製薬）']];
for (let r = 2; r <= 79; r++) output.push([]);
output[25] = ['年度合計（予測）', 900, 1000, 1100];
for (let i = 0; i < 12; i++) output[28 + i] = [D(2026, 4 + i), 70, 80, 90];
output[64] = ['年度合計（予測）', 800, 900, 1000];
for (let i = 0; i < 12; i++) output[67 + i] = [D(2026, 4 + i), 60, 70, 80];
const book = env.makeBook('クライアント別売上予測', {
  CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', 'テスト製薬'], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
  OUTPUT: { values: output },
  FORECAST_SNAPSHOT: { values: [H.FORECAST_SNAPSHOT], formats: { D: '@' } },
  AI_IMPACT_HISTORY: { values: [H.AI_IMPACT_HISTORY] },
  SUBJECTIVE_IMPACT_HISTORY: { values: [H.SUBJECTIVE_IMPACT_HISTORY] },
  AI_RESEARCH_STRUCTURED: { values: [H.AI_RESEARCH_STRUCTURED, ['テスト製薬', D(2026, 10, 3), 'Market', 'event', 'up', 0.4, 0.7, '新薬の採用が広がっている', 'short', '主力製品に関係', '', '', '', '', '', '', '', '', '', 0.4, 0.2, 0.3]] },
  CALIBRATION_STATE: { values: [H.CALIBRATION_STATE, ['テスト製薬', D(2026, 10, 1), 'owner', '', '', '', 0.97, '', '{"2026-05":-0.05}', '2026Q2', '', 1, '']] },
  PROCESS_STATUS: { values: [H.PROCESS_STATUS, ['step4_status', D(2026, 9, 1), 'owner', 'success', 'テスト製薬', 12, '']] },
});
const planId = env.seedPlan(book);

// 旧来の予測の代わり: OUTPUT と、記録のシート（旧来と同じ見出し）を 1 回分書く。2 回目は 5 月の見解が強くなる
env.run(`(() => {
  globalThis.__n = 0;
  const base = appLegacyEngine_;
  appLegacyEngine_ = function (svc) {
    const eng = base(svc);
    eng.runPhase1Forecast = function () {
      __n++;
      const ss = svc.SpreadsheetApp.getActiveSpreadsheet(), now = new svc.Date();
      const out = ss.getSheetByName('OUTPUT');
      out.getRange(26, 1, 1, 4).setValues([['年度合計（予測）', 900, 1000 + __n * 10, 1100]]);
      for (let i = 0; i < 12; i++) out.getRange(29 + i, 1, 1, 4).setValues([[new svc.Date(2026, 3 + i, 1), 70, 80 + (i === 1 ? __n * 10 : 0), 90]]);
      out.getRange(65, 1, 1, 4).setValues([['年度合計（予測）', 800, 900, 1000]]);
      for (let i = 0; i < 12; i++) out.getRange(68 + i, 1, 1, 4).setValues([[new svc.Date(2026, 3 + i, 1), 60, 70, 80]]);
      const ym = i => '2026/' + String(4 + i > 12 ? 4 + i - 12 : 4 + i).padStart(2, '0');
      const snap = ss.getSheetByName('FORECAST_SNAPSHOT');
      for (let i = 0; i < 12; i++) snap.getRange(snap.getLastRow() + 1, 1, 1, 15).setValues([['S' + __n + '-' + i, now, 'テスト製薬', ym(i), 'neutral', 80, 0, 0, i === 0 ? 3 : 0, 80, 70, 90, JSON.stringify({ opinion: i === 1 ? '鷹野:+' + (5 * __n) + '%(0.70)' : '' }), '', JSON.stringify({ bias_correction_factor: 0.97 })]]);
      const ai = ss.getSheetByName('AI_IMPACT_HISTORY');
      for (let i = 0; i < 12; i++) ai.getRange(ai.getLastRow() + 1, 1, 1, 12).setValues([['R' + __n, now, 'テスト製薬', ym(i), 1.02, 0.3, 'up', 80, 78, 0, 0, 'forecast_open']]);
      const sp = ss.getSheetByName('SUBJECTIVE_IMPACT_HISTORY');
      sp.getRange(sp.getLastRow() + 1, 1, 2, 11).setValues([
        ['R' + __n, now, 'テスト製薬', ym(1), 'opinion', '鷹野', 0.05 * __n, 'up', 0.8, '', 'forecast_open'],
        ['R' + __n, now, 'テスト製薬', ym(1), 'ai_topic', 'Market', 0.02, 'up', 1, '', 'forecast_open']]);
    };
    return eng;
  };
})()`);
const r1 = env.runJob('FORECAST.RUN', { planId });
assert.equal(r1.status, 'DONE', r1.error);
const r2 = env.runJob('FORECAST.RUN', { planId });
assert.equal(r2.status, 'DONE', r2.error);

const b = env.call('apiForecastBasis(__in)', { __in: { planId } });
assert.equal(b.monthly.length, 12);
const may = b.monthly[1];
assert.equal(may.month, '2026/05');
assert.equal(may.p50, 100, '最終の P50（2 回目）');
assert.equal(may.objective, 70, '過去の売上だけの P50');
assert.equal(may.diff, 30);
assert.equal(may.prevP50, 90);
assert.equal(may.change, 10, '前回から');
assert.deepEqual(may.sources.map((s) => [s.label, Math.round(s.step * 100)]), [['見解', 10], ['AI の話題', 2]], '最新の回の押しだけ');
assert.match(may.opinion, /鷹野:\+10%/);
assert.equal(b.monthly[0].spot, 3, 'スポットの確定分');
assert.equal(b.monthly[0].kAi, 1.02);
assert.equal(b.calibration.factor, 0.97);
assert.equal(b.calibration.monthBias['2026-05'], -0.05);
assert.ok(b.applied && b.applied.bias_correction_factor === 0.97);
assert.equal(b.research.length, 1);
assert.equal(b.research[0].direction, 'up');
assert.match(b.research[0].evidence, /新薬/);
assert.equal(b.annual.p50 - b.annual.prevP50, 10);
assert.equal(b.inputChanged, false, 'この間に計画への操作は無い');
// テストの予測は、同じ入力・同じ種でも回ごとに数字を変える（本物は変えない）。種が同じ = 入力・版・地域が同じなので、版の違いとは言わない（other）
assert.deepEqual([b.changeCause, b.changeCauses, b.changeVersion], ['other', ['other'], { engine: 'same', app: 'unknown' }]);

// ==== 前回の予測からの変化のわけ（changeCause。2026-10-08 F4）: 決まった順で 1 つ選ぶ ====
{
  const SEED = (c) => c.repeat(64);
  const run = (o) => Object.assign({ run_id: 'R', seed: SEED('a'), as_of: '2026-10-08T10:00:00+0900', engine_version: '2.4.0-dev', engine_sha256: SEED('e'), annual_p50: 1000 }, o);
  const mon = (v) => ({ '2026/04': { p50: v }, '2026/05': { p50: 500 } });
  const cause = (latest, prev, cur, before, between) => env.call('appBasisChangeCause_(__l, __p, __c, __b, __w)', { __l: latest, __p: prev, __c: cur, __b: before, __w: between || [] });
  const op = (action, changed) => ({ action, changed });
  // 前の回が無い
  assert.deepEqual(cause(run({}), null, mon(500), {}), { changeCause: '', changeCauses: [], changeVersion: null });
  // none: 年度と月の P50 がまったく同じ（操作があっても、月や版が違っても none だけ）
  const n = cause(run({ seed: SEED('b'), as_of: '2026-11-02T09:00:00+0900' }), run({}), mon(500), mon(500), [op('INPUT.SAVE', ['OPINIONS'])]);
  assert.deepEqual([n.changeCause, n.changeCauses], ['none', ['none']]);
  assert.equal(cause(run({ annual_p50: 1000 }), run({}), mon(501), mon(500)).changeCause !== 'none', true, '月の P50 が違えば none ではない');
  assert.notEqual(cause(run({ annual_p50: 1001 }), run({}), mon(500), mon(500)).changeCause, 'none', '年度の P50 が違えば none ではない');
  assert.notEqual(cause(run({}), run({}), mon(500), { '2026/04': { p50: 500 } }).changeCause, 'none', '片方にしか無い月があれば none ではない');
  // inputs: 間の操作が予測の読む表を変えた（月や版が違っても inputs が先）。予算・記入・実績の取り込みだけなら数えない
  const i = cause(run({ seed: SEED('b'), annual_p50: 1100, as_of: '2026-11-02T09:00:00+0900', engine_sha256: SEED('f') }), run({}), mon(600), mon(500), [op('BUDGET.SAVE', ['OUTPUT']), op('INPUT.SAVE', ['OPINIONS'])]);
  assert.deepEqual([i.changeCause, i.changeCauses, i.changeVersion], ['inputs', ['inputs', 'calendar', 'version'], { engine: 'changed', app: 'unknown' }]);
  for (const [a, ch] of [['IMPORT.SALES', ['SALES_INPUT', 'PROCESS_STATUS']], ['AI.RESEARCH', ['AI_RESEARCH_STRUCTURED']], ['EVAL.REPORT', ['EVAL_LOG', 'CALIBRATION_STATE']], ['SETUP.PEOPLE', ['CONFIG']]]) {
    assert.equal(cause(run({ seed: SEED('b'), annual_p50: 1100 }), run({}), mon(600), mon(500), [op(a, ch)]).changeCause, 'inputs', a);
  }
  for (const [a, ch] of [['BUDGET.SAVE', ['OUTPUT']], ['INSIGHT.SAVE', ['EVAL_INSIGHTS']], ['IMPORT.ACTUALS', ['ACTUAL_EVAL_MONTHLY', 'PROCESS_STATUS', 'RUN_LOG']], ['EVAL.REPORT', ['EVAL_LOG', 'EVAL_COMPARE_MONTHLY']]]) {
    assert.notEqual(cause(run({ seed: SEED('b'), annual_p50: 1100 }), run({}), mon(600), mon(500), [op(a, ch)]).changeCause, 'inputs', a + ' は予測の読む表を変えない');
  }
  // 両方とも中身から作った同じ種なら、操作があっても予測の読んだ表は同じ（inputs にしない）。月が違えば calendar
  const ss = cause(run({ annual_p50: 1100, as_of: '2026-11-02T09:00:00+0900' }), run({}), mon(600), mon(500), [op('INPUT.SAVE', ['OPINIONS'])]);
  assert.deepEqual([ss.changeCause, ss.changeCauses], ['calendar', ['calendar']]);
  // calendar: 「今」の月が違う（同じ月の別の日は calendar ではない）。版が違っても calendar が先
  const cal = cause(run({ seed: SEED('b'), annual_p50: 1100, as_of: '2026-11-01T00:30:00+0900', engine_sha256: SEED('f') }), run({ as_of: '2026-10-31T23:30:00+0900' }), mon(600), mon(500));
  assert.deepEqual([cal.changeCause, cal.changeCauses, cal.changeVersion.engine], ['calendar', ['calendar', 'version'], 'changed']);
  assert.notEqual(cause(run({ seed: SEED('b'), annual_p50: 1100, as_of: '2026-10-31T23:00:00+0900' }), run({ as_of: '2026-10-01T09:00:00+0900' }), mon(600), mon(500)).changeCause, 'calendar');
  // version: 旧来の計算の中身が違う（同じ月・操作なし）。どちらの回も中身から作った種
  const ver = cause(run({ seed: SEED('b'), annual_p50: 1100, engine_sha256: SEED('f') }), run({}), mon(600), mon(500));
  assert.deepEqual([ver.changeCause, ver.changeCauses, ver.changeVersion], ['version', ['version'], { engine: 'changed', app: 'unknown' }]);
  // 中身のハッシュの記録が無ければ、版の名前で比べる（名前も同じ・無ければ分からない）
  assert.deepEqual(cause(run({ seed: SEED('b'), annual_p50: 1100, engine_sha256: '', engine_version: '2.5.0' }), run({}), mon(600), mon(500)).changeVersion, { engine: 'changed', app: 'unknown' });
  assert.deepEqual(cause(run({ seed: SEED('b'), annual_p50: 1100, engine_sha256: '' }), run({}), mon(600), mon(500)).changeVersion, { engine: 'unknown', app: 'unknown' });
  // jitter: どちらかの種が中身から作った種でない（v0.29.0 より前の実行の ID）。同じ月・同じ旧来の計算。アプリの版も前と後で違う
  const jit = cause(run({ annual_p50: 1100 }), run({ seed: 'ACT-20261001-abc' }), mon(600), mon(500));
  assert.deepEqual([jit.changeCause, jit.changeCauses, jit.changeVersion], ['jitter', ['jitter', 'version'], { engine: 'same', app: 'changed' }]);
  assert.equal(cause(run({ seed: 'RUN-1', annual_p50: 1100 }), run({ seed: 'RUN-0' }), mon(600), mon(500)).changeCause, 'jitter', '両方とも前の種');
  assert.equal(cause(run({ annual_p50: 1100, as_of: '2026-11-02T09:00:00+0900' }), run({ seed: 'RUN-0' }), mon(600), mon(500)).changeCause, 'calendar', '月が違えば jitter ではない');
  assert.equal(cause(run({ annual_p50: 1100, engine_sha256: SEED('f') }), run({ seed: 'RUN-0' }), mon(600), mon(500)).changeCause, 'version', '旧来の計算が違えば jitter ではない');
  // version（ほかのわけが無い）: 両方とも中身から作った種で、月も旧来の計算も同じ。アプリの版は記録していないので unknown
  const fb = cause(run({ seed: SEED('b'), annual_p50: 1100 }), run({}), mon(600), mon(500), [op('BUDGET.SAVE', ['OUTPUT'])]);
  assert.deepEqual([fb.changeCause, fb.changeCauses, fb.changeVersion], ['version', ['version'], { engine: 'same', app: 'unknown' }]);
  // 同じ入力なら何度でも同じ答え（決まった順）
  assert.deepEqual(cause(run({ seed: SEED('b'), annual_p50: 1100 }), run({}), mon(600), mon(500)), cause(run({ seed: SEED('b'), annual_p50: 1100 }), run({}), mon(600), mon(500)));
}
console.log('app-basis: all tests passed');
