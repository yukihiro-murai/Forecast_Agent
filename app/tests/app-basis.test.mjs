/**
 * app-basis.test.mjs — 予測の根拠（月ごとの内訳・入力と AI の押し・補正・AI 調査の根拠・前回からの変化）。
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
const url = 'https://docs.google.com/spreadsheets/d/' + book.getId() + '/edit';
const dry = env.runJob('MIGRATION.DRYRUN', { bookUrl: url });
const imp = env.runJob('MIGRATION.IMPORT', { bookUrl: url, contentHash: dry.result.contentHash });
assert.equal(imp.status, 'DONE', imp.error);
const planId = imp.result.planId;

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
assert.equal(b.inputChanged, false, 'この間に計画への操作は無い（変化は乱数の揺れ）');
console.log('app-basis: all tests passed');
