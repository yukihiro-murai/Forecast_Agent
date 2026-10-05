#!/usr/bin/env node
/*
 * forecast-weather.test.mjs — 空模様（天気キャラ）の表示判定と、画面へのキャラ配置の契約テスト。
 * Forecast_WebAppUI.html のクライアント JS から判定関数を取り出し vm で実行する（DOM・GAS 非依存の範囲）。
 * 判定は既存の検証データ（compare[] の ape / signedErr、insights[] の制約超過フラグ）を読むだけ。
 *
 *   node tests/forecast-weather.test.mjs
 */

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const html = await readFile(path.join(root, 'Forecast_WebAppUI.html'), 'utf8');
const js = (html.match(/<script>([\s\S]*?)<\/script>/) || [])[1];
assert.ok(js, 'Forecast_WebAppUI.html にクライアント script がある');

/** `prefix` で始まる文を、対応する閉じ括弧まで切り出す（function / var X = {...}; の両方）。 */
function extract(prefix) {
  const start = js.indexOf(prefix);
  assert.notEqual(start, -1, `${prefix} not found`);
  let depth = 0;
  for (let p = js.indexOf('{', start); p < js.length; p++) {
    if (js[p] === '{') depth++;
    else if (js[p] === '}' && --depth === 0) return js.slice(start, js[p + 1] === ';' ? p + 2 : p + 1);
  }
  throw new Error(`${prefix} not closed`);
}
const line = (re) => { const m = js.match(re); assert.ok(m, String(re)); return m[0]; };

const ctx = vm.createContext({});
vm.runInContext([
  line(/^var CHAR_SVG = .*;$/m), line(/^var YOMI_POSE = .*;$/m),
  extract('function esc('), extract('function fmtPct('),
  'function help(t){ return ""; }',
  extract('var WX = '), line(/^var WX_FLAGS = .*;$/m), line(/^var WX_MIN_MONTHS = .*;$/m),
  extract('function wxEvalMonths('), extract('function forecastWeather('), extract('function wxChange('), extract('function wxCardHtml('),
].join('\n'), ctx);
const W = (ev, upto) => vm.runInContext('forecastWeather(__ev, __upto)', Object.assign(ctx, { __ev: ev, __upto: upto }));
const C = (ev) => vm.runInContext('wxChange(__ev)', Object.assign(ctx, { __ev: ev }));

/** 月ごとの {ape, signedErr} と、最新月（または months で指定）の insights フラグから検証データを作る。 */
function ev(months, flagsByMonth = {}, opts = {}) {
  const compare = months.map(([month, ape, signedErr]) => ({ month, p50: 100, forecastTotal: 100, actualTotal: 100, ape, signedErr }));
  const insights = Object.entries(flagsByMonth).map(([month, f]) => ({
    month, annualBreach: !!f.annual, halfBreach: !!f.half, overBreach: !!f.over, rangeBreach: !!f.range }));
  return { hasActuals: opts.hasActuals !== false, compare, insights, metrics: [] };
}
const M = ['2026/06', '2026/07', '2026/08', '2026/09'];
const calm = [[M[0], 0.08, 5], [M[1], 0.06, -3], [M[2], 0.07, 4], [M[3], 0.05, 2]];
const none = {};

// ---- 1. 未確認・霧: 実績なし / 3 か月未満 / B-4 未実行 ----
{
  assert.equal(W(ev([], {}, { hasActuals: false })).id, 'mikakunin', '実績が未取込');
  assert.equal(W(null).id, 'mikakunin', '検証データ自体が無い');
  const two = W(ev(calm.slice(0, 2), { [M[1]]: none }));
  assert.equal(two.id, 'mikakunin'); assert.match(two.reason, /2 か月/);
  assert.equal(W(ev(calm, {})).id, 'mikakunin', '最新月の insights（B-4）が無い');
  const noFc = ev(calm, { [M[3]]: none }); noFc.compare[3].p50 = null; noFc.compare[3].forecastTotal = null;
  assert.equal(W(noFc).month, M[2], '予測の無い月は判定に使わない');
}

// ---- 2〜8. 優先順位どおりに決まる ----
{
  assert.equal(W(ev([...calm.slice(0, 3), [M[3], 1.6, -50]], { [M[3]]: { annual: 1, half: 1, over: 1 } })).id, 'tenpen', 'APE 150% 以上は天変地異が最優先');
  assert.equal(W(ev([[M[0], 0.1, 1], [M[1], 0.35, 30], [M[2], 0.1, 2], [M[3], 0.4, -40]], { [M[3]]: { over: 1 } })).id, 'taifuu', '向きが入れ替わり APE 30% 以上が 2 か月 → 台風（雪より先）');
  assert.notEqual(W(ev([[M[0], 0.1, 1], [M[1], 0.35, 30], [M[2], 0.1, 2], [M[3], 0.4, 40]], { [M[3]]: none })).id, 'taifuu', '向きが変わらなければ台風ではない');
  assert.notEqual(W(ev([[M[0], 0.1, 1], [M[1], 0.35, 30], [M[2], 0.1, -2], [M[3], 0.2, 40]], { [M[3]]: none })).id, 'taifuu', 'APE 30% 以上が 1 か月なら台風ではない');
  assert.equal(W(ev(calm, { [M[3]]: { over: 1, annual: 1 } })).id, 'sekka', '過大の超過は雪（雨より先）');
  const rain = W(ev(calm, { [M[3]]: { annual: 1, range: 1 } }));
  assert.equal(rain.id, 'ame'); assert.match(rain.reason, /2 つ/); assert.match(rain.reason, /年間・範囲外/);
  const cloud = W(ev(calm, { [M[3]]: { half: 1 } }));
  assert.equal(cloud.id, 'kumori'); assert.match(cloud.reason, /半期/);
  assert.equal(W(ev([[M[0], 0.03, 1], [M[1], 0.05, 1], [M[2], 0.07, 1], [M[3], 0.09, 1]], { [M[3]]: none })).id, 'harenochi', '超過なしで APE が 2 か月続けて悪化');
  assert.equal(W(ev([[M[0], 0.03, 1], [M[1], 0.05, 1], [M[2], 0.09, 1], [M[3], 0.09, 1]], { [M[3]]: none })).id, 'kaisei', '横ばいは悪化に数えない');
  const fine = W(ev(calm, { [M[3]]: none }));
  assert.equal(fine.id, 'kaisei'); assert.equal(fine.month, M[3]); assert.deepEqual([...fine.flags], []);
}

// ---- 9. 前の月からの変化（ホームの「変化発見」） ----
{
  const d = ev(calm, { [M[2]]: none, [M[3]]: { half: 1 } });
  assert.equal(W(d, M[2]).id, 'kaisei', 'upto で前の月の空模様');
  const c = C(d);
  assert.ok(c, '快晴 → 曇り は変化');
  assert.equal(c.from.id, 'kaisei'); assert.equal(c.to.id, 'kumori');
  assert.equal(C(ev(calm, { [M[2]]: none, [M[3]]: none })), null, '変わらなければ null');
  assert.equal(C(ev(calm.slice(0, 1), { [M[0]]: none })), null, '1 か月だけなら変化なし');
}

// ---- 10. 空模様カード: 理由と制約のチップ ----
{
  const card = vm.runInContext('wxCardHtml(forecastWeather(__ev))', Object.assign(ctx, { __ev: ev(calm, { [M[3]]: { half: 1 } }) }));
  assert.match(card, /2026\/09 の空模様/);
  assert.match(card, /半期 超過/); assert.match(card, /年間 OK/); assert.match(card, /過大 OK/); assert.match(card, /範囲外 OK/);
  assert.match(card, /<svg/);
  const fog = vm.runInContext('wxCardHtml(forecastWeather(__ev))', Object.assign(ctx, { __ev: ev([], {}, { hasActuals: false }) }));
  assert.doesNotMatch(fog, /超過|OK</, '判定できないときは制約のチップを出さない');
}

// ---- 11. 絵は正本 (assets/characters の SVG) と同じ ----
{
  const chars = vm.runInContext('CHAR_SVG', ctx), poses = vm.runInContext('YOMI_POSE', ctx);
  const ids = ['yomi', 'kaisei', 'harenochi', 'kumori', 'ame', 'sekka', 'taifuu', 'tenpen', 'mikakunin'];
  assert.deepEqual(Object.keys(chars), ids);
  for (const id of ids) {
    const svg = (await readFile(path.join(root, 'assets/characters/svg', id + '.svg'), 'utf8')).trim();
    assert.equal(chars[id], svg, `${id} が assets/characters/svg と違う（cd assets/characters/src && python3 build.py）`);
  }
  assert.deepEqual(Object.keys(poses), ['guide', 'observe', 'discover', 'explain', 'done']);
  for (const k of Object.keys(poses)) {
    const svg = (await readFile(path.join(root, 'assets/characters/yomi/svg', 'yomi_' + k + '.svg'), 'utf8')).trim();
    assert.equal(poses[k], svg, `yomi_${k} が assets と違う`);
  }
  const vm2 = vm.runInContext('Object.keys(WX)', ctx);
  for (const id of vm2) assert.ok(chars[id], `WX.${id} の絵がある`);
}

// ---- 12. 画面への配置（案のとおり） ----
{
  assert.doesNotMatch(js, /CHAR_GEAR|CHAR_ARROW/, '旧キャラ（歯車・矢印）は残さない');
  assert.match(extract('function emptyBox('), /CHAR_SVG\.mikakunin/, '空の表示は 未確認・霧');
  const home = extract('function renderHome(');
  assert.match(home, /YOMI_POSE\.guide/, 'ホーム「次の一手」は よみ（通常案内）');
  assert.match(home, /YOMI_POSE\.discover/, '空模様が変わった月は よみ（変化発見）');
  assert.match(extract('function renderEval('), /wxCardHtml\(forecastWeather\(B\.eval\)\)/, '検証画面の上部に最新月の空模様');
  assert.match(extract('function busyChar('), /YOMI_POSE\.observe/, '処理中は よみ（観測中）');
  const rpcSrc = extract('function rpc(');
  assert.equal((rpcSrc.match(/busyChar\(false\)/g) || []).length, 2, '成功・失敗のどちらでも処理中の表示を消す');
  assert.match(extract('function showToast('), /YOMI_POSE\.done/, '完了の通知は よみ（確認完了）');
}

process.stdout.write('PASS forecast-weather contract tests\n');
