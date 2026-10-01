#!/usr/bin/env node
/*
 * forecast-autolearn.test.mjs — B-5 月次ベイズ自動学習の純粋計算部を Node で検証する。
 * Forecast_Agent.js から対象関数と AUTOLEARN_* 定数を抽出し vm で実行（GAS API 非依存の範囲）。
 *
 *   node tests/forecast-autolearn.test.mjs
 */

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const src = await readFile(path.join(root, 'Forecast_Agent.js'), 'utf8');

/** `function name(` から対応する閉じ括弧までのソースを切り出す。 */
function extractFunction(name) {
  const start = src.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} not found in Forecast_Agent.js`);
  let i = src.indexOf('{', start);
  let depth = 0;
  let end = -1;
  for (let p = i; p < src.length; p++) {
    const ch = src[p];
    if (ch === '{') depth++;
    else if (ch === '}') { depth--; if (depth === 0) { end = p + 1; break; } }
  }
  assert.notEqual(end, -1, `${name} body not closed`);
  return src.slice(start, end);
}

// AUTOLEARN_*/VERTEX_ASSIST_* の const 宣言をそのまま取り込む（値の二重管理を防ぐ）
const constLines = src.match(/^const (AUTOLEARN_[A-Z_]+|VERTEX_ASSIST_[A-Z_]+) *= *.+;.*$/gm) || [];
assert.ok(constLines.length >= 8, 'AUTOLEARN constants should be extracted');

const harness = [
  ...constLines,
  extractFunction('clamp_'),
  extractFunction('parseYM_'),
  extractFunction('autoLearnComputeState_'),
  extractFunction('parseResidualMonthBiasJson_'),
  extractFunction('canonicalMonthBiasJson_'),
].join('\n');

const sandbox = { Math, JSON, Number, String, isFinite, console };
vm.createContext(sandbox);
vm.runInContext(`${harness}\n;this.__fns = { autoLearnComputeState_, parseResidualMonthBiasJson_, canonicalMonthBiasJson_ };`, sandbox);
const { autoLearnComputeState_, parseResidualMonthBiasJson_, canonicalMonthBiasJson_ } = sandbox.__fns;

const AUTOLEARN_MIN_EVAL_MONTHS = 3;
const pairs = n => Array.from({ length: n }, (_, i) => ({ ym: `2025/${String(12 - i).padStart(2, '0')}`, pred: 110, actual: 100 }));

// ---- 1. 評価月不足では更新しない ----
{
  const r = autoLearnComputeState_(pairs(2));
  assert.equal(r.ready, false);
  assert.equal(r.reason, 'insufficient_eval_months');
  assert.equal(r.n, 2);
}

// ---- 2. 過剰予測は係数を下げる方向に更新（maxDelta=±0.05 にクランプ） ----
{
  const r = autoLearnComputeState_(pairs(6));
  assert.equal(r.ready, true);
  assert.ok(r.postBias > 0 && r.postBias < 0.10, `postBias は生の平均 0.10 より事前分布側に縮む: ${r.postBias}`);
  assert.ok(r.targetFactor < 1 && r.targetFactor >= 0.75, `targetFactor: ${r.targetFactor}`);
  assert.equal(r.factor, 0.95, '1回の更新幅は +0.05/-0.05 まで');
}

// ---- 3. 過小予測は係数を上げる ----
{
  const r = autoLearnComputeState_(pairs(6).map(p => ({ ...p, pred: 90 })));
  assert.ok(r.postBias < 0);
  assert.equal(r.factor, 1.05);
}

// ---- 4. 極端な誤差でも係数は [0.75, 1.25] と 1更新あたり±0.05 を守る ----
{
  const r = autoLearnComputeState_(pairs(12).map(p => ({ ...p, pred: 600 })));
  assert.equal(r.targetFactor, 0.75);
  assert.equal(r.factor, 0.95, '暴走しない: 1更新は+0.05まで');
}

// ---- 5. 反復で目標係数へ収束する（学習ループのシミュレーション） ----
{
  let cur = 1.0;
  for (let k = 0; k < 20; k++) {
    const r = autoLearnComputeState_(pairs(6), { curFactor: cur });
    cur = r.factor;
  }
  const r = autoLearnComputeState_(pairs(6), { curFactor: cur });
  assert.ok(Math.abs(cur - r.targetFactor) < 0.02, `収束先 ${cur} が target ${r.targetFactor} に近い`);
}

// ---- 6. 暦月バイアス: 特定月だけ系統誤差があるとその暦月に学習される ----
{
  // 12月,1月,2月は誤差0 / 4月だけ +0.25 過剰（FY跨ぎで4月を複数回含む）
  const data = [
    { ym: '2026/04', pred: 125, actual: 100 },
    { ym: '2026/03', pred: 100, actual: 100 },
    { ym: '2026/02', pred: 100, actual: 100 },
    { ym: '2026/01', pred: 100, actual: 100 },
    { ym: '2025/12', pred: 100, actual: 100 },
    { ym: '2025/04', pred: 125, actual: 100 },
  ];
  const r = autoLearnComputeState_(data);
  assert.equal(r.ready, true);
  assert.ok(r.monthBias['4'] < -0.10, `4月の過剰予測 → 負の補正: ${r.monthBias['4']}`);
  assert.ok(Math.abs(r.monthBias['4']) <= 0.20, '暦月バイアスは±0.20にキャップ');
  assert.equal(r.monthBias['3'], undefined, '誤差のない月は登録されない');
}

// ---- 7. 微小な月バイアス (<0.005) は記録しない ----
{
  const data = pairs(6).map(p => ({ ...p, pred: 100.4 })); // e≈+0.004
  const r = autoLearnComputeState_(data);
  assert.equal(Object.keys(r.monthBias).length, 0);
}

// ---- 8. parseResidualMonthBiasJson_: 月範囲外・壊れた値を落とす ----
{
  const out = parseResidualMonthBiasJson_('{"4":-0.12,"13":0.5,"x":1,"7":"0.03"}');
  assert.equal(JSON.stringify(out), JSON.stringify({ '4': -0.12, '7': 0.03 }));
  assert.equal(JSON.stringify(parseResidualMonthBiasJson_('')), '{}');
  assert.equal(JSON.stringify(parseResidualMonthBiasJson_('not json')), '{}');
  assert.equal(JSON.stringify(parseResidualMonthBiasJson_('[1,2,3]')), '{}');
}

// ---- 9. canonicalMonthBiasJson_: 数値ソート・非数値除去で差分比較が安定 ----
{
  const j = canonicalMonthBiasJson_({ '10': 0.1, '2': -0.05, '4': 'bad' });
  assert.equal(j, '{"2":-0.05,"10":0.1}');
  assert.equal(canonicalMonthBiasJson_(parseResidualMonthBiasJson_(j)), j, '往復で冪等');
}

// ---- 10. actual=0 / 非数値の行は評価から除外 ----
{
  const data = pairs(4).concat([{ ym: '2025/11', pred: 5, actual: 0 }, { ym: '2025/10', pred: 'x', actual: 100 }]);
  const r = autoLearnComputeState_(data);
  assert.equal(r.n, 4, 'actual=0 と非数値 pred は除外');
}

process.stdout.write('PASS forecast-autolearn contract tests\n');
