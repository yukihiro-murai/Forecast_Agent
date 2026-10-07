#!/usr/bin/env node
/*
 * forecast-autolearn.test.mjs — B-5 月次ベイズ自動学習の純粋計算部を Node で検証する。
 * Forecast_Agent.js から対象関数と AUTOLEARN_* 定数を抽出し vm で実行（GAS API 非依存の範囲）。
 * 12〜17（2026-10-07）: 補正を掛ける前の予測で学ぶ。補正が無ければ今までと同じ・何か月回しても戻らない・上下限・回の選び直し。
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
const constLines = src.match(/^const (AUTOLEARN_[A-Z_]+|VERTEX_ASSIST_[A-Z_]+|EVAL_CALENDAR_TZ) *= *.+;.*$/gm) || [];
assert.ok(constLines.length >= 8, 'AUTOLEARN constants should be extracted');

// 2026-10-07 より前の autoLearnComputeState_（補正の掛かった予測をそのまま学んでいた）。「今まで」と比べるための写し（名前だけ変えた）
const OLD_COMPUTE_SRC = `function oldAutoLearnComputeState_(pairs, opts) {
  const o = opts || {};
  const minMonths = isFinite(o.minMonths) ? o.minMonths : AUTOLEARN_MIN_EVAL_MONTHS;
  const halfLife = isFinite(o.halfLife) && o.halfLife > 0 ? o.halfLife : AUTOLEARN_EWMA_HALFLIFE_MONTHS;
  const kG = isFinite(o.kG) ? o.kG : AUTOLEARN_GLOBAL_SHRINK_K;
  const kM = isFinite(o.kM) ? o.kM : AUTOLEARN_MONTH_SHRINK_K;
  const maxDelta = isFinite(o.maxDelta) ? o.maxDelta : AUTOLEARN_MAX_GLOBAL_DELTA;
  const monthCap = isFinite(o.monthCap) ? o.monthCap : AUTOLEARN_MONTH_BIAS_CAP;
  const fMin = isFinite(o.fMin) ? o.fMin : AUTOLEARN_GLOBAL_FACTOR_MIN;
  const fMax = isFinite(o.fMax) ? o.fMax : AUTOLEARN_GLOBAL_FACTOR_MAX;
  const curFactor = isFinite(o.curFactor) ? o.curFactor : 1.0;

  const usable = (pairs || [])
    .map(p => {
      const a = Number(p.actual); const pr = Number(p.pred);
      if (!isFinite(a) || a === 0 || !isFinite(pr)) return null;
      return { ym: String(p.ym || ''), e: (pr - a) / Math.abs(a) };
    })
    .filter(Boolean)
    .sort((a, b) => (a.ym < b.ym ? 1 : -1)); // 新しい月から順に
  if (usable.length < minMonths) {
    return { ready: false, reason: 'insufficient_eval_months', n: usable.length };
  }
  const wByIdx = usable.map((_, i) => Math.pow(0.5, i / halfLife));
  let wSum = 0; let weSum = 0;
  usable.forEach((p, i) => { wSum += wByIdx[i]; weSum += wByIdx[i] * p.e; });
  // e>0 は過剰予測 → 補正係数は 1 - 事後平均（事前平均0・精度kGで縮小）
  const postBias = weSum / Math.max(1e-9, wSum + kG);
  const targetFactor = clamp_(1 - postBias, fMin, fMax);
  const newFactor = clamp_(curFactor + clamp_(targetFactor - curFactor, -maxDelta, maxDelta), fMin, fMax);

  const byMonth = {};
  usable.forEach((p, i) => {
    const dt = parseYM_(p.ym);
    if (!dt) return;
    const key = String(dt.getMonth() + 1);
    if (!byMonth[key]) byMonth[key] = { n: 0, w: 0, we: 0 };
    byMonth[key].n += 1;
    byMonth[key].w += wByIdx[i];
    byMonth[key].we += wByIdx[i] * p.e;
  });
  const monthBias = {};
  Object.keys(byMonth).forEach(k => {
    const g = byMonth[k];
    const b = clamp_(-(g.we / Math.max(1e-9, g.w + kM)), -monthCap, monthCap);
    if (Math.abs(b) >= 0.005) monthBias[k] = Number(b.toFixed(4));
  });
  return {
    ready: true,
    n: usable.length,
    postBias: postBias,
    factor: newFactor,
    targetFactor: targetFactor,
    curFactor: curFactor,
    monthBias: monthBias
  };
}`;

const harness = [
  ...constLines,
  "const SHEETS = { FORECAST_SNAPSHOT: 'FORECAST_SNAPSHOT' };",
  extractFunction('clamp_'),
  extractFunction('parseYM_'),
  extractFunction('toMonthStart_'),
  extractFunction('fmtYM_'),
  extractFunction('ymKey_'),
  extractFunction('normalizeClientName_'),
  extractFunction('isSameClient_'),
  extractFunction('headerIndexMap_'),
  // 検証に使う回は、その月が始まる前の回（2026-10-07 D6）
  extractFunction('evalTimeMs_'),
  extractFunction('evalJstDay_'),
  extractFunction('isRunBeforeMonth_'),
  extractFunction('selectEvalSnapshotRows_'),
  extractFunction('autoLearnComputeState_'),
  extractFunction('attachAppliedCalibration_'),
  extractFunction('parseResidualMonthBiasJson_'),
  extractFunction('canonicalMonthBiasJson_'),
  OLD_COMPUTE_SRC,
].join('\n');

// fmtYM_ は Utilities.formatDate(d,TZ,'yyyy/MM') を使う — Node 側で最小実装を注入。
// 'yyyy-MM-dd' は検証の決まり（日本の暦の日）で使う
const sandbox = { Math, JSON, Number, String, isFinite, console, TZ: 'Asia/Tokyo',
  Utilities: { formatDate: (d, _tz, fmt) => {
    if (fmt === 'yyyy-MM-dd') { const t = new Date(d.getTime() + 9 * 3600e3); return t.toISOString().slice(0, 10); }
    if (fmt !== 'yyyy/MM') throw new Error('unexpected fmt ' + fmt);
    return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}`;
  } } };
vm.createContext(sandbox);
vm.runInContext(`${harness}\n;this.__fns = { autoLearnComputeState_, parseResidualMonthBiasJson_, canonicalMonthBiasJson_, ymKey_, attachAppliedCalibration_, oldAutoLearnComputeState_ };`, sandbox);
const { autoLearnComputeState_, parseResidualMonthBiasJson_, canonicalMonthBiasJson_, ymKey_, attachAppliedCalibration_, oldAutoLearnComputeState_ } = sandbox.__fns;

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

// ---- 11. ymKey_: Sheets自動Date変換セルと文字列セルを同一キーへ揃える（EVAL結合バグの回帰） ----
{
  assert.equal(ymKey_('2026/04'), '2026/04');
  assert.equal(ymKey_(new Date(2026, 3, 1)), '2026/04', 'Date セルは文字列と同じキーになる');
  assert.equal(ymKey_(new Date(2026, 3, 15)), '2026/04', '月初以外の Date も月初に揃う');
  assert.equal(ymKey_('2026-04'), '2026/04');
  assert.equal(ymKey_('2026/4'), '2026/04', 'ゼロ埋めなし表記も揃う');
  assert.equal(ymKey_(''), '', '空は空のまま');
  // FORECAST_SNAPSHOT(Date) × ACTUAL_EVAL_MONTHLY('yyyy/MM'文字列) の結合が成立する
  const snapYm = ymKey_(new Date(2026, 3, 1));
  const actYm = ymKey_('2026/04');
  assert.equal(snapYm, actYm, 'sheet 読み出し値が異なっても結合キー一致');
}

// ==== 2026-10-07: 補正を掛ける前の予測で学ぶ（B-5 が自分の補正を打ち消さない） ====
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const ymOf = (t) => { const y = 2023 + Math.floor((t + 3) / 12); const m = (t + 3) % 12 + 1; return `${y}/${String(m).padStart(2, '0')}`; };   // t=0 が 2023/04
/** 決まった種の乱数（テストの値を毎回同じにする） */
function rng(seed) { let s = seed >>> 0; return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; }; }
/** 今までの結果と、同じ項目が 1 ビットも違わないか */
function sameAsOld(a, b, msg) {
  Object.keys(b).forEach(k => assert.ok(k in a, `${msg}: ${k} がない`));
  for (const k of ['ready', 'reason', 'n', 'postBias', 'factor', 'targetFactor', 'curFactor']) assert.ok(Object.is(a[k], b[k]), `${msg}: ${k} ${a[k]} / ${b[k]}`);
  assert.equal(JSON.stringify(a.monthBias), JSON.stringify(b.monthBias), `${msg}: monthBias`);
}

// ---- 12. 補正が掛かっていない（係数 1・暦月バイアス 0）・分からない組は、今までとまったく同じ結果 ----
{
  const r = rng(7);
  for (let k = 0; k < 200; k++) {
    const n = 1 + Math.floor(r() * 30);
    const start = Math.floor(r() * 12);
    const pairs = Array.from({ length: n }, (_, i) => ({ ym: ymOf(start + i), pred: 50 + r() * 100, actual: r() < 0.05 ? 0 : 50 + r() * 100 }));
    if (k % 7 === 0) pairs.push({ ym: ymOf(start + n), pred: 'x', actual: 100 });
    const cur = 0.75 + r() * 0.5;
    const old = oldAutoLearnComputeState_(pairs, { curFactor: cur });
    const unknowns = [undefined, null, '', 'x', 0, -1, NaN];
    const variants = [
      pairs,                                                                              // 何も付いていない（今までの組）
      pairs.map(p => ({ ...p, factorUsed: 1, monthBiasUsed: 0 })),                          // 補正なしの回
      pairs.map((p, i) => ({ ...p, factorUsed: unknowns[i % 7], monthBiasUsed: unknowns[(i + 3) % 7] })),   // 読めない値は補正なし
    ];
    variants.forEach((v, j) => {
      const res = autoLearnComputeState_(v, { curFactor: cur });
      sameAsOld(res, old, `set ${k} variant ${j}`);
      if (res.ready) { assert.equal(res.method, 'raw'); assert.equal(res.corrected, 0); }
    });
  }
}

// ---- 13. 補正が掛かった予測は割り戻してから学ぶ（1 回分） ----
{
  // 全体: 係数 0.93・暦月バイアスを掛けた予測は、補正の前の予測（raw）で今までの式に入れたのと同じ
  // 暦月: 暦月バイアスだけを割り戻した予測（raw × 係数）で今までの式に入れたのと同じ（全体の係数を掛けた後に残る誤差を直す分）
  const r = rng(11);
  const raw = Array.from({ length: 14 }, (_, i) => ({ ym: ymOf(i), raw: 80 + r() * 60, actual: 80 + r() * 60, f: 0.85 + r() * 0.3, mb: -0.2 + r() * 0.4 }));
  const corrected = raw.map(p => ({ ym: p.ym, pred: p.raw * p.f * (1 + p.mb), actual: p.actual, factorUsed: p.f, monthBiasUsed: p.mb }));
  const res = autoLearnComputeState_(corrected, { curFactor: 0.97 });
  const g = oldAutoLearnComputeState_(raw.map(p => ({ ym: p.ym, pred: p.raw, actual: p.actual })), { curFactor: 0.97 });
  const m = oldAutoLearnComputeState_(raw.map(p => ({ ym: p.ym, pred: p.raw * p.f, actual: p.actual })), { curFactor: 0.97 });
  for (const k of ['postBias', 'targetFactor', 'factor']) assert.ok(Math.abs(res[k] - g[k]) < 1e-12, `${k}: ${res[k]} / ${g[k]}`);
  assert.equal(JSON.stringify(res.monthBias), JSON.stringify(m.monthBias));
  assert.equal(res.corrected, 14);

  // 報告の例: 6 か月とも 10% 多い予測に係数 0.95 が掛かっていた（pred 104.5・実績 100）
  const six = Array.from({ length: 6 }, (_, i) => ({ ym: ymOf(i), pred: 110 * 0.95, actual: 100, factorUsed: 0.95, monthBiasUsed: 0 }));
  const before = oldAutoLearnComputeState_(six, { curFactor: 0.95 });
  const after = autoLearnComputeState_(six, { curFactor: 0.95 });
  assert.ok(before.factor > 0.97, `今まで: 残りの誤差 4.5% だけを見て 1 の側へ戻る ${before.factor}`);
  assert.ok(after.factor < 0.95 && after.factor > 0.94, `今後: 補正の前の誤差 10% を見て下げ続ける ${after.factor}`);
  assert.ok(Math.abs(after.postBias - before.postBias / 0.45) < 1e-12, '縮め方（半減期・事前の重み）は同じ');
}

// ---- 14. 何か月も回す: 10% 多い予測が続くと、今後は縮めた目標に落ち着いて戻らない。今までは 1 の側へ戻っていた ----
/** 毎月: それまでに学んだ係数と暦月バイアス（A-9 と同じく ±0.25 で切る）で予測し、実績が出たら学ぶ（B-2 の後の B-5） */
function simulate(learn, rawBias, months, opts) {
  let F = 1; let MB = {};
  const pairs = []; const path = [];
  for (let t = 0; t < months; t++) {
    const ym = ymOf(t); const m = Number(ym.slice(5));
    const mbU = clamp(Number(MB[String(m)] || 0), -0.25, 0.25);
    const pred = 100 * (1 + rawBias(m)) * F * (1 + mbU);
    pairs.push({ ym, pred, actual: 100, factorUsed: F, monthBiasUsed: mbU });
    const prev = F;
    const res = learn(pairs, Object.assign({ curFactor: F }, opts || {}));
    if (res.ready) { F = res.factor; MB = res.monthBias; }
    path.push({ ym, m, prev, F, err: pred / 100 - 1, ready: res.ready });
  }
  return path;
}
{
  const N = 48;
  const nw = simulate(autoLearnComputeState_, () => 0.10, N);
  const od = simulate(oldAutoLearnComputeState_, () => 0.10, N);
  // 今後の落ち着き先: 今までと同じ縮め方 1 − 0.10 × S / (S + 3)（S = 48 か月の重みの和 ≈ 6.28）≈ 0.932
  const S = Array.from({ length: N }, (_, i) => Math.pow(0.5, i / 4)).reduce((s, x) => s + x, 0);
  const fStar = 1 - 0.10 * S / (S + 3);
  assert.ok(Math.abs(fStar - 0.9323) < 1e-4, String(fStar));
  const fN = nw[N - 1].F; const fO = od[N - 1].F;
  assert.ok(Math.abs(fN - fStar) < 1e-3, `今後は ${fStar.toFixed(4)} に落ち着く: ${fN}`);
  nw.forEach(p => assert.ok(p.F <= p.prev + 1e-12, `今後は一度も 1 の側へ戻らない: ${p.ym} ${p.prev} → ${p.F}`));
  const last12 = nw.slice(-12).map(p => p.F);
  assert.ok(Math.max(...last12) - Math.min(...last12) < 1e-3, '落ち着いたまま');
  // 今まで: 一度下げた係数が 1 の側へ戻り、0.96 あたりに止まる（10% の偏りの半分ほどしか直らない）
  const firstO = od.find(p => p.ready).F;
  assert.ok(Math.max(...od.map(p => p.F)) > firstO + 0.005, `今までは下げた後に戻る: ${firstO} → ${Math.max(...od.map(p => p.F))}`);
  assert.ok(fO > fStar + 0.02, `今まで ${fO} / 今後 ${fN}`);
  // 補正の後に残る誤差（最後の月の予測）: 今後 約 +2.6%、今まで 約 +5.7%
  assert.ok(nw[N - 1].err < 0.03 && od[N - 1].err > 0.05, `残る誤差 今後 ${nw[N - 1].err} / 今まで ${od[N - 1].err}`);
  // どちらも 1 回の幅 0.05 と上下限を守る
  [nw, od].forEach(path => path.forEach(p => { assert.ok(Math.abs(p.F - p.prev) <= 0.05 + 1e-12); assert.ok(p.F >= 0.75 && p.F <= 1.25); }));

  // 縮めなければ（kG=0。テストのためだけ）今後は 1 − 0.10 = 0.900 へ。今までは 0.952（= 2 / 2.1）の辺りで止まる。
  // 0.909（= 1 / 1.1）との差は 1 − postBias の近似によるもの（式は変えていない）
  const nw0 = simulate(autoLearnComputeState_, () => 0.10, N, { kG: 0 });
  const od0 = simulate(oldAutoLearnComputeState_, () => 0.10, N, { kG: 0 });
  assert.ok(Math.abs(nw0[N - 1].F - 0.9) < 1e-3, String(nw0[N - 1].F));
  assert.ok(od0[N - 1].F > 0.94, String(od0[N - 1].F));
}

// ---- 15. 暦月バイアスも自分を打ち消さない ----
{
  // 4 月だけ 20% 多い。前の 4 月の予測には暦月バイアス −0.1667 が掛かっていて、ちょうど直っていた（pred ≈ 実績）
  const data = [
    { ym: '2026/05', pred: 100, actual: 100 },
    { ym: '2026/04', pred: 120 * (1 - 0.1667), actual: 100, factorUsed: 1, monthBiasUsed: -0.1667 },
    { ym: '2026/03', pred: 100, actual: 100 },
    { ym: '2026/02', pred: 100, actual: 100 },
  ];
  const before = oldAutoLearnComputeState_(data);
  const after = autoLearnComputeState_(data);
  assert.equal(before.monthBias['4'], undefined, '今まで: 直った 4 月の誤差は 0 に見え、4 月のバイアスを忘れる');
  assert.ok(after.monthBias['4'] < -0.05, `今後: 割り戻した 4 月は 20% 多いので、負のバイアスが残る ${after.monthBias['4']}`);
  assert.equal(after.corrected, 1);

  // 何年も回す: 全体で 10% 多く、4 月はさらに多い（4 月は 32%）。今後は 4 月もほかの月も、残る誤差が小さく、年ごとに変わらない
  const bias = (m) => (m === 4 ? 0.32 : 0.10);
  const nw = simulate(autoLearnComputeState_, bias, 60);
  const od = simulate(oldAutoLearnComputeState_, bias, 60);
  const aprN = nw.filter(p => p.m === 4).map(p => p.err); const aprO = od.filter(p => p.m === 4).map(p => p.err);
  assert.equal(aprN.length, 5);
  assert.ok(Math.abs(aprN[4] - aprN[3]) < 0.002 && Math.abs(aprN[4] - aprN[2]) < 0.002, `今後の 4 月の誤差は年ごとに変わらない: ${aprN}`);
  assert.ok(aprN[4] < aprO[4] - 0.03, `4 月: 今後 ${aprN[4]} / 今まで ${aprO[4]}`);
  assert.ok(nw[58].err < od[58].err - 0.03, `ほかの月: 今後 ${nw[58].err} / 今まで ${od[58].err}`);
}

// ---- 16. 割り戻しても、上下限・1 回の幅・暦月の上限は今までどおり ----
{
  // 極端な過剰予測（さらに補正が下限まで掛かっていた）: 目標は下限 0.75、1 回は 0.05 まで
  const big = pairs(12).map(p => ({ ...p, pred: 600 * 0.75 * 0.75, factorUsed: 0.75, monthBiasUsed: -0.25 }));
  let r1 = autoLearnComputeState_(big);
  assert.equal(r1.targetFactor, 0.75);
  assert.equal(r1.factor, 0.95);
  r1 = autoLearnComputeState_(big, { curFactor: 0.77 });
  assert.equal(r1.factor, 0.75, '下限より下には行かない');
  // 極端な過小予測（補正が上限まで掛かっていた）: 目標は上限 1.25
  const small = pairs(12).map(p => ({ ...p, pred: 10 * 1.25 * 1.25, factorUsed: 1.25, monthBiasUsed: 0.25 }));
  const r2 = autoLearnComputeState_(small, { curFactor: 1.24 });
  assert.equal(r2.targetFactor, 1.25);
  assert.equal(r2.factor, 1.25);
  // ばらばらな値で: 係数は [0.75, 1.25]・1 回の幅は ±0.05・暦月は ±0.20（0.005 未満は書かない）
  const r = rng(23);
  for (let k = 0; k < 300; k++) {
    const n = 3 + Math.floor(r() * 30);
    const ps = Array.from({ length: n }, (_, i) => ({ ym: ymOf(i), pred: 1 + r() * 400, actual: 1 + r() * 200, factorUsed: 0.75 + r() * 0.5, monthBiasUsed: -0.25 + r() * 0.5 }));
    const cur = 0.75 + r() * 0.5;
    const res = autoLearnComputeState_(ps, { curFactor: cur });
    assert.ok(res.ready);
    assert.ok(res.targetFactor >= 0.75 && res.targetFactor <= 1.25);
    assert.ok(res.factor >= 0.75 && res.factor <= 1.25);
    assert.ok(Math.abs(res.factor - cur) <= 0.05 + 1e-12);
    Object.values(res.monthBias).forEach(b => assert.ok(Math.abs(b) <= 0.20 && Math.abs(b) >= 0.005, String(b)));
  }
}

// ---- 17. attachAppliedCalibration_: EVAL_LOG の pred を作った回の補正を、B-2 と同じ選び方で読む ----
{
  const HEAD = ['snapshot_id', 'run_date', 'client', 'target_month', 'scenario', 'base_pred', 'subjective_adj', 'ai_adj', 'deterministic_adj', 'final_pred',
    'confidence_interval_lower', 'confidence_interval_upper', 'key_factors_json', 'subjective_input_date', 'calibration_applied_json'];
  const row = (sid, client, ym, sc, pred, source, cal) => [sid, '2026-01-01', client, ym, sc, pred, 0, 0, 0, pred, 0, 0,
    JSON.stringify(source === undefined ? { opinion: '' } : { opinion: '', forecast_source: source }), '', typeof cal === 'string' ? cal : JSON.stringify(cal)];
  const run = (sid, client, ym, p50, source, cal) => [row(sid, client, ym, 'nega', p50 * 0.9, source, cal), row(sid, client, ym, 'neutral', p50, source, cal), row(sid, client, ym, 'posi', p50 * 1.1, source, cal)];
  const book = (rows, head) => ({ getSheetByName: (n) => (n === 'FORECAST_SNAPSHOT' && rows ? { getLastRow: () => rows.length + 1, getDataRange: () => ({ getValues: () => [head || HEAD, ...rows] }) } : null) });
  const used = (ss, pairs) => JSON.parse(JSON.stringify(attachAppliedCalibration_(ss, '甲製薬', pairs))).map(p => [p.ym, p.factorUsed, p.monthBiasUsed]);
  const c = (f, mb) => ({ version: 'x', bias_correction_factor: f, residual_month_bias_json: mb || '' });
  const rows = [
    // 2026/04: 1 回目（係数 0.95）→ 2 回目（係数 0.9・4 月 −0.3 は A-9 と同じく −0.25 で切る）→ 締まった後の回（使わない）
    ...run('S1', '甲製薬', '2026/04', 104.5, 'forecast_open', c(0.95, '{"4":-0.05}')),
    ...run('S2', '甲製薬', '2026/04', 110 * 0.9 * 0.75, 'forecast_open', c(0.9, '{"4":-0.3,"5":0.1}')),
    ...run('S3', '甲製薬', '2026/04', 100, 'actual_closed', c(0.8, '{"4":-0.1}')),
    // 2026/05: 前からの記録（forecast_source なし）。2 回目は実績と同じ P50（締まった後の回）なので 1 回目を使う
    ...run('T1', '甲製薬', '2026/05', 110 * 1.1, undefined, c(1, '{"5":0.1}')),
    ...run('T2', '甲製薬', '2026/05', 100, undefined, c(0.7, '')),
    // 2026/06: 別のクライアントの回だけ → 補正なし
    ...run('U1', '乙製薬', '2026/06', 90, 'forecast_open', c(0.8, '')),
    // 2026/07: 読めない JSON → 補正なし
    ...run('V1', '甲製薬', '2026/07', 95, 'forecast_open', '{broken'),
    // 2026/08: 2 回目は EVAL_LOG の後に予測し直した回（pred が合わない）→ 補正なし（今までと同じ扱い）
    ...run('W1', '甲製薬', '2026/08', 99, 'forecast_open', c(0.9, '')),
    ...run('W2', '甲製薬', '2026/08', 97, 'forecast_open', c(0.85, '')),
  ];
  const evalPairs = [
    { ym: '2026/04', pred: 110 * 0.9 * 0.75, actual: 100 },
    { ym: '2026/05', pred: 110 * 1.1, actual: 100 },
    { ym: '2026/06', pred: 90, actual: 100 },
    { ym: '2026/07', pred: 95, actual: 100 },
    { ym: '2026/08', pred: 99, actual: 100 },
    { ym: '2026/09', pred: 100, actual: 100 },   // 回が無い月
  ];
  const before = JSON.stringify(evalPairs);
  assert.deepEqual(used(book(rows), evalPairs), [
    ['2026/04', 0.9, -0.25], ['2026/05', 1, 0.1], ['2026/06', 1, 0], ['2026/07', 1, 0], ['2026/08', 1, 0], ['2026/09', 1, 0]]);
  assert.equal(JSON.stringify(evalPairs), before, '元の組は書き換えない');
  // 割り戻すと、2026/04 は補正の前の 110・2026/05 も 110 に戻る
  const res = autoLearnComputeState_(attachAppliedCalibration_(book(rows), '甲製薬', evalPairs));
  assert.equal(res.corrected, 2);
  // シートが無い・列が無い・空のとき: 全部補正なし（今までと同じ）
  const none = evalPairs.map(p => [p.ym, 1, 0]);
  assert.deepEqual(used(book(null), evalPairs), none);
  assert.deepEqual(used(book(rows, HEAD.slice(0, 14)), evalPairs), none);
  assert.deepEqual(used(book([]), evalPairs), none);
  assert.deepEqual(used(book(rows), []), []);
}

process.stdout.write('PASS forecast-autolearn contract tests\n');
