#!/usr/bin/env node
/*
 * forecast-closed-months.test.mjs — 検証の決まり（2026-10-07 村井さん承認 D4〜D6）を、旧来の計算（Forecast_Agent.js）の関数を取り出して Node で確かめる。
 *   1. 締まり（D5）: 月末 + 5 日 <= 取り込んだ日（日本の暦）。月末から 3 日後の取り込みは締まっていない・6 日後は締まった。境目・時差・年またぎ・2 月
 *   2. 測る予測（D6）: その月が始まる前（予測した日 < 月の 1 日）の最後の回。後から作った回は使わない。前の回が無い月は測らず、数える
 *   3. 締まりの数え直し: ACTUAL_EVAL_MONTHLY の取り込んだ日時から。前の決まりで付いた closed の印をうのみにしない
 *   4. EVAL_LOG の行の選び方（今の検証の版で書いた・締まった月の行だけ）と、C-1 の当たりに使う回の選び方
 *   5. 予測（A-9 runPhase1Forecast）からは、変えた関数・新しい数を呼ばず、検証の表も読まない（予測の数字は変わらない）
 * 本物の Apps Script での確認の代わりではない（モックも使わない純粋な関数の確認）。
 *
 *   node tests/forecast-closed-months.test.mjs
 */
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const src = await readFile(path.join(root, 'Forecast_Agent.js'), 'utf8');
function extractFunction(name) {
  const start = src.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} not found`);
  let depth = 0;
  for (let p = src.indexOf('{', start); p < src.length; p++) {
    if (src[p] === '{') depth++;
    else if (src[p] === '}') { depth--; if (depth === 0) return src.slice(start, p + 1); }
  }
  throw new Error(name + ' not closed');
}
const constLines = src.match(/^const (EVALUATION_POLICY_VERSION|ACTUAL_CLOSE_LAG_DAYS|EVAL_CALENDAR_TZ) = .+;.*$/gm) || [];
assert.equal(constLines.length, 3, '検証の決まりの数が 1 か所ずつ');
const FNS = ['normalizeClientName_', 'isSameClient_', 'toMonthStart_', 'fmtYM_', 'ymKey_', 'headerIndexMap_', 'evalTimeMs_', 'evalJstDay_', 'isActualMonthClosed_',
  'isRunBeforeMonth_', 'closedActualMonthsOf_', 'isEvalLogRowCurrent_', 'keepPreMonthRunRows_', 'selectEvalSnapshotRows_'];

/** Utilities.formatDate の代わり（日本の時刻。'yyyy/MM' と 'yyyy-MM-dd' だけ） */
const formatDate = (d, tz, fmt) => {
  assert.equal(tz, 'Asia/Tokyo');
  const t = new Date(d.getTime() + 9 * 3600e3);
  const p = (n) => String(n).padStart(2, '0');
  if (fmt === 'yyyy/MM') return t.getUTCFullYear() + '/' + p(t.getUTCMonth() + 1);
  if (fmt === 'yyyy-MM-dd') return t.getUTCFullYear() + '-' + p(t.getUTCMonth() + 1) + '-' + p(t.getUTCDate());
  throw new Error('unexpected fmt ' + fmt);
};
const ctx = vm.createContext({ TZ: 'Asia/Tokyo', Utilities: { formatDate } });
vm.runInContext([...constLines, ...FNS.map(extractFunction)].join('\n') + '\n;this.__c = { EVALUATION_POLICY_VERSION, ACTUAL_CLOSE_LAG_DAYS, EVAL_CALENDAR_TZ };', ctx);
const fn = (name) => (...args) => { ctx.__a = args; return vm.runInContext(`${name}(...__a)`, ctx); };
/** vm の中で作った値を、こちらの値に（配列の比べ方を合わせる） */
const run_ = (code) => JSON.parse(vm.runInContext(`JSON.stringify(${code})`, ctx));
const isClosed = fn('isActualMonthClosed_');
const isBefore = fn('isRunBeforeMonth_');
/** 日本の時刻（月は 1 始まり） */
const jst = (y, m, d, H = 0, M = 0) => new Date(Date.UTC(y, m - 1, d, H - 9, M));

// ==== 1. 締まり（D5） ====
assert.equal(ctx.__c.ACTUAL_CLOSE_LAG_DAYS, 5, '月末から 5 日');
assert.equal(ctx.__c.EVAL_CALENDAR_TZ, 'Asia/Tokyo', '日本の暦');
assert.equal(ctx.__c.EVALUATION_POLICY_VERSION, 'policy-2026H1-v3', '検証の版を上げた（新しい行を見分ける）');
assert.equal(isClosed('2026/09', jst(2026, 10, 3, 12)), false, '月末から 3 日後に取り込んだ 9 月は締まっていない');
assert.equal(isClosed('2026/09', jst(2026, 10, 6, 9)), true, '6 日後に取り込んだ 9 月は締まった');
assert.equal(isClosed('2026/09', jst(2026, 10, 5, 0, 0)), true, '境目: 月末 + 5 日（10/05 0:00）は締まった');
assert.equal(isClosed('2026/09', jst(2026, 10, 4, 23, 59)), false, '境目: 10/04 23:59 は締まっていない');
assert.equal(isClosed('2026/09', new Date(Date.UTC(2026, 9, 4, 15, 0))), true, '時差: UTC 10/04 15:00 は日本の 10/05');
assert.equal(isClosed('2026/09', new Date(Date.UTC(2026, 9, 4, 14, 59))), false, '時差: UTC 10/04 14:59 は日本の 10/04');
assert.equal(isClosed('2026/08', jst(2026, 10, 3, 12)), true, '前の月はとうに締まった');
assert.equal(isClosed('2026/10', jst(2026, 10, 30)), false, '取り込んだ月そのものは締まっていない');
assert.equal(isClosed('2026/12', jst(2027, 1, 5)), true, '年またぎ: 12 月は 1/05 から');
assert.equal(isClosed('2026/12', jst(2027, 1, 4, 23, 59)), false);
assert.equal(isClosed('2028/02', jst(2028, 3, 5)), true, 'うるう年の 2 月（29 日 + 5 日 = 3/05）');
assert.equal(isClosed('2028/02', jst(2028, 3, 4, 23, 59)), false);
assert.equal(isClosed('2027/02', jst(2027, 3, 5)), true, '2 月（28 日 + 5 日 = 3/05）');
assert.equal(isClosed('2027/02', jst(2027, 3, 4, 23, 59)), false);
assert.equal(isClosed(new Date(2026, 8, 1), jst(2026, 10, 6)), true, '月が日付のセルでも同じ');
assert.equal(isClosed('2026/09', '2026/10/05 00:00'), true, "文字の日時（'yyyy/MM/dd HH:mm'）は日本の時刻");
assert.equal(isClosed('2026/09', '2026/10/04 23:59:59'), false);
assert.equal(isClosed('2026/09', '2026-10-04'), false, "文字の日（'yyyy-MM-dd'）も日本の暦");
assert.equal(isClosed('2026/09', '2026-10-04T15:00:00.000Z'), true, 'ISO の時刻は時差を見る');
assert.equal(isClosed('2026/09', '2026-10-05T00:00:00+0900'), true);
for (const bad of ['', null, undefined, 'abc', new Date(NaN)]) assert.equal(isClosed('2026/09', bad), false, '取り込んだ日が読めなければ締まっていない: ' + bad);
assert.equal(isClosed('x', jst(2026, 10, 6)), false, '月が読めなければ締まっていない');

// ==== 2. 測る予測（D6）: 予測した日 < 月の 1 日（日本の暦） ====
assert.equal(isBefore(jst(2026, 9, 30, 23, 59), '2026/10'), true, '前の月の最後の日');
assert.equal(isBefore(jst(2026, 10, 1, 0, 0), '2026/10'), false, 'その月の 1 日に作った予測は前ではない');
assert.equal(isBefore('2026-09-30T15:00:00Z', '2026/10'), false, 'UTC 9/30 15:00 は日本の 10/01');
assert.equal(isBefore('2026-01-01', '2026/04'), true);
assert.equal(isBefore(jst(2026, 6, 19), new Date(2026, 6, 1)), true, '月が日付のセルでも同じ');
assert.equal(isBefore('', '2026/07'), false, '予測した日が読めなければ前とは言えない');

// FORECAST_SNAPSHOT の行（予測した日は日付の型。シートから読んだ値と同じ）
const snapRow = (sid, at, client, ym, sc, pred, source) => [sid, at, client, ym, sc, pred, 0, 0, 0, pred, 0, 0,
  JSON.stringify(source === undefined ? { opinion: '' } : { opinion: '', forecast_source: source }), '', '{}'];
const run = (sid, at, client, ym, p50, source) => ['nega', 'neutral', 'posi'].map((sc, i) => snapRow(sid, at, client, ym, sc, p50 * [0.9, 1, 1.1][i], source));
const pick = (rows, actual) => {
  ctx.__rows = rows; ctx.__act = new Map(Object.entries(actual || {})); ctx.__st = {};
  const out = run_('selectEvalSnapshotRows_(__rows, __act, __st).map(r => r[0] + ":" + r[3] + ":" + r[4])');
  return { sids: out, noPre: run_('__st.noPreMonth').sort() };
};
{
  const S1 = jst(2026, 3, 20), S2 = jst(2026, 6, 19), S3 = jst(2026, 7, 15), S4 = jst(2026, 10, 2);
  // 2026/07: 3/20・6/19 の回は月が始まる前、7/15・10/02 の回は後 → 6/19（前の回のうち最後）。後の回は新しくても使わない
  let r = pick([...run('S1', S1, '甲製薬', '2026/07', 110), ...run('S2', S2, '甲製薬', '2026/07', 120), ...run('S3', S3, '甲製薬', '2026/07', 130),
    ...run('S4', S4, '甲製薬', '2026/07', 140)]);
  assert.deepEqual(r.sids, ['S2:2026/07:nega', 'S2:2026/07:neutral', 'S2:2026/07:posi'], '月が始まる前の最後の回');
  assert.deepEqual(r.noPre, [], '前の回があれば数えない');
  // シートの順が前後しても、予測した日で選ぶ
  r = pick([...run('S2', S2, '甲製薬', '2026/07', 120), ...run('S1', S1, '甲製薬', '2026/07', 110)]);
  assert.deepEqual([...new Set(r.sids.map((s) => s.split(':')[0]))], ['S2']);
  // 予測した日が同じなら、今までどおりシートの下の行
  r = pick([...run('T1', S1, '甲製薬', '2026/07', 110), ...run('T2', S1, '甲製薬', '2026/07', 111)]);
  assert.deepEqual([...new Set(r.sids.map((s) => s.split(':')[0]))], ['T2']);
  // 2026/06: 6/19 と 10/02 の回だけ（どちらも月が始まった後）→ 測らない。数える
  r = pick([...run('S2', S2, '甲製薬', '2026/06', 120), ...run('S4', S4, '甲製薬', '2026/06', 140), ...run('S1', S1, '甲製薬', '2026/07', 110)]);
  assert.deepEqual([...new Set(r.sids.map((s) => s.split(':')[0] + ':' + s.split(':')[1]))], ['S1:2026/07']);
  assert.deepEqual(r.noPre, ['甲製薬|2026/06'], '月が始まる前の予測が無い月');
  // 前の回が「締まった後の回」（今までの印）だけなら、それも使わない
  r = pick([...run('U1', S1, '甲製薬', '2026/07', 100), ...run('U2', S4, '甲製薬', '2026/07', 90)], { '甲製薬|2026/07': 100 });
  assert.deepEqual(r.sids, [], '予測 = 実績の回は使わない（今までどおり）');
  assert.deepEqual(r.noPre, ['甲製薬|2026/07']);
  r = pick([...run('V1', S1, '甲製薬', '2026/07', 95, 'forecast_open'), ...run('V2', S2, '甲製薬', '2026/07', 99, 'actual_closed')]);
  assert.deepEqual([...new Set(r.sids.map((s) => s.split(':')[0]))], ['V1'], 'actual_closed の回は使わない（今までどおり）');
  // クライアントごと
  r = pick([...run('A1', S1, '甲製薬', '2026/07', 110), ...run('B1', S2, '乙製薬', '2026/07', 220), ...run('B2', S4, '乙製薬', '2026/08', 230)]);
  assert.deepEqual([...new Set(r.sids.map((s) => s.split(':')[0]))].sort(), ['A1', 'B1']);
  assert.deepEqual(r.noPre, ['乙製薬|2026/08'], '乙製薬の 8 月の回は 10/02（月が始まった後）だけ');
  // 予測した日が読めない回は使わない
  r = pick(run('W1', 'not a date', '甲製薬', '2026/07', 110));
  assert.deepEqual(r.sids, []);
}

// ==== 3. 締まりの数え直し（ACTUAL_EVAL_MONTHLY） ====
{
  const A = (client, ym, flag, at, product = '製品A') => [client, 'BASE', product, ym, 100, flag, at];
  ctx.__rows = [
    A('甲製薬', '2026/08', 'closed', jst(2026, 10, 3, 9)),
    A('甲製薬', '2026/09', 'closed', jst(2026, 10, 3, 9)),   // 前の決まり（取り込んだ日の暦）では closed だった
    A('甲製薬', '2026/10', 'open', jst(2026, 10, 3, 9)),
    A('乙製薬', '2026/09', 'closed', jst(2026, 10, 6, 9)),
    A('乙製薬', '2026/08', 'open', jst(2026, 10, 6, 9)),       // 印が open なら締まっていない
    A('丙製薬', '2026/08', 'closed', jst(2026, 10, 6, 9)),
    A('丙製薬', '2026/08', 'closed', '', '製品B'),               // 日時の無い行が 1 つでもあれば、その月は締まっていない
    A('丁製薬', '2026/07', 1, jst(2026, 8, 6)),                  // 数の印（1）
    A('丁製薬', '2026/06', 0, jst(2026, 8, 6)),                  // 数の印（0）は締まっていない
    A('丁製薬', new Date(2026, 4, 1), 'closed', jst(2026, 8, 6)),   // 月が日付のセル
  ];
  const res = run_('(() => { const r = closedActualMonthsOf_(__rows); return { keys: [...r.keys].sort(), months: [...r.months].sort() }; })()');
  assert.deepEqual(res.keys, ['乙製薬|2026/09', '甲製薬|2026/08', '丁製薬|2026/05', '丁製薬|2026/07'].sort());
  assert.deepEqual(res.months, ['2026/05', '2026/07'], '月だけで見るときは、どのクライアントの行も締まっていること');
  ctx.__rows = [];
  assert.deepEqual(run_('[...closedActualMonthsOf_(__rows).keys]'), []);
}

// ==== 4. EVAL_LOG の行の選び方 ====
{
  const HEAD = ['eval_id', 'evaluated_at', 'client', 'target_month', 'scenario', 'pred', 'actual', 'evaluation_policy_version'];
  ctx.__idx = vm.runInContext(`headerIndexMap_(${JSON.stringify(HEAD)})`, ctx);
  ctx.__closed = new Set(['甲製薬|2026/07']);
  const cur = (row, closed = '__closed') => { ctx.__r = row; return vm.runInContext(`isEvalLogRowCurrent_(__r, __idx, ${closed})`, ctx); };
  const R = (ym, ver, client = '甲製薬') => ['E', jst(2026, 10, 6), client, ym, 'neutral', 110, 100, ver];
  assert.equal(cur(R('2026/07', 'policy-2026H1-v3')), true, '今の版・締まった月');
  assert.equal(cur(R(new Date(2026, 6, 1), 'policy-2026H1-v3')), true, '月が日付のセルでも');
  assert.equal(cur(R('2026/07', 'policy-2026H1-v2')), false, '前の版の行は使わない（後から作った予測で測っている）');
  assert.equal(cur(R('2026/07', '')), false);
  assert.equal(cur(R('2026/10', 'policy-2026H1-v3')), false, '締まっていない月は使わない');
  assert.equal(cur(R('2026/07', 'policy-2026H1-v3', '乙製薬')), false, 'クライアントごと');
  assert.equal(cur(R('2026/10', 'policy-2026H1-v3'), 'null'), true, 'ACTUAL_EVAL_MONTHLY が無いブックは版だけで判断');
  ctx.__idx = vm.runInContext(`headerIndexMap_(${JSON.stringify(HEAD.slice(0, 7))})`, ctx);
  assert.equal(cur(R('2026/07', 'policy-2026H1-v3')), false, '版の列が無ければ使わない');
}

// ==== 4b. C-1 の当たりに使う回（AI_IMPACT_HISTORY・SUBJECTIVE_IMPACT_HISTORY） ====
{
  const HEAD = ['run_id', 'run_at', 'client', 'target_month', 'k_ai', 'ai_direction'];
  const R = (id, at, ym, dir) => [id, at, '甲製薬', ym, 1, dir];
  ctx.__rows = [
    R('R0', jst(2026, 3, 20), new Date(2026, 6, 1), 'down'), R('R1', jst(2026, 6, 19), new Date(2026, 6, 1), 'up'), R('R2', jst(2026, 10, 2), new Date(2026, 6, 1), 'down'),
    R('R0', jst(2026, 3, 20), new Date(2026, 5, 1), 'down'), R('R1', jst(2026, 6, 19), new Date(2026, 5, 1), 'up'),
    R('R2', jst(2026, 10, 2), new Date(2026, 8, 1), 'up'),
    R('R1', 'x', new Date(2026, 7, 1), 'up'),
  ];
  const keep = (head) => run_(`keepPreMonthRunRows_(__rows, headerIndexMap_(${JSON.stringify(head)})).map(r => r[0] + ':' + ymKey_(r[3]) + ':' + r[5])`);
  assert.deepEqual(keep(HEAD), ['R1:2026/07:up', 'R0:2026/06:down'], '月ごとに、始まる前の最後の回だけ（後の回・日時の読めない回は使わない）');
  // run_id が無ければ run_at で回を分ける
  ctx.__rows = ctx.__rows.map((r) => ['', ...r.slice(1)]);
  assert.deepEqual(keep(HEAD), [':2026/07:up', ':2026/06:down']);
  assert.deepEqual(keep(HEAD.filter((h) => h !== 'run_at')), [], 'run_at の列が無ければ、前の回とは言えない');
}

// ==== 5. 予測（A-9）は変わらない: runPhase1Forecast から呼ぶ関数に、今回変えた関数・新しい数・検証の表の読み取りが無い ====
{
  // 関数の呼び出し（名前のあとに '('、または引数として渡す）だけをたどる。コメントの行は除く
  const code = src.split('\n').filter((l) => !/^\s*(\/\*|\*|\/\/)/.test(l)).map((l) => l.replace(/\s\/\/\s.*$/, '')).join('\n');
  const decl = [...code.matchAll(/^function (\w+)\s*\(/gm)];
  const bodies = new Map(decl.map((m, k) => [m[1], code.slice(m.index, k + 1 < decl.length ? decl[k + 1].index : code.length)]));
  const names = [...bodies.keys()];
  const calls = (body, n) => new RegExp('\\b' + n + '\\s*\\(|[(,]\\s*' + n + '\\s*[),]').test(body);
  const reach = (from) => {
    const seen = new Set(); const stack = [from];
    while (stack.length) { const f = stack.pop(); if (seen.has(f)) continue; seen.add(f); const b = bodies.get(f) || ''; names.forEach((n) => { if (n !== f && calls(b, n)) stack.push(n); }); }
    return seen;
  };
  const a9 = reach('runPhase1Forecast');
  assert.ok(a9.has('runForecastFYCore_') && a9.has('applyCalibrationToTuning_') && a9.has('writeForecastArtifacts_'), 'たどり方の確認（予測の中身の関数をたどれている）');
  const CHANGED = ['importMonthlyFromExternal_', 'selectEvalSnapshotRows_', 'updatePhase1EvaluationReport', 'writeEvalCompareMonthly_', 'updatePhase1Dashboard',
    'updatePhase1LearningInsights', 'collectNeutralEvalPairs_', 'attachAppliedCalibration_', 'runMonthlyAutoLearn_', 'collectQuarterlyReviewData_',
    'generateQuarterlyProposals_', 'computeReliabilityHitStats_', ...FNS.slice(6), 'readClosedActualMonths_', 'readCurrentEvalMonths_'];
  CHANGED.forEach((n) => assert.ok(bodies.has(n), n));
  assert.deepEqual(CHANGED.filter((n) => a9.has(n)), [], '予測からは、検証の決まりの関数を呼ばない');
  const a9code = [...a9].map((f) => bodies.get(f)).join('\n');
  assert.deepEqual(['EVALUATION_POLICY_VERSION', 'ACTUAL_CLOSE_LAG_DAYS', 'EVAL_CALENDAR_TZ'].filter((c) => new RegExp('\\b' + c + '\\b').test(a9code)), [], '新しい数も使わない');
  assert.deepEqual(['EVAL_LOG', 'EVAL_COMPARE_MONTHLY', 'ACTUAL_EVAL_MONTHLY', 'EVAL_INSIGHTS'].filter((s) => new RegExp('SHEETS\\.' + s + '\\b').test(a9code)), [],
    '予測は検証の表を読まない（締まっていない月の古い行が残っていても、予測には効かない）');
  // （参考）A-4 の Vertex アシストは「最近の外れ」に評価の組を使う。今後は締まった月の、月が始まる前の予測で測った組だけになる
  assert.ok(reach('runVertexAIResearch').has('collectNeutralEvalPairs_'));
}

console.log('forecast-closed-months: all tests passed');
