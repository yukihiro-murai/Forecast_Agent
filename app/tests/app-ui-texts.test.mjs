#!/usr/bin/env node
/*
 * app-ui-texts.test.mjs — 2026-10-07 の所有者の決定（D1〜D8）の後の、画面の言葉（UI.html の script を vm で読む）。
 *   1. 補正を学び直す: 計画の自動の学びの旗が 0 なら、見直し案を作る（止めている操作）と同じ押せないボタンにし、理由をカーソルで出す。
 *      旗が 1 なら押せる。サーバーが止めた理由があれば、それを先に出す
 *   2. 振り返り: 見直し案を作る操作を止めている間は、案内・人の記入の説明で見直し案が出るとは言わず、更新の説明で止めている操作を言う
 *   3. 根拠の補正: 所有者が承認した値で書いたときは「補正を学び直す」で決まるとは言わない。自動の学びを止めていれば（所有者の書いた跡が無くても）、
 *      補正を学び直す・見直しを反映では変わらないと言う
 *   4. AI の効き: 0 は 0%、小さい値（0.0008）は 0.08%（整数の % で 0% にしない）
 *   5. 学び > 人の学び: 比べた月がまだ無いときは「外れた月はありません」と言わず、当たり具合を計算するよう案内する
 *   6. 分析の案内: 外れ幅・読みのクセは、各計画の「当たり具合を計算」で出ると言う
 *   7. 出した直後と同じ形（どの計画も今の版の検証の行が無く、外れ幅が無い）: ホーム・分析・学びを本物の中身で描き、undefined・NaN を出さない
 *   8. 人の学びのまとめ: 実績と比べた月（summary.scoredMonths）が無い / あるが振り返りの記録が無い / 外れなし / 外れあり。scoredMonths が無いサーバーでも描ける
 *   9. 動かせない人（閲覧・情報提供）には「計算されると出ます」。締め済みの年度の分析は「計算し直しません」
 *  10. 荒れの文は出ている印だけで作り、比べた月が無ければ出し方も言う（学びのまとめ・ホームの学びの進み）
 *  11. 根拠の補正: 出している予測の補正と今の補正が違えば、見える一文は出している予測のこと
 *  12. 振り返り: 人の記入・見直し案があって比べた月が無いときは、精度の天気の場所に 1 行
 *  13. 承認して反映: 自動の学びを止めている間は、判断を記録するだけで補正の値は変わらないと言う（ボタンは押せる）
 *  14. AI の効きの端の値・人の学びの「のべ」・分析の案内の空白・スマホの説明が欄の focusout で消えない
 * モックの上の確かめで、本物の Apps Script・ブラウザの上では動かしていない。
 *
 *   node app/tests/app-ui-texts.test.mjs
 */
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { setUpEnv, uiHtml, OWNER } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const env = setUpEnv();
const PAUSED = env.run('APP_REVIEW_GENERATE_PAUSED');
const POLICY = env.run('APP_EVAL_POLICY_VERSION');

/**
 * 画面（UI.html の script）を vm で読み込む。よみの絵は、どの姿か分かる印にする。
 * listen: 渡すと document の addEventListener を種類ごとに控え、getElementById は同じ id に同じ要素を返す（説明の箱を確かめる）
 */
function loadUi(listen) {
  const js = uiHtml.slice(uiHtml.indexOf('<script>') + 8, uiHtml.lastIndexOf('</script>'))
    .replace('<?!= charsJs ?>', 'var YOMI_POSE = new Proxy({}, { get: (t, k) => "<i data-pose=\\"" + String(k) + "\\"></i>" }); var CHAR_SVG = new Proxy({}, { get: () => "" });')
    .replace('<?!= bootJson ?>', JSON.stringify({ app: { name: 'T', version: 'x' }, user: { email: OWNER, isOwner: true, isAdmin: true, roles: [{ role: 'ADMIN', scopeType: 'ALL' }] }, setUp: true, allowed: true }));
  const els = {};
  const el = (id) => (listen ? (els[id] = els[id] || fakeEl()) : { innerHTML: '', classList: { add() {}, remove() {} }, set outerHTML(v) {} });
  const on = listen ? (type, fn) => { (listen[type] = listen[type] || []).push(fn); } : () => {};
  const ui = vm.createContext({ document: { getElementById: el, querySelector: () => null, addEventListener: on }, setTimeout: () => 0, clearTimeout() {}, confirm: () => false,
    window: { addEventListener() {}, innerWidth: 1280, innerHeight: 800 },
    google: { script: { run: new Proxy({}, { get: (t, k) => (k === 'withSuccessHandler' || k === 'withFailureHandler' ? () => ui.google.script.run : () => {}) }) } } });
  vm.runInContext(js, ui);
  return ui;
}
/**
 * 要素の代わり（属性・クラス・親子だけ）。tag: 'span' など、cls: クラスの並び、attrs: 属性、kids: 子。
 * closest / querySelector は、この画面が使う形（[data-tip]・.act[data-tip]・button,a,select,input,textarea）だけ読む
 */
function fakeEl(tag = 'div', cls = [], attrs = {}, kids = []) {
  const on = new Set();
  const e = { tag, cls, attrs: Object.assign({}, attrs), kids, parent: null, disabled: !!attrs.disabled, textContent: '', style: {}, offsetWidth: 120, offsetHeight: 24,
    classList: { add: (c) => on.add(c), remove: (c) => on.delete(c), contains: (c) => on.has(c) },
    getAttribute: (k) => (k in e.attrs ? e.attrs[k] : null), setAttribute: (k, v) => { e.attrs[k] = String(v); }, removeAttribute: (k) => { delete e.attrs[k]; },
    getBoundingClientRect: () => ({ left: 10, top: 10, width: 40, height: 20, bottom: 30 }),
    contains: (x) => { for (let n = x; n; n = n.parent) if (n === e) return true; return false; },
    matches: (sel) => sel.split(',').some((one) => { const m = /^([a-z]*)((?:\.[\w-]+)*)((?:\[[\w-]+\])*)$/.exec(one.trim()); if (!m) return false;
      return (!m[1] || m[1] === e.tag) && (m[2].match(/[\w-]+/g) || []).every((c) => e.cls.includes(c)) && (m[3].match(/[\w-]+/g) || []).every((a) => a in e.attrs); }),
    closest: (sel) => { for (let n = e; n; n = n.parent) if (n.matches(sel)) return n; return null; },
    querySelector: (sel) => { for (const k of e.kids) { if (k.matches(sel)) return k; const x = k.querySelector(sel); if (x) return x; } return null; } };
  kids.forEach((k) => { k.parent = e; });
  return e;
}
const ui = loadUi();
const run = (code, vars) => { Object.assign(ui, vars || {}); return vm.runInContext(code, ui); };
/** 役割を変えて描く（閲覧・情報提供だけの人など）。終わったら管理者に戻す */
const VIEWER_ROLES = [{ role: 'VIEWER', scopeType: 'ALL' }, { role: 'CONTRIBUTOR', scopeType: 'ALL' }];
const asRoles = (roles, fn) => { const keep = run('B.user.roles'); run('B.user.roles = __r', { __r: roles }); try { return fn(); } finally { run('B.user.roles = __r', { __r: keep }); } };
/** 計画の画面の中身（振り返り・根拠に使うところだけ）。paused: 見直し案を作るを止めているか。learning: 旧来の画面の学びの中身。can: 計画を動かせるか。frozen: 締め済みの年度 */
const planView = ({ paused = true, learning = { autoUpdate: true, biasFactor: 1, note: '' }, insights = [], proposals = [], can = true, frozen = false } = {}) => ({
  plan: { planId: 'P1', clientName: 'テスト製薬', fy: 2026, frozen }, inputHash: 'h', sourceReady: true, can: { plan: can, approve: can, admin: can }, recent: [],
  actions: ['IMPORT.ACTUALS', 'EVAL.REPORT', 'EVAL.INSIGHTS', 'LEARN.MONTHLY', 'REVIEW.GENERATE', 'REVIEW.APPLY']
    .map((a) => ({ action: a, kind: 'run', label: a, minRole: 'PLANNER', paused: a === 'REVIEW.GENERATE' && paused ? PAUSED : '' })),
  boot: { eval: { insights }, quarterly: { proposals }, learning }
});
/** 振り返りのタブ（当たり具合は読み終わった・月はまだ無い） */
const review = (v) => run(`S.fc.planId = 'P1'; S.fc.view = __v; S.fc.data = { plan: __v.plan, latest: null, runs: [], stored: null }; S.fc.learn = { planId: 'P1', accuracy: { months: [] } }; S.fc.ins = {};
  fcReviewTab(S.fc.data, S.fc.view)`, { __v: v });
const tipOf = (html, h2) => (new RegExp('<h2>' + h2 + '<span class="tip"[^>]*data-tip="([^"]*)"').exec(html) || [])[1];
const OWNER_NOTE = 'owner-approved 2026-10-07 09:30: 自動の学びを止める';

// ==== 1. 補正を学び直す: 旗が 0 なら押せないボタン（見直し案を作ると同じ形・同じ幅のボタン） ====
{
  const off = review(planView({ learning: { autoUpdate: false, biasFactor: 1, note: OWNER_NOTE } }));
  assert.doesNotMatch(off, /fcRun\('LEARN\.MONTHLY'\)/);
  assert.match(off, /<span class="act" data-tip="実績から補正を学び直す\n自動の学びを止めています（2026-10-07 所有者の決定）。今は動かしても補正は変わりません。"><button class="btn btn-ghost" disabled aria-disabled="true">補正を学び直す<\/button><\/span>/);
  assert.match(off, /<span class="act" data-tip="AI の見直し案を作る\n[^"]*"><button class="btn btn-ghost" disabled aria-disabled="true">見直し案を作る<\/button><\/span>/, '見直し案を作ると同じ形');
  assert.match(off, /fcRun\('EVAL\.REPORT'\)/, 'ほかの実行は押せる');
  // 所有者が書いたのでなければ、日付と「所有者の決定」は言わない
  const off2 = review(planView({ learning: { autoUpdate: false, biasFactor: 1, note: 'auto-learned' } }));
  assert.match(off2, /data-tip="実績から補正を学び直す\n自動の学びを止めています。今は動かしても補正は変わりません。"><button class="btn btn-ghost" disabled/);
  // 旗が 1・学びの中身が無い（古い控え）なら、今までどおり押せる
  for (const learning of [{ autoUpdate: true, biasFactor: 0.9, note: OWNER_NOTE }, undefined]) {
    const on = review(planView({ learning }));
    assert.match(on, /<span class="act" data-tip="実績から補正を学び直す（次の予測に効く）\n実績から補正の係数を学び直します（次の予測に反映）。"><button class="btn btn-ghost" onclick="fcRun\('LEARN\.MONTHLY'\)">補正を学び直す<\/button>/);
  }
  // サーバーが止めた理由があれば、それを出す（旗より先）
  const v = planView({ learning: { autoUpdate: false, note: OWNER_NOTE } });
  v.actions.filter((a) => a.action === 'LEARN.MONTHLY')[0].paused = 'B-5 は止めています。';
  assert.equal(run(`S.fc.view = __v; fcPaused('LEARN.MONTHLY')`, { __v: v }), '補正を学び直す は止めています。', '中の番号はふつうの言葉に');
  // ほかの操作は旗では止めない
  assert.equal(run(`fcPaused('EVAL.REPORT') + fcPaused('REVIEW.APPLY')`), '');
}

// ==== 2. 振り返り: 見直し案を作るを止めている間の言葉 ====
{
  const insights = [{ row: 5, month: '2026/04', insight: '外れ', nextAction: '', hypothesis: '', actionType: 'update', reflection: '', owner: '', status: 'open' }];
  // 止めている（旗は 1）: 案内は見直し案が出るとは言わない・更新の説明で、止めているのは見直し案を作るだけと言う
  const p = review(planView());
  assert.match(p, /まだ振り返りの記録がありません。月の実績が締まったら、下の「振り返りの更新」を左から順に実行してください。当たり具合・外れの原因がここに出ます。/);
  assert.doesNotMatch(p, /AI の見直し案がここに出ます/);
  assert.equal(tipOf(p, '振り返りの更新'), '左から順に実行します。「見直し案を作る」は、今は止めています');
  // 旗も 0: 止めている 2 つを並べる
  assert.equal(tipOf(review(planView({ learning: { autoUpdate: false, note: OWNER_NOTE } })), '振り返りの更新'), '左から順に実行します。「補正を学び直す」と「見直し案を作る」は、今は止めています');
  // 止めていない: 今までどおり
  const u = review(planView({ paused: false }));
  assert.match(u, /当たり具合・外れの原因・AI の見直し案がここに出ます。/);
  assert.equal(tipOf(u, '振り返りの更新'), '左から順に実行します');
  // 人の記入の説明・保存の説明
  const n = review(planView({ insights }));
  assert.match(tipOf(n, '人の記入'), /^外れた月の原因の仮説・対応・担当を書きます。次の予測の材料になります。/);
  assert.match(n, /data-tip="原因の仮説・対応・担当・状態を保存します。次の予測の材料になります"/);
  assert.doesNotMatch(n, /見直し案の材料/);
  const n2 = review(planView({ paused: false, insights }));
  assert.match(tipOf(n2, '人の記入'), /次の予測と AI の見直し案の材料になります。/);
  assert.match(n2, /data-tip="原因の仮説・対応・担当・状態を保存します。次の予測と AI の見直し案の材料になります"/);
  // 承認待ちの見直し案は、止めている間も承認して反映できる（ボタンを消さない）
  const props = [{ row: 8, pid: 'P1', target: 'ai_weight_override', current: '0.0008', proposed: '0.0004', conf: '中', rationale: '', impact: '', decision: '', rollback: '' }];
  const a = review(planView({ proposals: props }));
  assert.match(a, /onclick="fcRvGo\(\)">承認して反映<\/button>/);
  // ホームの学びの進み: 承認待ちの見直し案は「四半期ごとに作る」とは言わない
  assert.doesNotMatch(uiHtml, /四半期ごとに作る見直し案/);
}

// ==== 3. 根拠の補正: 所有者が承認した値・自動の学びを止めている ====
{
  const basis = (learning, cal) => {
    const v = planView({ learning });
    return tipOf(run(`S.fc.view = __v; S.fc.basis = __b; fcBasisTab(S.fc.data, __v)`, { __v: v, __b: { planId: 'P1', annual: {}, monthly: [], research: [], applied: null, calibration: cal } }), '補正');
  };
  const cal = { factor: 1, aiWeight: 0, aiMax: null, monthBias: {}, updatedAt: '2026-10-07T09:30:00+09:00', quarter: '' };
  const OFF = '\n自動の学びを止めているので、「補正を学び直す」と「見直しを反映」では変わりません';
  assert.equal(basis({ autoUpdate: false, biasFactor: 1, note: OWNER_NOTE }, cal), '所有者が承認した値（2026-10-07）で決めた、予測の直し方です' + OFF);
  assert.equal(basis({ autoUpdate: true, biasFactor: 1, note: OWNER_NOTE }, cal), '所有者が承認した値（2026-10-07）で決めた、予測の直し方です');
  assert.equal(basis({ autoUpdate: true, biasFactor: 0.9, note: 'auto-learned' }, cal), '「補正を学び直す」と「見直しを反映」で決まる、予測の直し方です');
  assert.equal(basis(undefined, cal), '「補正を学び直す」と「見直しを反映」で決まる、予測の直し方です', '学びの中身が無くても描ける');
  // 旗が 0 なら、所有者の書いた跡が無くても「補正を学び直す」と「見直しを反映」で決まるとは言わない（どちらでも値は変わらない）
  for (const note of ['', 'auto-learned']) {
    const t = basis({ autoUpdate: false, biasFactor: 0.9, note }, cal);
    assert.equal(t, '今の値のまま使う、予測の直し方です' + OFF);
    assert.doesNotMatch(t, /で決まる/);
  }
}

// ==== 4. AI の効き: 0 は 0%、小さい値は 2 けた ====
{
  const q = (x) => run('qrValue("ai_weight_override", __x)', { __x: x });
  assert.deepEqual([0, '0', 0.0008, '0.0004', 0.002, 0.0015, 0.01, 0.5, 1, '', null].map(q), ['0%', '0%', '0.08%', '0.04%', '0.2%', '0.15%', '1%', '50%', '100%', '既定', '既定']);
  assert.equal(q('abc'), 'abc', '数でない値はそのまま');
  // 見直し案の表: 0.0008 → 0.0004 が「0% → 0%」にならない
  const props = [{ row: 8, pid: 'P1', target: 'ai_weight_override', current: '0.0008', proposed: '0.0004', conf: '中', rationale: '', impact: '', decision: '', rollback: '' }];
  assert.match(run(`S.fc.view = __v; fcRvProposals({ title: '', period: '', reviewId: 'R-1', applied: false, logRecent: [], proposals: __p })`, { __v: planView(), __p: props }),
    /AI の効き<\/td><td class="num"[^>]*>0\.08%<\/td><td class="num"[^>]*>0\.04%<\/td>/);
  // 根拠の補正の説明
  const tipLines = (w) => (/<p class="fc-say" data-tip="([^"]*)"/.exec(run(`S.fc.basis = __b; fcBasisTab(S.fc.data, S.fc.view)`,
    { __b: { planId: 'P1', annual: {}, monthly: [], research: [], applied: null, calibration: { factor: 1, aiWeight: w, aiMax: 0.05, monthBias: {}, updatedAt: '', quarter: '' } } })) || [])[1];
  assert.match(tipLines(0.0004), /AI の効き 0\.04%・上限 ±5%/);
  assert.match(tipLines(0), /AI の効き 0%・上限 ±5%/);
  assert.match(tipLines(null), /AI の効き 既定・上限 ±5%/);
}

// ==== 5. 学び > 人の学び: 比べた月がまだ無い ====
const hero = (d) => run('lrPeopleHero(__d)', { __d: d });
const pose = (h) => (/data-pose="([^"]*)"/.exec(h) || [])[1];
const heroSay = (h) => (/<span data-tip="[^"]*">([^<]*)<\/span>/.exec(h) || [])[1];
const NONE_SAY = 'まだ実績と比べた月がありません。各計画の振り返りで「当たり具合を計算」と「外れの原因を整理」を動かすと、大きく外れた月がここに出ます。';
{
  const none = hero({ summary: { months: 0, misses: 0, withNotes: 0 }, openActions: [], repeats: [], noPreMonth: 3 });
  assert.equal(heroSay(none), NONE_SAY);
  assert.doesNotMatch(none, /大きく外れた月はありません|予報どおり/);
  assert.equal(pose(none), 'explain');
  assert.match(none, /data-tip="[^"]*締まった月（月末から 5 日たってから実績を取り込んだ月）を、その月が始まる前の最後の予測と比べます\n締まった月のうち のべ 3 か月は、その月が始まる前の予測が無いので比べていません"/);
  assert.equal(pose(hero({})), 'explain', 'まとめが無いときも、まだ比べていない');
  // 比べた月があって、外れが無いときだけ「外れた月はありません」
  const calm = hero({ summary: { months: 4, misses: 0, withNotes: 0 }, openActions: [], repeats: [], noPreMonth: 0 });
  assert.match(calm, /大きく外れた月はありません。予報どおりの空模様が続いています。/);
  assert.equal(pose(calm), 'done');
  assert.doesNotMatch(calm, /締まった月のうち/);
  const miss = hero({ summary: { months: 4, misses: 2, withNotes: 1 }, openActions: [{}], repeats: [] });
  assert.match(miss, /大きく外れた月を 2 件観測しました。人の振り返りは 1 件、残っている対応は 1 件です。/);
  assert.equal(pose(miss), 'discover');
}

// ==== 6. 分析の案内: 外れ幅・読みのクセは、当たり具合を計算で出る ====
{
  const plan = { planId: 'P1', clientName: 'テスト製薬', fy: '2026', budget: 1000, officialFinal: null, budgetUsed: null, landing: 1050, landingP10: 950, landingP90: 1150, landingSd: 78, ratio: 1.05, pAbove: 0.7,
    sky: 'harenochi', skyReason: 'ratio', accuracy: { mape: null, bias: null, n: 0, coverage: null, coverageN: 0 }, revisions: [{ at: '2026-09-01', p50: 1000 }, { at: '2026-10-01', p50: 1050 }], topics: [] };
  const guideOf = (plans) => (/<div class="guide">[\s\S]*?<div class="bubble">([\s\S]*?)<\/div><\/div><\/div>/.exec(run(`S.an.data = __d; S.an.tab = 'overview'; viewAnalysis()`,
    { __d: { fy: '2026', fys: ['2026'], totals: { budget: 1000, budgetPlans: 1, landing: 1050, ratioLanding: 1.05 }, plans } })) || [])[1];
  assert.equal(guideOf([plan]), 'まだ数字がそろっていないため、外れ幅・読みのクセは出していません。各計画の振り返りで「当たり具合を計算」を動かすと、ここに出てきます。');
  const noRev = Object.assign({}, plan, { revisions: [] });
  assert.equal(guideOf([noRev]), 'まだ数字がそろっていないため、外れ幅・読みのクセ・予測の見直しの流れは出していません。予測を動かし、実績を取り込むと、ここに出てきます。'
    + '外れ幅・読みのクセは、各計画の振り返りで「当たり具合を計算」を動かすと出ます。');
  const empty = { planId: 'P2', clientName: '別の製薬', fy: '2026', sky: 'mikakunin', skyReason: 'no_forecast', accuracy: null, revisions: [] };
  assert.equal(guideOf([empty]), 'まだ着地の推定や、予測と実績の比べがないため、グラフは出していません。予測を動かし、実績を取り込むと、ここに出てきます。'
    + '外れ幅・読みのクセは、各計画の振り返りで「当たり具合を計算」を動かすと出ます。');
  const withAcc = Object.assign({}, plan, { accuracy: { mape: 0.12, bias: 0.04, n: 5, coverage: 0.8, coverageN: 5 } });
  assert.equal(guideOf([withAcc]), undefined, 'そろっていれば案内は出さない');
}

// ==== 7. 出した直後と同じ形: 古い版の検証の行だけがある計画（外れ幅は無い）で、ホーム・分析・学びを描く ====
{
  const H = env.run('(() => { const o = {}; Object.keys(APP_ENGINE_SHEETS).forEach(k => { o[k] = APP_ENGINE_SHEETS[k].header || null; }); return o; })()');
  const row = (sheet, o) => H[sheet].map((h) => (o[h] === undefined ? '' : o[h]));
  const output = [['FY2026 売上予測（テスト製薬）']];
  for (let r = 2; r <= 25; r++) output.push([]);
  output.push(['年度合計（予測）', 900, 1000, 1100], [], []);
  for (let i = 0; i < 12; i++) { const m = 4 + i; output.push([(m > 12 ? 2027 : 2026) + '/' + String(m > 12 ? m - 12 : m).padStart(2, '0'), 90, 100, 110, '', '', '', 100, 10]); }
  const OLD = 'policy-2026H1-v2';
  assert.notEqual(OLD, POLICY);
  const evalRows = [['2026/04', 1250], ['2026/05', 900], ['2026/06', 1000]].flatMap(([ym, act], i) => ['nega', 'neutral', 'posi']
    .map((sc, j) => row('EVAL_LOG', { eval_id: 'E' + i + sc, evaluated_at: D(2026, 9, 1), client: 'テスト製薬', target_month: ym, scenario: sc, pred: 900 + 100 * j, actual: act,
      evaluation_policy_version: OLD, constraint_relevant_flag: 1 })));
  const planId = env.seedPlan(env.makeBook('テスト', {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', 'テスト製薬'], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
    OUTPUT: { values: output },
    PROCESS_STATUS: { values: [H.PROCESS_STATUS, ['step2_status', new Date(2026, 9, 6, 10), 'owner', 'success', 'テスト製薬', 10, ''],
      ['step5_status', new Date(2026, 9, 6, 11), 'owner', 'success', 'テスト製薬', 10, '']] },
    EVAL_LOG: { values: [H.EVAL_LOG].concat(evalRows), formats: { D: '@' } },
    EVAL_INSIGHTS: { values: [H.EVAL_INSIGHTS, row('EVAL_INSIGHTS', { client: 'テスト製薬', evaluated_at: D(2026, 6, 10), target_month: D(2026, 4), actual_total: 1250, pred_p50: 1000,
      cause_bucket: 'over_forecast', action_type: 'update', status: 'open', range_breach: 0 })] },
    QUARTERLY_REVIEW_LOG: { values: [H.QUARTERLY_REVIEW_LOG, row('QUARTERLY_REVIEW_LOG', { review_id: 'R1', proposal_id: 'P-A-1', reviewed_at: D(2026, 7, 10), quarter_label: 'FY2026-Q1', phase: 'A',
      target_field: 'ai_weight_override', current_value: '0.0008', proposed_value: '0.0004', confidence: '中', rationale: 'x', approval_status: '取り下げ', approval_decided_at: D(2026, 10, 7),
      approval_decided_by: OWNER, applied: 0 })] },
    CALIBRATION_HISTORY: { values: [H.CALIBRATION_HISTORY, row('CALIBRATION_HISTORY', { change_id: 'H-1', changed_at: D(2026, 10, 7), changed_by: OWNER, client: 'テスト製薬', quarter_label: 'FY2026-Q3',
      source_review_id: 'OWNER-APPROVED', factor_name: 'bias_correction_factor', old_value: '0.8', new_value: '1', rollback_hint: '戻す' })] }
  }));
  const home = env.call('apiHome()');
  const an = env.call('apiCrossMaker(__in)', { __in: {} });
  const people = env.call('apiPeopleLearning()');
  const ai = env.call('apiAiLearning()');
  // 本物の中身が、出した直後と同じ形になっている（古い版の行は数えない・外れ幅は無い・判断の履歴はある）
  const hp = home.plans.filter((p) => p.planId === planId)[0];
  assert.ok(hp && hp.mape === null, 'ホームの外れ幅は無い: ' + JSON.stringify(hp && hp.mape));
  const ap = an.plans.filter((p) => p.planId === planId)[0];
  assert.ok(ap && !(ap.accuracy && typeof ap.accuracy.mape === 'number'), '分析の外れ幅は無い');
  assert.equal(people.summary.months, 0);
  assert.ok((people.decisions || []).length >= 1, '判断の履歴はある（人の学びは案内 1 枚にならない）');
  const clean = (html, what) => { assert.ok(!/undefined|NaN/.test(html), what + 'に undefined や NaN を出さない'); return html; };
  // ホーム: 外れ幅は「-」、説明で当たり具合を計算を案内する。精度の天気は霧
  const h = clean(run(`S.view = 'home'; B.home = __h; S.lr.ai = __ai; viewHome()`, { __h: home, __ai: ai }), 'ホーム');
  assert.match(h, /<div class="kpi" data-tip="[^"]*まだ実績と比べたメーカーがありません。各計画の振り返りで「当たり具合を計算」を動かすと出ます"><div class="k">外れ幅<\/div><div class="v">-<\/div><\/div>/);
  assert.match(h, /data-tip="霧：比べられる実績がまだない\n各計画の振り返りで「当たり具合を計算」を動かすと出ます\n[^"]*"><span class="skych"[^>]*aria-label="霧"[\s\S]*精度の天気[\s\S]*霧[\s\S]*比べる実績はまだです/);
  // 分析: 外れ幅・読みのクセのカードは出さず、案内に当たり具合を計算を出す
  const a = clean(run(`S.view = 'analysis'; S.an.data = __a; S.an.tab = 'overview'; viewAnalysis()`, { __a: an }), '分析');
  assert.match(a, /外れ幅・読みのクセ[^<]*は出していません。[^<]*各計画の振り返りで「当たり具合を計算」を動かすと/);
  assert.doesNotMatch(a, /<h2>外れ幅<span|<h2>読みのクセ<span/);
  // 学び: 人の学びは「まだ比べた月がありません」、AI の学びも外れ幅が無いことを言う
  const lp = clean(run(`S.view = 'learning'; S.lr.tab = 'people'; S.lr.people = __p; viewLearning()`, { __p: people }), '人の学び');
  assert.match(lp, /まだ実績と比べた月がありません。/);
  assert.doesNotMatch(lp, /今の決まり/, '中の言葉（今の決まり）は出さない');
  assert.doesNotMatch(lp, /大きく外れた月はありません/);
  const la = clean(run(`S.lr.tab = 'ai'; S.lr.ai = __ai; viewLearning()`, { __ai: ai }), 'AI の学び');
  assert.match(la, /<div class="bubble">まだ実績と比べた月がありません。各計画の振り返りで「当たり具合を計算」を動かすと、外れ幅がここに出ます。<\/div>/);
  assert.match(la, /<span class="nm">霧<\/span><span class="v">比べる実績なし<\/span>/);
  assert.match(lp, /締まった月のうち のべ 3 か月は、その月が始まる前の予測が無いので比べていません/, '月が始まる前の予測が無い月の数（本物の中身）');
}

// ==== 8. 人の学びのまとめ: 実績と比べた月（scoredMonths）と、振り返りの記録（months） ====
{
  const h = (summary, more) => hero(Object.assign({ openActions: [], repeats: [], noPreMonth: 0, summary }, more || {}));
  // 比べた月が無い（scoredMonths = 0）: 当たり具合を計算から
  const n0 = h({ months: 0, misses: 0, withNotes: 0, scoredMonths: 0 });
  assert.equal(heroSay(n0), NONE_SAY);
  assert.equal(pose(n0), 'explain');
  // 比べた月はあるが、振り返りの記録が無い: 外れの原因を整理を案内する（「外れた月はありません」とは言わない）
  const n1 = h({ months: 0, misses: 0, withNotes: 0, scoredMonths: 5 });
  assert.equal(heroSay(n1), '実績と比べた月が のべ 5 か月あります。各計画の振り返りで「外れの原因を整理」を動かすと、大きく外れた月がここに出ます。');
  assert.equal(pose(n1), 'explain');
  assert.doesNotMatch(n1, /大きく外れた月はありません|予報どおり/);
  [n0, n1].forEach((x) => assert.doesNotMatch(heroSay(x), /今の決まり/, '中の言葉（今の決まり）は見える文に出さない'));
  // 振り返りの記録があれば、scoredMonths があっても無くても今までどおり（外れなし・外れあり）
  for (const scoredMonths of [6, undefined]) {
    const calm = h({ months: 4, misses: 0, withNotes: 0, scoredMonths });
    assert.equal(heroSay(calm), '大きく外れた月はありません。予報どおりの空模様が続いています。');
    assert.equal(pose(calm), 'done');
    const miss = h({ months: 4, misses: 2, withNotes: 1, scoredMonths }, { openActions: [{}] });
    assert.equal(heroSay(miss), '大きく外れた月を 2 件観測しました。人の振り返りは 1 件、残っている対応は 1 件です。');
    assert.equal(pose(miss), 'discover');
  }
  // scoredMonths を返さない（前の）サーバー: 振り返りの記録が無ければ、今までどおり当たり具合を計算から
  assert.equal(heroSay(h({ months: 0, misses: 0, withNotes: 0 })), NONE_SAY);
  assert.equal(heroSay(h(undefined)), NONE_SAY, 'まとめが無くても描ける');
  // 何も無いときの案内 1 枚も、外れの原因を整理を言う（大きく外れた月の振り返りは、そこで出る）
  assert.match(run('lrPeopleView({ summary: { months: 0 } })'),
    /<div class="bubble">実績が締まって、各計画の振り返りで「当たり具合を計算」と「外れの原因を整理」を動かすと、大きく外れた月の振り返りと、読みの当たりがここに並びます。<\/div>/);
}

// ==== 9. 動かせない人（閲覧・情報提供）には「計算されると出ます」。締め済みの年度の分析は「計算し直しません」 ====
const AN_PLAN = { planId: 'P1', clientName: 'テスト製薬', fy: '2026', budget: 1000, officialFinal: null, budgetUsed: null, landing: 1050, landingP10: 950, landingP90: 1150, landingSd: 78, ratio: 1.05, pAbove: 0.7,
  sky: 'harenochi', skyReason: 'ratio', accuracy: { mape: null, bias: null, n: 0, coverage: null, coverageN: 0 }, revisions: [{ at: '2026-09-01', p50: 1000 }, { at: '2026-10-01', p50: 1050 }], topics: [] };
const anGuide = (plans, fy = '2026') => (/<div class="guide">[\s\S]*?<div class="bubble">([\s\S]*?)<\/div><\/div><\/div>/.exec(run(`S.an.data = __d; S.an.fy = __d.fy; S.an.tab = 'overview'; viewAnalysis()`,
  { __d: { fy, fys: [fy], totals: { budget: 1000, budgetPlans: 1, landing: 1050, ratioLanding: 1.05 }, plans } })) || [])[1];
{
  const home = { fy: '2026', plans: [{ planId: 'P1', clientName: 'テスト製薬', fy: '2026', mape: null }] };
  asRoles(VIEWER_ROLES, () => {
    assert.equal(run('runWhen()'), '当たり具合が計算されると');
    // 人の学び・AI の学び
    assert.equal(heroSay(hero({ summary: { months: 0, scoredMonths: 0 } })), 'まだ実績と比べた月がありません。当たり具合が計算され、外れの原因が整理されると、大きく外れた月がここに出ます。');
    assert.equal(heroSay(hero({ summary: { months: 0, scoredMonths: 2 } })), '実績と比べた月が のべ 2 か月あります。外れの原因が整理されると、大きく外れた月がここに出ます。');
    assert.match(run('lrPeopleView({})'), /<div class="bubble">実績が締まって、当たり具合が計算され、外れの原因が整理されると、大きく外れた月の振り返りと、読みの当たりがここに並びます。<\/div>/);
    assert.match(run('lrAiView({})'), /<div class="bubble">実績が締まって、当たり具合が計算されると、AI が何を学んだか（重み・偏り・外れ幅・補正）がここに出ます。<\/div>/);
    assert.match(run('lrAiHero({ health: [], timeline: [] })'), /<div class="bubble">まだ実績と比べた月がありません。当たり具合が計算されると、外れ幅がここに出ます。/);
    // ホーム: 外れ幅の箱・精度の天気の説明
    assert.match(run('homeKpis(__h, __h.plans)', { __h: home }), /data-tip="[^"]*\nまだ実績と比べたメーカーがありません。当たり具合が計算されると出ます"><div class="k">外れ幅<\/div>/);
    assert.match(run('S.lr.ai = { health: [], pending: [] }; homeLearnCard(__h, __h.plans)', { __h: home }), /data-tip="霧：比べられる実績がまだない\n当たり具合が計算されると出ます\n精度の推移は/);
    // 分析
    assert.equal(anGuide([AN_PLAN]), 'まだ数字がそろっていないため、外れ幅・読みのクセは出していません。当たり具合が計算されると、ここに出てきます。');
    // 見える文・説明のどこにも「当たり具合を計算」を動かすようには言わない
    for (const html of [hero({ summary: { months: 0 } }), run('lrAiView({})'), run('homeKpis(__h, __h.plans)', { __h: home })]) assert.doesNotMatch(html, /「当たり具合を計算」を動かす/);
  });
  // メーカー単位の予算策定担当は動かせる（どこかの計画で動かせれば、動かし方を言う）
  asRoles(VIEWER_ROLES.concat([{ role: 'PLANNER', scopeType: 'CLIENT', clientId: 'C1' }]), () => assert.equal(run('runWhen()'), '各計画の振り返りで「当たり具合を計算」を動かすと'));
  assert.equal(run('runWhen(["EVAL.INSIGHTS"])'), '各計画の振り返りで「外れの原因を整理」を動かすと');
  // 締め済みの年度（計画の一覧の frozen）: 当たり具合は計算し直さないと言い、動かし方は言わない
  const fy25 = Object.assign({}, AN_PLAN, { fy: '2025' });
  run(`S.an.plans = [{ planId: 'P1', fy: 2025, frozen: true }, { planId: 'P9', fy: 2026, frozen: false }]`);
  assert.equal(anGuide([fy25], '2025'), 'まだ数字がそろっていないため、外れ幅・読みのクセは出していません。FY2025 は締め済みのため、当たり具合は計算し直しません。');
  assert.equal(anGuide([AN_PLAN]), 'まだ数字がそろっていないため、外れ幅・読みのクセは出していません。各計画の振り返りで「当たり具合を計算」を動かすと、ここに出てきます。', '締めていない年度はそのまま');
  // 分析の中身に frozen があれば、それを使う（グラフが何も描けないときは、予測を動かすようにも言わない）
  run('S.an.plans = null');
  const empty = { planId: 'P2', clientName: '別の製薬', fy: '2025', sky: 'mikakunin', skyReason: 'no_forecast', accuracy: null, revisions: [], frozen: true };
  assert.equal(anGuide([empty], '2025'), 'まだ着地の推定や、予測と実績の比べがないため、グラフは出していません。FY2025 は締め済みのため、予測や当たり具合は計算し直しません。');
  assert.equal(anGuide([fy25], '2025'), 'まだ数字がそろっていないため、外れ幅・読みのクセは出していません。各計画の振り返りで「当たり具合を計算」を動かすと、ここに出てきます。', '分からなければ締めていないとみなす');
  // 計画の一覧は、前の年度を開いたときだけ 1 回読む（今の年度・frozen のある中身では読まない）
  const keepHome = run('B.home');
  const loads = (r) => run(`B.home = { fy: '2026', plans: [] }; S.an.plans = null; delete S.ld.anPlans; anLoadPlans(__r); var __l = !!S.ld.anPlans; delete S.ld.anPlans; __l`, { __r: r });
  assert.deepEqual([loads({ fy: '2025', plans: [fy25] }), loads({ fy: '2026', plans: [AN_PLAN] }), loads({ fy: '2027', plans: [] }), loads({ fy: '2025', plans: [empty] })], [true, false, false, false]);
  run('B.home = __k', { __k: keepHome });
}

// ==== 10. 荒れの文は、出ている印だけで作る。比べた月が無ければ出し方も言う ====
{
  const fy = run('lrFy()');
  const clamp = { planId: 'P1', clientName: 'テスト製薬', fy, key: 'factor_clamp', label: '', value: 0.75 };
  const cov = { planId: 'P2', clientName: '別の製薬', fy, key: 'coverage_low', label: '', value: { coverage: 0.5, n: 6 } };
  const old = Object.assign({}, cov, { planId: 'P3', fy: String(Number(fy) - 1) });
  const say = (health, timeline) => (/<div class="bubble">([^<]*)/.exec(run('lrAiHero(__d)', { __d: { health, timeline: timeline || [], pending: [] } })) || [])[1];
  const NOT_YET = 'まだ実績と比べた月がありません。各計画の振り返りで「当たり具合を計算」を動かすと、外れ幅がここに出ます。';
  // 補正が限度に張りついただけ（比べた月は無い）: 見通しの幅の外れは言わず、出し方を言う
  assert.equal(say([clamp]), '予測の当たり方が荒れています（1 メーカー）。補正が限度に張りついていて、外れを補正で追いきれていない状態です。' + NOT_YET);
  assert.doesNotMatch(say([clamp]), /見通しの幅/);
  assert.equal(say([cov, old]), '予測の当たり方が荒れています（1 メーカー）。実績が見通しの幅から外れることが多く、幅が狭すぎる状態です。' + NOT_YET, '前の年度の印は数えない');
  assert.equal(say([clamp, cov]), '予測の当たり方が荒れています（2 メーカー）。補正が限度に張りついたり、実績が見通しの幅から外れたりしていて、予測がぶれやすい状態です。' + NOT_YET);
  // 比べた月があれば、出し方は言わない
  const tl = [{ ym: '2026/05', n: 2, mape: 0.12, rolling3: 0.12 }];
  assert.equal(say([clamp], tl), '予測の当たり方が荒れています（1 メーカー）。補正が限度に張りついていて、外れを補正で追いきれていない状態です。');
  asRoles(VIEWER_ROLES, () => assert.match(say([clamp]), /状態です。まだ実績と比べた月がありません。当たり具合が計算されると、外れ幅がここに出ます。$/));
  // ホームの学びの進み: 印が 1 種類ならその名前、比べた月が無ければ説明で出し方を言う
  const home = { fy, plans: [{ planId: 'P1', clientName: 'テスト製薬', fy, mape: null }] };
  const card = (health) => run('S.lr.ai = { health: __hl, pending: [] }; homeLearnCard(__h, __h.plans)', { __hl: health, __h: home });
  assert.match(card([clamp]), /<span class="t1">台風<\/span><span class="t2">補正が限度（1 メーカー）<\/span>/);
  assert.match(card([clamp]), /data-tip="台風：[^"\n]*\nまだ実績と比べたメーカーがありません。各計画の振り返りで「当たり具合を計算」を動かすと出ます\n荒れている印: 補正が上限か下限に張りついたメーカー 1\n/);
  assert.match(card([cov]), /<span class="t2">幅が狭すぎる（1 メーカー）<\/span>/);
  assert.match(card([clamp, cov]), /<span class="t2">荒れています（2 メーカー）<\/span>/);
  assert.match(card([]), /<span class="t2">比べる実績はまだです<\/span>/);
}

// ==== 11. 根拠の補正: 出している予測の補正と今の補正が違うとき（所有者が戻した直後など） ====
{
  const cal = (applied, factor) => {
    const html = run(`S.fc.view = __v; S.fc.basis = __b; fcBasisTab(S.fc.data, __v)`, { __v: planView(), __b: { planId: 'P1', annual: {}, monthly: [], research: [], applied,
      calibration: { factor, aiWeight: null, aiMax: null, monthBias: {}, updatedAt: '', quarter: '' } } });
    const m = /<h2>補正<span[\s\S]*?<p class="fc-say" data-tip="([^"]*)">([^<]*)<\/p>/.exec(html);
    return { tip: m[1], say: m[2] };
  };
  const reset = cal({ bias_correction_factor: 0.97 }, 1);
  assert.equal(reset.say, '最新の予測は 3.0% 下げています（次の予測から補正なし）');
  assert.match(reset.tip, /最新の予測は 3\.0% 下げて計算しました\n今の補正: 補正なし（次の予測から効きます）/);
  assert.doesNotMatch(reset.say, /かけていません/, '見える一文は、出している予測と食い違わない');
  assert.equal(cal({ bias_correction_factor: '1' }, 0.98).say, '最新の予測は、補正なしです（次の予測から 2.0% 下げ）');
  assert.equal(cal({ bias_correction_factor: '0.97' }, 1.05).say, '最新の予測は 3.0% 下げています（次の予測から 5.0% 上げ）');
  // 同じなら今までどおり（今の補正の一文だけ）
  const same = cal({ bias_correction_factor: 0.97 }, 0.97);
  assert.equal(same.say, 'これまでの外れ方から 3.0% 下げています');
  assert.doesNotMatch(same.tip, /次の予測から効きます/);
  assert.equal(cal(null, 1).say, 'これまでの外れ方による補正は、かけていません');
}

// ==== 12. 振り返り: 人の記入・見直し案があって、比べた月がまだ無い ====
{
  const insights = [{ row: 5, month: '2026/04', insight: '外れ', nextAction: '', hypothesis: '', actionType: 'update', reflection: '', owner: '', status: 'open' }];
  const props = [{ row: 8, pid: 'P1', target: 'ai_weight_override', current: '0.0008', proposed: '0.0004', conf: '中', rationale: '', impact: '', decision: '', rollback: '' }];
  const NOTE = (s) => '<div class="card"><div class="card-head"><h2>精度の天気</h2></div><p class="note">まだ実績と比べた月がありません。' + s + '</p></div>';
  const a = review(planView({ insights }));
  assert.ok(a.includes(NOTE('「当たり具合を計算」で出ます。')), a.slice(0, 300));
  assert.ok(a.indexOf(NOTE('')) < 0 && a.indexOf('<h2>精度の天気</h2>') < a.indexOf('<h2>人の記入'), '人の記入より先（精度の天気の場所）');
  assert.doesNotMatch(a, /まだ振り返りの記録がありません/, '案内 1 枚とは重ねない');
  assert.ok(review(planView({ proposals: props, can: false })).includes(NOTE('当たり具合が計算されると出ます。')), '動かせない人');
  assert.ok(review(planView({ insights, can: false, frozen: true })).includes(NOTE('締め済みの年度は、当たり具合を計算し直しません。')), '締め済みの年度');
  // 何も無いときは案内 1 枚だけ（1 行は出さない）。締め済みなら案内も「計算されると」とは言わない
  assert.doesNotMatch(review(planView()), /<h2>精度の天気<\/h2>/);
  const f = review(planView({ can: false, frozen: true }));
  assert.match(f, /まだ振り返りの記録がありません。締め済みの年度は、当たり具合を計算し直しません。/);
  assert.doesNotMatch(f, /計算されると/);
  // 比べた月があれば 1 行は出さない（精度の天気のカード）
  const withAcc = run(`S.fc.learn = { planId: 'P1', accuracy: { months: [{ month: '2026/04', p10: 1, p50: 2, p90: 3, actual: 2, err: 0, inside: true }], mape: 0.05, n: 1, coverage: 1, coverageN: 1 } }; fcReviewTab(S.fc.data, S.fc.view)`);
  assert.doesNotMatch(withAcc, /まだ実績と比べた月がありません/);
  assert.match(withAcc, /精度の天気[\s\S]*精度の推移/);
}

// ==== 13. 承認して反映: 自動の学びを止めている間は、判断を記録するだけ（ボタンは押せる） ====
{
  const props = [{ row: 8, pid: 'P1', target: 'ai_weight_override', current: '0.0008', proposed: '0.0004', conf: '中', rationale: '', impact: '', decision: '', rollback: '' }];
  const OFFN = '自動の学びを止めている間は、判断を記録するだけで、補正の値は変わりません';
  const BTN = '判断を保存してから、承認した案を設定に反映します（3〜6 分ほど）。判断を選んでいない案は承認になります';
  for (const note of ['', OWNER_NOTE]) {
    const off = review(planView({ proposals: props, learning: { autoUpdate: false, biasFactor: 1, note } }));
    assert.ok(off.includes('data-tip="' + BTN + '\n' + OFFN + '"><button class="btn" onclick="fcRvGo()">承認して反映</button>'), '押せるまま、説明で言う');
    assert.ok(off.includes('<span class="chip warn" data-tip="' + OFFN + '">未反映</span>'));
  }
  const on = review(planView({ proposals: props }));
  assert.ok(on.includes('data-tip="' + BTN + '"><button class="btn" onclick="fcRvGo()">承認して反映</button>'));
  assert.ok(on.includes('<span class="chip warn" data-tip="承認者が判断して反映すると、次の予測から効きます">未反映</span>'));
  // 確かめる文
  const asked = (v, dec) => run(`var __m = null, __ask = ask; ask = function(m){ __m = m; }; S.fc.view = __v; S.fc.dec = __dec; fcRvGo(); ask = __ask; S.fc.dec = {}; __m`, { __v: v, __dec: dec || {} });
  assert.equal(asked(planView({ proposals: props, learning: { autoUpdate: false, note: OWNER_NOTE } })), '1 件を承認します。\n' + OFFN + '（3〜6 分ほど）。');
  assert.equal(asked(planView({ proposals: props, learning: { autoUpdate: false, note: '' } }), { 8: '保留' }), '承認する案はありません。判断だけを記録します。保留 1 件は反映しません。\n' + OFFN + '（3〜6 分ほど）。');
  assert.equal(asked(planView({ proposals: props })), '1 件を承認して反映します。\n判断を保存してから反映します（3〜6 分ほど）。反映した補正は、次の予測から効きます。');
}

// ==== 14. 小さな直し: AI の効きの端の値・分析の案内の空白・スマホの説明 ====
{
  const p = (x) => run('aiPct(__x)', { __x: x });
  assert.deepEqual([-0.001, -1, NaN, '', '  ', null, undefined, 'abc', Infinity].map(p), Array(9).fill('-'), '空・数でない・マイナスは -');
  assert.deepEqual([1e-9, 1e-7, 0.0000099].map(p), ['0.001% 未満', '0.001% 未満', '0.001% 未満']);
  assert.deepEqual([0, '0', 0.00001, 0.0000123, 0.0008, '0.0004', 0.01, 1].map(p), ['0%', '0%', '0.001%', '0.0012%', '0.08%', '0.04%', '1%', '100%']);
  for (let e = -15; e <= 0; e++) assert.doesNotMatch(p(Math.pow(10, e)), /\de/i, '指数の書き方にしない: 1e' + e);
  assert.equal(run('qrValue("ai_weight_override", -0.01)'), '-');
  assert.equal(run('qrValue("ai_weight_override", "")'), '既定', '空は今までどおり既定');
  // 分析の案内: 「は出していません」の前に半角の空白を入れない
  assert.doesNotMatch(uiHtml, / は出していません/);
  // スマホ: 欄を選んだまま押せないボタンに触れたとき、欄の focusout で説明を消さない（消すのは説明の元とその中から外れたときだけ）
  const L = {};
  const t = loadUi(L);
  const box = vm.runInContext("$('tipbox')", t);
  const fire = (type, e) => (L[type] || []).forEach((fn) => fn(e));
  const on = () => box.classList.contains('on');
  const input = fakeEl('input'), btn = fakeEl('button', ['btn'], { disabled: 'disabled' });
  const act = fakeEl('span', ['act'], { 'data-tip': '保存するものがありません' }, [btn]);
  fakeEl('div', [], {}, [input, act]);
  fire('focusin', { target: input });
  assert.equal(on(), false);
  fire('pointerdown', { pointerType: 'touch', target: btn });
  assert.equal(on(), true);
  assert.equal(box.textContent, '保存するものがありません');
  fire('focusout', { target: input });
  assert.equal(on(), true, '欄の focusout では消さない');
  fire('focusout', { target: btn });
  assert.equal(on(), false, '説明の元の中から外れたら消す');
  // キーボード: 説明の印を選ぶと出て、そこから外れると消える
  const icon = fakeEl('span', ['tip'], { 'data-tip': '説明' });
  fire('focusin', { target: icon });
  assert.equal(on(), true);
  fire('focusout', { target: icon });
  assert.equal(on(), false);
}

console.log('app-ui-texts: all tests passed');
