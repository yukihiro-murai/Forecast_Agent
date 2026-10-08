#!/usr/bin/env node
/*
 * app-ui-v10-fix3.test.mjs — 版 10（v0.30.0）の最後の点検の直し 3（画面。UI.html の script を vm で読む）。
 *   H3. 自信だけ選び直した保存（表は変わらず、入力の記録に CONF を足した）は「保存しました」と知らせる
 *       （サーバーの返り値 inputLogRows。何も変わらない保存は今までどおり「変更はありませんでした」）
 * 名前と数字はテスト用の架空のもの。モックの上の確かめで、本物の Apps Script・ブラウザの上では動かしていない。
 *
 *   node app/tests/app-ui-v10-fix3.test.mjs
 */
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { setUpEnv, uiHtml, OWNER } from './gas-mock.mjs';

function loadUi() {
  const js = uiHtml.slice(uiHtml.indexOf('<script>') + 8, uiHtml.lastIndexOf('</script>'))
    .replace('<?!= charsJs ?>', 'var YOMI_POSE = new Proxy({}, { get: (t, k) => "<i data-pose=\\"" + String(k) + "\\"></i>" }); var CHAR_SVG = new Proxy({}, { get: () => "" });')
    .replace('<?!= bootJson ?>', JSON.stringify({ app: { name: 'T', version: 'x' }, user: { email: OWNER, isOwner: true, isAdmin: true, roles: [{ role: 'ADMIN', scopeType: 'ALL' }] }, setUp: true, allowed: true }));
  const ui = vm.createContext({ document: { getElementById: () => ({ innerHTML: '', classList: { add() {}, remove() {} }, set outerHTML(v) {} }), querySelector: () => null, addEventListener() {} },
    setTimeout: () => 0, clearTimeout() {}, confirm: () => false, window: { addEventListener() {}, innerWidth: 1280, innerHeight: 800 },
    google: { script: { run: new Proxy({}, { get: (t, k) => (k === 'withSuccessHandler' || k === 'withFailureHandler' ? () => ui.google.script.run : () => {}) }) } } });
  vm.runInContext(js, ui);
  return (code, vars) => { Object.assign(ui, vars || {}); return vm.runInContext(code, ui); };
}
const run = loadUi();
run('var __toasts = []; toast = function(m){ __toasts.push(m); }; jobReload = function(){};');
/** 保存（PLAN.EDIT）が終わったときの知らせ（返り値 r） */
const doneToast = (r) => JSON.parse(run(`__toasts = []; JOB_DONE['PLAN.EDIT'](__r, { planId: 'P1', action: 'INPUT.SAVE', args: { kind: 'product' } }); JSON.stringify(__toasts)`,
  { __r: Object.assign({ planId: 'P1', action: 'INPUT.SAVE' }, r) }));

// ==== H3. 知らせの文 ====
assert.deepEqual(doneToast({ changed: [], inputLogRows: 1 }), ['保存しました'], '自信だけの保存（CONF を足した）は保存した');
assert.deepEqual(doneToast({ changed: [], inputLogRows: 0 }), ['変更はありませんでした'], '何も変わらない保存');
assert.deepEqual(doneToast({ changed: ['PRODUCT'], inputLogRows: 2 }), ['保存しました']);
assert.deepEqual(doneToast({ changed: ['CONFIG'] }), ['保存しました'], '記録の数を返さない保存（担当者など）は今までどおり');
assert.deepEqual(doneToast({ changed: [] }), ['変更はありませんでした']);

// ==== H3. サーバーの返り値のまま（モックの上で自信だけ選び直して保存） ====
{
  const E = setUpEnv();
  const fy = E.run('appFy_(new Date())');
  const output = [['FY' + fy + ' 売上予測（自信製薬）']];
  for (let r = 2; r <= 25; r++) output.push([]);
  output.push(['年度合計（予測）', 900, 1000, 1100], [], []);
  for (let i = 0; i < 12; i++) output.push([new Date(fy, 3 + i, 1), 70, 80, 90, '', '', '', 80, '']);
  const P = E.seedPlan(E.makeBook('自信製薬', {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', '自信製薬'], ['[必須] 予測年度FY（YYYY）', fy], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
    OUTPUT: { values: output },
    PRODUCT: { values: [['Person', 'ProductName', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Reason'], ['鷹野', '製品A', new Date(fy, 3, 1), '+5%', 'a']] },
  }));
  E.call('apiListPlans()');
  const v = E.call('apiPlanView(__in)', { __in: { planId: P } });
  const rows = v.boot.input.product.map((r, i) => Object.assign({}, r, { fromRow: i, selfConf: '高い' }));
  const st = E.runJob('PLAN.EDIT', { planId: P, action: 'INPUT.SAVE', args: { kind: 'product', rows, fromRows: true }, inputHash: v.inputHash });
  assert.equal(st.status, 'DONE', st.error);
  assert.deepEqual(st.result.changed, [], '前提: 表は変わらない');
  assert.deepEqual(doneToast(st.result), ['保存しました'], '「変更はありませんでした」と言わない');
}

console.log('app-ui-v10-fix3: ok');
