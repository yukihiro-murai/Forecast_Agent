#!/usr/bin/env node
/*
 * app-plan.test.mjs — 計画の画面（段階2-3）の契約テスト。
 * - 読むだけのブック（データ本体から組み立てる）が、旧来の画面の読み取りに Sheet と同じ値を返すこと（OUTPUT の予算の数式も）
 * - 画面の中身は旧来の webGetBootstrap_ を本物のまま動かしたもの（入力・予測・四半期レビュー）
 * - 保存（入力・予算・四半期レビューの承認）は旧来の webSave* を本物のまま動かし、変わったシートだけをデータ本体へ戻すこと
 * - 実行（B-3 など）は計算 → 保存の 2 段で戻すこと（旧来の計算の代わりで確かめる）
 * - 権限（予算策定担当・承認者）・画面を開いた後にデータが変わったときの止め方・記録（PLAN_ACTIONS・監査）
 *
 *   node app/tests/app-plan.test.mjs
 */

import assert from 'node:assert/strict';
import vm from 'node:vm';
import { J, MEMBER, OWNER, setUpEnv, uiHtml } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
function legacyBook(env) {
  const output = [['FY2026 売上予測（テスト製薬）']];
  for (let r = 2; r <= 23; r++) output.push([]);
  output.push([' 混合（主要）']);   // 24 行目（年度合計の 2 行上）
  output.push([]);
  output.push(['年度合計（予測）', 900, 1000, 1100, '', '', '', '', '', '']);   // 26 行目
  output.push([], []);
  const months = ['2026/04', '2026/05', '2026/06', '2026/07', '2026/08', '2026/09', '2026/10', '2026/11', '2026/12', '2027/01', '2027/02', '2027/03'];
  months.forEach((m) => output.push([m, 70, 80, 90, '', '', '', '', '', '']));
  const formulas = { H26: '=SUM(H29:H40)', I26: '=SUM(I29:I40)', J26: '=SUM(J29:J40)' };
  for (let i = 0; i < 12; i++) formulas['J' + (29 + i)] = `=H${29 + i}+I${29 + i}`;
  return env.makeBook('クライアント別売上予測', {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', 'テスト製薬'], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野,佐藤']] },
    OUTPUT: { values: output, formulas, formats: { A: '@' } },
    PRODUCT: { values: [['Person', 'ProductName', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Reason'], ['鷹野', '製品A', D(2026, 5), 5, '新規']] },
    CLIENT: { values: [['Person', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Reason']] },
    OPINIONS: { values: [['Person', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Confidence(0..1)', 'Note']] },
    DEV_SPOT: { values: [['Person', 'Month(yyyy/mm/dd)', 'Project', 'Amount(JPY)', 'Confidence(0..1)']] },
    PROCESS_STATUS: {
      values: [['step_key', 'last_run_date', 'last_run_by', 'status', 'target_client', 'record_count', 'error_summary'],
        ['step4_status', D(2026, 9, 1), 'owner', 'success', 'テスト製薬', 12, '']],
    },
    SALES_INPUT: { values: [['client', 'service_type', 'product', 'target_month', 'input_amount', 'status', 'source_updated_at'],
      ['テスト製薬', 'BASE', '製品A', '2025/04', 1000000, 'closed', D(2026, 9, 1)]], formats: { D: '@' } },
    RUN_LOG: { values: [['run_id', 'run_at', 'run_by', 'function_name', 'client', 'status', 'count', 'model_version', 'parameters_snapshot_json', 'input_data_hash', 'execution_duration_sec', 'error_summary']] },
    QUARTERLY_REVIEW: {
      values: [['四半期レビュー（テスト製薬）'], ['2026/04〜2026/06'], [], [], [], [], ['proposal_id', 'target', 'current', 'proposed', 'conf', 'rationale', 'impact', '承認', 'rollback', 'review_id'],
        ['P1', 'bias_correction_factor', 1, 0.98, 0.7, '過大', '-2%', '', 1, 'R-1'],
        ['P2', 'ai_weight_override', '', 0.4, 0.5, '外れ', '-1%', '', '', '']],
      formats: { G: '@' },
    },
  });
}
function imported() {
  const env = setUpEnv();
  const book = legacyBook(env);
  return { env, book, planId: env.seedPlan(book) };
}
const engRows = (env, sheet, planId) => env.table('ENG_' + sheet).filter((r) => r.plan_id === planId);

// ==== 1. 読むだけのブック: Sheet と同じ値（数式は予算の列の SUM と足し算だけ）、書き込みは止める ====
{
  const { env, book, planId } = imported();
  env.ctx.__plan = env.run(`appPlanOf_('${planId}')`);
  const v = (expr) => J(env.run(`(() => { const b = appStoreBook_(__plan); return ${expr}; })()`));
  const real = book.getSheetByName('OUTPUT');
  assert.deepEqual(v(`b.getSheetByName('OUTPUT').getDataRange().getValues().length`), real.getLastRow());
  assert.equal(v(`b.getSheetByName('OUTPUT').getRange(26, 3).getValue()`), 1000);
  assert.equal(v(`b.getSheetByName('OUTPUT').getRange('C26').getValue()`), 1000, 'A1 でも読める');
  assert.deepEqual(v(`b.getSheetByName('OUTPUT').getRange('A29:B30').getValues()`), [['2026/04', 70], ['2026/05', 70]]);
  assert.equal(v(`b.getSheetByName('OUTPUT').getRange(29, 10).getValue()`), 0, '空の H+I は 0（Sheets と同じ）');
  assert.equal(v(`b.getSheetByName('OUTPUT').getRange(26, 10).getValue()`), 0);
  assert.equal(v(`b.getSheetByName('OUTPUT').getRange(29, 10).getFormula()`), '=H29+I29');
  assert.equal(v(`b.getSheetByName('OUTPUT').getRange(500, 30).getValue()`), '', '範囲の外は空');
  assert.equal(v(`b.getSheetByName('NOPE')`), null);
  assert.equal(v(`b.getSheetByName('PRODUCT').getRange(2, 3).getValue() instanceof Date`), true, '日付は日付のまま');
  assert.throws(() => env.run(`appStoreBook_(__plan).getSheetByName('OUTPUT').getRange(1, 1).setValue('x')`), /画面の表示では使えない操作です（Range\.setValue）/);
  assert.throws(() => env.run(`appStoreBook_(__plan).insertSheet('X')`), /Spreadsheet\.insertSheet/);
  assert.throws(() => env.run(`appStoreBook_(__plan).getSheetByName('OUTPUT').getRange(1, 1, 0, 1)`), /at least 1/);
  assert.equal(env.run(`appEvalFormula_('=AVERAGE(B1:B2)', () => 1)`), '', 'ほかの数式は空として読む');
}

// ==== 2. 画面の中身: 旧来の webGetBootstrap_ を本物のまま動かす（データ本体から・計算用ブックなし） ====
{
  const { env, planId } = imported();
  const scratchBefore = env.props.APP_SCRATCH_SPREADSHEET_ID;
  const view = env.call('apiPlanView(__in)', { __in: { planId } });
  assert.equal(view.plan.planId, planId);
  assert.equal(view.boot.setup.client, 'テスト製薬');
  assert.equal(view.boot.setup.fy, 2026);
  assert.deepEqual(view.boot.setup.people, ['鷹野', '佐藤']);
  assert.equal(view.boot.input.product.length, 1);
  assert.deepEqual(view.boot.input.product[0], { person: '鷹野', product: '製品A', ym: '2026-05', step: view.boot.input.product[0].step, reason: '新規' });
  assert.equal(view.boot.output.sections[0].annual.p50, 1000);
  assert.equal(view.boot.output.sections[0].monthly.length, 12);
  assert.equal(view.boot.output.sections[0].monthly[0].row, 29, '予算を書く行番号');
  assert.equal(view.boot.quarterly.proposals.length, 2);
  assert.equal(view.boot.quarterly.reviewId, 'R-1');
  assert.equal(view.boot.steps.find((s) => s.key === 'step4_status').status, 'success');
  assert.equal(view.boot.bookUrl, undefined, '旧ブックの URL は出さない');
  assert.equal(view.boot.access, undefined);
  assert.equal(view.can.plan, true);
  assert.equal(env.props.APP_SCRATCH_SPREADSHEET_ID, scratchBefore, '計算用ブックは使わない');
  const again = env.call('apiPlanView(__in)', { __in: { planId } });
  assert.deepEqual(again.boot, view.boot, '入力が同じなら覚えておいたものを返す');
  // 日本語の多い大きな結果も覚えておける（CacheService の上限はバイト数。2026-10-02 画面が読み込み中のまま）
  env.run(`appJobPutResult_('BIG', { t: '日本語'.repeat(40000) })`);
  assert.equal(J(env.run(`appJobGetResult_('BIG')`)).value.t.length, 120000);
}

// ==== 3. 保存: 旧来の webSave* を本物のまま動かし、変わったシートだけを戻す ====
{
  const { env, planId } = imported();
  const view = env.call('apiPlanView(__in)', { __in: { planId } });
  const before = Object.fromEntries(env.table('ENG_SHEETS').filter((r) => r.plan_id === planId).map((r) => [r.sheet, r.content_hash]));
  // 入力（製品）を全面書き換え
  const rows = [{ person: '鷹野', product: '製品A', ym: '2026-06', step: '10', reason: '増産' }, { person: '佐藤', product: '製品B', ym: '2026-07', step: '-5', reason: '' }];
  const st = env.runJob('PLAN.EDIT', { planId, action: 'INPUT.SAVE', args: { kind: 'product', rows }, inputHash: view.inputHash });
  assert.equal(st.status, 'DONE', st.error);
  assert.deepEqual(st.result.changed, ['PRODUCT'], '変わったシートだけ');
  assert.equal(st.result.result.saved, 2);
  const prod = engRows(env, 'PRODUCT', planId);
  assert.equal(prod.length, 2);
  assert.equal(prod[0].ProductName, '製品A');
  assert.equal(prod[0]._types.charAt(2), 'd', '月は日付で持つ（旧来と同じ）');
  assert.equal(prod[0]._types.charAt(3), 'n', '増減率は数値に変わる（Sheets の自動変換と同じ）');
  const after = Object.fromEntries(env.table('ENG_SHEETS').filter((r) => r.plan_id === planId).map((r) => [r.sheet, r.content_hash]));
  Object.keys(before).filter((k) => k !== 'PRODUCT').forEach((k) => assert.equal(after[k], before[k], k + ' は変えない'));
  const row = env.table('PLAN_ACTIONS').slice(-1)[0];
  assert.equal(row.action, 'INPUT.SAVE'); assert.equal(row.actor_email, OWNER); assert.equal(row.changed_sheets_json, '["PRODUCT"]');
  const audit = env.audit().filter((a) => a.action === 'PLAN.INPUT.SAVE');
  assert.equal(audit.length, 2, '開始と終了');
  assert.match(audit[0].detail_json, /製品B/, '保存した中身を記録に残す');
  const v2 = env.call('apiPlanView(__in)', { __in: { planId } });
  assert.equal(v2.boot.input.product.length, 2, '画面にすぐ出る');
  assert.equal(v2.recent[0].action, 'INPUT.SAVE');
  assert.deepEqual(v2.recent[0].changed, ['PRODUCT'], '変わったシートは一覧で渡す（文字列のままだと「進み」が表示できない）');
  // 画面を開いた後にデータが変わったら止める（古い画面のまま上書きしない）
  const stale = env.runJob('PLAN.EDIT', { planId, action: 'INPUT.SAVE', args: { kind: 'product', rows: [] }, inputHash: view.inputHash });
  assert.equal(stale.status, 'FAILED');
  assert.match(stale.error, /画面を開き直して/);
  assert.equal(engRows(env, 'PRODUCT', planId).length, 2);
  // 予算（OUTPUT の H/I 列）
  const b = env.runJob('PLAN.EDIT', { planId, action: 'BUDGET.SAVE', args: { rows: [{ row: 29, adopted: 85, uplift: 5 }, { row: 3, adopted: 1 }] }, inputHash: v2.inputHash });
  assert.equal(b.status, 'DONE', b.error);
  assert.deepEqual(b.result.changed, ['OUTPUT']);
  const v3 = env.call('apiPlanView(__in)', { __in: { planId } });
  const m = v3.boot.output.sections[0].monthly[0];
  assert.deepEqual([m.adopted, m.uplift, m.final], [85, 5, 90], 'J 列（数式）は H+I');
  assert.equal(v3.boot.output.sections[0].annual.final, 90, '年度合計は SUM');
  // 大きい入力（処理の記録の上限を超える）も保存できる
  const many = Array.from({ length: 120 }, (_, i) => ({ person: '鷹野', product: '製品' + i, ym: '2026-08', step: '1', reason: '理由'.repeat(10) }));
  const big = env.runJob('PLAN.EDIT', { planId, action: 'INPUT.SAVE', args: { kind: 'product', rows: many }, inputHash: v3.inputHash });
  assert.equal(big.status, 'DONE', big.error);
  assert.equal(engRows(env, 'PRODUCT', planId).length, 120);
}

// ==== 4. 実行: 計算 → 保存の 2 段（旧来の計算の代わりで確かめる） ====
{
  const { env, planId } = imported();
  env.run(`appLegacyEngine_ = function (svc) {
    return { VERSION: 'stub-1', SOURCE_SHA256: 'stub', WEB_SOURCE_SHA256: 'stub-web',
      webRunDashboard() {
        const ss = svc.SpreadsheetApp.getActiveSpreadsheet();
        let sh = ss.getSheetByName('DASHBOARD');
        if (!sh) sh = ss.insertSheet('DASHBOARD');
        sh.getRange(1, 1, 2, 3).setValues([['metric', 'value', 'note'], ['sMAPE', 0.12, svc.Session.getActiveUser().getEmail()]]);
        return { boot: { big: 'x'.repeat(100000) } };
      } };
  };`);
  const st = env.runJob('PLAN.RUN', { planId, action: 'EVAL.DASHBOARD' });
  assert.equal(st.status, 'DONE', st.error);
  assert.deepEqual(st.result.changed, ['DASHBOARD']);
  assert.deepEqual(st.result.result, {}, '画面の中身（boot）は記録に残さない');
  const dash = engRows(env, 'DASHBOARD', planId);
  assert.equal(dash.length, 1);
  assert.equal(dash[0].note, OWNER, '頼んだ人を「操作した人」として見せる');
  const row = env.table('PLAN_ACTIONS').slice(-1)[0];
  assert.equal(row.action, 'EVAL.DASHBOARD'); assert.equal(row.web_sha256, 'stub-web');
  assert.ok(env.audit().some((a) => a.action === 'PLAN.EVAL.DASHBOARD.CALC') && env.audit().some((a) => a.action === 'PLAN.EVAL.DASHBOARD.SAVE'));
}

// ==== 5. 権限: 保存・実行は予算策定担当（その計画のクライアント）以上、承認は承認者。未定義の操作は止める ====
{
  const { env, planId } = imported();
  const clientId = env.table('PLANS')[0].client_id;
  env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'M' })`);
  env.as(MEMBER);
  const view = env.call('apiPlanView(__in)', { __in: { planId } });
  assert.equal(view.can.plan, false, '見るのは社内全員');
  assert.throws(() => env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.EDIT', payload: { planId, action: 'INPUT.SAVE', args: { kind: 'product', rows: [] } } } }), /権限がありません/);
  env.as(OWNER);
  env.call('apiGrantRole(__in)', { __in: { email: MEMBER, role: 'PLANNER', scopeType: 'CLIENT', clientId } });
  env.as(MEMBER);
  const v = env.call('apiPlanView(__in)', { __in: { planId } });
  assert.deepEqual([v.can.plan, v.can.approve], [true, false]);
  assert.throws(() => env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.EDIT', payload: { planId, action: 'REVIEW.DECIDE', args: { rows: [{ row: 8, decision: '承認' }] } } } }), /権限がありません/, '承認は承認者だけ');
  assert.throws(() => env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.RUN', payload: { planId, action: 'REVIEW.APPLY' } } }), /権限がありません/);
  assert.throws(() => env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.EDIT', payload: { planId, action: 'EVAL.REPORT' } } }), /未定義の操作です/, '実行の操作を保存として始められない');
  assert.throws(() => env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.RUN', payload: { planId, action: 'FORECAST.RUN' } } }), /未定義の操作です/);
  assert.throws(() => env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.RUN_SAVE', payload: { planId, action: 'EVAL.REPORT' } } }), /画面から始められません/);
  env.as(OWNER);
  env.call('apiGrantRole(__in)', { __in: { email: MEMBER, role: 'APPROVER', scopeType: 'ALL' } });
  env.as(MEMBER);
  const st = env.runJob('PLAN.EDIT', { planId, action: 'REVIEW.DECIDE', args: { rows: [{ row: 8, decision: '承認' }, { row: 9, decision: '保留' }] }, inputHash: v.inputHash });
  assert.equal(st.status, 'DONE', st.error);
  assert.deepEqual(st.result.changed, ['QUARTERLY_REVIEW']);
  const q = env.call('apiPlanView(__in)', { __in: { planId } }).boot.quarterly.proposals;
  assert.deepEqual(q.map((x) => x.decision), ['承認', '保留']);
  const bad = env.runJob('PLAN.EDIT', { planId, action: 'REVIEW.DECIDE', args: { rows: [{ row: 8, decision: 'OK' }] } });
  assert.equal(bad.status, 'FAILED');
  assert.match(bad.error, /承認 \/ 却下 \/ 保留/, '旧来と同じ確かめ方');
}

// ==== 6. A-2 売上の取り込み（旧来の関数を本物のまま。元は設定の「ZAC の実績のスプレッドシート」）と担当者の保存 ====
{
  const { env, planId } = imported();
  const ext = Array.from({ length: 70 }, () => '');
  const rec = (client, cat, product, month, amount) => { const r = ext.slice(); r[40] = client; r[45] = cat; r[49] = product; r[56] = month; r[65] = amount; return r; };
  const zac = env.makeBook('Veeva 売上分析ツール', { '*2025_actual_value': { cols: 70, values: [ext.map((_, i) => 'c' + (i + 1)),
    rec('テスト製薬', 'ベース', '製品A', D(2025, 5, 15), 1200000), rec('テスト製薬', 'スポット', '製品B', D(2025, 6, 10), 300000),
    rec('別の製薬', 'ベース', '製品X', D(2025, 5, 15), 999)] } });
  const noSource = env.runJob('PLAN.RUN', { planId, action: 'IMPORT.SALES' });
  assert.equal(noSource.status, 'FAILED');
  assert.match(noSource.error, /ZAC の実績のスプレッドシート/, '元を決めていなければ取り込まない（旧来の既定に頼らない）');
  assert.equal(env.call('apiPlanView(__in)', { __in: { planId } }).sourceReady, false);
  env.call('apiSaveSetting(__in)', { __in: { key: 'source.zac_spreadsheet', value: 'https://docs.google.com/spreadsheets/d/' + zac.getId() + '/edit' } });
  const st = env.runJob('PLAN.RUN', { planId, action: 'IMPORT.SALES' });
  assert.equal(st.status, 'DONE', st.error);
  assert.ok(st.result.changed.includes('SALES_INPUT'));
  const sales = engRows(env, 'SALES_INPUT', planId);
  assert.deepEqual(sales.map((r) => [r.product, r.target_month, r.input_amount]), [['製品A', '2025/05', '1200000'], ['製品B', '2025/06', '300000']], 'そのクライアントの分だけ、旧来と同じ形で');
  assert.equal(st.result.result.count, 2);
  const calc = env.audit().filter((a) => a.action === 'PLAN.IMPORT.SALES.CALC' && a.phase === 'END').slice(-1)[0];
  assert.ok(calc, '計算の記録');
  const scratchBook = env.sheetsById[env.props.APP_SCRATCH_SPREADSHEET_ID];
  assert.deepEqual(scratchBook.getSheets().map((x) => x.getName()).filter((n) => !/^_/.test(n)).sort(),
    ['CLIENT', 'CONFIG', 'DEV_SPOT', 'OPINIONS', 'PROCESS_STATUS', 'PRODUCT', 'RUN_LOG', 'SALES_INPUT'], '使うシートだけを組み立てる（6 分の上限）');
  const steps = env.call('apiPlanView(__in)', { __in: { planId } }).boot.steps;
  assert.equal(steps.find((x) => x.key === 'step1_status').status, 'success');
  // 担当者（管理者）。クライアントと年度は変えない
  const v = env.call('apiPlanView(__in)', { __in: { planId } });
  const pe = env.runJob('PLAN.EDIT', { planId, action: 'SETUP.PEOPLE', args: { peopleCsv: '鷹野、佐藤 , 田中' }, inputHash: v.inputHash });
  assert.equal(pe.status, 'DONE', pe.error);
  assert.deepEqual(pe.result.changed, ['CONFIG']);
  const v2 = env.call('apiPlanView(__in)', { __in: { planId } });
  assert.deepEqual(v2.boot.setup.people, ['鷹野', '佐藤', '田中']);
  assert.equal(v2.boot.setup.client, 'テスト製薬');
  assert.equal(env.table('PLANS')[0].people_csv, '鷹野,佐藤,田中', '計画の表にも残す');
  const empty = env.runJob('PLAN.EDIT', { planId, action: 'SETUP.PEOPLE', args: { peopleCsv: ' , ' } });
  assert.match(empty.error, /1 人以上/);
  env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'M' })`);
  env.call('apiGrantRole(__in)', { __in: { email: MEMBER, role: 'PLANNER', scopeType: 'ALL' } });
  env.as(MEMBER);
  assert.throws(() => env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.RUN', payload: { planId, action: 'IMPORT.SALES' } } }), /権限がありません/, '取り込みは管理者');
  assert.throws(() => env.call('apiStartJob(__in)', { __in: { kind: 'PLAN.EDIT', payload: { planId, action: 'SETUP.PEOPLE', args: { peopleCsv: 'x' } } } }), /権限がありません/);
}

// ==== 7. 画面: 実際の中身で 5 つのタブをすべて描ける（描く途中で止まらない） ====
{
  const { env, planId } = imported();
  env.runJob('PLAN.EDIT', { planId, action: 'BUDGET.SAVE', args: { rows: [{ row: 29, adopted: 85, uplift: 5 }] }, inputHash: env.call('apiPlanView(__in)', { __in: { planId } }).inputHash });
  const view = env.call('apiPlanView(__in)', { __in: { planId } });
  const latest = env.call('apiForecastLatest(__in)', { __in: { planId } });
  const js = uiHtml.slice(uiHtml.indexOf('<script>') + 8, uiHtml.lastIndexOf('</script>'))
    .replace('<?!= charsJs ?>', 'var YOMI_POSE = new Proxy({}, { get: () => "" }); var CHAR_SVG = new Proxy({}, { get: () => "" });')
    .replace('<?!= bootJson ?>', JSON.stringify({ app: { name: 'T', version: 'x' }, user: { email: OWNER, isOwner: true, isAdmin: true, roles: [] }, setUp: true, allowed: true }));
  const el = () => ({ innerHTML: '', classList: { add() {}, remove() {} }, set outerHTML(v) {} });
  const ui = vm.createContext({ document: { getElementById: el, querySelector: () => null, addEventListener() {} }, setTimeout: () => 0, clearTimeout() {}, confirm: () => false, window: { addEventListener() {}, innerWidth: 1280, innerHeight: 800 },
    google: { script: { run: new Proxy({}, { get: (t, k) => (k === 'withSuccessHandler' || k === 'withFailureHandler' ? () => ui.google.script.run : () => {}) }) } } });
  vm.runInContext(js, ui);
  ui.__v = view; ui.__d = latest;
  vm.runInContext(`S.view = 'forecast'; S.fc.plans = [{ planId: __v.plan.planId, clientName: 'x', fy: '2026' }]; S.fc.planId = __v.plan.planId; S.fc.view = __v; S.fc.data = __d;`, ui);
  for (const tab of ['forecast', 'input', 'eval', 'quarterly', 'steps']) {
    const html = vm.runInContext(`S.fc.tab = '${tab}'; viewForecast()`, ui);
    assert.ok(!/このタブを表示できませんでした/.test(html), tab + ' タブを描ける: ' + (html.match(/<p class="note" style="margin-top:6px">([^<]*)/) || [])[1]);
  }
  assert.match(vm.runInContext(`S.fc.tab = 'steps'; viewForecast()`, ui), /予算の保存[\s\S]*OUTPUT/, '最近の操作に変わったシートが出る');
  // 旧来の計算が確認を求めたら、確認の窓（ブラウザが黙って「やめる」にすることがある）ではなく、画面の中のカードで聞く
  vm.runInContext(`S.fc.tab = 'eval'; toast = function(){}; JOB_DONE['FORECAST.RUN_CALC']({ needConfirm: { key: 'extreme', title: '極端な入力', message: '増減率 +80%' } }, { planId: S.fc.planId, confirms: [] })`, ui);
  const cf = vm.runInContext(`viewForecast()`, ui);
  assert.match(cf, /極端な入力[\s\S]*増減率 \+80%[\s\S]*このまま実行/, '確認のカードが予測と予算のタブに出る');
  assert.equal(vm.runInContext(`S.fc.tab`, ui), 'forecast');
  assert.doesNotMatch(vm.runInContext(`viewForecast()`, ui), /value="[0-9]+\.[0-9]+"/, '採用予測の欄は小数を見せない');
  // 検証のタブの、精度の推移と学習の影
  ui.__learn = env.call('apiLearningView(__in)', { __in: { planId } });
  const evalHtml = vm.runInContext(`S.fc.learn = __learn; S.fc.tab = 'eval'; viewForecast()`, ui);
  assert.match(evalHtml, /月の誤差（平均）[\s\S]*精度の推移[\s\S]*学習の影/);
  assert.ok(!/undefined|NaN/.test(evalHtml.split('検証の更新')[0]), '精度の推移に undefined や NaN を出さない');
  ui.__pool = env.call('apiPoolPreview()');
  assert.match(vm.runInContext(`S.pfPool = __pool; pfPoolCard()`, ui), /全計画での学習[\s\S]*各計画に入れる/);
  // 旧ブックから移す画面・旧ブックの片付けは無い（このアプリだけを使う。2026-10-05）
  assert.equal(vm.runInContext(`[typeof viewMigrate, typeof legacyBooksCard, typeof NAV_OWNER].join()`, ui), 'undefined,undefined,undefined');
  assert.doesNotMatch(vm.runInContext(`JSON.stringify(NAV_ADMIN) + Object.keys(NAV_ICON)`, ui), /旧ブック|migrate/, 'メニューに旧ブックを出さない');
  // ホーム: 数字・計画の状態・最近の動きと、よみのセリフ（状況から選ぶ）
  ui.__home = env.call('apiHome()');
  const homeHtml = vm.runInContext(`S.view = 'home'; B.home = __home; viewHome()`, ui);
  assert.match(homeHtml, /年間予算[\s\S]*暫定実績[\s\S]*着地見込み[\s\S]*見通しの空模様[\s\S]*データからわかること/);
  // 2026-10-06 村井さん: ホームは全体の概要だけ（個社の表・名前は出さない。最近の動き・あなたの役割も出さない）
  assert.doesNotMatch(homeHtml, /最近の動き|あなたの役割|クライアント|計画の状態/);
  const visible = homeHtml.replace(/data-tip="[^"]*"/g, '');
  assert.doesNotMatch(visible, /テスト製薬|別の製薬/, 'ホームの見える文字に個社の名前を出さない（カーソルの説明だけ）');
  // 2026-10-06 村井さん: 「よみが観測しました（時刻）」は出さない・1 カラム（左右に分けない）・「状態を見る」のボタンは出さない
  assert.doesNotMatch(homeHtml, /よみが観測しました|class="dash"|状態を見る/);
  // 2026-10-06 村井さん: 読み込みの待ちは、右下ではなく画面の中で、よみが「観測中…」と話す
  const waitHtml = vm.runInContext(`S.homeLoading = true; B.home = null; var __w = viewHome(); B.home = __home; S.homeLoading = false; __w`, ui);
  assert.match(waitHtml, /class="card loading"[\s\S]*観測中…[\s\S]*読み込んでいます/);
  assert.doesNotMatch(waitHtml, /読み込み中…/);
  assert.doesNotMatch(String(vm.runInContext(`busy.toString()`, ui)), /観測中…/, '右下に「観測中…」を出さない');
  // 保存中・裏の処理の間も、画面の中（見出しの下）でよみが話す（2026-10-06 村井さん）
  const saving = vm.runInContext(`S.writing = true; S.writingFn = 'apiSaveSetting'; S.writeShow = true; var __s = jobBanner(); S.writing = false; S.writeShow = false; __s`, ui);
  assert.match(saving, /class="card loading"[\s\S]*保存中…[\s\S]*保存しています/);
  assert.equal(vm.runInContext(`S.writing = true; S.writeShow = false; var __q = jobBanner(); S.writing = false; __q`, ui), '', '0.6 秒までは出さない');
  const jobSave = vm.runInContext(`S.job = { id: 'j', kind: 'PLAN.EDIT', payload: { action: 'INPUT.SAVE' }, status: 'RUNNING', started: Date.now() - 5000 }; var __j = jobBanner(); S.job = null; __j`, ui);
  assert.match(jobSave, /class="card loading"[\s\S]*保存中…[\s\S]*処理中・5 秒/);
  assert.doesNotMatch(String(vm.runInContext(`busy.toString()`, ui)), /innerHTML|classList/, '右下には何も出さない');
  assert.ok(!/undefined|NaN/.test(homeHtml), 'ホームに undefined や NaN を出さない');
  assert.doesNotMatch(homeHtml, /旧アプリ/, 'ホームに旧アプリの案内を出さない');
  const says = JSON.parse(vm.runInContext(`JSON.stringify(yomiSays(__home))`, ui));
  assert.ok(says.length >= 1 && says.every((x) => x.text && x.pose && x.lvl), JSON.stringify(says));
  // 承認待ちがあれば、承認者に確認を頼む（いちばん上）
  const h2 = JSON.parse(JSON.stringify(ui.__home));
  h2.approvals = [{ planId: h2.plans[0].planId, clientName: 'テスト製薬', fy: h2.fy, versionNo: 2, submittedBy: 'planner' }];
  ui.__h2 = h2;
  assert.match(JSON.parse(vm.runInContext(`JSON.stringify(yomiSays(__h2)[0])`, ui)).text, /承認待ちの公式版が 1 件[\s\S]*確認をお願いします/);
  // 仕組みの異常（バックアップが止まった等）はセリフで知らせ、「状態を見る」のボタンは付けない（2026-10-06）
  const h3 = JSON.parse(JSON.stringify(ui.__home));
  h3.system = { backup: { enabled: false }, housekeeping: { ok: false, problems: ['x'] }, journal: null };
  ui.__h3 = h3;
  const says3 = JSON.parse(vm.runInContext(`JSON.stringify(yomiSays(__h3))`, ui));
  assert.ok(says3.some((x) => /バックアップ/.test(x.text)) && says3.some((x) => /手入れ/.test(x.text)), JSON.stringify(says3));
  assert.ok(says3.every((x) => !x.act || x.act[0] !== '状態を見る'), JSON.stringify(says3));
  assert.doesNotMatch(vm.runInContext(`B.home = __h3; viewHome()`, ui), /状態を見る/);
  vm.runInContext(`B.home = __home`, ui);
  vm.runInContext(`S.view = 'forecast'`, ui);
  // 根拠のタブ
  ui.__basis = env.call('apiForecastBasis(__in)', { __in: { planId } });
  const basisHtml = vm.runInContext(`S.fc.basis = __basis; S.fc.tab = 'basis'; viewForecast()`, ui);
  assert.match(basisHtml, /月ごとの内訳[\s\S]*補正[\s\S]*AI 調査の根拠[\s\S]*前回の予測からの変化/);
  assert.ok(!/undefined|NaN/.test(basisHtml), '根拠のタブに undefined や NaN を出さない');
  // 公式版のタブ（版がまだ無いとき・出したとき）
  ui.__ver = env.call('apiVersionList(__in)', { __in: { planId } });
  const verHtml = vm.runInContext(`S.fc.ver = __ver; S.fc.tab = 'version'; viewForecast()`, ui);
  assert.match(verHtml, /公式版の最終予算[\s\S]*今の内容を公式版として出す[\s\S]*まだ版がありません/);
  env.call('apiVersionSubmit(__in)', { __in: { planId, note: 'テスト' } });
  ui.__ver = env.call('apiVersionList(__in)', { __in: { planId } });
  const verHtml2 = vm.runInContext(`S.fc.ver = __ver; S.fc.verCmp = __ver.versions[0].versionId; viewForecast()`, ui);
  assert.match(verHtml2, /承認待ち（v1）[\s\S]*v1 と今の比べ/);
  assert.ok(!/undefined|NaN/.test(verHtml2), '公式版のタブに undefined や NaN を出さない');
  // メーカー（計画の一覧とクライアントをまとめた画面。2026-10-06）: わかる範囲のメーカーと、年度ごとの予算策定の状態だけ（数字は出さない）
  ui.__pf = env.call('apiPortfolio()').plans;
  ui.__cand = { candidates: [{ zac: 'ｼﾝｷ製薬(株)', display: 'シンキ製薬' }, { zac: 'テスト製薬', display: 'テスト製薬' }], defaultFy: 2026, existing: [] };
  const mkHtml = vm.runInContext(`S.pf = __pf; S.pfCand = __cand; S.mkFy = String(__pf[0].fy); B.user.isAdmin = true; viewMakers()`, ui);
  assert.match(mkHtml, /<h1>メーカー<\/h1><select[^>]*data-near[\s\S]*ステータス[\s\S]*年間予算[\s\S]*暫定実績[\s\S]*着地見込み[\s\S]*シンキ製薬[\s\S]*未着手[\s\S]*テスト製薬[\s\S]*(策定中|承認待ち|承認済み)[\s\S]*合計[\s\S]*計画を作る/);
  assert.equal((mkHtml.match(/>テスト製薬</g) || []).length, 1, '計画と候補の同じメーカーは 1 行にまとめる');
  assert.doesNotMatch(mkHtml, /ZAC コード|絞り込|並べ方|undefined|NaN/, '絞り込み・並べ替え・ZAC コードは出さない');
  // 年度は題のすぐ横に出す
  assert.match(vm.runInContext(`viewTitle('<div class="page-head"><h1>メーカー</h1><select data-near="1"></select></div>')`, ui), /<h1[^>]*>メーカー<\/h1><span class="vnear"><select/);
  // 予測は、まずメーカーを選ぶ（計画を勝手に開かない）
  const pick = vm.runInContext(`var __keep = S.fc.planId; S.fc.planId = null; S.fc.plans = [{ planId: 'P1', clientName: 'テスト製薬', fy: '2026' }]; var __p = viewForecast(); S.fc.planId = __keep; __p`, ui);
  assert.match(pick, /メーカーを選ぶ[\s\S]*テスト製薬[\s\S]*FY2026[\s\S]*開く/);
  // 計画のないメーカーの「計画を作る」は、下の欄に名前と年度を入れる
  vm.runInContext(`mkCreate('ｼﾝｷ製薬(株)', '2026')`, ui);
  assert.equal(vm.runInContext(`S.pcClient + '|' + S.pcFy`, ui), 'ｼﾝｷ製薬(株)|2026');
  // メニュー: ホーム・メーカー・予測・設定・記録と状態（メンバーは設定へ、操作の記録と状態はまとめる）
  assert.deepEqual(JSON.parse(vm.runInContext(`JSON.stringify(NAV_ADMIN.map(function(n){ return n[1]; }))`, ui)), ['ホーム', 'メーカー', '予測', '設定', '記録と状態']);
  ui.__dir = env.call('apiListDirectory()'); ui.__set = env.call('apiListSettings()').settings;
  const setHtml = vm.runInContext(`S.dir = __dir; S.settings = __set; S.setTab = 'members'; viewSettings()`, ui);
  assert.match(setHtml, /<h1>設定<\/h1>[\s\S]*業務[\s\S]*メンバー[\s\S]*表示名[\s\S]*学習[\s\S]*役割を付ける/);
  assert.doesNotMatch(vm.runInContext(`S.setTab = 'names'; viewSettings()`, ui), /ZAC コード/);
  assert.match(vm.runInContext(`S.rec = 'health'; S.health = null; viewRecords()`, ui), /<h1>記録と状態<\/h1>[\s\S]*状態[\s\S]*操作の記録[\s\S]*点検しています/);
  vm.runInContext(`S.setTab = 'biz'`, ui);
}

console.log('app-plan: all tests passed');
