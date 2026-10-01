#!/usr/bin/env node
/*
 * app-migration.test.mjs — 旧ブックの取り込み（段階2）の契約テスト。
 * 旧ブックのシート ↔ データ本体の ENG_* の相互変換が値・型・数式・表示形式を失わないこと、
 * 計算用ブックに組み立て直したときに自動変換で値が変わるセルを直せること、取り込みが冪等で所有者だけのものであることを確かめる。
 * GAS 上での実行の代わりではない（本物のシートの自動変換は、試しの取り込みの結果で確かめる）。
 *
 *   node app/tests/app-migration.test.mjs
 */

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { repoRoot, OWNER, MEMBER, J, isDate, makeEnv, setUpEnv } from './gas-mock.mjs';

const legacySrc = await readFile(path.join(repoRoot, 'Forecast_Agent.js'), 'utf8');

// ==== 1. 移すシートの一覧と見出しは旧来のコードと同じ ====
{
  const env = makeEnv();
  const reg = J(env.run('APP_ENGINE_SHEETS'));
  const notMigrated = J(env.run('APP_ENGINE_NOT_MIGRATED'));
  // 旧来の SHEETS 定数のシート名をすべて「移す」か「移さない」に分けてある
  const sheetsBlock = /const SHEETS = \{([\s\S]*?)\n\};/.exec(legacySrc)[1];
  const legacyNames = [...sheetsBlock.matchAll(/:\s*'([A-Z_]+)'/g)].map((m) => m[1]);
  assert.ok(legacyNames.length >= 30);
  for (const n of legacyNames) assert.ok(reg[n] || notMigrated.includes(n), `旧来のシート ${n} の扱いが決まっていない`);
  for (const n of Object.keys(reg)) assert.ok(legacyNames.includes(n), `${n} は旧来のシートにない`);
  // 表の見出しは旧来のコードにそのままの並びで書かれている
  for (const [name, d] of Object.entries(reg)) {
    if (d.mode !== 'table') continue;
    const pattern = new RegExp('\\[\\s*' + d.header.map((h) => "'" + h.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "'").join('\\s*,\\s*') + '\\s*\\]');
    assert.match(legacySrc, pattern, `${name} の見出しが Forecast_Agent.js の定義と違う`);
  }
  // ENG_ の表の列 = 計画・行番号・見出し・型の並び（文字列のまま持つ）
  const tables = J(env.run('APP_TABLES'));
  for (const [name, d] of Object.entries(reg)) {
    if (d.mode !== 'table') { assert.ok(!tables['ENG_' + name]); continue; }
    assert.deepEqual(tables['ENG_' + name].columns, ['plan_id', 'seq', ...d.header, '_types']);
    assert.equal(tables['ENG_' + name].raw, true);
  }
  for (const [name, d] of Object.entries(tables)) assert.equal(new Set(d.columns).size, d.columns.length, `${name} の列名が重複`);
}

// ==== 2. セル 1 つの変換（型ごと元に戻る） ====
{
  const env = makeEnv();
  const vals = ['', 'abc', '2026/04', '0012', '=not formula', 0, -0, 0.1 + 0.2, 1e21, -1234.5678, true, false, new Date(2026, 3, 1), new Date('2026-10-01T10:20:30.456Z')];
  for (const v of vals) {
    const e = env.run('appCellEncode_(__v, "")', { __v: v });
    const back = env.run('appCellDecode_(__e[0], __e[1])', { __e: e });
    assert.ok(env.run('appCellSame_(__a, __b)', { __a: v, __b: back }), `戻らない: ${String(v)}`);
    if (typeof v === 'number') assert.ok(Object.is(v, back), `数値がずれる: ${v}`);
    if (isDate(v)) assert.equal(back.getTime(), v.getTime());
  }
  assert.deepEqual(J(env.run(`appCellEncode_('', '=SUM(A1:A3)')`)), ['f', '=SUM(A1:A3)']);
  assert.equal(env.run(`appCellSame_('1', 1)`), false, '文字列と数値は別');
  assert.equal(env.run(`appCellSame_('', null)`), true);
  assert.deepEqual(J(env.run('[appA1_(0, 0), appA1_(9, 25), appA1_(0, 26), appA1_(1, 701)]')), ['A1', 'Z10', 'AA1', 'ZZ2']);
}

// ---- 旧ブックの見本（値・表示形式・数式。中身は作り物） ----
const D = (y, m, d = 1) => new Date(y, m - 1, d);
function legacySpec() {
  return {
    GUIDE: { values: [['売上予測ツール ガイド']] },
    CONFIG: {
      values: [
        ['項目', '値'],
        ['[必須] メーカー名（外部集計キー）', 'テスト製薬'],
        ['[必須] 予測年度FY（YYYY）', 2027],
        ['[必須] 担当者（カンマ区切り）', '鷹野,鶴田'],
        ['RELIABILITY_APPLY_ENABLED（信頼度を反映）', 1],
        ['AI_SCORE_BASIS（level/momentum）', 'level'],
        ['メモ', '0012'],   // 自動の書式なのに文字列のまま残っているセル（作り物の難しい例）
      ],
    },
    SALES_INPUT: {
      values: [
        ['client', 'service_type', 'product', 'target_month', 'input_amount', 'status', 'source_updated_at'],
        ['テスト製薬', 'BASE', '製品A', D(2023, 4), 1200000, 'closed', D(2026, 9, 30)],
        ['テスト製薬', 'SPOT', '開発', D(2023, 5), 300000, 'closed', D(2026, 9, 30)],
        ['テスト製薬', 'BASE', '製品A', '2023/06', -5000, 'closed', ''],
      ],
      formats: { D: '@', E: '#,##0' },   // 旧来は書いた後に D 列を書式なしテキストにする（日付のまま残る）
    },
    PRODUCT: {
      values: [
        ['Person', 'ProductName', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Reason'],
        ['鷹野', '製品A', D(2026, 4), '+10%', '新しい適応'],
        ['', '製品B', D(2026, 4), '0%', ''],
      ],
      formats: { C: 'yyyy/MM/dd', D: '@' },
    },
    OPINIONS: {
      values: [
        ['Person', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Confidence(0..1)', 'Note'],
        ['鷹野', D(2026, 4), '+5%', 0.7, ''],
        ['鶴田', D(2026, 4), '-2.5%', 0.5, '慎重'],
      ],
      formats: { C: '@' },
    },
    EVAL_LOG: {
      values: [
        ['eval_id', 'evaluated_at', 'client', 'target_month', 'scenario', 'pred', 'actual', 'ape', 'was_overridden', 'error_category', 'forecast_role',
          'is_planning_point_estimate', 'signed_error', 'abs_error', 'bias_direction', 'range_contains_actual', 'quarter_label', 'half_label', 'fy_label',
          'model_version', 'evaluation_policy_version', 'constraint_relevant_flag'],
        ['E1', D(2026, 5, 2), 'テスト製薬', '2026/04', 'neutral', 1000, 900, 0.111, 0, 'model_limitation', 'P50', 1, 100, 100, 'over', true, 'Q1', 'H1', 'FY2027', '2.4.0', 'policy-2026H1-v1', 1],
      ],
      formats: { D: '@' },
    },
    EVAL_COMPARE_MONTHLY: {
      values: [
        ['target_month', 'forecast_base', 'forecast_spot', 'forecast_total', 'actual_base', 'actual_spot', 'actual_total', 'gap_total', 'forecast_total_p10',
          'forecast_total_p50', 'forecast_total_p90', 'signed_error_p50', 'abs_error_p50', 'ape_p50', 'quarter_label', 'half_label', 'fy_label', 'over_flag',
          'under_flag', 'range_outside_flag', 'note_for_investigation', 'planning_point_estimate_label', 'range_label', '', '年間制約サマリー', '値'],
        [D(2026, 4), 800, 200, 1000, 700, 200, 900, -100, 800, 1000, 1200, 100, 100, 0.111, 'Q1', 'H1', 'FY2027', 1, 0, 0, '', 'P50', 'P10-P90', '', 'annual_actual_total', 900],
      ],
    },
    DASHBOARD: { values: [['metric', 'value', 'memo'], ['plan_point_estimate', 'P50', '']] },   // 見出しが旧来と違う
    OUTPUT: {
      values: [['FY2027 売上予測（テスト製薬）'], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [],
        ['年度合計（予測）', 9000, 10000, 11000, '', '', 2000, '', 10000, 500]],
      formulas: { J26: '=H26+I26' },
    },
    CALIBRATION_STATE: {
      values: [
        ['client', 'updated_at', 'updated_by', 'ai_weight_override', 'ai_max_abs_effect_override', 'ai_topic_disable_json', 'bias_correction_factor',
          'qual_scale_override', 'residual_month_bias_json', 'last_applied_quarter', 'last_applied_review_id', 'auto_update_enabled', 'note'],
        ['テスト製薬', D(2026, 9, 1), 'auto', '', '', '[]', 0.98, '', '{"1":0.01}', '', '', 1, ''],
      ],
    },
    'メモ': { values: [['人が足したシート']] },
  };
}
function bookUrl(ss) { return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=0'; }

// ==== 3. シート 1 枚: 保存の形にして戻すと元と同じ（値・数式・表示形式・大きさ） ====
{
  const env = setUpEnv();
  const book = env.makeBook('クライアント別売上予測', legacySpec());
  const snap = (name) => env.run('appSheetSnapshot_(__sh, __n)', { __sh: book.getSheetByName(name), __n: name === 'SALES_INPUT' ? 7 : 0 });
  for (const name of ['CONFIG', 'SALES_INPUT', 'PRODUCT', 'OPINIONS', 'EVAL_LOG', 'EVAL_COMPARE_MONTHLY', 'DASHBOARD', 'OUTPUT', 'CALIBRATION_STATE']) {
    const s = snap(name);
    const enc = env.run('appEngEncodeSheet_("PL-T", __s)', { __s: s });
    const dec = env.run('appEngDecodeSheet_(__e.sheetRow, __e.tableRows, __e.rowSegs, __e.formatRows)', { __e: enc });
    assert.deepEqual(J(env.run('appEngCompare_(__s, __d)', { __s: s, __d: dec })), [], `${name} が元に戻らない`);
  }
  const si = env.run('appEngEncodeSheet_("PL-T", __s)', { __s: snap('SALES_INPUT') });
  assert.equal(si.sheetRow.mode, 'table');
  assert.equal(si.tableRows.length, 3, '見出しの次の行から 1 行 = 1 件');
  assert.equal(si.tableRows[0]._types, 'sssdnsd');
  assert.equal(si.tableRows[2]._types, 'ssssnse', '同じ列でも行ごとに型を持つ（文字列の月・空のセル）');
  assert.equal(si.tableRows[0].target_month, D(2023, 4).toISOString());
  const fmtD = JSON.parse(si.formatRows.find((f) => f.col === '4').runs_json);
  assert.deepEqual(fmtD, [[1, 1000, '@']], '表示形式は列ごとの続いた範囲で持つ');
  // 表の外のセル（右側の集計の欄）は行ごとに持つ
  const cmp = env.run('appEngEncodeSheet_("PL-T", __s)', { __s: snap('EVAL_COMPARE_MONTHLY') });
  assert.equal(cmp.sheetRow.mode, 'table');
  assert.deepEqual(J(cmp.rowSegs.map((r) => [r.row_no, r.col_from, JSON.parse(r.cells_json)])), [
    ['1', '24', ['e', 's年間制約サマリー', 's値']],
    ['2', '24', ['e', 'sannual_actual_total', 'n900']],
  ]);
  // 見出しが違う表は、値を失わないよう行ごとに持つ
  const dash = env.run('appEngEncodeSheet_("PL-T", __s)', { __s: snap('DASHBOARD') });
  assert.equal(dash.sheetRow.mode, 'rows');
  assert.equal(dash.tableRows.length, 0);
  assert.match(dash.warnings[0], /見出しが旧来の定義と違う/);
  // 数式は数式として持つ
  const out = env.run('appEngEncodeSheet_("PL-T", __s)', { __s: snap('OUTPUT') });
  assert.equal(out.formulaCount, 1);
  assert.ok(out.rowSegs.find((r) => r.row_no === '26').cells_json.includes('"f=H26+I26"'));
  // 内容のハッシュは計画の ID に左右されない
  assert.equal(env.run('appEngEncodeSheet_("PL-OTHER", __s)', { __s: snap('SALES_INPUT') }).sheetRow.content_hash, si.sheetRow.content_hash);
}

// ==== 4. 試しの取り込み: データ本体に書かず、計算用ブックで組み立てを確かめる ====
{
  const env = setUpEnv();
  const book = env.makeBook('クライアント別売上予測', legacySpec());
  const before = { tables: env.data().getSheets().map((s) => [s.name, s.rows.length]), plans: env.table('PLANS').length };
  const r = env.call('apiMigrationInspect(__in)', { __in: { bookUrl: bookUrl(book) } });
  assert.deepEqual([r.client, r.fy, r.peopleCount], ['テスト製薬', 2027, 2]);
  assert.equal(r.lossless, true, '保存の形にして戻すと元と同じ');
  assert.equal(r.faithful, false, '文字列に固定したセルがあるので「そのまま」ではない');
  const by = Object.fromEntries(r.sheets.map((x) => [x.sheet, x]));
  assert.equal(by.SALES_INPUT.mismatch, 0);
  assert.equal(by.SALES_INPUT.repaired, 2, '書式なしテキストの列の日付は「書いてから形式を付ける」で直る');
  assert.equal(by.CONFIG.forcedText, 1, '自動の書式で文字列のまま残っていたセルは文字列に固定する');
  assert.equal(by.CONFIG.mismatch, 0);
  assert.equal(by.EVAL_LOG.mismatch + by.EVAL_LOG.repaired, 0, '書式なしテキストの列の文字列の月はそのまま');
  assert.equal(by.OUTPUT.mismatch, 0, '数式は計算後の値が一致する');
  assert.equal(by.DASHBOARD.mode, 'rows');
  assert.ok(r.missing.includes('POOL_PRIOR') && r.missing.includes('FORECAST_SNAPSHOT'));
  assert.deepEqual(r.notMigrated, ['GUIDE']);
  assert.deepEqual(r.unknown, ['メモ']);
  assert.match(r.contentHash, /^[0-9a-f]{64}$/);
  assert.equal(r.existingPlan, null);
  // 計算用ブックには旧ブックと同じ値が入っている
  const scratch = env.scratch();
  assert.ok(scratch && scratch.name === '売上予測アプリ 計算用（自動）');
  assert.equal(env.files[scratch.id].parent, env.props.APP_FOLDER_ID, '計算用ブックはシステムのフォルダの中');
  for (const name of Object.keys(by)) {
    const a = book.getSheetByName(name); const b = scratch.getSheetByName(name);
    const va = a.getRange(1, 1, a.getLastRow(), a.getLastColumn()).getValues();
    const vb = b.getRange(1, 1, a.getLastRow(), a.getLastColumn()).getValues();
    va.forEach((row, i) => row.forEach((v, j) => assert.ok(env.run('appCellSame_(__a, __b)', { __a: v, __b: vb[i][j] }), `${name} ${i + 1},${j + 1}`)));
    assert.equal(b.getMaxRows(), a.getMaxRows());
    assert.equal(b.getMaxColumns(), a.getMaxColumns());
  }
  assert.ok(!scratch.getSheets().some((s) => s.name.startsWith('_EMPTY_')), '空のシートは残さない');
  // データ本体は変わらない
  assert.deepEqual(env.data().getSheets().map((s) => [s.name, s.rows.length]), before.tables);
  assert.equal(env.table('PLANS').length, before.plans);
  const au = env.audit().filter((a) => a.action === 'MIGRATION.DRYRUN');
  assert.deepEqual(au.map((a) => [a.phase, a.result]), [['START', ''], ['END', 'OK']]);
  assert.equal(JSON.parse(au[1].after_json).contentHash, r.contentHash);
  // 旧ブックは何も変えない
  assert.deepEqual(book.getSheets().map((s) => s.name), Object.keys(legacySpec()));
  // URL ではないもの・旧ブックでないもの
  assert.throws(() => env.call(`apiMigrationInspect({ bookUrl: 'abc' })`), /URL を入れてください/);
  const other = env.makeBook('別のファイル', { Sheet1: { values: [['x']] } });
  assert.throws(() => env.call('apiMigrationInspect(__in)', { __in: { bookUrl: bookUrl(other) } }), /CONFIG シートがありません/);
  // 所有者だけ（管理者でも所有者でなければ拒否し、記録する）
  env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'M' })`);
  env.call(`apiGrantRole({ email: '${MEMBER}', role: 'ADMIN' })`);
  env.as(MEMBER);
  assert.throws(() => env.call('apiMigrationInspect(__in)', { __in: { bookUrl: bookUrl(book) } }), /権限がありません/);
  assert.throws(() => env.call('apiMigrationImport(__in)', { __in: { bookUrl: bookUrl(book), contentHash: r.contentHash } }), /権限がありません/);
  assert.deepEqual(env.audit().slice(-2).map((a) => [a.action, a.result]), [['MIGRATION.DRYRUN', 'DENIED'], ['MIGRATION.IMPORT', 'DENIED']]);
  assert.equal(env.state.lockHeld, false);
}

// ==== 5. 取り込み: 試しと同じ内容のときだけ・計画 1 つ分を入れ替える・同じ内容なら書かない ====
{
  const env = setUpEnv();
  const book = env.makeBook('クライアント別売上予測', legacySpec());
  const url = bookUrl(book);
  assert.throws(() => env.call('apiMigrationImport(__in)', { __in: { bookUrl: url } }), /先に「試しに読む」/);
  const dry = env.call('apiMigrationInspect(__in)', { __in: { bookUrl: url } });
  assert.throws(() => env.call('apiMigrationImport(__in)', { __in: { bookUrl: url, contentHash: 'x'.repeat(64) } }), /変わりました/);
  assert.equal(env.table('PLANS').length, 0, '内容が違えば何も書かない');
  assert.equal(env.table('CLIENTS').length, 0);
  // ほかの計画の行は残る
  env.run(`appInsertRows_('ENG_SALES_INPUT', [{ plan_id: 'PL-OTHER', seq: '2', client: 'x', _types: 's' }])`);
  const res = env.call('apiMigrationImport(__in)', { __in: { bookUrl: url, contentHash: dry.contentHash } });
  assert.equal(res.unchanged, false);
  assert.deepEqual([res.verified, res.verify], [true, []], '書いた後にデータ本体から読み戻して照合する');
  assert.equal(env.table('IMPORT_BATCHES')[0].status, 'OK');
  const plans = env.table('PLANS');
  assert.equal(plans.length, 1);
  const plan = plans[0];
  assert.equal(plan.plan_id, res.planId);
  assert.deepEqual([plan.fy, plan.client_label, plan.people_csv, plan.source_book_id, plan.locale, plan.time_zone, plan.state],
    ['2027', 'テスト製薬', '鷹野,鶴田', book.getId(), 'ja_JP', 'Asia/Tokyo', 'ACTIVE']);
  const clients = env.table('CLIENTS');
  assert.deepEqual(clients.map((c) => [c.client_id, c.client_name]), [[plan.client_id, 'テスト製薬']]);
  assert.equal(env.table('ENG_SALES_INPUT').filter((x) => x.plan_id === plan.plan_id).length, 3);
  assert.equal(env.table('ENG_SALES_INPUT').filter((x) => x.plan_id === 'PL-OTHER').length, 1, 'ほかの計画の行は残す');
  assert.equal(env.table('ENG_DASHBOARD').length, 0, '見出しが違う表は ENG_ROWS に持つ');
  assert.ok(env.table('ENG_ROWS').some((x) => x.sheet === 'DASHBOARD'));
  assert.equal(env.table('ENG_SHEETS').length, dry.sheets.length);
  assert.equal(env.table('IMPORT_BATCHES').length, 1);
  assert.equal(env.table('IMPORT_BATCHES')[0].content_hash, dry.contentHash);
  const imports = env.audit().filter((a) => a.action === 'MIGRATION.IMPORT' && a.phase === 'END');
  assert.deepEqual(imports.map((a) => a.result), ['FAILED', 'FAILED', 'OK'], '止めた取り込みも記録する');
  const end = imports[2];
  assert.deepEqual([end.result, end.entity_id, end.client_id], ['OK', plan.plan_id, plan.client_id]);
  // データ本体から組み立て直すと、旧ブックと同じ（保存した文字列から型ごと戻る）
  const stored = (table, f) => env.run(`appReadTable_('${table}')`).filter(f);
  for (const s of env.run('appReadTable_("ENG_SHEETS")').filter((x) => x.plan_id === plan.plan_id)) {
    const name = s.sheet;
    const tableRows = s.mode === 'table' ? stored('ENG_' + name, (x) => x.plan_id === plan.plan_id) : [];
    const segs = stored('ENG_ROWS', (x) => x.plan_id === plan.plan_id && x.sheet === name);
    const fmts = stored('ENG_FORMATS', (x) => x.plan_id === plan.plan_id && x.sheet === name);
    const dec = env.run('appEngDecodeSheet_(__s, __t, __g, __f)', { __s: s, __t: tableRows, __g: segs, __f: fmts });
    const snap = env.run('appSheetSnapshot_(__sh, __n)', { __sh: book.getSheetByName(name), __n: Number(s.fmt_columns) });
    assert.deepEqual(J(env.run('appEngCompare_(__a, __b)', { __a: snap, __b: dec })), [], `${name} がデータ本体から元に戻らない`);
  }
  // 同じ内容をもう一度: 書かない
  const rows0 = env.data().getSheets().map((s) => [s.name, s.rows.length]);
  const again = env.call('apiMigrationImport(__in)', { __in: { bookUrl: url, contentHash: dry.contentHash } });
  assert.equal(again.unchanged, true);
  assert.equal(env.table('IMPORT_BATCHES').length, 1);
  assert.deepEqual(env.data().getSheets().map((s) => [s.name, s.rows.length]), rows0);
  // 旧ブックが変わったら、試しからやり直してから取り込む（同じ計画を入れ替える）
  book.getSheetByName('SALES_INPUT').load({ values: [[], [], [], [], ['テスト製薬', 'BASE', '製品C', D(2023, 7), 1, 'closed', '']] });
  assert.throws(() => env.call('apiMigrationImport(__in)', { __in: { bookUrl: url, contentHash: dry.contentHash } }), /変わりました/);
  const dry2 = env.call('apiMigrationInspect(__in)', { __in: { bookUrl: url } });
  assert.notEqual(dry2.contentHash, dry.contentHash);
  assert.equal(dry2.existingPlan.planId, plan.plan_id);
  assert.equal(dry2.existingPlan.unchanged, false);
  const res2 = env.call('apiMigrationImport(__in)', { __in: { bookUrl: url, contentHash: dry2.contentHash } });
  assert.equal(res2.planId, plan.plan_id, '同じ旧ブックは同じ計画');
  assert.equal(env.table('ENG_SALES_INPUT').filter((x) => x.plan_id === plan.plan_id).length, 4, '入れ替えなので重ならない');
  assert.equal(env.table('ENG_SALES_INPUT').filter((x) => x.plan_id === 'PL-OTHER').length, 1);
  assert.equal(env.table('IMPORT_BATCHES').length, 2);
  assert.equal(env.table('CLIENTS').length, 1);
  const upd = env.audit().filter((a) => a.action === 'MIGRATION.IMPORT' && a.phase === 'END').slice(-1)[0];
  assert.equal(JSON.parse(upd.before_json).plan_id, plan.plan_id, '計画の変更前を記録する');
  assert.equal(env.state.lockHeld, false);
}

// ==== 5b. 読み戻しで違いが出たら、取り込みは「照合に失敗」として記録し、次の取り込みを止めない ====
{
  const env = setUpEnv();
  const book = env.makeBook('クライアント別売上予測', legacySpec());
  const url = bookUrl(book);
  const dry = env.call('apiMigrationInspect(__in)', { __in: { bookUrl: url } });
  // データ本体に書くときに文字が変わる（本物のシートでしか起きない差）を真似する
  const orig = env.run('appWriteBody_');
  env.run(`appWriteBody_ = function (sh, rows, width) { return __orig(sh, rows.map(r => r.map(v => String(v).replace('"slevel"', '"sLEVEL"'))), width); }`, { __orig: orig });
  const res = env.call('apiMigrationImport(__in)', { __in: { bookUrl: url, contentHash: dry.contentHash } });
  assert.equal(res.verified, false);
  assert.equal(res.verify[0].sheet, 'CONFIG');
  assert.equal(env.table('IMPORT_BATCHES')[0].status, 'VERIFY_FAILED');
  env.run('appWriteBody_ = __orig', { __orig: orig });
  const again = env.call('apiMigrationImport(__in)', { __in: { bookUrl: url, contentHash: dry.contentHash } });
  assert.deepEqual([again.unchanged, again.verified], [false, true], '照合に失敗した取り込みは「前回」に数えず、もう一度取り込める');
}

// ==== 6. 表を足した版: 足りない表は「状態」に出て、初期設定をもう一度実行すると作られる ====
{
  const env = setUpEnv();
  env.data().deleteSheet(env.data().getSheetByName('ENG_RUN_LOG'));
  const h = env.call('apiHealth()');
  const t = h.tables.find((x) => x.name === 'ENG_RUN_LOG');
  assert.deepEqual([t.ok, t.missing, t.engine], [false, true, true]);
  assert.match(t.note, /未作成/);
  assert.ok(h.tables.filter((x) => x.name !== 'ENG_RUN_LOG').every((x) => x.ok));
  assert.deepEqual(env.call('apiSetup()').tables, ['ENG_RUN_LOG']);
  assert.ok(env.call('apiHealth()').tables.every((x) => x.ok));
  // 列の多い表もシートの列数を合わせて作る（既定の 26 列を超える）
  assert.equal(env.data().getSheetByName('ENG_EVAL_INSIGHTS').getMaxColumns(), 27);
  assert.equal(env.data().getSheetByName('MEMBERS').getMaxColumns(), 10, '余った列は消してセルの上限を節約する');
}

// ==== 7. 入れ替え（appReplaceRows_）: 残す行・足す行・余りの行を消す・キーの重複は止める ====
{
  const env = setUpEnv();
  const row = (plan, seq) => ({ plan_id: plan, seq: String(seq), client: plan + seq, _types: 's' });
  env.run(`appInsertRows_('ENG_RUN_LOG', __r)`, { __r: [row('A', 2), row('A', 3), row('B', 2), row('A', 4)] });
  const r = env.call(`appReplaceRows_('ENG_RUN_LOG', (x) => x.plan_id !== 'A', __n)`, { __n: [row('A', 2)] });
  assert.deepEqual(r, { removed: 3, kept: 1, added: 1 });
  assert.deepEqual(env.table('ENG_RUN_LOG').map((x) => x.client), ['B2', 'A2']);
  assert.equal(env.data().getSheetByName('ENG_RUN_LOG').getLastRow(), 3, '余った古い行は消す');
  assert.throws(() => env.run(`appReplaceRows_('ENG_RUN_LOG', () => true, __n)`, { __n: [row('B', 2)] }), /同じキー/);
  assert.deepEqual(env.table('ENG_RUN_LOG').map((x) => x.client), ['B2', 'A2'], '止めたときは書かない');
}

console.log('app-migration: all tests passed');
