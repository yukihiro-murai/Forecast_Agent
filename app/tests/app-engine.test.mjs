#!/usr/bin/env node
/*
 * app-engine.test.mjs — 旧来の計算を新アプリで動かす土台（段階2-2）の契約テスト。
 * - LegacyEngine.js が Forecast_Agent.js をそのまま包んだもの（1 文字も違わない）であること
 * - 「今」の固定・種から決まる ID・外部への通信を止める差し替えが効くこと
 * - 計算の一致の確認（旧ブックの写し ↔ データ本体から組み立て）が 2 段の処理で最後まで動き、違いを見つけること
 * 旧来の予測そのもの（runPhase1Forecast）はモックでは動かさない（本物の GAS の上の一致の確認で確かめる）。
 *
 *   node app/tests/app-engine.test.mjs
 */

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { appDir, repoRoot, sources, sha, J, makeEnv, setUpEnv, OWNER } from './gas-mock.mjs';
import { buildEngine, ENGINE_EXPORTS, ENGINE_SERVICES } from '../tools/build-engine.mjs';

const legacySrc = await readFile(path.join(repoRoot, 'Forecast_Agent.js'), 'utf8');

// ==== 1. LegacyEngine.js は Forecast_Agent.js をそのまま包んだもの ====
{
  const built = buildEngine(legacySrc);
  assert.equal(sources['LegacyEngine.js'], built.text, 'LegacyEngine.js が古い（node app/tools/build-engine.mjs で作り直す）');
  const text = sources['LegacyEngine.js'];
  const begin = text.indexOf('// ===== Forecast_Agent.js ここから（変更しない） =====\n') + '// ===== Forecast_Agent.js ここから（変更しない） =====\n'.length;
  const end = text.lastIndexOf('\n  // ===== Forecast_Agent.js ここまで =====');
  assert.equal(text.slice(begin, end), legacySrc, '包んだ中身は元のファイルと 1 文字も違わない');
  assert.ok(text.includes('SHA-256 ' + sha(legacySrc)));
  assert.ok(text.startsWith('/**') && text.trimEnd().endsWith('}'));
  // 外に出るのは appLegacyEngine_ だけ（読み込んだ後のグローバルの関数で確かめる）
  const env = makeEnv();
  const fromEngine = Object.keys(env.ctx).filter((k) => typeof env.ctx[k] === 'function' && /^(runPhase1Forecast|onOpen|onEdit|setupForecastBook)$/.test(k));
  assert.deepEqual(fromEngine, [], '旧来の関数はグローバルにない');
  assert.equal(typeof env.ctx.appLegacyEngine_, 'function');
}

// ==== 2. 差し替え: 「今」の固定・種から決まる ID・画面と外部への通信を止める ====
{
  const env = setUpEnv();
  const asOf = Date.UTC(2026, 9, 1, 3, 0, 0);
  const r = J(env.run(`(() => {
    const D = appFrozenDate_(${asOf});
    const real = new Date(2020, 0, 2);
    return {
      now: new D().getTime(), nowFn: D.now(), args: new D(2026, 3, 1).getTime(), utc: D.UTC(2026, 3, 1), parse: D.parse('2026-04-01T00:00:00Z'),
      inst: real instanceof D, inst2: new D() instanceof Date, inst3: new D(5) instanceof D, str: typeof D(), proto: D.prototype === Date.prototype
    };
  })()`));
  assert.deepEqual(r, { now: asOf, nowFn: asOf, args: new Date(2026, 3, 1).getTime(), utc: Date.UTC(2026, 3, 1), parse: Date.parse('2026-04-01T00:00:00Z'),
    inst: true, inst2: true, inst3: true, str: 'string', proto: true });
  const ids = J(env.run(`(() => { const a = appSeededUuidMaker_('S'); const b = appSeededUuidMaker_('S'); const c = appSeededUuidMaker_('T'); return [a(), a(), b(), c()]; })()`));
  assert.equal(ids[0], ids[2], '同じ種なら同じ ID');
  assert.notEqual(ids[0], ids[1]);
  assert.notEqual(ids[0], ids[3]);
  ids.forEach((id) => assert.match(id, /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-a[0-9a-f]{3}-[0-9a-f]{12}$/));
  const book = env.makeBook('計算用', { CONFIG: { values: [['項目', '値']] } });
  env.props.APP_ZAC_SOURCE_SPREADSHEET_ID = 'ZAC-SOURCE-ID';
  const svc = env.run(`appLegacyServices_(__b, { asOfMs: ${asOf}, seed: 'S' })`, { __b: book });
  env.ctx.__svc = svc;
  assert.equal(env.run('__svc.SpreadsheetApp.getActiveSpreadsheet().getId()'), book.getId(), '開いているスプレッドシート = 計算用ブック');
  assert.throws(() => env.run('__svc.SpreadsheetApp.getUi()'), /画面/);
  // 画面のない実行では本物のトーストは止められるが、旧来の計算から呼ばれても何もしない
  assert.throws(() => book.toast('x'), /showNotification/);
  assert.equal(env.run(`__svc.SpreadsheetApp.getActiveSpreadsheet().toast('x', 'y', 5)`), undefined);
  assert.equal(env.run(`__svc.SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(__s)`, { __s: book.getSheetByName('CONFIG') }), book.getSheetByName('CONFIG'));
  assert.equal(env.run(`__svc.SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG')`), book.getSheetByName('CONFIG'), 'ほかはそのまま');
  assert.equal(env.run('__svc.SpreadsheetApp.openById(__id)', { __id: book.getId() }), book, 'ほかの機能は本物を通す');
  assert.equal(env.run('__svc.Utilities.getUuid()'), ids[0]);
  assert.equal(env.run('__svc.Utilities.sleep(450)'), undefined, '待たない');
  assert.equal(env.run(`__svc.Utilities.formatDate(new Date(${asOf}), 'Asia/Tokyo', 'yyyy-MM-dd')`), '2026-10-01');
  assert.equal(env.run(`__svc.PropertiesService.getScriptProperties().getProperty('FORECAST_SOURCE_SPREADSHEET_ID')`), 'ZAC-SOURCE-ID');
  assert.equal(env.run(`__svc.PropertiesService.getScriptProperties().getProperty('APP_DATA_SPREADSHEET_ID')`), null, '新アプリの設定は見せない');
  assert.throws(() => env.run(`__svc.PropertiesService.getScriptProperties().setProperty('X', '1')`), /書きません/);
  assert.throws(() => env.run(`__svc.UrlFetchApp.fetch('https://example.com')`), /UrlFetchApp\.fetch を使いません/);
  assert.throws(() => env.run(`__svc.HtmlService.createHtmlOutput('x')`), /HtmlService/);
  // 本物の旧来の計算を、差し替えのもとで読み込める（トップレベルの文が通り、使う関数がそろう）
  const eng = env.run('appLegacyEngine_(__svc)');
  for (const n of ENGINE_EXPORTS) assert.notEqual(eng[n], undefined, `${n} がない`);
  assert.equal(typeof eng.runPhase1Forecast, 'function');
  assert.equal(eng.SOURCE_SHA256, sha(legacySrc));
  assert.equal(eng.VERSION, /const VERSION = '([^']+)'/.exec(legacySrc)[1]);
  assert.deepEqual(Object.keys(J(eng.SHEETS)).length >= 30, true);
  assert.deepEqual(ENGINE_SERVICES, ['SpreadsheetApp', 'Date', 'Utilities', 'PropertiesService', 'UrlFetchApp', 'HtmlService']);
  // 旧来の年度の数え方（FY N = N 年 4 月〜）を、差し替えた Date のもとでも同じに使う
  assert.equal(J(eng.getForecastFYStart_(2026)), new Date(2026, 3, 1).toISOString());
  assert.equal(env.run('appParseIso_("2026-10-01T22:40:00+0900").toISOString()'), '2026-10-01T13:40:00.000Z');
  assert.equal(env.run('appParseIso_("x")'), null);
}

// ==== 3. 計算の一致の確認（2 段の処理）: 同じなら「完全に一致」、データ本体が違えばそのシートを示す ====
const D = (y, m, d = 1) => new Date(y, m - 1, d);
function legacyBook(env) {
  return env.makeBook('クライアント別売上予測', {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', 'テスト製薬'], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
    SALES_INPUT: {
      values: [['client', 'service_type', 'product', 'target_month', 'input_amount', 'status', 'source_updated_at'],
        ['テスト製薬', 'BASE', '製品A', D(2023, 4), 1200000, 'closed', D(2026, 9, 30)]],
      formats: { D: '@' },
    },
    OUTPUT: { values: [['FY2026 売上予測（テスト製薬）']] },
    CALIBRATION_STATE: {
      values: [['client', 'updated_at', 'updated_by', 'ai_weight_override', 'ai_max_abs_effect_override', 'ai_topic_disable_json', 'bias_correction_factor',
        'qual_scale_override', 'residual_month_bias_json', 'last_applied_quarter', 'last_applied_review_id', 'auto_update_enabled', 'note'],
      ['テスト製薬', D(2026, 9, 1), 'auto', '', '', '[]', 0.98, '', '{}', '', '', 1, '']],
    },
  });
}
/** 旧来の予測の代わり（モックでは本物を動かさない）: 入力と乱数・「今」・ID から決まる値を書く */
const STUB_ENGINE = `appLegacyEngine_ = function (svc) {
  return {
    WEB_UI_CONFIRMS_: {}, VERSION: 'stub', SOURCE_SHA256: 'stub',
    runPhase1Forecast() {
      if (!this.WEB_UI_CONFIRMS_.extreme) throw Object.assign(new Error('confirm'), { webConfirm: { key: 'extreme' } });
      const ss = svc.SpreadsheetApp.getActiveSpreadsheet();
      ss.toast('予測を更新しています', 'Forecast Agent', 5);   // 旧来の toastProgress_ と同じ（画面のない実行でも止まらない）
      ss.setActiveSheet(ss.getSheetByName('OUTPUT'));
      const fy = Number(ss.getSheetByName('CONFIG').getRange('B3').getValue());
      const bias = Number(ss.getSheetByName('CALIBRATION_STATE').getRange(2, 7).getValue());
      const r = () => Math.round(Math.random() * 1000);
      const out = ss.getSheetByName('OUTPUT');
      out.getRange(26, 1, 1, 4).setValues([['年度合計（予測）', fy * bias + r(), fy * 2 + r(), fy * 3 + r()]]);
      for (let i = 0; i < 12; i++) out.getRange(29 + i, 1, 1, 4).setValues([[new svc.Date(fy, 3 + i, 1), r(), r(), r()]]);
      const cal = ss.getSheetByName('CALIBRATION_STATE');
      cal.getRange(cal.getLastRow() + 1, 1, 1, 3).setValues([[svc.Utilities.getUuid(), new svc.Date(), 'stub']]);
    }
  };
};`;
{
  const env = setUpEnv();
  const book = legacyBook(env);
  const url = 'https://docs.google.com/spreadsheets/d/' + book.getId() + '/edit';
  const dry = env.runJob('MIGRATION.DRYRUN', { bookUrl: url });
  assert.equal(dry.status, 'DONE', dry.error);
  const imp = env.runJob('MIGRATION.IMPORT', { bookUrl: url, contentHash: dry.result.contentHash });
  assert.equal(imp.status, 'DONE', imp.error);
  const planId = imp.result.planId;
  const plans = env.call('apiListPlans()').plans;
  assert.deepEqual(plans.map((p) => [p.planId, p.clientName, p.fy]), [[planId, 'テスト製薬', '2026']]);
  env.run(STUB_ENGINE);
  const dataBefore = env.data().getSheets().map((s) => [s.name, s.rows.length]);
  const bookBefore = JSON.stringify(book.getSheets().map((s) => [s.name, s.rows]));
  const st = env.runJob('FORECAST.PARITY', { planId });
  assert.equal(st.status, 'DONE', st.error);
  const r = st.result;
  assert.equal(r.same, true, JSON.stringify(r.diff));
  assert.deepEqual(r.diff, []);
  assert.equal(r.planId, planId);
  assert.ok(r.store.annual.p50 > 0);
  assert.deepEqual(r.store, r.legacy, '同じ種・同じ時刻なら旧ブックの写しとデータ本体で同じ結果');
  assert.equal(r.previous, null, '旧ブックの前回の結果（OUTPUT が短いので無し）');
  assert.equal(r.legacyChangedAfterImport, false);
  assert.match(r.seed, /^PARITY:/);
  assert.equal(r.store.monthly.length, 12);
  assert.equal(r.store.monthly[0].month, '2026/04');
  // データ本体と旧ブックには書かない
  assert.deepEqual(env.data().getSheets().map((s) => [s.name, s.rows.length]), dataBefore);
  assert.equal(JSON.stringify(book.getSheets().map((s) => [s.name, s.rows])), bookBefore);
  // 記録: A と B の両方を残す
  const au = env.audit().filter((a) => /^FORECAST\.PARITY/.test(a.action) && a.phase === 'END');
  assert.deepEqual(au.map((a) => [a.action, a.result]), [['FORECAST.PARITY.A', 'OK'], ['FORECAST.PARITY.B', 'OK']]);
  assert.equal(JSON.parse(au[1].after_json).same, true);
  // データ本体の値が旧ブックと違えば、そのシートを示す
  const cal = env.data().getSheetByName('ENG_CALIBRATION_STATE');
  const col = cal.rows[0].indexOf('bias_correction_factor');
  cal.rows[1][col] = '0.5';
  const st2 = env.runJob('FORECAST.PARITY', { planId });
  assert.equal(st2.status, 'DONE', st2.error);
  assert.equal(st2.result.same, false);
  const sheetsWithDiff = st2.result.diff.map((d) => d.sheet).sort();
  assert.deepEqual(sheetsWithDiff, ['CALIBRATION_STATE', 'OUTPUT']);
  assert.ok(st2.result.diff.find((d) => d.sheet === 'CALIBRATION_STATE').rows.includes(2));
  assert.notEqual(st2.result.store.annual.p10, st2.result.legacy.annual.p10);
  // 取り込みの後に旧ブックが変わったら知らせる
  env.files[book.getId()].updated = new Date(Date.now() + 3600e3);
  const st3 = env.runJob('FORECAST.PARITY', { planId });
  assert.equal(st3.result.legacyChangedAfterImport, true);
  // 続きの処理は画面から始められない・計画が無ければ止まる・所有者だけ
  assert.throws(() => env.call(`apiStartJob({ kind: 'FORECAST.PARITY_B', payload: {} })`), /画面から始められません/);
  const bad = env.runJob('FORECAST.PARITY', { planId: 'PL-none' });
  assert.deepEqual([bad.status, /計画が見つかりません/.test(bad.error)], ['FAILED', true]);
  assert.equal(env.state.lockHeld, false);
  assert.equal(env.run('Math.random === __r', { __r: env.run('Math.random') }), true);
}

// ==== 4. 控えの比べ方 ====
{
  const env = makeEnv();
  const d = (rows, extra) => Object.assign({ rows, formats: 'f', size: '10x5x' + rows.length + 'x5', hash: rows.join('|') }, extra || {});
  const a = { S1: d(['a', 'b', 'c']), S2: d(['x']), S3: d(['y']) };
  const b = { S1: d(['a', 'B', 'c', 'd']), S2: d(['x']), S4: d(['z']) };
  const diff = J(env.run('appDigestDiff_(__a, __b)', { __a: a, __b: b }));
  assert.deepEqual(diff, [
    { sheet: 'S1', size: '10x5x3x5 / 10x5x4x5', formats: false, rows: [2, 4], rowCount: 2 },
    { sheet: 'S3', missing: 'B' },
    { sheet: 'S4', missing: 'A' },
  ]);
}

// ==== 5. 表示形式の違いをセルで示す・計算の前から違うのか計算で違ったのかを分ける ====
{
  const env = makeEnv();
  const a = { 1: [[1, 10, 'General']], 2: [[1, 10, '@']] };
  const b = { 1: [[1, 4, 'General'], [5, 6, '0.0%'], [7, 10, 'General']], 2: [[1, 10, '@']] };
  assert.deepEqual(J(env.run('appFormatCellsDiff_(__a, __b, 8)', { __a: a, __b: b })), { count: 2, samples: [{ cells: 'A5:A6', a: 'General', b: '0.0%' }] });
  assert.deepEqual(J(env.run('appFormatCellsDiff_(__a, __a, 8)', { __a: a })), { count: 0, samples: [] });
  // データ本体の表示形式が旧ブックと違えば、計算の前の違いとして、そのセルと両方の形式を示す
  const env2 = setUpEnv();
  const book = legacyBook(env2);
  const url = 'https://docs.google.com/spreadsheets/d/' + book.getId() + '/edit';
  const dry = env2.runJob('MIGRATION.DRYRUN', { bookUrl: url });
  const imp = env2.runJob('MIGRATION.IMPORT', { bookUrl: url, contentHash: dry.result.contentHash });
  env2.run(STUB_ENGINE);
  const fm = env2.data().getSheetByName('ENG_FORMATS');
  const H = fm.rows[0];
  const row = fm.rows.findIndex((r, i) => i > 0 && r[H.indexOf('sheet')] === 'CALIBRATION_STATE' && r[H.indexOf('col')] === '7');
  const orig = JSON.parse(fm.rows[row][H.indexOf('runs_json')]);
  assert.deepEqual(orig, [[1, 1000, '0.###############']], '何もしていない列は「自動」（本物は 0.############### と返す）');
  fm.rows[row][H.indexOf('runs_json')] = JSON.stringify([[1, 1, '0.###############'], [2, 2, '0.00'], [3, 1000, '0.###############']]);
  const st = env2.runJob('FORECAST.PARITY', { planId: imp.result.planId });
  assert.equal(st.status, 'DONE', st.error);
  assert.equal(st.result.preSame, false);
  const pre = st.result.preDiff.find((d) => d.sheet === 'CALIBRATION_STATE');
  assert.deepEqual(pre.formatCells.samples[0], { cells: 'G2', a: '0.###############', b: '0.00' });
  assert.equal(pre.rowCount, 0, '値は同じで表示形式だけが違う');
}

// ==== 6. 形式を消したセル（空の表示形式）も、組み立て直すと同じになる（2026-10-02 DASHBOARD!B17・PROCESS_STATUS!B8） ====
{
  const env = setUpEnv();
  const book = env.makeBook('クライアント別売上予測', {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', 'テスト製薬'], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
    OUTPUT: { values: [['FY2026 売上予測（テスト製薬）']] },
    CALIBRATION_STATE: {
      values: [['client', 'updated_at', 'updated_by', 'ai_weight_override', 'ai_max_abs_effect_override', 'ai_topic_disable_json', 'bias_correction_factor',
        'qual_scale_override', 'residual_month_bias_json', 'last_applied_quarter', 'last_applied_review_id', 'auto_update_enabled', 'note'],
      ['テスト製薬', D(2026, 9, 1), 'auto', '', '', '[]', 0.98, '', '{}', '', '', 1, '']],
    },
    PROCESS_STATUS: {
      values: [['step_key', 'last_run_date', 'last_run_by', 'status', 'target_client', 'record_count', 'error_summary'],
        ['step1_status', D(2026, 9, 30), 'owner', 'success', 'テスト製薬', 10, ''],
        ['step5_status', D(2026, 9, 30), 'owner', 'success', 'テスト製薬', 3, '']],
      formats: { B: 'yyyy/MM/dd H:mm:ss' },
      cellFormats: { B3: null },   // 形式を付けていないセルに日付を書いた（旧来がシートを消してから書いた）セル。本物は空の形式と返す
    },
  });
  const url = 'https://docs.google.com/spreadsheets/d/' + book.getId() + '/edit';
  const dry = env.runJob('MIGRATION.DRYRUN', { bookUrl: url });
  assert.equal(dry.status, 'DONE', dry.error);
  const ps = dry.result.sheets.find((x) => x.sheet === 'PROCESS_STATUS');
  assert.deepEqual([ps.mismatch, ps.formatMismatches], [0, 0], '形式を消したセルも組み立て直すと同じ（値も表示形式も）');
  assert.equal(dry.result.faithful, true);
  const scratch = env.scratch();
  assert.equal(scratch.getSheetByName('PROCESS_STATUS').getRange(3, 2).getNumberFormats()[0][0], '', '組み立てでも空の形式になる');
  assert.equal(scratch.getSheetByName('PROCESS_STATUS').getRange(2, 2).getNumberFormats()[0][0], 'yyyy/MM/dd H:mm:ss');
  // 直す前のやり方（「自動」を明示してから書く・形式を消してから書く）では空にならないことも確かめる
  const t = scratch.insertSheet('T');
  t.getRange(1, 1).setNumberFormat('General'); t.getRange(1, 1).setValues([[new Date(2026, 8, 30)]]);
  t.getRange(2, 1).setNumberFormat('General'); t.getRange(2, 1).clearFormat(); t.getRange(2, 1).setValues([[new Date(2026, 8, 30)]]);
  t.getRange(3, 1).setValues([[new Date(2026, 8, 30)]]);
  assert.deepEqual(t.getRange(1, 1, 3, 1).getNumberFormats().map((r) => r[0]), ['0.###############', '0.###############', '']);
  scratch.deleteSheet(t);
  const imp = env.runJob('MIGRATION.IMPORT', { bookUrl: url, contentHash: dry.result.contentHash });
  assert.equal(imp.status, 'DONE', imp.error);
  env.run(STUB_ENGINE);
  const st = env.runJob('FORECAST.PARITY', { planId: imp.result.planId });
  assert.equal(st.status, 'DONE', st.error);
  assert.deepEqual([st.result.preSame, st.result.same], [true, true], JSON.stringify(st.result.preDiff));
  assert.deepEqual(st.result.build, [], '組み立てで直せなかったセルはない');
}

console.log('app-engine: all tests passed');
