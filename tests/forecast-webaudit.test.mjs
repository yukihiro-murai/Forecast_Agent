#!/usr/bin/env node
/*
 * forecast-webaudit.test.mjs — Web アプリの操作の記録（監査ログ）と実行者の確認の契約テスト（段階0）。
 * Forecast_WebApp.js の webAudited_ ほかを GAS のモックの上で vm 実行する。
 *
 *   node tests/forecast-webaudit.test.mjs
 */

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const webSrc = await readFile(path.join(root, 'Forecast_WebApp.js'), 'utf8');
const coreSrc = await readFile(path.join(root, 'VNext_Core.js'), 'utf8');
const manifest = JSON.parse(await readFile(path.join(root, 'appsscript.json'), 'utf8'));

function extractFunction(src, name) {
  const start = src.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} not found`);
  let depth = 0;
  for (let p = src.indexOf('{', start); p < src.length; p++) {
    if (src[p] === '{') depth++;
    else if (src[p] === '}' && --depth === 0) return src.slice(start, p + 1);
  }
  throw new Error(`${name} not closed`);
}
const auditSection = webSrc.slice(webSrc.indexOf('// ===== 操作の記録（監査ログ）と実行者の確認'));
assert.ok(auditSection.length > 1000, '操作の記録の区間がある');

// ---- GAS のモック ----
function makeSheet(name) {
  const rows = [];
  let maxRows = 1000;
  const sh = {
    rows, name,
    getName: () => name,
    getLastRow: () => rows.length,
    getMaxRows: () => maxRows,
    insertRowsAfter: (_, n) => { maxRows += n; },
    setFrozenRows: () => {},
    failWrites: false,
    getRange(r, c, nr = 1, nc = 1) {
      if (typeof r === 'string') { // A1 形式（'B2' など）
        const m = /^([A-Z]+)(\d+)$/.exec(r);
        c = [...m[1]].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0); r = Number(m[2]);
      }
      const range = {
        setNumberFormat: () => range,
        setValues(vals) {
          if (sh.failWrites) throw new Error('write failed');
          if (r > maxRows) throw new Error('out of grid');
          vals.forEach((v, i) => { rows[r - 1 + i] = v.slice(); });
          return range;
        },
        getValues: () => Array.from({ length: nr }, (_, i) => (rows[r - 1 + i] || []).slice(c - 1, c - 1 + nc)),
        getValue: () => (rows[r - 1] || [])[c - 1] ?? '',
      };
      return range;
    },
  };
  return sh;
}
function makeSpreadsheet(id, sheetNames = ['シート1']) {
  const sheets = sheetNames.map(makeSheet);
  return {
    id, sheets,
    getId: () => id,
    getSheetByName: (n) => sheets.find((s) => s.name === n) || null,
    getSheets: () => sheets.slice(),
    insertSheet(n) { const s = makeSheet(n); sheets.push(s); return s; },
    deleteSheet(s) { sheets.splice(sheets.indexOf(s), 1); },
  };
}
function makeEnv({ active = 'owner@bigm2y.com', owner = 'owner@bigm2y.com', props = {} } = {}) {
  const store = Object.assign({}, props);
  const files = {};
  let created = 0, uuid = 0;
  const config = makeSheet('CONFIG');
  config.rows[1] = ['クライアント', 'テスト製薬'];
  config.rows[2] = ['FY', '2027'];
  const activeSs = { getSheetByName: (n) => (n === 'CONFIG' ? config : null) };
  const env = {
    store, files, created: () => created, Logger: { log() {} },
    Session: { getActiveUser: () => ({ getEmail: () => active }), getEffectiveUser: () => ({ getEmail: () => owner }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in store ? store[k] : null), setProperty: (k, v) => { store[k] = String(v); } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    SpreadsheetApp: {
      create(name) { created++; const ss = makeSpreadsheet('LOG-' + created); ss.title = name; files[ss.id] = ss; return ss; },
      openById(id) { if (!files[id]) throw new Error('not found ' + id); return files[id]; },
      getActiveSpreadsheet: () => activeSs,
      flush() {},
    },
    DriveApp: { getFileById: () => ({ moveTo() {} }), getFolderById: () => ({}) },
    Utilities: {
      getUuid: () => 'uuid-' + (++uuid),
      formatDate: (d, tz, fmt) => (fmt === 'yyyy_MM' ? '2026_10' : '2026-10-01T10:00:00+0900'),
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (alg, s) => Array.from(createHash('sha256').update(String(s), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b)),
    },
    TZ: 'Asia/Tokyo', VERSION: 'test', BUILD_STAGE: 'stage', SHEETS: { CONFIG: 'CONFIG' },
  };
  const ctx = vm.createContext(env);
  vm.runInContext(extractFunction(coreSrc, 'vNextSha256Hex_') + '\n' + auditSection, ctx);
  return ctx;
}
const sha = (s) => createHash('sha256').update(s, 'utf8').digest('hex');
const logOf = (ctx) => {
  const map = JSON.parse(ctx.store.FORECAST_LOG_SPREADSHEETS_JSON || '{}');
  const id = Object.values(map)[0];
  return id ? ctx.files[id] : null;
};

// ---- 1. 管理者の操作: 開始と終了を記録し、ハッシュの鎖がつながる ----
{
  const ctx = makeEnv();
  let ran = 0;
  const res = vm.runInContext(`webAudited_('TEST.SAVE', () => { __ran(); return { saved: 3 }; },
    { detail: { rows: 3 }, before: () => ({ v: 'old' }), after: (r) => ({ saved: r.saved }) })`, Object.assign(ctx, { __ran: () => { ran++; } }));
  assert.equal(ran, 1, '処理を 1 回実行');
  assert.deepEqual({ ...res }, { saved: 3 });
  const log = logOf(ctx);
  assert.ok(log, 'ログのファイルができる');
  assert.match(log.title, /^売上予測 ログ FY\d{4}$/);
  const fyKey = Object.keys(JSON.parse(ctx.store.FORECAST_LOG_SPREADSHEETS_JSON))[0];
  const now = new Date();
  assert.equal(fyKey, 'FY' + (now.getMonth() >= 3 ? now.getFullYear() + 1 : now.getFullYear()), '年度は 4 月始まり・終わる年で呼ぶ');
  assert.equal(log.sheets.length, 1, '最初からある空のシートは消す');
  const sh = log.getSheetByName('AUDIT_2026_10');
  assert.ok(sh, '月ごとのシート');
  const H = [...sh.rows[0]];
  assert.deepEqual(H, ['audit_id', 'occurred_at', 'actor_email', 'action', 'phase', 'result', 'client', 'fy', 'detail_json', 'before_json', 'after_json', 'error', 'request_id', 'app_version', 'prev_hash', 'row_hash']);
  const col = (row, name) => row[H.indexOf(name)];
  const [start, end] = [sh.rows[1], sh.rows[2]];
  assert.equal(col(start, 'phase'), 'START'); assert.equal(col(end, 'phase'), 'END'); assert.equal(col(end, 'result'), 'OK');
  assert.equal(col(start, 'actor_email'), 'owner@bigm2y.com');
  assert.equal(col(start, 'client'), 'テスト製薬'); assert.equal(col(start, 'fy'), '2027');
  assert.equal(col(start, 'before_json'), '{"v":"old"}'); assert.equal(col(end, 'after_json'), '{"saved":3}');
  assert.equal(col(start, 'request_id'), col(end, 'request_id'), '開始と終了は同じ request_id');
  // 鎖: 1 行目の prev は空、2 行目の prev は 1 行目の hash。hash は「prev + 改行 + 行の内容」の SHA-256
  assert.equal(col(start, 'prev_hash'), '');
  assert.equal(col(end, 'prev_hash'), col(start, 'row_hash'));
  for (const r of [start, end]) assert.equal(col(r, 'row_hash'), sha(col(r, 'prev_hash') + '\n' + JSON.stringify(r.slice(0, 14))));
  assert.equal(ctx.store.FORECAST_AUDIT_LAST_HASH, col(end, 'row_hash'), '最新の hash を控える');
  // 2 回目は同じファイルに追記し、鎖が続く
  vm.runInContext(`webAudited_('TEST.RUN', () => ({ ok: 1 }))`, ctx);
  assert.equal(ctx.created(), 1, 'ログのファイルは年度で 1 つ');
  assert.equal(col(sh.rows[3], 'prev_hash'), col(end, 'row_hash'));
}

// ---- 2. 記録できなければ処理しない（fail-closed） ----
{
  const ctx = makeEnv();
  vm.runInContext(`webAudited_('WARMUP', () => 1)`, ctx);
  logOf(ctx).getSheetByName('AUDIT_2026_10').failWrites = true;
  let ran = 0;
  assert.throws(() => vm.runInContext(`webAudited_('TEST.SAVE', () => { __ran(); })`, Object.assign(ctx, { __ran: () => { ran++; } })), /write failed/);
  assert.equal(ran, 0, '開始を記録できないときは実行しない');
  const ctx2 = makeEnv({ props: { FORECAST_LOG_SPREADSHEETS_JSON: JSON.stringify({ FY2027: 'missing', FY2026: 'missing', FY2028: 'missing' }) } });
  assert.throws(() => vm.runInContext(`webAudited_('TEST.SAVE', () => { __ran(); })`, Object.assign(ctx2, { __ran: () => { ran++; } })), /not found/);
  assert.equal(ran, 0, 'ログのファイルを開けないときも実行しない');
}

// ---- 3. 管理者以外は拒否し、拒否も記録する ----
{
  const ctx = makeEnv({ active: 'someone@bigm2y.com' });
  let ran = 0;
  assert.throws(() => vm.runInContext(`webAudited_('TEST.SAVE', () => { __ran(); })`, Object.assign(ctx, { __ran: () => { ran++; } })), /管理者だけ/);
  assert.equal(ran, 0);
  const row = logOf(ctx).getSheetByName('AUDIT_2026_10').rows[1];
  assert.equal(row[4], 'DENIED'); assert.equal(row[5], 'DENIED'); assert.equal(row[2], 'someone@bigm2y.com');
  const ctxBlank = makeEnv({ active: '' });
  assert.throws(() => vm.runInContext(`webAudited_('TEST.SAVE', () => 1)`, ctxBlank), /管理者だけ/, 'メールが取れなければ拒否');
  const ctxAdmin = makeEnv({ active: 'Planner@BIGM2Y.com', props: { FORECAST_WEB_ADMIN_EMAILS: 'planner@bigm2y.com, other@bigm2y.com' } });
  assert.equal(vm.runInContext(`webAudited_('TEST.SAVE', () => 7)`, ctxAdmin), 7, 'Script Property の管理者も実行できる（大文字小文字は区別しない）');
}

// ---- 4. 失敗・確認待ちも記録する ----
{
  const ctx = makeEnv();
  assert.throws(() => vm.runInContext(`webAudited_('TEST.RUN', () => { throw new Error('boom'); })`, ctx), /boom/);
  const sh = logOf(ctx).getSheetByName('AUDIT_2026_10');
  assert.equal(sh.rows[2][5], 'FAILED'); assert.equal(sh.rows[2][11], 'boom');
  vm.runInContext(`webAudited_('FORECAST.RUN', () => ({ needConfirm: { key: 'short_history' } }))`, ctx);
  assert.equal(sh.rows[4][5], 'NEEDS_CONFIRM');
}

// ---- 5. 長い JSON はハッシュと先頭だけにする ----
{
  const ctx = makeEnv();
  const big = 'x'.repeat(50000);
  const j = JSON.parse(vm.runInContext(`webAuditJson_(__big)`, Object.assign(ctx, { __big: { big } })));
  assert.equal(j.truncated, true); assert.equal(j.sha256, sha(JSON.stringify({ big }))); assert.equal(j.head.length, 2000);
  assert.equal(vm.runInContext(`webAuditJson_('')`, ctx), '');
}

// ---- 6. 列が想定と違うシートには書かない ----
{
  const ctx = makeEnv();
  vm.runInContext(`webAudited_('WARMUP', () => 1)`, ctx);
  logOf(ctx).getSheetByName('AUDIT_2026_10').rows[0][3] = 'changed';
  assert.throws(() => vm.runInContext(`webAudited_('TEST.SAVE', () => 1)`, ctx), /列が想定と違います/);
}

// ---- 7. 書き込み・実行の 17 関数はすべて記録つき。読むだけの関数は包まない ----
{
  const wrapped = ['webSaveSetup', 'webRunImportSales', 'webRunAggregate', 'webRunAiResearch', 'webSaveInputs', 'webRunForecast',
    'webSaveBudget', 'webRunImportActuals', 'webRunEvalReport', 'webRunDashboard', 'webRunInsights', 'webSaveEvalInsights',
    'webRunQuarterly', 'webSaveQuarterlyDecisions', 'webApplyQuarterly', 'webRunMonthlyLearn', 'webRunVertexAssist'];
  for (const n of wrapped) assert.match(extractFunction(webSrc, n), new RegExp(`^function ${n}\\([^)]*\\) \\{\\n  return webAudited_\\('`), `${n} は webAudited_ を通す`);
  for (const n of ['webGetBootstrap', 'webGetClientCandidates', 'webRunLearningBacktest']) assert.doesNotMatch(extractFunction(webSrc, n), /webAudited_/, `${n} は読むだけ`);
  const pub = [...webSrc.matchAll(/^function (web[A-Za-z]+)\(/gm)].map((m) => m[1]);
  assert.deepEqual(pub.filter((n) => !wrapped.includes(n) && !['webGetBootstrap', 'webGetClientCandidates', 'webRunLearningBacktest'].includes(n)), [],
    '新しい公開関数を足したら、記録つきにするか読むだけかをここに足す');
}

// ---- 8. Web アプリの公開範囲は「自分のみ」 ----
{
  assert.equal(manifest.webapp.access, 'MYSELF', '社内全員にすると、名前の末尾が _ でない全関数をデプロイした人の権限で呼べてしまう');
  assert.equal(manifest.webapp.executeAs, 'USER_DEPLOYING');
}

process.stdout.write('PASS forecast-webaudit contract tests\n');
