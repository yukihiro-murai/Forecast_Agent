#!/usr/bin/env node
/*
 * app-contract.test.mjs — 売上予測アプリ（段階1: 土台）の契約テスト。
 * app/src の *.js を GAS のモックの上で vm 実行し、入口・権限・記録・保存・バックアップの約束を確かめる。
 * GAS 上での実行の代わりではない（最終確認は GAS 側で行う）。
 *
 *   node app/tests/app-contract.test.mjs
 */

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile, readdir } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const appDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const srcDir = path.join(appDir, 'src');
const repoRoot = path.resolve(appDir, '..');
const srcNames = (await readdir(srcDir)).sort();
const jsFiles = srcNames.filter((n) => n.endsWith('.js'));
const sources = Object.fromEntries(await Promise.all(jsFiles.map(async (n) => [n, await readFile(path.join(srcDir, n), 'utf8')])));
const uiHtml = await readFile(path.join(srcDir, 'UI.html'), 'utf8');
const manifest = JSON.parse(await readFile(path.join(srcDir, 'appsscript.json'), 'utf8'));

const OWNER = 'owner@bigm2y.com';
const MEMBER = 'member@bigm2y.com';
const OTHER = 'other@bigm2y.com';
const OUTSIDER = 'someone@gmail.com';
const SHEETS_MIME = 'application/vnd.google-apps.spreadsheet';

/** 画面（ブラウザ）から呼べる関数。足すときはここにも足す */
const PUBLIC = ['doGet', 'apiBootstrap', 'apiSetup', 'apiListDirectory', 'apiSaveMember', 'apiGrantRole', 'apiRevokeRole',
  'apiSaveClient', 'apiListSettings', 'apiSaveSetting', 'apiListAudit', 'apiHealth', 'apiEnableBackup', 'apiRunBackup',
  'triggerDailyBackup'];

const sha = (s) => createHash('sha256').update(s, 'utf8').digest('hex');
const J = (v) => (v === undefined ? undefined : JSON.parse(JSON.stringify(v)));
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
/** Utilities.formatDate の代わり（Asia/Tokyo 固定。使う書式だけ） */
function fmtDate(d, tz, fmt) {
  assert.equal(tz, 'Asia/Tokyo');
  const t = new Date(d.getTime() + 9 * 3600e3);
  const p = (n) => String(n).padStart(2, '0');
  const map = { yyyy: String(t.getUTCFullYear()), MM: p(t.getUTCMonth() + 1), dd: p(t.getUTCDate()), HH: p(t.getUTCHours()),
    mm: p(t.getUTCMinutes()), ss: p(t.getUTCSeconds()), Z: '+0900' };
  let out = '';
  for (let i = 0; i < fmt.length;) {
    if (fmt[i] === "'") { const j = fmt.indexOf("'", i + 1); out += fmt.slice(i + 1, j); i = j + 1; continue; }
    const tok = ['yyyy', 'MM', 'dd', 'HH', 'mm', 'ss', 'Z'].find((k) => fmt.startsWith(k, i));
    if (tok) { out += map[tok]; i += tok.length; } else { out += fmt[i++]; }
  }
  return out;
}
const jstDay = (offsetDays) => fmtDate(new Date(Date.now() + offsetDays * 86400e3), 'Asia/Tokyo', 'yyyy-MM-dd');
const TODAY = jstDay(0);
const YESTERDAY = jstDay(-1);
const TOMORROW = jstDay(1);
const FY = (() => { const d = new Date(); return 'FY' + (d.getMonth() >= 3 ? d.getFullYear() + 1 : d.getFullYear()); })();
const MONTH = fmtDate(new Date(), 'Asia/Tokyo', 'yyyy_MM');

// ---- GAS のモック ----
function makeSheet(name) {
  const rows = [];
  let maxRows = 1000;
  const sh = {
    rows, name, failWrites: false,
    getName: () => name,
    getLastRow: () => rows.length,
    getMaxRows: () => maxRows,
    insertRowsAfter: (_, n) => { maxRows += n; },
    setFrozenRows: () => {},
    getRange(r, c, nr = 1, nc = 1) {
      const range = {
        setNumberFormat: (f) => { assert.equal(f, '@', 'セルは書式なしテキストで書く'); return range; },
        setValues(vals) {
          if (sh.failWrites) throw new Error('write failed');
          if (r + vals.length - 1 > maxRows) throw new Error('out of grid');
          vals.forEach((v, i) => {
            assert.equal(v.length, nc);
            v.forEach((x) => assert.equal(typeof x, 'string', 'すべて文字列で書く'));
            const row = rows[r - 1 + i] || [];
            v.forEach((x, j) => { row[c - 1 + j] = x; });
            rows[r - 1 + i] = row;
          });
          return range;
        },
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (rows[r - 1 + i] || [])[c - 1 + j] ?? '')),
      };
      return range;
    },
  };
  return sh;
}
function makeEnv({ owner = OWNER, active = owner, order = 'name' } = {}) {
  const state = { active, owner, locks: 0, lockHeld: false, uuid: 0, clock: Date.now() - 1e9, seq: 0 };
  const props = {};
  const files = {};
  const sheetsById = {};
  const triggers = [];
  const logs = [];
  const newId = (p) => p + '-' + (++state.seq);
  const iter = (arr) => { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; };
  function makeSpreadsheet(id, name) {
    const sheets = [makeSheet('シート1')];
    const ss = {
      id, name, sheets,
      getId: () => id, getName: () => name, getUrl: () => 'https://docs.google.com/spreadsheets/d/' + id,
      getSheetByName: (n) => sheets.find((s) => s.name === n) || null,
      getSheets: () => sheets.slice(),
      insertSheet(n) { assert.ok(!sheets.some((s) => s.name === n), 'シート名の重複'); const s = makeSheet(n); sheets.push(s); return s; },
      deleteSheet(s) { if (sheets.length === 1) throw new Error('最後のシートは消せない'); sheets.splice(sheets.indexOf(s), 1); },
    };
    sheetsById[id] = ss;
    return ss;
  }
  function makeFolder(name, parent) {
    const f = { kind: 'folder', id: newId('FOLDER'), name, parent, trashed: false };
    Object.assign(f, {
      getId: () => f.id, getName: () => f.name, getUrl: () => 'https://drive.google.com/drive/folders/' + f.id,
      createFolder: (n) => makeFolder(n, f.id),
      // 本物と同じく、ゴミ箱のファイルも一覧に出す
      getFilesByType: (mime) => iter(Object.values(files).filter((x) => x.kind === 'file' && x.parent === f.id && x.mime === mime)),
    });
    files[f.id] = f;
    return f;
  }
  function makeFile(id, name, parent) {
    const f = { kind: 'file', id, name, mime: SHEETS_MIME, parent, trashed: false, created: new Date(state.clock += 60000) };
    Object.assign(f, {
      getId: () => f.id, getName: () => f.name, getDateCreated: () => f.created, isTrashed: () => f.trashed,
      setTrashed: (b) => { f.trashed = !!b; return f; },
      moveTo: (folder) => { f.parent = folder.getId(); return f; },
      makeCopy: (n, folder) => {
        const src = sheetsById[id];
        const copy = makeSpreadsheet(newId('SS'), n);
        copy.sheets.splice(0, copy.sheets.length, ...src.sheets.map((s) => { const c = makeSheet(s.name); s.rows.forEach((r) => c.rows.push(r.slice())); return c; }));
        return makeFile(copy.id, n, folder.getId());
      },
    });
    files[id] = f;
    return f;
  }
  const env = {
    Logger: { log: (m) => logs.push(String(m)) },
    Session: { getActiveUser: () => ({ getEmail: () => state.active }), getEffectiveUser: () => ({ getEmail: () => state.owner }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); } }) },
    LockService: {
      getScriptLock: () => ({
        waitLock() { if (state.lockHeld) throw new Error('ロックを二重に取ろうとした'); state.lockHeld = true; state.locks++; },
        releaseLock() { state.lockHeld = false; },
      }),
    },
    SpreadsheetApp: {
      create(name) { const ss = makeSpreadsheet(newId('SS'), name); makeFile(ss.id, name, 'ROOT'); return ss; },
      openById(id) { if (!sheetsById[id]) throw new Error('not found ' + id); return sheetsById[id]; },
      flush() {},
    },
    DriveApp: {
      createFolder: (name) => makeFolder(name, 'ROOT'),
      getFolderById: (id) => { if (!files[id] || files[id].kind !== 'folder') throw new Error('no folder ' + id); return files[id]; },
      getFileById: (id) => { if (!files[id] || files[id].kind !== 'file') throw new Error('no file ' + id); return files[id]; },
    },
    MimeType: { GOOGLE_SHEETS: SHEETS_MIME },
    ScriptApp: {
      getProjectTriggers: () => triggers.map((t) => ({ getUniqueId: () => t.uid, getHandlerFunction: () => t.handler })),
      newTrigger(handler) {
        const t = { handler, uid: String(9000000000 + triggers.length) };
        const b = { timeBased: () => b, everyDays: (n) => { t.everyDays = n; return b; }, atHour: (h) => { t.atHour = h; return b; },
          create: () => { triggers.push(t); return t; } };
        return b;
      },
    },
    HtmlService: {
      createTemplateFromFile(name) {
        assert.equal(name, 'UI');
        const t = {
          evaluate() {
            const html = uiHtml.replace(/<\?!=\s*(\w+)\s*\?>/g, (_, k) => { if (!(k in t)) throw new Error('template: ' + k); return String(t[k]); });
            const out = { html, title: '', favicon: '', metas: [] };
            Object.assign(out, {
              setTitle: (x) => { out.title = x; return out; },
              addMetaTag: (n, c) => { out.metas.push([n, c]); return out; },
              // 本物は画像の拡張子で終わらない URL を受け付けない
              setFaviconUrl: (u) => { if (!/\.(png|ico|gif|jpe?g)$/i.test(u)) throw new Error('Invalid argument: url'); out.favicon = u; return out; },
              getContent: () => html,
            });
            return out;
          },
        };
        return t;
      },
    },
    Utilities: {
      getUuid: () => (++state.uuid).toString(16).padStart(8, '0') + '-0000-4000-8000-000000000000',
      formatDate: fmtDate,
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (alg, s, cs) => {
        assert.equal(alg, 'sha256'); assert.equal(cs, 'utf8');
        return Array.from(createHash('sha256').update(String(s), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b));
      },
    },
  };
  const ctx = vm.createContext(env);
  const names = order === 'reverse' ? [...jsFiles].reverse() : jsFiles;
  for (const n of names) vm.runInContext(sources[n], ctx, { filename: n });
  const run = (code, extra) => { Object.assign(ctx, extra || {}); return vm.runInContext(code, ctx); };
  const call = (code, extra) => J(run(code, extra));
  const as = (email) => { state.active = email; };
  const data = () => sheetsById[props.APP_DATA_SPREADSHEET_ID];
  const log = () => sheetsById[JSON.parse(props.APP_LOG_SPREADSHEETS_JSON || '{}')[FY]];
  const objects = (sh) => (sh ? sh.rows.slice(1).map((r) => Object.fromEntries(sh.rows[0].map((h, j) => [h, r[j] ?? '']))) : []);
  return {
    ctx, state, props, files, sheetsById, triggers, logs, run, call, as, data, log,
    table: (name) => objects(data().getSheetByName(name)),
    auditSheet: () => log().getSheetByName('AUDIT_' + MONTH),
    audit: () => objects(log() && log().getSheetByName('AUDIT_' + MONTH)),
    runLog: () => objects(log() && log().getSheetByName('RUN_' + MONTH)),
    errors: () => objects(log() && log().getSheetByName('ERROR_' + MONTH)),
    backups: () => Object.values(files).filter((f) => f.kind === 'file' && f.parent === props.APP_BACKUP_FOLDER_ID),
  };
}
function setUpEnv(opts) {
  const env = makeEnv(opts);
  env.as(OWNER);
  env.call('apiSetup()');
  return env;
}
const tables = makeEnv().run('APP_TABLES');
const auditCols = makeEnv().run('APP_LOG_TABLES.AUDIT');

// ==== 1. 入口の一覧（ブラウザから呼べる関数を増やさない） ====
{
  for (const n of srcNames) assert.match(n, /^(appsscript\.json|[A-Za-z]+\.(js|html))$/, `src に置けるのは GAS のファイルだけ: ${n}`);
  assert.deepEqual(srcNames.filter((n) => n.endsWith('.html')), ['UI.html']);
  const declared = [];
  for (const [file, src] of Object.entries(sources)) {
    for (const m of src.matchAll(/^(?:async\s+)?function\s+([A-Za-z0-9_$]+)\s*\(/gm)) declared.push({ file, name: m[1] });
    assert.doesNotMatch(src, /^(?:var|let|const)\s+[A-Za-z0-9_$]+\s*=\s*(?:async\b|function\b|\([^)]*\)\s*=>|[A-Za-z0-9_$]+\s*=>)/m,
      `${file}: トップレベルで関数を変数に入れない（入口が増える）`);
  }
  const pub = declared.filter((d) => !d.name.endsWith('_'));
  assert.deepEqual(pub.map((d) => d.name).sort(), [...PUBLIC].sort(), '公開する関数は一覧のものだけ（ほかは名前の末尾を _ にする）');
  // 実際に読み込んだ後のグローバルでも確かめる（関数の中に包んだコードは外から呼べない）
  const probe = makeEnv();
  const globalFns = Object.keys(probe.ctx).filter((k) => typeof probe.ctx[k] === 'function');
  assert.deepEqual(globalFns.filter((k) => !k.endsWith('_')).sort(), [...PUBLIC].sort(), '読み込んだ後に外から呼べる関数も一覧のものだけ');
  assert.ok(pub.every((d) => d.file === 'Api.js'), '公開する関数は Api.js にだけ置く');
  const names = declared.map((d) => d.name);
  assert.deepEqual(names.filter((n, i) => names.indexOf(n) !== i), [], '同じ名前の関数がない（GAS では黙って上書きされる）');
  for (const name of PUBLIC.filter((n) => n !== 'doGet')) {
    const body = extractFunction(sources['Api.js'], name);
    assert.ok(body.slice(body.indexOf('{') + 1).trim().startsWith('return api_('), `${name} は api_ だけを通す`);
  }
  assert.match(extractFunction(sources['Api.js'], 'doGet'), /api_\('APP\.OPEN', \{ minRole: 'VIEWER', audit: false, allowAnonymousView: true \}/);
  // 業務のコードは SpreadsheetApp を Store / Audit / Setup の外で触らない
  for (const [file, src] of Object.entries(sources)) {
    if (!['Store.js', 'Audit.js', 'Setup.js', 'Backup.js', 'Engine.js'].includes(file)) assert.doesNotMatch(src, /SpreadsheetApp\./, `${file} は保存の層を通す`);
  }
}

// ==== 2. マニフェスト（段階1 は所有者だけが開ける） ====
{
  assert.equal(manifest.runtimeVersion, 'V8');
  assert.equal(manifest.timeZone, 'Asia/Tokyo');
  assert.deepEqual(manifest.webapp, { executeAs: 'USER_DEPLOYING', access: 'MYSELF' },
    '社内に開くのは段階2で所有者が決めてから（ここを DOMAIN に変えるときは設計文書 12 章を更新する）');
  assert.deepEqual([...manifest.oauthScopes].sort(), [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/userinfo.email',
  ]);
  const clasp = JSON.parse(await readFile(path.join(appDir, '.clasp.json'), 'utf8'));
  assert.equal(clasp.rootDir, 'src');
  const rootClasp = JSON.parse(await readFile(path.join(repoRoot, '.clasp.json'), 'utf8'));
  assert.notEqual(clasp.scriptId, rootClasp.scriptId, '旧アプリとは別のプロジェクト');
  assert.match(await readFile(path.join(repoRoot, '.claspignore'), 'utf8'), /^app\/\*\*$/m, '旧アプリの clasp push に app/ を含めない');
}

// ==== 3. 画面（テンプレートと呼び出し） ====
{
  const scriptlets = [...uiHtml.matchAll(/<\?([\s\S]*?)\?>/g)].map((m) => m[1].trim()).sort();
  assert.deepEqual(scriptlets, ['!= bootJson', '!= charsJs'], 'テンプレートで差し込むのは初期データとキャラの絵だけ');
  const called = new Set([...uiHtml.matchAll(/call\(\\?'([A-Za-z0-9_]+)\\?'/g)].map((m) => m[1]));
  assert.ok(called.size >= 10);
  for (const n of called) assert.ok(PUBLIC.includes(n) && n.startsWith('api'), `画面から呼ぶのは公開の入口だけ: ${n}`);
  // 画面が「読むだけ」として扱う入口は、サーバーでも記録しない読み取り
  const reads = Object.keys(JSON.parse(/var READS = (\{[^}]*\});/.exec(uiHtml)[1].replace(/(\w+):/g, '"$1":')));
  for (const n of reads) assert.match(extractFunction(sources['Api.js'], n), /audit: false/, `${n} は読むだけ`);
  for (const n of called) {
    if (!reads.includes(n) && n !== 'apiSetup') assert.doesNotMatch(extractFunction(sources['Api.js'], n), /audit: false/, `${n} は書き込みなので記録する`);
  }
  // キャラの絵は assets/characters と同じ（cd assets/characters/src && python3 build.py で作り直す）
  const env = makeEnv();
  const charsJs = env.run('APP_UI_CHARS_JS');
  assert.ok(!charsJs.includes('<?') && !/<\/script/i.test(charsJs));
  const box = vm.createContext({});
  vm.runInContext(charsJs + '\n;globalThis.__c = { CHAR_SVG, YOMI_POSE };', box);
  const { CHAR_SVG, YOMI_POSE } = J(box.__c);
  for (const id of Object.keys(CHAR_SVG)) {
    assert.equal(CHAR_SVG[id], (await readFile(path.join(repoRoot, 'assets/characters/svg', id + '.svg'), 'utf8')).trim(), `${id} が assets と違う`);
  }
  assert.equal(Object.keys(CHAR_SVG).length, 9);
  assert.deepEqual(Object.keys(YOMI_POSE).sort(), ['discover', 'done', 'explain', 'guide', 'observe']);
  for (const k of Object.keys(YOMI_POSE)) {
    assert.equal(YOMI_POSE[k], (await readFile(path.join(repoRoot, 'assets/characters/yomi/svg', 'yomi_' + k + '.svg'), 'utf8')).trim(), `yomi_${k} が assets と違う`);
  }
  for (const k of ['guide', 'observe', 'explain', 'done']) assert.ok(uiHtml.includes('YOMI_POSE.' + k) || uiHtml.includes("'" + k + "'"), `画面で ${k} を使う`);
  assert.match(env.run('APP_FAVICON_URL'), /^data:image\/png;base64,[A-Za-z0-9+/=]+#favicon\.png$/);
}

// ==== 4. 初期設定（所有者だけ・何度実行しても同じ） ====
{
  const env = makeEnv();
  let boot = env.call('apiBootstrap()');
  assert.deepEqual([boot.setUp, boot.allowed, boot.user.isOwner, boot.user.isAdmin], [false, true, true, true]);
  assert.throws(() => env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'M' })`), /初期設定がまだです/);
  env.as(MEMBER);
  assert.throws(() => env.call('apiSetup()'), /権限がありません/);
  assert.equal(Object.keys(env.files).length, 0, '所有者以外の初期設定では何も作らない');
  env.as(OWNER);
  const r1 = env.call('apiSetup()');
  const folders = Object.values(env.files).filter((f) => f.kind === 'folder');
  const sheetsFiles = Object.values(env.files).filter((f) => f.kind === 'file');
  assert.deepEqual(folders.map((f) => f.name).sort(), ['アーカイブ', 'バックアップ', '売上予測アプリ（システム）']);
  const top = folders.find((f) => f.name === '売上予測アプリ（システム）');
  assert.equal(top.parent, 'ROOT');
  assert.ok(folders.filter((f) => f !== top).every((f) => f.parent === top.id), 'バックアップとアーカイブはシステムのフォルダの中');
  assert.deepEqual(sheetsFiles.map((f) => f.name).sort(), ['売上予測アプリ データ', '売上予測アプリ ログ ' + FY].sort(), 'スプレッドシートはデータ本体と今年度のログの 2 つだけ');
  assert.ok(sheetsFiles.every((f) => f.parent === top.id));
  assert.equal(env.props.APP_INTERNAL_DOMAIN, 'bigm2y.com');
  assert.deepEqual(env.data().getSheets().map((s) => s.name).sort(), Object.keys(tables).sort(), 'データ本体は表のシートだけ（最初の空のシートは消す）');
  for (const [name, def] of Object.entries(tables)) assert.deepEqual([...env.data().getSheetByName(name).rows[0]], [...def.columns], `${name} の 1 行目は列名だけ`);
  const schema = env.table('_SCHEMA');
  assert.equal(schema.length, Object.keys(tables).length);
  for (const s of schema) {
    assert.equal(s.schema_version, '1');
    assert.equal(s.columns_hash, sha(tables[s.table].columns.join('|')));
  }
  assert.deepEqual(env.table('MEMBERS').map((m) => [m.email, m.is_active, m.row_version]), [[OWNER, 'TRUE', '1']]);
  assert.deepEqual(env.log().getSheets().map((s) => s.name), ['AUDIT_' + MONTH], 'ログは月ごとのシートだけ');
  assert.deepEqual(env.audit().map((a) => [a.action, a.phase, a.result, a.actor_email]), [['SETUP.INIT', 'END', 'OK', OWNER]]);
  assert.ok(r1.created.length >= 6 && r1.dataNew === true);
  assert.match(r1.dataUrl, /^https:\/\/docs\.google\.com\/spreadsheets\/d\//);
  // 2 回目: 何も増えない
  const fileCount = Object.keys(env.files).length;
  const r2 = env.call('apiSetup()');
  assert.deepEqual([r2.created, r2.tables, r2.dataNew], [[], [], false]);
  assert.equal(Object.keys(env.files).length, fileCount);
  assert.equal(env.table('_SCHEMA').length, Object.keys(tables).length);
  assert.equal(env.table('MEMBERS').length, 1);
  assert.equal(env.audit().length, 2, '実行のたびに記録は残す');
  // 表が 1 つ消えていたら作り直す（ほかの表はそのまま）
  env.data().deleteSheet(env.data().getSheetByName('CLIENTS'));
  assert.deepEqual(env.call('apiSetup()').tables, ['CLIENTS']);
  boot = env.call('apiBootstrap()');
  assert.equal(boot.setUp, true);
  assert.equal(env.state.lockHeld, false);
}

// ==== 5. 本人確認と役割 ====
{
  const env = setUpEnv();
  env.as(MEMBER);
  const boot = env.call('apiBootstrap()');
  assert.equal(boot.allowed, true);
  assert.equal(boot.user.isAdmin, false);
  assert.deepEqual(boot.user.roles.map((r) => r.role), ['VIEWER', 'CONTRIBUTOR'], '社内の人は閲覧と情報提供を自動で持つ');
  const n0 = env.audit().length;
  assert.throws(() => env.call('apiListDirectory()'), /権限がありません/);
  assert.throws(() => env.call(`apiSaveMember({ email: '${OTHER}', displayName: 'X' })`), /権限がありません/);
  assert.deepEqual(env.audit().slice(n0).map((a) => [a.action, a.phase, a.result, a.actor_email, a.actor_roles]), [
    ['DIRECTORY.LIST', 'DENIED', 'DENIED', MEMBER, 'VIEWER,CONTRIBUTOR'],
    ['MEMBER.SAVE', 'DENIED', 'DENIED', MEMBER, 'VIEWER,CONTRIBUTOR'],
  ], '拒否も記録する');
  assert.equal(env.table('MEMBERS').length, 1, '拒否した操作は実行しない');
  for (const who of [OUTSIDER, '']) {
    env.as(who);
    const b = env.call('apiBootstrap()');
    assert.deepEqual([b.allowed, b.user.isAdmin, b.user.roles], [false, false, []], `社外・メールが取れない人には何も許さない: "${who}"`);
    assert.throws(() => env.call('apiHealth()'), /権限がありません/);
  }
  assert.equal(env.audit().slice(-1)[0].actor_email, '(unknown)');
  env.as('Owner@BIGM2Y.com');
  assert.equal(env.call('apiBootstrap()').user.isAdmin, true, 'メールの大文字・小文字は区別しない');
  // 社外のメールに付与の行があっても効かない
  env.as(OWNER);
  env.run(`appInsertRows_('ROLES', [{ role_id: 'RL-X', email: '${OUTSIDER}', role: 'ADMIN', scope_type: 'ALL', client_id: '', valid_from: '', valid_to: '', is_active: true, row_version: 1 }])`);
  env.as(OUTSIDER);
  assert.deepEqual(env.call('appRolesOf_(appCurrentUser_())'), []);
  // 役割の強さとクライアント単位
  const has = (roles, min, client) => env.run(`appHasRole_(__r, '${min}'${client ? `, '${client}'` : ''})`, { __r: roles });
  const planner = [{ role: 'PLANNER', scope_type: 'CLIENT', client_id: 'CL-1' }];
  assert.equal(has(planner, 'PLANNER'), false, 'クライアント単位の役割は全体の操作に効かない');
  assert.equal(has(planner, 'PLANNER', 'CL-1'), true);
  assert.equal(has(planner, 'PLANNER', 'CL-2'), false);
  assert.equal(has(planner, 'CONTRIBUTOR', 'CL-1'), true, '強い役割は弱い役割を含む');
  assert.equal(has(planner, 'APPROVER', 'CL-1'), false);
  assert.equal(has([{ role: 'ADMIN', scope_type: 'ALL', client_id: '' }], 'APPROVER'), true);
  assert.throws(() => has(planner, 'BOSS'), /未定義の役割/);
}

// ==== 6. メンバー・役割の付与と取り消し ====
{
  const env = setUpEnv();
  assert.throws(() => env.call(`apiSaveMember({ email: '${OUTSIDER}', displayName: 'X' })`), /社内（@bigm2y\.com）のメールアドレスだけ/);
  assert.throws(() => env.call(`apiSaveMember({ email: 'bad', displayName: 'X' })`), /形式が正しくありません/);
  assert.throws(() => env.call(`apiSaveMember({ email: 'x@evil-bigm2y.com', displayName: 'X' })`), /社内/);
  assert.throws(() => env.call(`apiSaveMember({ email: '${MEMBER}', displayName: '  ' })`), /名前を入力/);
  assert.throws(() => env.call(`apiGrantRole({ email: '${MEMBER}', role: 'PLANNER' })`), /先にメンバーとして登録/);
  const m = env.call(`apiSaveMember({ email: ' Member@BIGM2Y.com ', displayName: 'メンバー', department: '営業' })`);
  assert.deepEqual([m.member.email, m.created], [MEMBER, true]);
  const cl = env.call(`apiSaveClient({ clientName: '株式会社テスト製薬', zacCode: 'Z001' })`).client;
  assert.match(cl.client_id, /^CL-\d{14}-[0-9A-F]{8}$/);
  assert.throws(() => env.call(`apiGrantRole({ email: '${MEMBER}', role: 'VIEWER' })`), /付与できる役割は 予算策定担当・承認者・管理者/);
  assert.throws(() => env.call(`apiGrantRole({ email: '${MEMBER}', role: 'APPROVER', scopeType: 'CLIENT', clientId: '${cl.client_id}' })`), /予算策定担当だけ/);
  assert.throws(() => env.call(`apiGrantRole({ email: '${MEMBER}', role: 'PLANNER', scopeType: 'CLIENT', clientId: 'CL-none' })`), /クライアントを選んで/);
  assert.throws(() => env.call(`apiGrantRole({ email: '${MEMBER}', role: 'PLANNER', validFrom: '2026/10/01' })`), /yyyy-MM-dd/);
  assert.throws(() => env.call(`apiGrantRole({ email: '${MEMBER}', role: 'PLANNER', validFrom: '${TODAY}', validTo: '${YESTERDAY}' })`), /終わりの日が始まりの日より前/);
  const g = env.call(`apiGrantRole({ email: '${MEMBER}', role: 'PLANNER', scopeType: 'CLIENT', clientId: '${cl.client_id}' })`).role;
  assert.deepEqual([g.scope_type, g.client_id, g.valid_from, g.is_active], ['CLIENT', cl.client_id, TODAY, true]);
  assert.throws(() => env.call(`apiGrantRole({ email: '${MEMBER}', role: 'PLANNER', scopeType: 'CLIENT', clientId: '${cl.client_id}' })`), /すでに付いています/);
  const grantEnd = env.audit().filter((a) => a.action === 'ROLE.GRANT' && a.result === 'OK')[0];
  assert.deepEqual([grantEnd.entity_type, grantEnd.entity_id, grantEnd.client_id], ['ROLE', g.role_id, cl.client_id], '作った ID を終了の行に残す');
  const rolesOf = () => env.call('appRolesOf_(appCurrentUser_())');
  env.as(MEMBER);
  assert.deepEqual(rolesOf().map((r) => [r.role, r.scope_type, r.client_id]).slice(-1), [['PLANNER', 'CLIENT', cl.client_id]]);
  // 外す: 行は残して無効にする
  env.as(OWNER);
  const rv = env.call(`apiRevokeRole({ roleId: '${g.role_id}', rowVersion: 1 })`).role;
  assert.deepEqual([rv.is_active, rv.valid_to, rv.row_version, rv.updated_by], [false, TODAY, 2, OWNER]);
  const end = env.audit().filter((a) => a.action === 'ROLE.REVOKE' && a.phase === 'END')[0];
  assert.equal(JSON.parse(end.before_json).is_active, true);
  assert.equal(JSON.parse(end.after_json).is_active, false);
  assert.deepEqual([end.entity_id, end.client_id], [g.role_id, cl.client_id]);
  assert.throws(() => env.call(`apiRevokeRole({ roleId: '${g.role_id}' })`), /すでに外れています/);
  assert.throws(() => env.call(`apiRevokeRole({ roleId: 'RL-none' })`), /見つかりません/);
  assert.equal(env.table('ROLES').length, 1, '外した役割の行も消さない');
  env.as(MEMBER);
  assert.equal(rolesOf().length, 2, '外した役割は効かない');
  // 期間の外は効かない（始まる前・終わった後）
  env.as(OWNER);
  env.call(`apiGrantRole({ email: '${MEMBER}', role: 'APPROVER', validFrom: '${TOMORROW}' })`);
  env.call(`apiGrantRole({ email: '${MEMBER}', role: 'ADMIN', validFrom: '2026-01-01', validTo: '${YESTERDAY}' })`);
  env.as(MEMBER);
  assert.deepEqual(rolesOf().map((r) => r.role), ['VIEWER', 'CONTRIBUTOR']);
  env.as(OWNER);
  assert.throws(() => env.call(`apiGrantRole({ email: '${MEMBER}', role: 'APPROVER' })`), /同じ期間にすでに付いています/, '期間が重なる付与は二重にしない');
  // 所有者の管理者は外せない
  env.as(OWNER);
  const ga = env.call(`apiGrantRole({ email: '${OWNER}', role: 'ADMIN' })`).role;
  assert.throws(() => env.call(`apiRevokeRole({ roleId: '${ga.role_id}' })`), /所有者は管理者のまま/);
  // 期限が切れた付与の後には付け直せる。管理者を付けた人は管理の入口を使える
  env.call(`apiGrantRole({ email: '${MEMBER}', role: 'ADMIN' })`);
  env.as(MEMBER);
  const dir = env.call('apiListDirectory()');
  assert.deepEqual([dir.members.length, dir.owner], [2, OWNER]);
  assert.ok(dir.members.every((x) => !('_row' in x)), '行番号は画面に出さない');
  // 役割の表が壊れていたら付与なし（最小の権限）にする
  env.data().getSheetByName('ROLES').rows[0][2] = 'kind';
  assert.throws(() => env.call('apiListDirectory()'), /権限がありません/);
  env.as(OWNER);
  assert.equal(env.call('apiBootstrap()').user.isAdmin, true, '所有者は表が壊れていても管理者（直すため）');
}

// ==== 7. 楽観ロック（ほかの人の更新を上書きしない） ====
{
  const env = setUpEnv();
  env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'A' })`);
  const u1 = env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'B', rowVersion: 1 })`);
  assert.deepEqual([u1.created, u1.member.row_version, u1.before.display_name], [false, 2, 'A']);
  assert.throws(() => env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'C', rowVersion: 1 })`), /ほかの人が先に更新しました/);
  assert.equal(env.table('MEMBERS').find((x) => x.email === MEMBER).display_name, 'B');
  assert.deepEqual(env.audit().slice(-2).map((a) => [a.action, a.phase, a.result]), [['MEMBER.SAVE', 'START', ''], ['MEMBER.SAVE', 'END', 'FAILED']]);
  assert.match(env.audit().slice(-1)[0].error, /ほかの人が先に更新/);
  assert.ok(env.errors().some((e) => e.where === 'MEMBER.SAVE' && /ほかの人が先に更新/.test(e.message)), 'エラーのログにも残す');
  const ok = env.audit().filter((a) => a.action === 'MEMBER.SAVE' && a.result === 'OK');
  assert.equal(JSON.parse(ok[1].before_json).display_name, 'A');
  assert.equal(JSON.parse(ok[1].after_json).display_name, 'B');
  assert.equal(ok[0].before_json, '', '新規は変更前なし');
  const startRow = env.audit().filter((a) => a.action === 'MEMBER.SAVE' && a.phase === 'START')[0];
  assert.equal(JSON.parse(startRow.detail_json).displayName, 'A', '入力を開始の行に残す');
  assert.equal(env.state.lockHeld, false, '失敗してもロックは返す');
}

// ==== 8. クライアント（表記ゆれを吸収して重複を防ぐ） ====
{
  const env = setUpEnv();
  const norm = (s) => env.run('appNormalizeName_(__s)', { __s: s });
  for (const s of ['テスト製薬', '（株）テスト製薬', '㈱テスト製薬', 'テスト製薬株式会社', 'ﾃｽﾄ製薬', ' テスト 製薬 ']) {
    assert.equal(norm(s), norm('株式会社テスト製薬'), s);
  }
  assert.equal(norm('ABC Pharma'), norm('ａｂｃ　ｐｈａｒｍａ'));
  assert.equal(norm('株主製薬'), '株主製薬', '会社の種類は語として除く（1 文字ずつは消さない）');
  assert.equal(norm('合同製薬'), '合同製薬');
  const c = env.call(`apiSaveClient({ clientName: '株式会社テスト製薬' })`).client;
  assert.throws(() => env.call(`apiSaveClient({ clientName: 'テスト製薬（株）' })`), /同じ名前のクライアント/);
  const u = env.call(`apiSaveClient({ clientId: '${c.client_id}', clientName: 'テスト製薬', zacCode: 'Z9', rowVersion: 1 })`);
  assert.deepEqual([u.client.client_name, u.client.zac_code, u.client.row_version, u.created], ['テスト製薬', 'Z9', 2, false]);
  assert.throws(() => env.call(`apiSaveClient({ clientName: ' ' })`), /クライアント名を入力/);
  assert.equal(env.table('CLIENTS').length, 1);
  assert.equal(env.table('CLIENTS')[0].aliases_json, '[]');
}

// ==== 9. 設定（既定値は旧来の値・範囲の確認・上書きしない履歴） ====
{
  const env = setUpEnv();
  const legacy = await readFile(path.join(repoRoot, 'Forecast_Agent.js'), 'utf8');
  const legacyUi = await readFile(path.join(repoRoot, 'Forecast_WebAppUI.html'), 'utf8');
  const num = (src, name) => Number(new RegExp(name + '\\s*=\\s*([0-9.]+)').exec(src)[1]);
  const cur = () => Object.fromEntries(env.call('apiListSettings()').settings.map((s) => [s.key, s]));
  let s = cur();
  assert.equal(s['eval.annual_abs_error_max'].value, num(legacy, 'ANNUAL_ABS_ERROR_CONSTRAINT'));
  assert.equal(s['eval.half_wape_max'].value, num(legacy, 'HALF_WAPE_CONSTRAINT'));
  assert.equal(s['eval.overforecast_rate_max'].value, num(legacy, 'OVERFORECAST_RATE_CONSTRAINT'));
  assert.equal(s['weather.min_months'].value, num(legacyUi, 'WX_MIN_MONTHS'));
  assert.equal(s['weather.tenpen_ape'].value, num(legacyUi, 'WX_TENPEN_APE'));
  assert.equal(s['weather.taifuu_ape'].value, num(legacyUi, 'WX_TAIFUU_APE'));
  assert.equal(s['audit.retention_years'].value, 7);
  assert.ok(Object.values(s).every((x) => x.isDefault));
  for (const [key, value, re] of [
    ['eval.half_wape_max', '1.5', /0〜1 の範囲/], ['eval.half_wape_max', 'abc', /数値で入力/], ['eval.half_wape_max', '', /数値で入力/],
    ['weather.min_months', '2.5', /整数で入力/], ['audit.log_views', 'maybe', /する \/ しない/], ['no.such', '1', /未定義の設定/],
  ]) {
    assert.throws(() => env.call('apiSaveSetting(__in)', { __in: { key, value } }), re, `${key}=${value}`);
  }
  assert.throws(() => env.call(`apiSaveSetting({ key: 'eval.half_wape_max', value: '0.2', effectiveFrom: '2026/10/01' })`), /yyyy-MM-dd/);
  assert.equal(env.table('SETTINGS').length, 0, 'だめな値は保存しない');
  env.call(`apiSaveSetting({ key: 'eval.half_wape_max', value: '0.15' })`);
  env.call(`apiSaveSetting({ key: 'eval.half_wape_max', value: 0.11 })`);
  s = cur();
  assert.deepEqual([s['eval.half_wape_max'].value, s['eval.half_wape_max'].isDefault, s['eval.half_wape_max'].effectiveFrom], [0.11, false, TODAY],
    '同じ日に 2 回変えたら後の値');
  assert.equal(env.table('SETTINGS').length, 2, '上書きせず行を足す');
  const end = env.audit().filter((a) => a.action === 'SETTING.SAVE' && a.result === 'OK')[1];
  assert.equal(JSON.parse(end.before_json).value, 0.15);
  assert.equal(JSON.parse(end.after_json).value, '0.11');
  env.call(`apiSaveSetting({ key: 'weather.min_months', value: 6, effectiveFrom: '${TOMORROW}' })`);
  s = cur();
  assert.equal(s['weather.min_months'].value, 3, '先の日付の設定はその日まで効かない');
  assert.deepEqual(s['weather.min_months'].scheduled, [{ value: '6', effectiveFrom: TOMORROW }]);
  env.call(`apiSaveSetting({ key: 'audit.log_views', value: 'true' })`);
  assert.equal(cur()['audit.log_views'].value, true);
  assert.equal(env.run(`appSettingValue_('eval.half_wape_max')`), 0.11);
  // 表に紛れ込んだ不正な値は使わず既定値に戻す
  env.run(`appInsertRows_('SETTINGS', [{ setting_id: 'ST-X', key: 'eval.overforecast_rate_max', value: '9', scope: 'GLOBAL', scope_id: '', effective_from: '${TODAY}', note: '', created_at: '', created_by: '' }])`);
  assert.equal(cur()['eval.overforecast_rate_max'].value, 0.05);
}

// ==== 10. 操作の記録（鎖・改ざんの検出・一覧） ====
{
  const env = setUpEnv();
  env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'A' })`);
  env.call(`apiSaveClient({ clientName: 'テスト製薬' })`);
  const sh = env.auditSheet();
  assert.deepEqual([...sh.rows[0]], [...auditCols]);
  const ip = auditCols.indexOf('prev_hash');
  sh.rows.slice(1).forEach((r, i) => {
    assert.equal(r[ip], i === 0 ? '' : sh.rows[i][ip + 1], '前の行の hash につながる');
    assert.equal(r[ip + 1], sha(r[ip] + '\n' + JSON.stringify(r.slice(0, ip))), 'hash は「前の hash + 改行 + 行の内容」の SHA-256');
  });
  assert.equal(env.props.APP_AUDIT_LAST_HASH, sh.rows.slice(-1)[0][ip + 1], '最新の hash を控える');
  const verify = () => env.call('appVerifyAuditSheet_(__sh)', { __sh: sh });
  let v = verify();
  assert.deepEqual([v.ok, v.rows, v.lastHash], [true, sh.rows.length - 1, env.props.APP_AUDIT_LAST_HASH]);
  const h = env.call('apiHealth()');
  assert.ok(h.setUp && h.tables.every((t) => t.ok), JSON.stringify(h.tables));
  assert.deepEqual([h.audit.ok, h.audit.matchesLatest], [true, true]);
  assert.match(h.files.data, /spreadsheets/);
  // 一覧: 新しい順・絞り込み
  const list = env.call(`apiListAudit({ limit: 3 })`).rows;
  assert.equal(list.length, 3);
  assert.ok(list[0].occurred_at >= list[2].occurred_at);
  const q = env.call(`apiListAudit({ query: 'client.save' })`).rows;
  assert.ok(q.length === 2 && q.every((a) => a.action === 'CLIENT.SAVE'), '大文字小文字を区別せず絞り込む');
  // 書き換え・削除を見つける
  const saved = sh.rows.map((r) => r.slice());
  sh.rows[2][4] = 'TAMPERED';
  v = verify();
  assert.deepEqual([v.ok, v.brokenAt], [false, 3]);
  assert.equal(env.call('apiHealth()').audit.ok, false);
  sh.rows.splice(0, sh.rows.length, ...saved);
  sh.rows.splice(2, 1);
  assert.deepEqual([verify().ok, verify().brokenAt], [false, 3]);
  sh.rows.splice(0, sh.rows.length, ...saved);
  sh.rows.pop();
  v = verify();
  assert.equal(v.ok, true);
  assert.equal(env.call('apiHealth()').audit.matchesLatest, false, '末尾を消すと最新の hash と合わない');
  // 定義と違う表を見つける
  sh.rows.splice(0, sh.rows.length, ...saved);
  const schemaSheet = env.data().getSheetByName('_SCHEMA');
  schemaSheet.rows[1][2] = 'x';
  const bad = env.call('apiHealth()').tables.filter((t) => !t.ok);
  assert.equal(bad.length, 1);
  assert.match(bad[0].note, /_SCHEMA と違う/);
  // 数値などが来ても文字列にして書く（読み戻した値で鎖が合う）
  sh.rows.splice(0, sh.rows.length, ...saved);
  schemaSheet.rows[1][2] = sha(tables[schemaSheet.rows[1][0]].columns.join('|'));
  env.run(`appAuditAppend_({ actor: '${OWNER}', action: 'TEST.NUM', phase: 'END', result: 'OK', entityId: 123, clientId: 0 })`);
  const lastRow = sh.rows.slice(-1)[0];
  assert.equal(lastRow[auditCols.indexOf('entity_id')], '123');
  assert.equal(verify().ok, true);
  // 長い JSON はハッシュと先頭だけ
  const big = { big: 'x'.repeat(50000) };
  const j = JSON.parse(env.run('appJson_(__b)', { __b: big }));
  assert.deepEqual([j.truncated, j.sha256, j.head.length], [true, sha(JSON.stringify(big)), 2000]);
  // 年度は 4 月始まり・終わる年で呼ぶ
  assert.deepEqual([env.run('appFy_(new Date(2026, 2, 31))'), env.run('appFy_(new Date(2026, 3, 1))'), env.run('appFy_(new Date(2026, 9, 1))')], [2026, 2027, 2027]);
}

// ==== 11. 記録できなければ書き込まない（fail-closed）・列が違う表には書かない ====
{
  const env = setUpEnv();
  env.auditSheet().failWrites = true;
  assert.throws(() => env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'A' })`), /write failed/);
  assert.equal(env.table('MEMBERS').length, 1, '開始を記録できなければ実行しない');
  env.auditSheet().failWrites = false;
  env.data().getSheetByName('MEMBERS').rows[0][1] = 'name';
  assert.throws(() => env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'A' })`), /表の列が定義と違います（MEMBERS）/);
  assert.equal(env.data().getSheetByName('MEMBERS').rows.length, 2);
  env.data().getSheetByName('MEMBERS').rows[0][1] = 'display_name';
  env.auditSheet().rows[0][5] = 'step';
  assert.throws(() => env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'A' })`), /ログの列が想定と違います/);
  assert.equal(env.table('MEMBERS').length, 1);
  env.auditSheet().rows[0][5] = 'phase';
  // キーの重複・空のキーは足さない
  assert.throws(() => env.run(`appInsertRows_('MEMBERS', [{ email: '${OWNER}', display_name: 'dup' }])`), /同じキーの行がすでにあります/);
  assert.throws(() => env.run(`appInsertRows_('MEMBERS', [{ email: '${MEMBER}' }, { email: '${MEMBER}' }])`), /同じキーの行/);
  assert.throws(() => env.run(`appInsertRows_('ROLES', [{ email: '${MEMBER}', role: 'ADMIN' }])`), /キーが空の行/);
  assert.equal(env.table('MEMBERS').length, 1);
  env.props.APP_LOG_SPREADSHEETS_JSON = JSON.stringify({ [FY]: 'missing' });
  assert.throws(() => env.call(`apiSaveMember({ email: '${MEMBER}', displayName: 'A' })`), /not found/);
  assert.equal(env.table('MEMBERS').length, 1, 'ログのファイルを開けなければ実行しない');
  assert.equal(env.state.lockHeld, false);
}

// ==== 12. バックアップ（14 世代・古いものはゴミ箱へ・データ本体は触らない） ====
{
  const env = setUpEnv();
  for (let i = 0; i < 16; i++) env.call('apiRunBackup()');
  const all = env.backups().sort((a, b) => a.created - b.created);
  assert.equal(all.length, 16);
  assert.deepEqual(all.map((f) => f.trashed), [true, true, ...Array(14).fill(false)], '新しい 14 世代を残し、古いものはゴミ箱へ（消さない）');
  assert.ok(all.every((f) => f.name.startsWith('売上予測アプリ データ バックアップ ')));
  assert.ok(env.sheetsById[all[15].id].getSheetByName('MEMBERS').rows.length === 2, '複製にはデータ本体の表が入る');
  assert.equal(env.files[env.props.APP_DATA_SPREADSHEET_ID].trashed, false);
  assert.equal(env.files[env.props.APP_DATA_SPREADSHEET_ID].parent, env.files[env.props.APP_BACKUP_FOLDER_ID].parent, 'データ本体は動かさない');
  assert.equal(env.runLog().filter((r) => r.kind === 'BACKUP' && r.status === 'OK').length, 16, '実行ログに残す');
  const h = env.call('apiHealth()');
  assert.deepEqual([h.backup.count, h.backup.enabled], [14, false], 'ゴミ箱のものは数えない');
  assert.equal(h.backup.latest, all[15].name);
  // 毎日のトリガー（二重に作らない）
  assert.equal(env.call('apiEnableBackup()').already, false);
  assert.equal(env.call('apiEnableBackup()').already, true);
  assert.equal(env.triggers.length, 1);
  assert.deepEqual([env.triggers[0].handler, env.triggers[0].everyDays, env.triggers[0].atHour], ['triggerDailyBackup', 1, 3]);
  assert.equal(env.call('apiHealth()').backup.enabled, true);
  // トリガーの中ではメールが空でも、このプロジェクトのトリガーなら所有者として動く
  const uid = env.triggers[0].uid;
  env.as('');
  env.call(`triggerDailyBackup({ triggerUid: '${uid}' })`);
  assert.equal(env.backups().length, 17);
  const last = env.audit().slice(-1)[0];
  assert.deepEqual([last.action, last.result, last.actor_email], ['BACKUP.DAILY', 'OK', OWNER]);
  assert.throws(() => env.call(`triggerDailyBackup({ triggerUid: 'forged' })`), /権限がありません/);
  assert.throws(() => env.call('triggerDailyBackup()'), /権限がありません/);
  // ブラウザから社内の人が呼んでも、本人の権限で判定する（UID を知っていても所有者にならない）
  env.as(MEMBER);
  assert.throws(() => env.call(`triggerDailyBackup({ triggerUid: '${uid}' })`), /権限がありません/);
  assert.equal(env.backups().length, 17);
}

// ==== 13. 画面を開く（doGet） ====
{
  const env = setUpEnv();
  const out = env.run('doGet({})');
  assert.equal(out.title, '売上予測アプリ');
  assert.match(out.favicon, /^data:image\/png;base64,.+#favicon\.png$/);
  const html = out.getContent();
  const boot = JSON.parse(/var B = (.*);\n/.exec(html)[1]);
  assert.deepEqual([boot.allowed, boot.setUp, boot.user.isAdmin, boot.app.name], [true, true, true, '売上予測アプリ']);
  assert.match(html, /var CHAR_SVG = \{/);
  assert.match(html, /var YOMI_POSE = \{/);
  assert.doesNotMatch(html, /<\?/);
  assert.equal(env.audit().length, 1, '画面を開くだけでは記録しない（設定 audit.log_views は既定で しない）');
  env.as(OUTSIDER);
  const b2 = JSON.parse(/var B = (.*);\n/.exec(env.run('doGet({})').getContent())[1]);
  assert.deepEqual([b2.allowed, b2.user.roles, b2.user.isAdmin], [false, [], false], '社外の人には役割も出さない');
  // 初期データの「<」は \u003c にする（</script> で画面が壊れない）
  env.run(`appBootstrap_ = function () { return { app: { name: '</script><b>x', version: '0' }, user: { roles: [] }, setUp: true, allowed: true }; }`);
  env.as(OWNER);
  const html3 = env.run('doGet({})').getContent();
  assert.ok(html3.includes('"\\u003c/script>\\u003cb>x"'));
  assert.equal((html3.match(/<\/script>/g) || []).length, 1, '閉じタグは本来の 1 つだけ');
  // ファイルの読み込み順に依存しない（GAS はプロジェクトのファイル順に実行する）
  const rev = setUpEnv({ order: 'reverse' });
  assert.equal(JSON.parse(/var B = (.*);\n/.exec(rev.run('doGet({})').getContent())[1]).setUp, true);
}

// ==== 14. 乱数の固定（同じ種なら同じ結果・終われば元に戻す） ====
{
  const env = makeEnv();
  const seq = (seed, n) => env.run(`(() => { const r = appSeededRandom_(__s); return Array.from({ length: ${n} }, () => r()); })()`, { __s: seed });
  const a = seq('RUN-1', 5);
  assert.deepEqual([...a], [...seq('RUN-1', 5)], '同じ種なら同じ並び');
  assert.notDeepEqual([...a], [...seq('RUN-2', 5)], '種が違えば違う並び');
  const many = [...seq('dist', 100000)];
  assert.ok(many.every((x) => x >= 0 && x < 1));
  const mean = many.reduce((s, x) => s + x, 0) / many.length;
  assert.ok(Math.abs(mean - 0.5) < 0.005, `平均 ${mean}`);
  const buckets = Array(10).fill(0);
  many.forEach((x) => { buckets[Math.floor(x * 10)]++; });
  assert.ok(buckets.every((b) => Math.abs(b - 10000) < 500), `偏り ${buckets}`);
  const original = env.run('Math.random');
  const inside = env.run(`appWithSeededRandom_('S', () => [Math.random(), Math.random()])`);
  assert.deepEqual([...inside], [...seq('S', 2)], 'fn の中の Math.random は種つき');
  assert.equal(env.run('Math.random'), original, '終われば元に戻す');
  assert.throws(() => env.run(`appWithSeededRandom_('S', () => { throw new Error('boom'); })`), /boom/);
  assert.equal(env.run('Math.random'), original, '失敗しても元に戻す');
  assert.throws(() => env.run(`appWithSeededRandom_('A', () => appWithSeededRandom_('B', () => 1))`), /入れ子/);
  assert.equal(env.run('Math.random'), original);
}

console.log('app-contract: all tests passed');
