/**
 * Verify year selection and exact archival text using synthetic raw tables.
 * This checks preservation, not forecast accuracy or production capacity.
 */
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { setUpEnv } from './gas-mock.mjs';

function writeRaw(env, name, objects) {
  const def = env.run(`APP_TABLES['${name}']`);
  const rows = objects.map(o => def.columns.map(c => String(o[c] ?? '')));
  const sh = env.data().getSheetByName(name);
  sh.getRange(2, 1, rows.length, def.columns.length).setNumberFormat('@').setValues(rows);
  env.run('APP_STORE_CACHE_ = {}');
}
function fixture() {
  const env = setUpEnv();
  writeRaw(env, 'PLANS', [
    { plan_id: 'P-B', client_id: 'C-B', fy: '2025', state: 'ACTIVE' },
    { plan_id: 'P-C', client_id: 'C-C', fy: '2026', state: 'ACTIVE' },
    { plan_id: 'P-A', client_id: 'C-A', fy: '2025', state: 'ACTIVE' },
  ]);
  writeRaw(env, 'CLIENTS', [
    { client_id: 'C-C', client_name: 'Gamma', is_active: 'TRUE' },
    { client_id: 'C-B', client_name: 'Beta', is_active: 'FALSE' },
    { client_id: 'C-A', client_name: 'Alpha', is_active: 'TRUE' },
  ]);
  writeRaw(env, 'CLIENT_NAMES', [
    { client_id: 'C-C', display_name: 'Other' },
    { client_id: 'C-A', display_name: 'Alpha display' },
  ]);
  writeRaw(env, 'ENG_SHEETS', [
    { plan_id: 'P-A', sheet: 'CONFIG', mode: 'rows', max_rows: '1000', max_columns: '26', last_row: '4', last_column: '2', content_hash: 'hash-config' },
    { plan_id: 'P-B', sheet: 'OUTPUT', mode: 'rows', max_rows: '1000', max_columns: '26', last_row: '40', last_column: '10', content_hash: 'hash-output' },
  ]);
  writeRaw(env, 'ENG_ROWS', [
    { plan_id: 'P-B', sheet: 'OUTPUT', row_no: '29', col_from: '1', cells_json: '["s2025/04","n0","n001","e","f=SUM(H1:H2)","bFALSE"]' },
    { plan_id: 'P-A', sheet: 'CONFIG', row_no: '2', col_from: '1', cells_json: '["sメーカー名","sAlpha","d2025-04-01T00:00:00.000Z"]' },
    { plan_id: 'P-C', sheet: 'CONFIG', row_no: '2', col_from: '1', cells_json: '["sOther year"]' },
  ]);
  writeRaw(env, 'ENG_FORMATS', [
    { plan_id: 'P-A', sheet: 'CONFIG', col: '2', runs_json: '[[1,1000,"@"]]' },
    { plan_id: 'P-C', sheet: 'CONFIG', col: '2', runs_json: '[[1,1000,"0.00"]]' },
  ]);
  writeRaw(env, 'ENG_PRODUCT', [
    { plan_id: 'P-B', seq: '2', Person: 'B', ProductName: '001', _types: 'ssdss' },
    { plan_id: 'P-A', seq: '2', Person: 'A', ProductName: '0', _types: 'ssdss' },
    { plan_id: 'P-C', seq: '2', Person: 'C', ProductName: 'other', _types: 'ssdss' },
  ]);
  writeRaw(env, 'FORECAST_RUNS', [
    { run_id: 'R-A', plan_id: 'P-A', status: 'DONE', annual_p50: '001' },
    { run_id: 'R-C', plan_id: 'P-C', status: 'DONE', annual_p50: '999' },
  ]);
  writeRaw(env, 'FORECAST_MONTHLY', [
    { run_id: 'R-A', plan_id: 'P-A', ym: '2025-05', p50: '0', p10: '', p90: '001' },
    { run_id: 'R-C', plan_id: 'P-C', ym: '2026-04', p50: '999' },
    { run_id: 'R-A', plan_id: 'P-A', ym: '2025-04', p50: '001', p10: '0', p90: '' },
  ]);
  writeRaw(env, 'PLAN_ACTIONS', [
    { action_id: 'A-A', plan_id: 'P-A', action: 'INPUT.SAVE', status: 'OK' },
    { action_id: 'A-C', plan_id: 'P-C', action: 'INPUT.SAVE', status: 'OK' },
  ]);
  writeRaw(env, 'PLAN_VERSIONS', [
    { version_id: 'V-A', plan_id: 'P-A', version_no: '001', state: 'APPROVED', monthly_json: '[{"value":"001"}]' },
    { version_id: 'V-C', plan_id: 'P-C', version_no: '1', state: 'APPROVED' },
  ]);
  return env;
}
const snapshot = env => env.call('appYearSnapshot_(2025)');
const tables = s => Object.fromEntries(JSON.parse(s.text).tables.map(t => [t.name, t]));
const field = (t, row, col) => t.rows[row][t.columns.indexOf(col)];
const hash = text => createHash('sha256').update(text, 'utf8').digest('hex');

{
  const env = fixture();
  const before = JSON.stringify(env.data().getSheets().map(sh => [sh.name, sh.rows]));
  const s = snapshot(env);
  assert.equal(s.planCount, 2);
  assert.equal(s.rowCount, 17);
  assert.deepEqual(s.planIds, ['P-A', 'P-B']);
  assert.equal(s.sha256, hash(s.text));
  assert.equal(s.bytes, Buffer.byteLength(s.text, 'utf8'));
  const ts = tables(s);
  const counts = { PLANS: 2, CLIENTS: 2, CLIENT_NAMES: 1, ENG_SHEETS: 2, ENG_ROWS: 2,
    ENG_FORMATS: 1, ENG_PRODUCT: 2, FORECAST_RUNS: 1, FORECAST_MONTHLY: 2, PLAN_ACTIONS: 1, PLAN_VERSIONS: 1 };
  for (const [name, t] of Object.entries(ts)) assert.equal(t.rows.length, counts[name] ?? 0, name);
  for (const name of ['MEMBERS', 'ROLES', 'SETTINGS', '_SCHEMA', 'AUDIT_ANCHORS', 'YEAR_CLOSURES']) {
    assert.equal(ts[name], undefined, `not annual data: ${name}`);
  }
  assert.equal(field(ts.PLANS, 0, 'plan_id'), 'P-A');
  assert.equal(field(ts.CLIENTS, 1, 'is_active'), 'FALSE');
  assert.equal(field(ts.FORECAST_RUNS, 0, 'annual_p50'), '001');
  assert.equal(field(ts.FORECAST_MONTHLY, 0, 'ym'), '2025-04');
  assert.equal(field(ts.FORECAST_MONTHLY, 0, 'p50'), '001');
  assert.equal(field(ts.FORECAST_MONTHLY, 0, 'p90'), '');
  assert.equal(field(ts.ENG_FORMATS, 0, 'runs_json'), '[[1,1000,"@"]]');
  assert.equal(field(ts.ENG_ROWS, 1, 'cells_json'), '["s2025/04","n0","n001","e","f=SUM(H1:H2)","bFALSE"]');
  assert.equal(JSON.stringify(env.data().getSheets().map(sh => [sh.name, sh.rows])), before);
  assert.equal(snapshot(env).text, s.text);
  const file = { isTrashed: () => false, getBlob: () => ({ getDataAsString: encoding => {
    assert.equal(encoding, 'UTF-8');
    return s.text;
  } }) };
  assert.equal(env.run('appYearVerifySnapshot_(__s, __file)', { __s: s, __file: file }), true);
  const bad = { isTrashed: () => false, getBlob: () => ({ getDataAsString: () => s.text.replace('"001"', '"1"') }) };
  assert.throws(() => env.run('appYearVerifySnapshot_(__s, __file)', { __s: s, __file: bad }), /一致しません/);
  assert.throws(() => env.run('appYearVerifySnapshot_(__s, __file)', {
    __s: s, __file: { ...file, isTrashed: () => true },
  }), /ゴミ箱/);
  const sh = env.data().getSheetByName('PLANS');
  [sh.rows[1], sh.rows[3]] = [sh.rows[3], sh.rows[1]];
  env.run('APP_STORE_CACHE_ = {}');
  assert.equal(snapshot(env).text, s.text, 'physical row order must not change the fingerprint');
}
{
  const env = fixture();
  env.run(`APP_TABLES.FUTURE_PLAN_DATA = { key: ['plan_id', 'seq'], columns: ['plan_id', 'seq', 'value'] }; appTableSheet_('FUTURE_PLAN_DATA', true);`);
  writeRaw(env, 'FUTURE_PLAN_DATA', [
    { plan_id: 'P-A', seq: '1', value: 'future-owned-data' },
    { plan_id: 'P-C', seq: '1', value: 'other-year' },
  ]);
  const s = snapshot(env);
  assert.equal(s.rowCount, 18);
  assert.deepEqual(tables(s).FUTURE_PLAN_DATA.rows, [['P-A', '1', 'future-owned-data']]);
}
{
  const env = fixture();
  const sh = env.data().getSheetByName('ENG_ROWS');
  sh.rows.push(sh.rows[1].slice());
  env.run('APP_STORE_CACHE_ = {}');
  assert.throws(() => snapshot(env), /重複/);
}
{
  const env = fixture();
  writeRaw(env, 'PLANS', [{ plan_id: '', client_id: 'C-A', fy: '2025' }]);
  assert.throws(() => snapshot(env), /キーか列/);
}
{
  const env = fixture();
  const sh = env.data().getSheetByName('CLIENTS');
  sh.rows.splice(3, 1);
  env.run('APP_STORE_CACHE_ = {}');
  assert.throws(() => snapshot(env), /メーカーが見つかりません/);
}
{
  const env = fixture();
  for (const fy of [2025.5, 1999, 2101, 'invalid']) {
    assert.throws(() => env.call('appYearSnapshot_(__fy)', { __fy: fy }), /年度/);
  }
  assert.throws(() => env.call('appYearSnapshot_(2024)'), /計画はありません/);
  env.run('appUtf8Bytes_ = function () { return 8 * 1024 * 1024 + 1; };');
  assert.throws(() => snapshot(env), /8MB/);
}
{
  const env = fixture();
  assert.equal(env.run('appYearStoredClosure_(2025)'), null);
  writeRaw(env, 'YEAR_CLOSURES', [{ fy: '2025', state: 'CLOSED', file_id: 'ARCHIVE-1' }]);
  assert.equal(env.run('appYearStoredClosure_(2025).file_id'), 'ARCHIVE-1');
  writeRaw(env, 'YEAR_CLOSURES', [{ fy: 'broken', state: 'CLOSED', file_id: 'ARCHIVE-1' }]);
  assert.throws(() => env.run('appYearStoredClosure_(2025)'), /記録を確認できません/);
  writeRaw(env, 'YEAR_CLOSURES', [{ fy: '', state: 'CLOSED', file_id: 'ARCHIVE-1' }]);
  assert.throws(() => env.run('appYearStoredClosure_(2025)'), /キーか列/);
}
{
  const env = fixture();
  writeRaw(env, 'YEAR_CLOSURES', [
    { fy: '2025', state: 'CLOSED', file_id: 'ARCHIVE-1' },
    { fy: '2025', state: 'CLOSED', file_id: 'ARCHIVE-2' },
  ]);
  assert.throws(() => env.run('appYearStoredClosure_(2025)'), /重複/);
}
{
  const env = fixture();
  for (const fy of ['broken', '', 2025.5, 0]) {
    assert.throws(() => env.call('appYearStoredClosure_(__fy)', { __fy: fy }), /年度を確認できません/);
  }
  writeRaw(env, 'PLANS', [{ plan_id: 'P-B', client_id: 'C-B', fy: 'broken' }]);
  assert.throws(() => env.run('appYearStoredPlans_()'), /年度を確認できません/);
}
console.log('app-year-snapshot: all preservation and selection checks passed');
