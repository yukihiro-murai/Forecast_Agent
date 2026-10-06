#!/usr/bin/env node
/**
 * app-speed.test.mjs — 読むだけの処理を速くしても、結果は変わらない。
 *   表を読むのは 1 回（見出しと本文を一緒に読み、見出しが違えば今までどおり止める）・計画の行だけを読む・
 *   値だけを使うところは表示形式を読まない・検証の画面の控え・控えと処理の記録をまとめて読む・バックアップの要点を残す。
 *
 *   node app/tests/app-speed.test.mjs
 */
import assert from 'node:assert/strict';
import { setUpEnv, STATS, J, sha, sources, extractFunction } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const reset = () => { for (const k of Object.keys(STATS)) STATS[k] = k === 'bySheet' ? {} : 0; };
const reads = (name) => (STATS.bySheet[name] || { reads: 0 }).reads;
const env = setUpEnv();
const H = env.run('(() => { const o = {}; Object.keys(APP_ENGINE_SHEETS).forEach(k => { o[k] = APP_ENGINE_SHEETS[k].header || null; }); return o; })()');

/** 計画のブック: OUTPUT（年度合計・月ごと・採用予測と上乗せ・目標）と、検証の記録（予測は実績の over 倍） */
function book(client, base, over) {
  const out = [['FY2026 売上予測（' + client + '）']];
  for (let r = 2; r <= 25; r++) out.push([]);
  out.push(['年度合計（予測）', base * 0.9, base, base * 1.1]);
  out.push([], []);
  for (let i = 0; i < 12; i++) out.push([D(2026, 4 + i), base / 12 * 0.9, base / 12, base / 12 * 1.1, '', '', '', base / 12, i % 2 ? 5 : '']);
  for (let r = 41; r <= 64; r++) out.push([]);
  out.push(['目標', base * 0.8, base * 0.95, base * 1.05]);
  const ev = [H.EVAL_LOG];
  for (let i = 0; i < 6; i++) {
    const ym = '2025/' + String(i + 1).padStart(2, '0');
    const act = 1000 + i * 15;
    const p50 = act * over;
    for (const [sc, p] of [['nega', p50 * 0.9], ['neutral', p50], ['posi', p50 * 1.1]]) {
      ev.push(['E' + i + sc, D(2026, 1, 5), client, ym, sc, p, act, Math.abs(p - act) / act, 0, '', '', sc === 'neutral' ? 1 : 0, p - act, Math.abs(p - act), '', '', '', '', '', '', '', 1]);
    }
  }
  return env.makeBook(client, {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', client], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
    OUTPUT: { values: out, formats: { B: '#,##0', C: '#,##0' } },
    EVAL_LOG: { values: ev, formats: { D: '@' } },
    CALIBRATION_STATE: { values: [H.CALIBRATION_STATE, [client, D(2026, 1, 5), 'owner', '', '', '', 0.97, '', '{}', '', '', 1, '']] },
  });
}
const ids = [['甲製薬', 1200, 1.08], ['乙製薬', 3600, 0.96], ['丙製薬', 2400, 1.02]].map(([c, b, o]) => env.seedPlan(book(c, b, o)));

/** 今の関数の読み方だけを、前の読み方（表を全部読んで計画で絞る）に戻した写し（名前の末尾を Whole_ にする） */
function whole(file, name, from, to) {
  const src = extractFunction(sources[file], name);
  assert.ok(src.includes(from), name + ' の読み方が想定と違う');
  env.run(src.replace('function ' + name + '(', 'function ' + name.replace(/_$/, 'Whole_') + '(').split(from).join(to));
}
whole('Forecast.js', 'appStoredHeadline_', "appReadPlanTable_('ENG_ROWS', planId).filter(r => r.sheet === 'OUTPUT')", "appReadTable_('ENG_ROWS').filter(r => r.plan_id === planId && r.sheet === 'OUTPUT')");
whole('Versions.js', 'appPlanNumbers_', "appReadPlanTable_('ENG_ROWS', planId).filter(r => r.sheet === 'OUTPUT' && Number(r.col_from) === 1)", "appReadTable_('ENG_ROWS').filter(r => r.plan_id === planId && r.sheet === 'OUTPUT' && Number(r.col_from) === 1)");
whole('Forecast.js', 'appPlanInputHash_', "appReadPlanTable_('ENG_SHEETS', planId)", "appReadTable_('ENG_SHEETS').filter(r => r.plan_id === planId)");
whole('Forecast.js', 'appStoredHashes_', "appReadPlanTable_('ENG_SHEETS', planId)", "appReadTable_('ENG_SHEETS').filter(r => r.plan_id === planId)");
const fresh = (code, extra) => env.call(`(() => { APP_STORE_CACHE_ = {}; return ${code}; })()`, extra);

// ==== 1. 計画の行だけを読む: 全部読んで絞ったときと同じ結果（表が小さいとき・大きいとき） ====
for (const minRows of [3000, 0]) {
  env.run(`appPlanReadMinRows_ = () => ${minRows}`);
  for (const id of ids) {
    for (const fn of ['appStoredHeadline', 'appPlanNumbers', 'appPlanInputHash', 'appStoredHashes']) {
      const now = fresh(`${fn}_(__id)`, { __id: id });
      assert.deepEqual(now, fresh(`${fn}Whole_(__id)`, { __id: id }), `${fn}_ は前と同じ結果（${minRows}）`);
      assert.ok(now !== null && now !== undefined && JSON.stringify(now) !== '{}', `${fn}_ が中身を返す`);
    }
  }
  const h = fresh('appStoredHeadline_(__id)', { __id: ids[1] });
  assert.deepEqual([h.annual.p50, h.objective.p50, h.monthly.length, h.monthly[0].month], [3600, 3420, 12, '2026/04'], '計画ごとの数字');
  const nums = fresh('appPlanNumbers_(__id)', { __id: ids[0] });
  assert.deepEqual([nums.annual.p50, nums.budget.uplift, nums.monthly[1].adopted, nums.monthly[1].final], [1200, 30, 100, 105], '採用予測と上乗せ');
  const mine = env.table('ENG_SHEETS').filter((r) => r.plan_id === ids[2]);
  assert.equal(fresh('appPlanInputHash_(__id)', { __id: ids[2] }), sha(mine.map((r) => r.sheet + ':' + r.content_hash).sort().join('|')), '入力のハッシュはその計画の行だけから');
  assert.deepEqual(fresh('appStoredHashes_(__id)', { __id: ids[2] }), Object.fromEntries(mine.map((r) => [r.sheet, r.content_hash])));
}
env.run('appPlanReadMinRows_ = () => 0');
reset();
fresh('appStoredHeadline_(__id)', { __id: ids[2] });
const planCells = STATS.readCells;
reset();
fresh('appStoredHeadlineWhole_(__id)', { __id: ids[2] });
assert.ok(planCells < STATS.readCells / 2, '大きい表は計画の行だけを読む: ' + STATS.readCells + ' → ' + planCells);
env.run('appPlanReadMinRows_ = () => 3000');

// ==== 2. 表を読むのは 1 回（見出しと本文を一緒に）。最後の行も調べ直さない。見出しが違えば今までどおり止める ====
{
  reset();
  const plans = fresh("appReadTable_('PLANS').length");
  assert.equal(plans, 3);
  assert.equal(reads('PLANS'), 1, '見出しと本文を 1 回で読む');
  const sh = env.data().getSheetByName('ENG_ROWS');
  const orig = sh.getLastRow;
  let lastRows = 0;
  sh.getLastRow = () => { lastRows++; return orig(); };
  reset();
  assert.ok(fresh("appReadPlanTable_('ENG_ROWS', __id).length", { __id: ids[0] }) > 0);
  assert.deepEqual([reads('ENG_ROWS'), lastRows], [1, 1], '小さい表: 読むのも最後の行を調べるのも 1 回');
  sh.getLastRow = orig;
  // 書いた後に読み直すときは、見出しを確かめ直さない（本文だけ）
  reset();
  env.run("(() => { APP_STORE_CACHE_ = {}; appReadTable_('PLANS'); appStoreForget_('PLANS'); return appReadTable_('PLANS').length; })()");
  assert.equal(reads('PLANS'), 2);
  // 見出しが定義と違えば止める（全部読む・計画の行だけ読む・書く、のどれでも）
  for (const name of ['PLANS', 'ENG_SHEETS']) {
    const s = env.data().getSheetByName(name);
    const keep = s.rows[0][1];
    s.rows[0][1] = '壊れた見出し';
    const re = new RegExp('表の列が定義と違います（' + name + '）');
    assert.throws(() => fresh(`appReadTable_('${name}')`), re);
    if (name === 'ENG_SHEETS') {
      for (const minRows of [3000, 0]) {
        env.run(`appPlanReadMinRows_ = () => ${minRows}`);
        assert.throws(() => fresh(`appReadPlanTable_('ENG_SHEETS', __id)`, { __id: ids[0] }), re, '計画の行だけ読むときも止める: ' + minRows);
      }
      env.run('appPlanReadMinRows_ = () => 3000');
    }
    assert.throws(() => fresh(`appTableSheet_('${name}', false)`), re, '書く前の確かめも今までどおり');
    s.rows[0][1] = keep;
    assert.ok(fresh(`appReadTable_('${name}').length`) > 0, '直せば読める');
  }
  assert.equal(env.call('apiListPlans()').plans.length, 3, '直した後は画面も読める');
}

// ==== 3. 値だけを使うところは表示形式（ENG_FORMATS）を読まない。値と数式は同じ ====
{
  const full = fresh('appEngLoadPlanSheets_(__id)', { __id: ids[0] });
  const vals = fresh('appEngLoadPlanSheets_(__id, null, true)', { __id: ids[0] });
  assert.deepEqual(Object.keys(vals).sort(), Object.keys(full).sort());
  for (const n of Object.keys(full)) {
    assert.deepEqual([vals[n].values, vals[n].formulas, vals[n].lastRow, vals[n].lastCol, vals[n].maxRows], [full[n].values, full[n].formulas, full[n].lastRow, full[n].lastCol, full[n].maxRows], n + ' の値と数式は同じ');
    assert.equal(vals[n].fmtCols, 0);
  }
  assert.ok(Object.values(full).some((s) => s.formats.some((r) => r.some((f) => f !== ''))), '全部読むときは表示形式もある');
  reset();
  const t = fresh("appEngTableObjects_(__id, ['EVAL_LOG', 'CALIBRATION_STATE'])", { __id: ids[0] });
  assert.equal(reads('ENG_FORMATS'), 0, '表示形式は読まない');
  assert.equal(t.EVAL_LOG.length, 18);
  assert.equal(t.CALIBRATION_STATE[0].bias_correction_factor, 0.97);
}

// ==== 4. 検証の画面: 前と同じ結果・データ本体が変わらない間は表を読まない・精度の控えはまとめて読む ====
{
  // 前の作り（計画ごとに控えを読む・表示形式も読む）と比べる
  const src = extractFunction(sources['Learning.js'], 'appLearningViewData_');
  const from = "const accs = appAccuracyAll_(appReadTable_('PLANS').filter(p => p.state !== 'ARCHIVED').map(p => p.plan_id));";
  assert.ok(src.includes(from));
  env.run(src.replace('function appLearningViewData_(', 'function appLearningViewBefore_(').replace(from,
    "const accs = {}; appReadTable_('PLANS').filter(p => p.state !== 'ARCHIVED').forEach(p => { try { accs[p.plan_id] = appAccuracyCached_(p.plan_id); } catch (e) { } });"));
  whole('Basis.js', 'appEngTableObjects_', 'appEngLoadPlanSheets_(planId, names, true)', 'appEngLoadPlanSheets_(planId, names)');
  const clearCache = () => { for (const k of Object.keys(env.cache)) delete env.cache[k]; };
  for (const id of ids) {
    clearCache();
    const before = env.call(`(() => { APP_STORE_CACHE_ = {}; const keep = appEngTableObjects_; appEngTableObjects_ = appEngTableObjectsWhole_;
      try { return appLearningViewBefore_(__id); } finally { appEngTableObjects_ = keep; } })()`, { __id: id });
    clearCache();
    const now = env.call('apiLearningView(__in)', { __in: { planId: id } });
    assert.deepEqual(now, before, '前と同じ結果: ' + id);
    assert.equal(now.pooledPlans, 3);
  }
  // 2 回目は表を読まない
  env.call('apiLearningView(__in)', { __in: { planId: ids[1] } });
  reset();
  const again = env.call('apiLearningView(__in)', { __in: { planId: ids[1] } });
  assert.equal(STATS.reads, 0, '控えから返す');
  assert.equal(again.planId, ids[1]);
  // 精度の控え（ACC_）は、計画の数によらず getAll でまとめて読む
  const orig = env.ctx.CacheService.getScriptCache;
  const ops = { get: 0, getAll: 0 };
  env.ctx.CacheService.getScriptCache = () => { const c = orig(); return Object.assign({}, c, { get: (k) => (ops.get++, c.get(k)), getAll: (ks) => (ops.getAll++, c.getAll(ks)) }); };
  env.run(`CacheService.getScriptCache().remove('APP_DATA_GEN')`);   // ほかの人が何か保存した後と同じ（精度の控えは残る）
  ops.get = 0; ops.getAll = 0;
  reset();
  assert.deepEqual(env.call('apiLearningView(__in)', { __in: { planId: ids[1] } }), again);
  assert.ok(ops.get <= 1 && ops.getAll <= 3, '控えは計画の数によらず数回で読む: ' + JSON.stringify(ops));
  assert.equal(reads('ENG_EVAL_LOG'), 0, '精度は控えから（検証の記録は読まない）');
  env.ctx.CacheService.getScriptCache = orig;
  // 保存したら読み直す
  env.call(`apiSaveClient({ clientName: '別の製薬' })`);
  reset();
  env.call('apiLearningView(__in)', { __in: { planId: ids[1] } });
  assert.ok(STATS.reads > 0, '書いた後は読み直す');
  assert.throws(() => env.call('apiLearningView(__in)', { __in: { planId: 'PL-NONE' } }), /計画が見つかりません/, '無い計画は今までどおり止まる（控えない）');
}

// ==== 5. 分けて置いた控えは getAll でまとめて読む。欠けていれば読み直す ====
{
  const orig = env.ctx.CacheService.getScriptCache;
  const ops = { get: 0, getAll: 0 };
  env.ctx.CacheService.getScriptCache = () => { const c = orig(); return Object.assign({}, c, { get: (k) => (ops.get++, c.get(k)), getAll: (ks) => (ops.getAll++, c.getAll(ks)) }); };
  env.run('__calls = 0');
  const big = () => env.call(`(() => { APP_STORE_CACHE_ = {}; return appCachedRead_('BIG', () => { __calls++; return { t: 'あ'.repeat(95000) }; }); })()`);
  assert.equal(big().t.length, 95000);
  ops.get = 0; ops.getAll = 0;
  assert.equal(big().t.length, 95000);
  assert.deepEqual([env.run('__calls'), ops.get, ops.getAll], [1, 1, 2], '4 つに分けた控えを、版の印 1 回と getAll 2 回で読む');
  const k = Object.keys(env.cache).find((x) => /^RC_[0-9a-f]{32}_2$/.test(x) && env.cache[x].startsWith('あ'));
  delete env.cache[k];
  assert.equal(big().t.length, 95000);
  assert.equal(env.run('__calls'), 2, '欠けていれば読み直す');
  // 処理の結果（appJobGetResult_）も同じ
  env.run(`appJobPutResult_('SPEED', { t: 'い'.repeat(70000) })`);
  ops.get = 0; ops.getAll = 0;
  assert.equal(J(env.run(`appJobGetResult_('SPEED')`)).value.t.length, 70000);
  assert.deepEqual([ops.get, ops.getAll], [0, 2]);
  delete env.cache['APP_JOB_RESULT_SPEED_1'];
  assert.deepEqual(J(env.run(`appJobGetResult_('SPEED')`)), { found: false });
  const many = J(env.run(`appJobGetResults_(['SPEED', 'NONE'])`));
  assert.deepEqual(many, { SPEED: { found: false }, NONE: { found: false } });
  env.run(`appJobPutResult_('A1', { a: 1 }); appJobPutResult_('B2', { b: 'う'.repeat(40000) })`);
  ops.get = 0; ops.getAll = 0;
  assert.deepEqual(J(env.run(`appJobGetResults_(['A1', 'B2', 'NONE'])`)), { A1: { found: true, value: { a: 1 } }, B2: { found: true, value: { b: 'う'.repeat(40000) } }, NONE: { found: false } });
  assert.deepEqual([ops.get, ops.getAll], [0, 2]);
  env.ctx.CacheService.getScriptCache = orig;
}

// ==== 6. 処理の記録の一覧は、Script Properties を 1 回で読む（前と同じ一覧） ====
{
  for (let i = 0; i < 3; i++) assert.equal(env.runJob('SYSTEM.RECOVER', {}).status, 'DONE');
  env.call('apiStartJob(__in)', { __in: { kind: 'SYSTEM.RECOVER', payload: {} } });   // 待っている処理も 1 つ
  const before = J(env.run(`appProps_().getKeys().filter(k => k.indexOf(APP_JOB_PREFIX) === 0 && k.indexOf(APP_JOB_RESULT_PREFIX) !== 0)
    .map(k => { try { return JSON.parse(appProps_().getProperty(k)); } catch (e) { return null; } }).filter(Boolean).sort((a, b) => (a.createdAt < b.createdAt ? -1 : 1))`));
  const orig = env.ctx.PropertiesService.getScriptProperties;
  let gets = 0;
  env.ctx.PropertiesService.getScriptProperties = () => { const p = orig(); return Object.assign({}, p, { getProperty: (k) => (gets++, p.getProperty(k)) }); };
  const now = J(env.run('appJobList_()'));
  env.ctx.PropertiesService.getScriptProperties = orig;
  assert.ok(before.length >= 4, '処理の記録がある: ' + before.length);
  assert.deepEqual(now, before);
  assert.equal(gets, 0, '1 つずつは読まない');
  env.fireTriggers('triggerRunJob');
}

// ==== 7. バックアップ: 要点を残し、ホームはドライブを数えずに読む（状態の点検は今までどおり数える） ====
{
  const drive = env.ctx.DriveApp;
  const origFolder = drive.getFolderById;
  let listed = 0;
  drive.getFolderById = (id) => { const f = origFolder(id); return Object.assign(Object.create(f), { getFilesByType: (m) => { listed++; return f.getFilesByType(m); } }); };
  const fast = () => J(env.run('appBackupStatusFast_()'));
  const slow = () => J(env.run('appBackupStatus_()'));
  // 要点がまだ無い（この版の前のバックアップだけ）: 今までどおり数え、5 分控える
  for (const k of Object.keys(env.cache)) delete env.cache[k];
  delete env.props.APP_BACKUP_STATUS;
  env.call('apiRunBackup()');
  delete env.props.APP_BACKUP_STATUS;
  listed = 0;
  assert.deepEqual(fast(), slow());
  assert.equal(listed, 2);
  fast();
  assert.equal(listed, 2, '数えた結果は 5 分控える');
  // バックアップすると要点を残す。ホームは数えない
  for (const k of Object.keys(env.cache)) delete env.cache[k];
  env.call('apiRunBackup()');
  assert.ok(env.props.APP_BACKUP_STATUS, '要点を残す');
  listed = 0;
  assert.deepEqual(fast(), slow());
  assert.equal(listed, 1, '要点から返す（数えたのは比べた側だけ）');
  assert.deepEqual([fast().count, fast().enabled, fast().latest.startsWith('売上予測アプリ データ バックアップ ')], [2, false, true]);
  listed = 0;
  const home = env.call('apiHome()');
  assert.deepEqual(home.system.backup, slow());
  assert.equal(listed, 1, 'ホームはドライブを数えない');
  // 毎日のバックアップを有効にすると、要点の「有効」も合わせる
  env.call('apiEnableBackup()');
  assert.deepEqual(fast(), slow());
  assert.equal(fast().enabled, true);
  // 古い世代を移せずに止まっても、作った複製と残っている数を残す
  for (let i = 0; i < 12; i++) env.call('apiRunBackup()');
  assert.equal(fast().count, 14);
  const oldest = env.backups().filter((f) => !f.trashed).sort((a, b) => a.created - b.created)[0];
  const move = oldest.moveTo;
  oldest.moveTo = () => { throw new Error('move failed'); };
  assert.throws(() => env.call('apiRunBackup()'), /move failed/);
  assert.deepEqual(fast(), slow());
  assert.equal(fast().count, 15);
  oldest.moveTo = move;
  env.call('apiRunBackup()');
  assert.deepEqual(fast(), slow());
  assert.equal(fast().count, 14);
  // 状態の点検（apiHealth）は今までどおりドライブを数える
  listed = 0;
  const health = env.call('apiHealth()');
  assert.deepEqual(health.backup, slow());
  assert.equal(listed, 2, '状態の点検は数える');
  // 要点が壊れていれば数える
  env.props.APP_BACKUP_STATUS = '{壊れた';
  for (const k of Object.keys(env.cache)) delete env.cache[k];
  assert.deepEqual(fast(), slow());
  drive.getFolderById = origFolder;
}

console.log('app-speed: all tests passed');
