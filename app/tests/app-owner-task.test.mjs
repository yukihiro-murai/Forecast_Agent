#!/usr/bin/env node
/*
 * app-owner-task.test.mjs — 所有者が Apps Script のエディタから管理の操作をする入口（apiOwnerTask・OwnerTask.js）。
 * スクリプト プロパティ OWNER_TASK に JSON を置き、エディタと同じく引数なしで実行する。
 * 画面の入口と同じ中の関数・同じ監査（操作ごとの名前）で動くこと・結果は OWNER_TASK_RESULT（9KB まで）と実行ログ・
 * うまくいけば頼みを消し、失敗したら残すこと・所有者だけ・年度の締めは確認の指紋が要ることを確かめる。
 *
 *   node app/tests/app-owner-task.test.mjs
 */
import assert from 'node:assert/strict';
import { OWNER, MEMBER, OUTSIDER, sources, extractFunction, setUpEnv } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const RESULT_MAX = 8000;

/** OWNER_TASK に頼みを置いて、エディタからの実行と同じく引数なしで apiOwnerTask を動かす */
function task(env, t) {
  env.props.OWNER_TASK = typeof t === 'string' ? t : JSON.stringify(t);
  return env.call('apiOwnerTask()');
}
/** OWNER_TASK_RESULT（9KB の上限に収まる JSON） */
function result(env) {
  const raw = env.props.OWNER_TASK_RESULT;
  assert.ok(raw, 'OWNER_TASK_RESULT がある');
  assert.ok(Buffer.byteLength(raw, 'utf8') <= RESULT_MAX, '結果は 1 つの値の上限に収まる: ' + Buffer.byteLength(raw, 'utf8'));
  return JSON.parse(raw);
}
/** うまくいく: 結果が書かれ、頼みは消える。{ full: 返り値（全部）, rec: OWNER_TASK_RESULT } */
function ok(env, t) {
  const full = task(env, t);
  const rec = result(env);
  assert.equal(rec.ok, true, JSON.stringify(rec));
  assert.equal(rec.action, t.action);
  assert.equal(full.ok, true);
  assert.equal('OWNER_TASK' in env.props, false, 'うまくいったら頼みを消す: ' + t.action);
  return { full, rec };
}
/** 失敗する: 投げる・エラーが OWNER_TASK_RESULT に書かれ、頼みは残る */
function fails(env, t, re) {
  const raw = typeof t === 'string' ? t : JSON.stringify(t);
  assert.throws(() => task(env, t), re);
  const rec = result(env);
  assert.equal(rec.ok, false);
  assert.match(rec.error, re);
  assert.equal(env.props.OWNER_TASK, raw, '失敗したら頼みを残す（直してもう一度実行する）');
  return rec;
}
/** 裏の処理をトリガーで動かし、jobStatus で終わりの状態を受け取る */
function finishJob(env, jobId) {
  for (let i = 0; i < 30; i++) {
    env.fireTriggers('triggerRunJob');
    const st = ok(env, jobId ? { action: 'jobStatus', jobId } : { action: 'jobStatus' }).full.result;
    if (['QUEUED', 'RUNNING', 'CONTINUED'].includes(st.status)) continue;
    return st;
  }
  throw new Error('処理が終わらない');
}
/** 監査の行（操作の名前・段階で） */
const audits = (env, action, phase) => env.audit().filter((a) => a.action === action && (!phase || a.phase === phase));
/** 書く操作は、入口と同じ操作の名前で開始と終了（OK）が所有者の名前で残り、理由の欄にエディタからと残る */
function audited(env, action) {
  const start = audits(env, action, 'START').slice(-1)[0];
  const end = audits(env, action, 'END').slice(-1)[0];
  assert.ok(start && end, action + ' の開始と終了の記録');
  assert.equal(end.result, 'OK', action);
  assert.equal(end.actor_email, OWNER);
  assert.equal(start.request_id, end.request_id);
  assert.match(start.reason, /OWNER_TASK/, 'エディタからの操作と分かる');
  return { start, end };
}
const jobProps = (env) => Object.keys(env.props).filter((k) => /^APP_JOB_(?!RESULT_)/.test(k)).length;

// ==== 1. 入口: 所有者だけ・本体は api_ だけ・この入口そのものは記録しない ====
{
  const body = extractFunction(sources['Api.js'], 'apiOwnerTask');
  assert.match(body, /return api_\('OWNER\.TASK', \{ ownerOnly: true, audit: false \}, ctx => appOwnerTask_\(ctx\)\);/);
  assert.doesNotMatch(sources['OwnerTask.js'], /SpreadsheetApp\./, '保存の層を通す');
}

const env = setUpEnv();
const currentFy = env.run('appFy_(new Date())');
const past = currentFy - 1;

// ==== 2. 頼みが空・JSON が読めない・形が違う・未定義の操作・使えない項目: 何もせずにエラーを残す ====
{
  const members0 = JSON.stringify(env.table('MEMBERS'));
  const audit0 = env.audit().length;
  fails(env, '', /OWNER_TASK が空です/);
  fails(env, '   ', /OWNER_TASK が空です/);
  fails(env, '{"action":"health"', /JSON が読めません/);
  fails(env, "{'action':'health'}", /JSON が読めません/);
  fails(env, '[1,2]', /\{"action":"…"\} の形/);
  fails(env, '"health"', /\{"action":"…"\} の形/);
  fails(env, {}, /未定義の操作です: （action がありません）/);
  const unknown = fails(env, { action: 'dropAllTables' }, /未定義の操作です: dropAllTables/);
  assert.match(unknown.error, /listDirectory/, '使える操作の一覧を出す');
  fails(env, { action: 'toString' }, /未定義の操作です: toString/);
  fails(env, { action: 'saveMember', mail: MEMBER, displayName: 'X' }, /使えない項目があります: mail。saveMember で使える項目: email/);
  fails(env, { action: 'health', verbose: true }, /使えない項目があります: verbose/);
  assert.equal(JSON.stringify(env.table('MEMBERS')), members0, '何も書かない');
  assert.equal(env.audit().length, audit0, '操作の記録も増えない（何も動かしていない）');
  assert.ok(env.errors().some((e) => e.where === 'OWNER.TASK' && /未定義の操作/.test(e.message)), 'エラーのログに残る');
  assert.ok(env.runLog().some((r) => r.kind === 'OWNER.TASK' && r.status === 'FAILED'), '実行ログに残る');
  assert.ok(env.logs.some((l) => /OWNER_TASK dropAllTables: 失敗しました/.test(l)), 'エディタの実行ログに出る');
}

// ==== 3. メンバー・役割（画面の入口と同じ中の関数・同じ監査） ====
let memberRole;
{
  const { full } = ok(env, { action: 'saveMember', email: MEMBER, displayName: 'メンバー', department: '営業部' });
  assert.equal(full.auditAction, 'MEMBER.SAVE');
  assert.equal(full.result.member.email, MEMBER);
  assert.equal(env.table('MEMBERS').filter((m) => m.email === MEMBER)[0].department, '営業部');
  const a = audited(env, 'MEMBER.SAVE');
  assert.equal(a.end.entity_id, MEMBER);
  assert.equal(a.start.entity_type, 'MEMBER');

  ok(env, { action: 'grantRole', email: MEMBER, role: 'ADMIN', scopeType: 'ALL' });
  audited(env, 'ROLE.GRANT');
  assert.equal(env.table('ROLES').filter((r) => r.email === MEMBER && r.role === 'ADMIN').length, 1);
  // 社外のメールは中の関数が断る（入口と同じ判定。記録は FAILED）
  fails(env, { action: 'saveMember', email: OUTSIDER, displayName: 'X' }, /社内/);
  assert.equal(audits(env, 'MEMBER.SAVE', 'END').slice(-1)[0].result, 'FAILED');

  // 一覧: OWNER_TASK_RESULT には次の操作にそのまま写せる短い形、返り値と実行ログには全部
  const dir = ok(env, { action: 'listDirectory' });
  assert.equal(dir.rec.summarized, true);
  memberRole = dir.rec.result.roles.filter((r) => r.email === MEMBER)[0];
  assert.ok(memberRole.roleId && memberRole.rowVersion === 1, JSON.stringify(memberRole));
  assert.ok(dir.rec.result.members.some((m) => m.email === MEMBER && m.displayName === 'メンバー'));
  assert.ok(dir.full.result.members.some((m) => m.email === MEMBER && m.display_name === 'メンバー'), '全部の形');
  assert.equal(audits(env, 'DIRECTORY.LIST').length, 0, '読むだけの操作は記録しない（入口と同じ）');
  assert.ok(env.logs.some((l) => /^結果/.test(l) && l.includes('"display_name": "メンバー"')), '全部は実行ログに出る');
}

// ==== 4. 所有者だけ: 管理者の役割があっても、ほかの人は断る（何もせず、結果も書かず、断ったことを記録する） ====
{
  env.call('apiListDirectory()');   // 準備: MEMBER は全体の管理者
  const settings0 = JSON.stringify(env.table('SETTINGS'));
  const prevResult = env.props.OWNER_TASK_RESULT;
  const raw = JSON.stringify({ action: 'saveSetting', key: 'source.zac_spreadsheet', value: 'https://docs.google.com/spreadsheets/d/' + 'x'.repeat(30) + '/edit' });
  for (const who of [MEMBER, OUTSIDER, '']) {
    env.as(who);
    env.props.OWNER_TASK = raw;
    assert.throws(() => env.call('apiOwnerTask()'), /権限がありません/, who);
    assert.equal(env.props.OWNER_TASK, raw, '頼みはそのまま');
    assert.equal(env.props.OWNER_TASK_RESULT, prevResult, '結果も書かない');
  }
  env.as(MEMBER);
  assert.ok(env.call('apiListDirectory()').members.length >= 2, 'MEMBER は管理者（画面の入口は使える）');
  env.as(OWNER);
  assert.equal(JSON.stringify(env.table('SETTINGS')), settings0, '何も書かない');
  const denied = env.audit().filter((a) => a.phase === 'DENIED' && a.action === 'OWNER.TASK');
  assert.ok(denied.some((a) => a.actor_email === MEMBER), '断ったことは記録に残る');
  delete env.props.OWNER_TASK;
}

// ==== 5. 役割を外す（rowVersion つき）。外した役割をもう一度外すと止まる ====
{
  ok(env, { action: 'revokeRole', roleId: memberRole.roleId, rowVersion: memberRole.rowVersion });
  audited(env, 'ROLE.REVOKE');
  const row = env.table('ROLES').filter((r) => r.role_id === memberRole.roleId)[0];
  assert.equal(row.is_active, 'FALSE');
  fails(env, { action: 'revokeRole', roleId: memberRole.roleId, rowVersion: 2 }, /すでに外れています/);
  env.as(MEMBER);
  assert.throws(() => env.call('apiListDirectory()'), /権限がありません/, '外した役割はすぐ効く');
  env.as(OWNER);
}

// ==== 6. 業務の設定 ====
const ext = Array.from({ length: 70 }, () => '');
const rec = (client, month, amount) => { const r = ext.slice(); r[40] = client; r[45] = 'ベース'; r[49] = '製品A'; r[56] = month; r[65] = amount; return r; };
const zac = env.makeBook('Veeva 売上分析ツール', { ['*' + past + '_actual_value']: { cols: 70, values: [ext.map((_, i) => 'c' + (i + 1)),
  rec('テスト製薬', D(past, 5, 15), 1200000), rec('ｻﾝﾌﾟﾙ製薬(株)', D(past, 6, 10), 300000)] } });
{
  fails(env, { action: 'saveSetting', key: 'source.zac_spreadsheet', value: 'not a url' }, /URL で入力/);
  assert.equal(audits(env, 'SETTING.SAVE', 'END').slice(-1)[0].result, 'FAILED', '中の関数が断っても記録は残る');
  fails(env, { action: 'saveSetting', key: 'no.such.key', value: '1' }, /未定義の設定/);
  ok(env, { action: 'saveSetting', key: 'source.zac_spreadsheet', value: 'https://docs.google.com/spreadsheets/d/' + zac.getId() + '/edit' });
  const a = audited(env, 'SETTING.SAVE');
  assert.equal(a.end.entity_id, 'source.zac_spreadsheet');
  const s = ok(env, { action: 'listSettings' }).rec.result.settings.filter((x) => x.key === 'source.zac_spreadsheet')[0];
  assert.equal(s.value, zac.getId());
  assert.equal(s.display, 'Veeva 売上分析ツール');
  assert.equal(s.isDefault, false);
}

// ==== 7. 計画: 候補 → 作る（裏の処理）→ 一覧 → 担当者（裏の処理）→ メーカーの表示名 ====
let planId;
{
  const cand = ok(env, { action: 'planCandidates' });
  assert.ok(cand.rec.result.candidates.includes('テスト製薬'), JSON.stringify(cand.rec.result));
  assert.ok(Number.isInteger(cand.rec.result.defaultFy));
  assert.ok(cand.full.result.candidates.some((c) => c.zac === 'テスト製薬'), '全部の形は表示名つき');

  const started = ok(env, { action: 'createPlan', clientName: 'テスト製薬', fy: currentFy, peopleCsv: '鷹野, 佐藤' });
  assert.equal(started.full.auditAction, 'JOB.START');
  assert.match(started.full.result.jobId, /^JOB-/);
  assert.match(started.rec.hint, /jobStatus/);
  const st = finishJob(env);   // jobId を省くと、自分が最後に始めた処理
  assert.equal(st.status, 'DONE', st.error);
  planId = st.result.planId;
  const plan = env.table('PLANS').filter((p) => p.plan_id === planId)[0];
  assert.equal(plan.fy, String(currentFy));
  assert.equal(plan.people_csv, '鷹野,佐藤');
  assert.equal(audits(env, 'PLAN.CREATE.SAVE', 'END').slice(-1)[0].actor_email, OWNER, '裏の処理が頼んだ人の名前で記録する');
  assert.equal(audits(env, 'PLAN.CREATE.SAVE', 'END').slice(-1)[0].result, 'OK');
  // 同じメーカー × 年度は作らない（中の関数の判定のまま）
  ok(env, { action: 'createPlan', clientName: 'テスト製薬', fy: currentFy, peopleCsv: '鷹野' });
  assert.match(finishJob(env).error, /すでにあります/);

  assert.ok(ok(env, { action: 'listPlans' }).rec.result.plans.some((p) => p.planId === planId));

  ok(env, { action: 'setPeople', planId, peopleCsv: '鷹野、田中' });
  const sp = finishJob(env);
  assert.equal(sp.status, 'DONE', sp.error);
  assert.equal(env.table('PLANS').filter((p) => p.plan_id === planId)[0].people_csv, '鷹野,田中');
  assert.equal(audits(env, 'PLAN.SETUP.PEOPLE', 'END').slice(-1)[0].result, 'OK');
  // 画面を開いた後に変わった（指紋が違う）なら書かない（入口と同じ）
  ok(env, { action: 'setPeople', planId, peopleCsv: '佐藤', inputHash: 'f'.repeat(64) });
  assert.match(finishJob(env).error, /データが変わりました/);
  assert.equal(env.table('PLANS').filter((p) => p.plan_id === planId)[0].people_csv, '鷹野,田中');
  // 計画が無ければ、始める前に止まる（apiStartJob と同じ判定）
  fails(env, { action: 'setPeople', planId: 'PL-none', peopleCsv: '鷹野' }, /計画が見つかりません/);

  const dir = ok(env, { action: 'listDirectory' }).rec.result;
  const client = dir.clients.filter((c) => c.zacName === 'テスト製薬')[0];
  assert.ok(client && client.auto === true);
  ok(env, { action: 'saveClientName', clientId: client.clientId, displayName: 'テスト', rowVersion: client.rowVersion });
  audited(env, 'CLIENT.NAME');
  const after = ok(env, { action: 'listDirectory' }).rec.result.clients.filter((c) => c.clientId === client.clientId)[0];
  assert.equal(after.displayName, 'テスト');
  assert.equal(after.auto, false);
  assert.equal(after.rowVersion, 1);
}

// ==== 8. バックアップ・手入れ・監査の鎖・状態（要約） ====
{
  ok(env, { action: 'enableBackup' });
  audited(env, 'BACKUP.ENABLE');
  assert.equal(env.triggers.filter((t) => t.handler === 'triggerDailyBackup').length, 1);
  ok(env, { action: 'enableBackup' });
  assert.equal(env.triggers.filter((t) => t.handler === 'triggerDailyBackup').length, 1, '2 度目は何もしない');
  ok(env, { action: 'runBackup' });
  audited(env, 'BACKUP.RUN');
  assert.equal(env.backups().length, 1);
  const hk = ok(env, { action: 'runHousekeeping' });
  assert.equal(hk.full.result.ok, true, JSON.stringify(hk.full.result.problems));
  audited(env, 'HOUSEKEEPING.RUN');
  const v = ok(env, { action: 'verifyAudit' });
  assert.equal(v.full.result.ok, true, JSON.stringify(v.full.result.months));
  audited(env, 'AUDIT.VERIFY');

  const h = ok(env, { action: 'health' });
  assert.equal(h.rec.summarized, true, '状態は要約して置く');
  assert.equal(h.rec.result.setUp, true);
  assert.equal(h.rec.result.ok, true, JSON.stringify(h.rec.result));
  assert.ok(h.rec.result.tables.count > 10);
  assert.equal(h.rec.result.tables.ok, h.rec.result.tables.count);
  assert.deepEqual(h.rec.result.tables.problems, []);
  assert.equal(h.rec.result.audit.ok, true);
  assert.equal(h.rec.result.backup.enabled, true);
  assert.equal(h.rec.result.triggers.triggerDailyBackup, 1);
  assert.equal(h.rec.result.housekeeping.ok, true);
  assert.ok(Array.isArray(h.full.result.tables), '全部の形（返り値・実行ログ）は表ごと');
  assert.equal(audits(env, 'HEALTH.CHECK').length, 0, '読むだけの操作は記録しない');
}

// ==== 9. 事前分布: 見比べる → 書く（裏の処理。作れる情報源が無ければ中の関数が止める） ====
{
  const pv = ok(env, { action: 'poolPreview' }).rec.result;
  assert.equal(pv.ready, false);
  assert.ok(pv.types.length > 0);
  ok(env, { action: 'poolApply' });
  const st = finishJob(env);
  assert.equal(st.status, 'FAILED');
  assert.match(st.error, /事前分布を作れる情報源がまだありません/);
  assert.equal(audits(env, 'LEARN.POOL', 'END').slice(-1)[0].result, 'FAILED');
}

// ==== 10. 年度の締め: 確認の指紋（inputHash）が要る。違う指紋・指紋なし・今年度は締めない ====
{
  const planBook = (client, fy) => {
    const output = [['FY' + fy + ' 売上予測（' + client + '）']];
    for (let r = 2; r <= 25; r++) output.push([]);
    output.push(['年度合計（予測）', 900, 1000, 1100]);
    output.push([], []);
    for (let i = 0; i < 12; i++) output.push([D(fy, 4 + i, 1), 70, 80, 90, '', '', '', 80, '']);
    return env.makeBook('予測 ' + client, { CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', client], ['[必須] 予測年度FY（YYYY）', fy], ['[必須] 担当者（カンマ区切り）', '鷹野']] }, OUTPUT: { values: output } });
  };
  for (const c of ['過去製薬', '過去薬品']) {
    const id = env.seedPlan(planBook(c, past));
    const sub = env.call('apiVersionSubmit(__in)', { __in: { planId: id } });
    env.call('apiVersionDecide(__in)', { __in: { versionId: sub.version.versionId, decision: 'APPROVED', rowVersion: sub.version.rowVersion } });
  }
  const pv = ok(env, { action: 'yearPreview', fy: past });
  assert.equal(pv.rec.result.canClose, true, JSON.stringify(pv.rec.result));
  assert.match(pv.rec.result.inputHash, /^[a-f0-9]{64}$/, '指紋は OWNER_TASK_RESULT から写せる');
  assert.equal(pv.rec.result.planCount, 2);
  assert.equal(audits(env, 'YEAR.PREVIEW').length, 0, '確認は何も書かない・記録しない（入口と同じ）');
  assert.equal(ok(env, { action: 'yearPreview', fy: currentFy }).rec.result.canClose, false);

  const jobs0 = jobProps(env);
  fails(env, { action: 'yearClose', fy: past }, /「内容を確認」/);
  fails(env, { action: 'yearClose', fy: past, inputHash: 'abc' }, /「内容を確認」/);
  fails(env, { action: 'yearClose', fy: currentFy, inputHash: pv.rec.result.inputHash }, /今年度/);
  assert.equal(jobProps(env), jobs0, '始める前に止まる（処理を入れない）');
  // 形は正しいが確認と違う指紋: 裏の処理が今の控えと比べて止める（何も凍結しない）
  ok(env, { action: 'yearClose', fy: past, inputHash: '0'.repeat(64) });
  const bad = finishJob(env);
  assert.equal(bad.status, 'FAILED');
  assert.match(bad.error, /確認したあとに年度の内容が変わりました/);
  assert.equal(env.table('YEAR_CLOSURES').length, 0);
  assert.equal(audits(env, 'YEAR.CLOSE', 'END').slice(-1)[0].result, 'FAILED');
  // 確認した指紋なら締まる
  ok(env, { action: 'yearClose', fy: past, inputHash: pv.rec.result.inputHash });
  const done = finishJob(env);
  assert.equal(done.status, 'DONE', done.error);
  assert.equal(done.result.state, 'CLOSED');
  assert.equal(done.result.planCount, 2);
  const closures = env.table('YEAR_CLOSURES');
  assert.equal(closures.length, 1);
  assert.equal(closures[0].closed_by, OWNER);
  const end = audits(env, 'YEAR.CLOSE', 'END').slice(-1)[0];
  assert.equal(end.result, 'OK');
  assert.equal(end.actor_email, OWNER);
  assert.equal(ok(env, { action: 'health' }).rec.result.years.years.filter((y) => y.fy === past)[0].state, 'CLOSED');
  assert.equal(ok(env, { action: 'yearPreview', fy: past }).rec.result.canClose, false, '締めた年度はもう締めない');
}

// ==== 11. ログ: 種類・月・絞り込み。大きい結果は OWNER_TASK_RESULT に要約、全部は実行ログに分けて出す ====
{
  for (let i = 0; i < 30; i++) {
    ok(env, { action: 'saveMember', email: MEMBER, displayName: 'メンバー' + i, department: '営業部 第' + i + '課（ログを大きくするための長い部署の名前）' });
  }
  const before = env.logs.length;
  const la = ok(env, { action: 'listAudit', kind: 'AUDIT' });
  const total = la.full.result.rows.length;
  assert.ok(total > 60, '記録の行: ' + total);
  assert.equal(la.rec.summarized, true);
  assert.equal(la.rec.result.count, total, '件数は全部の数');
  assert.ok(la.rec.result.rows.length < total, '置く行は先頭（新しい方）だけ');
  assert.match(String(la.rec.result.rows[la.rec.result.rows.length - 1]), /ほか \d+ 件/);
  assert.deepEqual(Object.keys(la.rec.result.rows[0]), ['at', 'actor', 'action', 'phase', 'result', 'entity', 'error'], '要る列だけ');
  assert.match(la.rec.note, /全部は実行ログ/);
  // 実行ログを分けて出した全部をつなげると、元の結果になる
  const parts = env.logs.slice(before).filter((l) => /^結果（\d+\/\d+）: /.test(l));
  assert.ok(parts.length >= 2, '長い結果は分けて出す: ' + parts.length);
  const joined = JSON.parse(parts.map((l) => l.replace(/^結果（\d+\/\d+）: /, '')).join(''));
  assert.equal(joined.result.rows.length, total);
  assert.equal(joined.action, 'listAudit');

  const runs = ok(env, { action: 'listAudit', kind: 'RUN' }).rec.result;
  assert.equal(runs.kind, 'RUN');
  assert.ok(runs.rows.some((r) => r.kind === 'OWNER.TASK' && r.status === 'OK'), '操作ごとに実行ログ（RUN）に残る');
  const errs = ok(env, { action: 'listAudit', kind: 'ERROR' }).rec.result;
  assert.ok(errs.rows.some((r) => r.where === 'OWNER.TASK'));
  const q = ok(env, { action: 'listAudit', kind: 'AUDIT', query: 'ROLE.REVOKE' }).full.result;
  assert.ok(q.rows.length >= 2 && q.rows.every((r) => JSON.stringify(r).includes('ROLE.REVOKE')));
  // 月の指定（無い月は空）
  assert.equal(ok(env, { action: 'listAudit', kind: 'AUDIT', month: '2001_01' }).rec.result.count, 0);
}

// ==== 12. 結果を縮める仕組み: どんな大きさでも 1 つの値の上限に収まる ====
{
  const big = env.run(`(() => {
    const rows = []; for (let i = 0; i < 3000; i++) rows.push({ id: 'R' + i, text: '長い文字'.repeat(200), nested: { list: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], s: 'x'.repeat(1000) } });
    const rec = appOwnerResultRecord_({ ok: true, action: 'listAudit', at: 'now', requestId: 'r', result: { rows: rows } }, {});
    return { bytes: appUtf8Bytes_(JSON.stringify(rec)), summarized: rec.summarized, n: rec.result.rows.length };
  })()`);
  assert.ok(big.bytes <= RESULT_MAX, String(big.bytes));
  assert.equal(big.summarized, true);
  assert.ok(big.n < 3000);
  const wide = env.run(`(() => { const o = {}; for (let i = 0; i < 5000; i++) o['k' + i] = '値'.repeat(50);
    const rec = appOwnerResultRecord_({ ok: true, action: 'x', result: o }, {}); return appUtf8Bytes_(JSON.stringify(rec)); })()`);
  assert.ok(wide <= RESULT_MAX, String(wide));
  const small = env.run(`appOwnerResultRecord_({ ok: true, action: 'x', result: { a: 1 } }, {})`);
  assert.equal(small.summarized, undefined, '小さい結果はそのまま');
}

// ==== 13. 操作ごとの役割の判定は弱めない（役割が足りなければ中の関数を動かさず、断ったことを記録する） ====
{
  const ran = env.call(`(() => { let ran = false;
    try { appOwnerRun_({ roles: [{ role: 'PLANNER', scope_type: 'ALL', client_id: '' }], actor: 'x@bigm2y.com', requestId: 'REQ-DENY' }, 'SETTING.SAVE',
      { minRole: 'ADMIN', entityType: 'SETTING', entityId: 'k' }, () => { ran = true; }); } catch (e) { return { ran: ran, error: e.message }; }
    return { ran: ran, error: '' }; })()`);
  assert.deepEqual(ran, { ran: false, error: 'この操作をする権限がありません。' });
  assert.ok(env.audit().some((a) => a.phase === 'DENIED' && a.action === 'SETTING.SAVE' && a.request_id === 'REQ-DENY'));
}

console.log('app-owner-task: all tests passed');
