/**
 * app-versions.test.mjs — 公式版（確定した予測と予算）と承認。
 */
import assert from 'node:assert/strict';
import { OWNER, setUpEnv } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
const env = setUpEnv();
const output = [['FY2026 売上予測（テスト製薬）']];
for (let r = 2; r <= 25; r++) output.push([]);
output.push(['年度合計（予測）', 900, 1000, 1100]);
output.push([], []);
for (let i = 0; i < 12; i++) output.push([D(2026, 4 + i), 70, 80, 90, '', '', '', 80, i === 11 ? 5 : '', '']);
const book = env.makeBook('クライアント別売上予測', {
  CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', 'テスト製薬'], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
  OUTPUT: { values: output },
});
const url = 'https://docs.google.com/spreadsheets/d/' + book.getId() + '/edit';
const dry = env.runJob('MIGRATION.DRYRUN', { bookUrl: url });
const imp = env.runJob('MIGRATION.IMPORT', { bookUrl: url, contentHash: dry.result.contentHash });
assert.equal(imp.status, 'DONE', imp.error);
const planId = imp.result.planId;
const clientId = env.table('PLANS')[0].client_id;

// 予算策定担当（このクライアント）と承認者を登録する
const PLANNER = 'planner@bigm2y.com', APPROVER = 'approver@bigm2y.com';
for (const [email, role, scope] of [[PLANNER, 'PLANNER', 'CLIENT'], [APPROVER, 'APPROVER', 'ALL']]) {
  env.call('apiSaveMember(__in)', { __in: { email, displayName: email.split('@')[0], department: '営業' } });
  env.call('apiGrantRole(__in)', { __in: { email, role, scopeType: scope, clientId: scope === 'CLIENT' ? clientId : '' } });
}

// ==== 1. 今の数字（年度・月ごと・予算） ====
env.as(PLANNER);
let list = env.call('apiVersionList(__in)', { __in: { planId } });
assert.equal(list.current.annual.p50, 1000);
assert.equal(list.current.budget.adopted, 960);
assert.equal(list.current.budget.uplift, 5);
assert.equal(list.current.budget.final, 965);
assert.equal(list.current.monthly.length, 12);
assert.deepEqual(list.can, { submit: true, approve: false });
assert.equal(list.versions.length, 0);

// ==== 2. 出す → 承認待ち。出した本人は承認できない ====
const sub = env.call('apiVersionSubmit(__in)', { __in: { planId, note: '10 月の見直し', inputHash: list.inputHash } });
assert.equal(sub.version.no, 1);
assert.equal(sub.version.state, 'SUBMITTED');
assert.equal(sub.version.budget.final, 965);
assert.throws(() => env.call('apiVersionDecide(__in)', { __in: { versionId: sub.version.versionId, decision: 'APPROVED' } }), /権限/, '予算策定担当は承認できない');
assert.throws(() => env.call('apiVersionSubmit(__in)', { __in: { planId, note: '=HYPERLINK("x")' } }), /数式として動いてしまう文字/);

// ==== 3. 承認者が承認 → 公式版。却下は理由が要る ====
env.as(APPROVER);
assert.throws(() => env.call('apiVersionDecide(__in)', { __in: { versionId: sub.version.versionId, decision: 'REJECTED' } }), /理由/);
const ok = env.call('apiVersionDecide(__in)', { __in: { versionId: sub.version.versionId, decision: 'APPROVED', note: '確認しました', rowVersion: sub.version.rowVersion } });
assert.equal(ok.version.state, 'APPROVED');
assert.throws(() => env.call('apiVersionDecide(__in)', { __in: { versionId: sub.version.versionId, decision: 'APPROVED' } }), /承認待ちではありません/);

// ==== 4. 次の版: 前の承認待ちは取り下げ、承認すると前の公式版は置き換わる。出した版の数字は変わらない ====
env.as(PLANNER);
const s2 = env.call('apiVersionSubmit(__in)', { __in: { planId, note: '2 回目' } });
const s3 = env.call('apiVersionSubmit(__in)', { __in: { planId, note: '3 回目' } });
assert.deepEqual(s3.withdrawn, [2]);
env.as(APPROVER);
const ok3 = env.call('apiVersionDecide(__in)', { __in: { versionId: s3.version.versionId, decision: 'APPROVED' } });
assert.deepEqual(ok3.superseded, [1]);
list = env.call('apiVersionList(__in)', { __in: { planId } });
assert.deepEqual(list.versions.map((v) => [v.no, v.state]), [[3, 'APPROVED'], [2, 'WITHDRAWN'], [1, 'SUPERSEDED']]);
assert.equal(list.official.no, 3);
assert.equal(list.versions[2].monthly[11].final, 85, '出した版の月の数字が残る');

// ==== 4b. 承認が止められたとき（画面が古い）は、前の公式版をそのまま残す ====
env.as(PLANNER);
const s4 = env.call('apiVersionSubmit(__in)', { __in: { planId, note: '4 回目' } });
env.as(APPROVER);
assert.throws(() => env.call('apiVersionDecide(__in)', { __in: { versionId: s4.version.versionId, decision: 'APPROVED', rowVersion: 99 } }), /先に更新しました/);
assert.equal(env.call('apiVersionList(__in)', { __in: { planId } }).official.no, 3, '公式版は v3 のまま');
env.call('apiVersionDecide(__in)', { __in: { versionId: s4.version.versionId, decision: 'REJECTED', note: '見直し中' } });

// ==== 5. 一覧に公式版が出る。記録に残る ====
const pf = env.call('apiPortfolio()').plans[0];
assert.equal(pf.officialNo, 3);
assert.equal(pf.officialFinal, 965);
assert.ok(env.audit().some((a) => a.action === 'VERSION.DECIDE' && a.phase === 'END' && a.actor_email === APPROVER));

// ==== 6. 社内の人は見られるが、出せない ====
env.as('someone@bigm2y.com');
assert.equal(env.call('apiVersionList(__in)', { __in: { planId } }).can.submit, false);
assert.throws(() => env.call('apiVersionSubmit(__in)', { __in: { planId } }), /権限/);
env.as(OWNER);
// ==== 7. 取り込んだ旧ブックは、アーカイブのフォルダの「旧ブック」へ移せる（所有者だけ。消さない） ====
const books = env.call('apiLegacyBooks()').books;
assert.equal(books.length, 1);
assert.equal(books[0].archived, false);
const ar = env.call('apiArchiveLegacyBooks()');
assert.equal(ar.moved, 1);
assert.equal(env.files[book.getId()].trashed, false, '消さない');
assert.equal(env.call('apiLegacyBooks()').books[0].archived, true);
assert.equal(env.call('apiArchiveLegacyBooks()').moved, 0, '2 回目は何もしない');
env.as(APPROVER);
assert.throws(() => env.call('apiArchiveLegacyBooks()'), /権限/);
env.as(OWNER);
console.log('app-versions: all tests passed');
