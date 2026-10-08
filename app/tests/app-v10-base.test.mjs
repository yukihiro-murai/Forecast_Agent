/**
 * app-v10-base.test.mjs — 表の版 10 の土台（SCHEMA_PLAN_v10-12_JA.md の 3 章・7 章）。
 * 表と列の定義・一度だけの写しの動かし方・状態の点検のセルの数と年度の控えの目安・記録を見る人の範囲（人のつなぎ）・
 * 測る専用の計画と年度を締める条件。
 *
 *   node app/tests/app-v10-base.test.mjs
 */
import assert from 'node:assert/strict';
import { setUpEnv, STATS, OWNER, MEMBER, OTHER, sha } from './gas-mock.mjs';

const D = (y, m, d = 1) => new Date(y, m - 1, d);
function planBook(client, fy) {
  const output = [['FY' + fy + ' 売上予測（' + client + '）']];
  for (let r = 2; r <= 25; r++) output.push([]);
  output.push(['年度合計（予測）', 900, 1000, 1100]);
  output.push([], []);
  for (let i = 0; i < 12; i++) output.push([D(fy, 4 + i, 1), 70, 80, 90, '', '', '', 80, '']);
  return { CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', client], ['[必須] 予測年度FY（YYYY）', fy], ['[必須] 担当者（カンマ区切り）', '担当A']] }, OUTPUT: { values: output } };
}

// ==== 1. 表と列（3-1〜3-9 のとおり） ====
{
  const env = setUpEnv();
  const T = env.call('APP_TABLES');
  assert.equal(env.call('APP_SCHEMA_VERSION'), 10);
  const want = {
    INPUT_LOG: [['log_id'], ['plan_id', 'log_id', 'action_id', 'action', 'kind', 'change', 'row_key', 'person', 'before_json', 'after_json', 'self_conf', 'reason', 'signal_id', 'actor_email', 'saved_at']],
    AI_RESEARCH_LOG: [['research_id'], ['plan_id', 'research_id', 'action_id', 'started_by', 'as_of_date', 'topic', 'row_type', 'direction', 'impact_score', 'confidence', 'event_score', 'benchmark_score', 'blended_score', 'time_horizon', 'row_json', 'recorded_at']],
    LAYER_EFFECTS: [['effect_id'], ['plan_id', 'effect_id', 'run_id', 'ym', 'final_p50', 'stat_p50', 'spot_yen', 'human_yen', 'human_by_type_json', 'ai_yen', 'calib_yen', 'other_yen', 'method', 'source', 'recorded_at']],
    HIT_RECORDS: [['hit_id'], ['plan_id', 'hit_id', 'client_id', 'quarter', 'source_kind', 'source_key', 'person_email', 'months_json', 'push', 'actual_dir', 'hit', 'n_months', 'policy_version', 'calc_version', 'computed_at', 'computed_by']],
    LEARNING_LOG: [['learn_id'], ['plan_id', 'learn_id', 'proposal_id', 'event', 'origin', 'target', 'current_value', 'proposed_value', 'evidence_n', 'ci80_json', 'compare_json', 'decision', 'applied_value', 'review_quarter', 'note', 'actor_email', 'at']],
    PERSON_LINKS: [['link_id'], ['link_id', 'person_name', 'email', 'client_id', 'valid_from', 'valid_to', 'is_active', 'note', 'created_at', 'created_by', 'updated_at', 'updated_by', 'row_version']],
    BACKTEST: [['point_id'], ['plan_id', 'point_id', 'bt_id', 'cutoff_ym', 'target_ym', 'horizon', 'method', 'p10', 'p50', 'p90', 'actual', 'real_months', 'counted', 'engine_sha256', 'seed', 'calc_version', 'computed_at', 'computed_by']],
  };
  for (const [name, [key, columns]] of Object.entries(want)) {
    assert.deepEqual(T[name].key, key, name + ' のキー');
    assert.deepEqual(T[name].columns, columns, name + ' の列');
    assert.equal(!!T[name].appendOnly, name !== 'PERSON_LINKS', name + ' は追記だけ（つなぎは無効の印を変える）');
    assert.deepEqual([...env.data().getSheetByName(name).rows[0]], columns);
  }
  assert.equal(T.LEARNING_LOG.globalRows, true);
  const type = (name, col) => env.call('appColumnType_(__c, APP_TABLES[__n])', { __c: col, __n: name });
  for (const c of ['impact_score', 'confidence', 'event_score', 'benchmark_score', 'blended_score']) assert.equal(type('AI_RESEARCH_LOG', c), 'num');
  for (const c of ['final_p50', 'stat_p50', 'spot_yen', 'human_yen', 'ai_yen', 'calib_yen', 'other_yen']) assert.equal(type('LAYER_EFFECTS', c), 'num');
  assert.deepEqual(['push', 'actual_dir', 'hit', 'n_months'].map((c) => type('HIT_RECORDS', c)), ['num', 'int', 'int', 'int']);
  assert.deepEqual(['horizon', 'p10', 'p50', 'p90', 'actual', 'real_months', 'counted', 'seed'].map((c) => type('BACKTEST', c)), ['int', 'num', 'num', 'num', 'num', 'int', 'bool', 'text']);
  assert.equal(type('PERSON_LINKS', 'is_active'), 'bool');
  // 足した列: 定義の終わりと同じ並び。前の版の列はその前まで
  const added = env.call('APP_ADDED_COLUMNS');
  assert.deepEqual(Object.keys(added).sort(), ['FORECAST_RUNS', 'PLANS']);
  assert.deepEqual(T.FORECAST_RUNS.columns.slice(-3), ['app_version', 'seed_rule', 'fixes_json']);
  assert.deepEqual(T.PLANS.columns.slice(-1), ['purpose']);
  for (const [name, list] of Object.entries(added)) {
    let end = T[name].columns.length;
    for (let i = list.length - 1; i >= 0; i--) {
      assert.deepEqual(T[name].columns.slice(end - list[i].columns.length, end), list[i].columns, name + ' の足した列は定義の終わり');
      end -= list[i].columns.length;
    }
    assert.equal(env.call('appTableStages_(__n)', { __n: name }).length, list.length + 1);
  }
  assert.equal(env.call('appColumnsHash_("PLANS")'), sha(T.PLANS.columns.join('|')));
}

// ==== 2. 一度だけの写し: 済んだら記録して二度と動かさない・仮のものは記録しない・失敗は 10 分あけてやり直す ====
{
  const env = setUpEnv();
  assert.deepEqual(env.call('APP_V10_BACKFILLS'), ['appV10BackfillInputLog_', 'appV10BackfillAiResearchLog_', 'appV10BackfillHitRecords_']);
  for (const n of env.call('APP_V10_BACKFILLS').filter((n) => n !== 'appV10BackfillInputLog_')) assert.deepEqual(env.call(n + '({})'), { skipped: true }, n + ' は今は仮のもの');
  env.run('appV10BackfillInputLog_ = function(ctx) { return { skipped: true }; };');   // 入力の記録の写しの中身は app-v10-input で確かめる
  env.call('apiListPlans()');
  assert.equal(env.props.APP_BACKFILLS, undefined, '仮のものは記録しない');
  // 何もすることが無い操作ではロックを取らず、表も読まない
  env.call('apiListPlans()');
  const locks = env.state.locks;
  for (const k of Object.keys(STATS)) STATS[k] = k === 'bySheet' ? {} : 0;
  env.call('apiListPlans()');
  assert.equal(env.state.locks, locks, 'ロックを取らない');
  assert.equal(STATS.reads, 0, '表を読まない（一覧は控えから）');
  // 済む写し・続きがある写し・投げる写し
  env.run(`__calls = { a: 0, b: 0, c: 0 };
    appV10BackfillInputLog_ = function(ctx) { __calls.a++; return { rows: 4 }; };
    appV10BackfillAiResearchLog_ = function(ctx) { __calls.b++; return __calls.b < 3 ? { more: true, rows: 1 } : { rows: 1 }; };
    appV10BackfillHitRecords_ = function(ctx) { __calls.c++; throw new Error('まだ読めない'); };`);
  env.call('apiListPlans()');
  let st = JSON.parse(env.props.APP_BACKFILLS);
  assert.deepEqual(Object.keys(st.done), ['appV10BackfillInputLog_']);
  assert.equal(st.done.appV10BackfillInputLog_.rows, 4);
  assert.equal(st.failed.appV10BackfillHitRecords_.tries, 1);
  assert.match(st.failed.appV10BackfillHitRecords_.error, /まだ読めない/);
  assert.ok(env.errors().some((e) => e.where === 'SCHEMA.BACKFILL'), '失敗はエラーのログに残す（操作は止めない）');
  assert.ok(env.runLog().some((r) => r.kind === 'SCHEMA.BACKFILL' && r.status === 'OK' && JSON.parse(r.detail_json).name === 'appV10BackfillInputLog_'));
  env.call('apiListPlans()');
  env.call('apiListPlans()');
  assert.deepEqual(env.call('__calls'), { a: 1, b: 3, c: 1 }, '済んだものは動かさない・続きは次の操作で・失敗は 10 分あける');
  st = JSON.parse(env.props.APP_BACKFILLS);
  assert.deepEqual(Object.keys(st.done).sort(), ['appV10BackfillAiResearchLog_', 'appV10BackfillInputLog_']);
  st.failed.appV10BackfillHitRecords_.at = Date.now() - 11 * 60 * 1000;   // 10 分たった
  env.props.APP_BACKFILLS = JSON.stringify(st);
  env.run('appV10BackfillHitRecords_ = function(ctx) { __calls.c++; return { rows: 0 }; };');
  env.call('apiListPlans()');
  st = JSON.parse(env.props.APP_BACKFILLS);
  assert.equal(Object.keys(st.done).length, 3);
  assert.deepEqual(st.failed, {}, 'やり直して済んだら失敗の控えを消す');
  const h = env.call('apiHealth()');
  assert.deepEqual([h.backfills.pending, h.backfills.done.length, h.backfills.failed], [[], 3, []]);
  // 表の版がそろっていなければ動かさない（移行の後に動かす）
  env.props.APP_BACKFILLS = '';
  env.props.APP_TABLES_VERSION = '9';
  env.data().getSheetByName('PLANS').rows[0][0] = 'broken';   // 版をそろえられない
  env.call('apiBootstrap()');
  assert.equal(env.props.APP_BACKFILLS, '', '版がそろう前は写さない');
}

// ==== 3. 状態の点検: セルの数と上限に対する割合・年度の控えの目安 ====
{
  const env = setUpEnv();
  const fy = env.run('appFy_(new Date())');
  const planId = env.seedPlan(env.makeBook('A', planBook('テスト製薬', fy)));
  const h = env.call('apiHealth()');
  const cells = env.data().sheets.reduce((n, s) => n + s.getMaxRows() * s.getMaxColumns(), 0);
  assert.equal(h.size.cells, cells, '全部のシートの 行 × 列');
  assert.equal(h.size.limit, 10000000);
  assert.equal(h.size.ratio, Math.round(cells / 1e7 * 1000) / 1000);
  assert.ok(h.size.usedCells > 0 && h.size.usedCells <= cells);
  assert.ok(h.size.sheets.length <= 8 && h.size.sheets[0].cells >= h.size.sheets[h.size.sheets.length - 1].cells);
  const y = h.size.years.find((x) => x.fy === fy);
  assert.equal(y.plans.length, 1);
  assert.equal(y.plans[0].planId, planId);
  const snap = env.call('appYearSnapshot_(__f)', { __f: fy });
  assert.ok(y.bytes > 0 && y.bytes <= snap.bytes, `目安は実際の控えより少し小さい（${y.bytes} / ${snap.bytes}）`);
  assert.equal(y.ratio, Math.round(y.bytes / (8 * 1024 * 1024) * 1000) / 1000);
  // 所有者の要約: 数と、上限に近いときの要確認
  const brief = env.call('appOwnerHealthBrief_(apiHealth())');
  assert.equal(brief.size.cells, cells);
  assert.equal(brief.size.years.find((x) => x.fy === fy).plans, 1);
  assert.ok(!brief.warnings.some((w) => /セル|年度の控え/.test(w)));
  env.run('appCellLimit_ = function() { return 1000; };');
  assert.ok(env.call('appOwnerHealthBrief_(apiHealth())').warnings.some((w) => /データ本体のセルが上限の \d+% です/.test(w)), '半分を超えたら要確認');
  const fake = env.call(`appOwnerHealthBrief_({ setUp: true, tables: [], audit: { ok: true }, backup: { enabled: true, ageHours: 1 },
    size: { cells: 1, ratio: 0.1, years: [{ fy: 2025, plans: [{}], bytes: 7000000, ratio: 0.835 }] }, backfills: { failed: [{ name: 'x' }] } })`);
  assert.ok(fake.warnings.some((w) => /FY2025 の年度の控えが上限（8MB）の約 84% です/.test(w)));
  assert.ok(fake.warnings.some((w) => /一度だけの写しに失敗があります/.test(w)));
}

// ==== 4. 記録を見る人の範囲（予算策定担当以上は全部。ほかの人は、つなぎで本人と分かる自分の行と、種類ごとの件数だけ） ====
{
  const env = setUpEnv();
  env.seedPlan(env.makeBook('X', planBook('テスト製薬', env.run('appFy_(new Date())'))));
  const X = env.table('CLIENTS')[0].client_id;   // メーカー X（計画のあるメーカー）。Y はつなぎだけに出る別のメーカー
  env.call('apiSaveMember(__in)', { __in: { email: MEMBER, displayName: 'M', department: '営業' } });
  env.call('apiSaveMember(__in)', { __in: { email: OTHER, displayName: 'O', department: '営業' } });
  const link = (id, name, email, clientId, extra) => Object.assign({ link_id: id, person_name: name, email, client_id: clientId || '', valid_from: '', valid_to: '',
    is_active: true, note: '', created_at: 'now', created_by: OWNER, updated_at: 'now', updated_by: OWNER, row_version: 1 }, extra || {});
  env.call('appWithLock_(() => appInsertRows_("PERSON_LINKS", __r))', { __r: [
    link('PLK-1', '担当A', MEMBER, ''),                     // 全部のメーカーで 担当A = MEMBER
    link('PLK-2', '担当A', OTHER, 'CL-Y'),                 // メーカー Y では 担当A = OTHER（メーカーで分かれる）
    link('PLK-3', '担当B', MEMBER, '', { is_active: false }),   // 外したつなぎ（行は残る）
    link('PLK-4', '担当C', MEMBER, '', { valid_to: '2000-01-01' }),   // 期間の外
    link('PLK-5', '担当D', MEMBER, ''), link('PLK-6', '担当D', OTHER, ''),   // 同じ名前が 2 人: 決めない
  ] });
  const emailOf = (n, c) => env.call('appPersonEmailOf_(__n, __c)', { __n: n, __c: c });
  assert.deepEqual([emailOf('担当A', X), emailOf('担当A', 'CL-Y'), emailOf(' 担当A ', ''), emailOf('担当B', X), emailOf('担当C', X), emailOf('担当D', X), emailOf('誰か', X)],
    [MEMBER, OTHER, MEMBER, '', '', '', '']);
  assert.equal(env.call('appPersonEmailOf_("担当C", __c, "1999-06-01")', { __c: X }), MEMBER, '期間の中の日なら効く');
  const viewer = (email, clientId) => { env.as(email); const v = env.call('(() => { const u = appCurrentUser_(); return appLogViewer_({ user: u, roles: appRolesOf_(u) }, __c); })()', { __c: clientId }); env.as(OWNER); return v; };
  assert.deepEqual(viewer(OWNER, X), { full: true, email: OWNER, clientId: X, names: [] }, '管理者は全部');
  assert.deepEqual(viewer(MEMBER, X), { full: false, email: MEMBER, clientId: X, names: ['担当A'] }, '閲覧の人は自分につながる名前だけ');
  assert.deepEqual(viewer(MEMBER, 'CL-Y').names, [], 'メーカー Y の 担当A はほかの人');
  assert.deepEqual(viewer(OTHER, 'CL-Y').names, ['担当A']);
  env.call('apiGrantRole(__in)', { __in: { email: MEMBER, role: 'PLANNER', scopeType: 'CLIENT', clientId: X } });
  assert.equal(viewer(MEMBER, X).full, true, 'そのメーカーの予算策定担当は全部');
  assert.equal(viewer(MEMBER, 'CL-Y').full, false, 'ほかのメーカーでは自分の行だけ');
  const rows = [{ person: '担当A', kind: 'product' }, { person: '担当B', kind: 'product' }, { person: '担当B', kind: 'client' }, { person_email: OTHER, kind: 'hit' }];
  const vis = (v) => env.call('appLogVisible_(__v, __r, { personOf: r => r.person, emailOf: r => r.person_email, typeOf: r => r.kind })', { __v: v, __r: rows });
  assert.deepEqual(vis({ full: true }), { rows, counts: {}, hidden: 0 });
  assert.deepEqual(vis({ full: false, email: MEMBER, names: ['担当A'] }), { rows: [rows[0]], counts: { product: 1, client: 1, hit: 1 }, hidden: 3 });
  assert.deepEqual(vis({ full: false, email: OTHER, names: [] }), { rows: [rows[3]], counts: { product: 2, client: 1 }, hidden: 3 }, 'メールで本人と分かる行');
  assert.deepEqual(vis({ full: false, email: '', names: [] }).rows, [], 'つなぎの無い人には件数だけ');
}

// ==== 5. 測る専用の計画は、年度を締める条件（公式版）から外す。最後の四半期の当たりの見張りの入口 ====
{
  const env = setUpEnv();
  const past = env.run('appFy_(new Date())') - 1;
  const planA = env.seedPlan(env.makeBook('A', planBook('テスト製薬', past)));
  const planM = env.seedPlan(env.makeBook('M', planBook('テスト薬品', past)));
  const sub = env.call('apiVersionSubmit(__in)', { __in: { planId: planA } });
  env.call('apiVersionDecide(__in)', { __in: { versionId: sub.version.versionId, decision: 'APPROVED', rowVersion: sub.version.rowVersion } });
  let pv = env.call('apiYearPreview(__in)', { __in: { fy: past } });
  assert.equal(pv.canClose, false);
  assert.match(pv.reason, /公式版のない計画/);
  assert.equal(env.call('appPlanIsMeasure_({ purpose: "MEASURE" })'), true);
  assert.equal(env.call('appPlanIsMeasure_({ purpose: "" })'), false);
  assert.equal(env.call('APP_MEASURE_PLAN_MAX'), 5);
  env.call(`appWithLock_(() => appUpdateByKey_('PLANS', { plan_id: '${planM}' }, { purpose: 'MEASURE' }, undefined, '${OWNER}'))`);
  pv = env.call('apiYearPreview(__in)', { __in: { fy: past } });
  assert.equal(pv.canClose, true, '測る専用の計画は公式版が無くても締められる: ' + pv.reason);
  assert.equal(pv.planCount, 2, '年度の控えには入る');
  assert.equal(env.call('appYearHitsPending_(2025, [])'), '', '今は止めない（当たりの記録を作るときに中身を書く）');
  env.run('appYearHitsPending_ = function(fy, plans) { return "最後の四半期の当たりを数え終えてから締めてください。"; };');
  pv = env.call('apiYearPreview(__in)', { __in: { fy: past } });
  assert.equal(pv.canClose, false);
  assert.match(pv.reason, /最後の四半期の当たり/, '締める前の見張りの入口がつながっている');
}

console.log('app-v10-base: ok');
