#!/usr/bin/env node

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const sourceDir = path.resolve(testDir, '..', 'src');
let uuid = 0;
const cacheValues = new Map();
const sandbox = {
  console,
  Logger: { log() {} },
  SpreadsheetApp: { getActiveSpreadsheet: () => ({}) },
  Session: { getActiveUser: () => ({ getEmail: () => 'creator@example.com' }) },
  CacheService: {
    getDocumentCache: () => ({
      get: key => cacheValues.get(key) || null,
      put: (key, value) => cacheValues.set(key, value),
      remove: key => cacheValues.delete(key)
    })
  },
  Utilities: {
    getUuid: () => `12345678-test-${++uuid}`,
    Charset: { UTF_8: 'UTF_8' },
    DigestAlgorithm: { SHA_256: 'SHA_256' },
    computeDigest(algorithm, value) { return [...createHash('sha256').update(value, 'utf8').digest()]; }
  },
  LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) }
};
vm.createContext(sandbox);
vm.runInContext(await readFile(path.join(sourceDir, 'Portal_Core.js'), 'utf8'), sandbox, { filename: 'Portal_Core.js' });

testCanonicalJson();
testV2AndLegacySchemas();
testMemberAndFiscalYearValidation();
testCatalogCacheAndStrictSelection();
testDuplicateCandidates();
testRequestRowProjection();
testStatusEventValidation();
testAppendRequestContract();
testCreateModel();
testRequestProgress();
testEntryModel();
await testStaticUxContracts();
process.stdout.write('PASS portal runtime behavior tests (12)\n');

function v2Payload(overrides = {}) {
  return {
    catalogKey: 'ZAC-C-001', clientName: 'テスト株式会社', fiscalYear: 2027,
    relatedMemberNames: ['山田 太郎'], requestId: 'PORTAL-REQ-12345678-test',
    requestType: 'CREATE_CLIENT_FY_BOOK', requestedAt: '2026-08-12T00:00:00.000Z',
    requestedBy: 'creator@example.com', schemaVersion: 'vnext-portal-request-2', ...overrides
  };
}

function legacyPayload(overrides = {}) {
  return {
    clientId: 'CLIENT-1', clientName: 'テスト株式会社', fiscalYear: 2027,
    forecastOwnerEmail: 'owner@example.com', relatedMemberEmails: ['member@example.com'],
    requestId: 'PORTAL-REQ-12345678-test', requestType: 'CREATE_CLIENT_FY_BOOK',
    requestedAt: '2026-08-12T00:00:00.000Z', requestedBy: 'creator@example.com',
    schemaVersion: 'vnext-portal-request-1', ...overrides
  };
}

function testCanonicalJson() {
  assert.equal(sandbox.vNextPortalCanonicalJson_({ b: 2, a: 1 }), '{"a":1,"b":2}');
  assert.equal(sandbox.vNextPortalCanonicalJson_({ z: undefined, a: [1, undefined] }), '{"a":[1,null]}');
  assert.equal(sandbox.vNextPortalSha256Hex_('{"a":1}'), createHash('sha256').update('{"a":1}').digest('hex'));
}

function testV2AndLegacySchemas() {
  const current = v2Payload();
  assert.equal(sandbox.vNextPortalValidateRequestPayload_(current), current);
  assert.deepEqual(Object.keys(current).sort(), [...sandbox.VNEXT_PORTAL.PAYLOAD_KEYS].sort());
  assert.throws(() => sandbox.vNextPortalValidateRequestPayload_({ ...current, clientId: 'tampered' }), /項目が契約と一致/);
  assert.throws(() => sandbox.vNextPortalValidateRequestPayload_(v2Payload({ relatedMemberNames: [] })), /1名以上/);

  const legacy = legacyPayload();
  assert.equal(sandbox.vNextPortalValidateRequestPayload_(legacy), legacy);
  const legacyModel = sandbox.vNextPortalRequestPayloadToModel_(legacy);
  assert.equal(legacyModel.clientId, 'CLIENT-1');
  assert.deepEqual([...legacyModel.relatedMemberEmails], ['member@example.com']);
  assert.deepEqual([...legacyModel.relatedMemberNames], []);
  const currentModel = sandbox.vNextPortalRequestPayloadToModel_(current);
  assert.equal(currentModel.clientId, 'ZAC-C-001');
  assert.equal(currentModel.forecastOwnerEmail, 'creator@example.com');
  assert.deepEqual([...currentModel.relatedMemberNames], ['山田 太郎']);
}

function testMemberAndFiscalYearValidation() {
  assert.throws(() => sandbox.vNextPortalNormalizeMemberNames_([], true), /1名以上/);
  assert.deepEqual([...sandbox.vNextPortalNormalizeMemberNames_(['山田 太郎'], true)], ['山田 太郎']);
  assert.deepEqual(
    [...sandbox.vNextPortalNormalizeMemberNames_(['A', 'B', 'C', 'D', 'E'], true)],
    ['A', 'B', 'C', 'D', 'E']
  );
  assert.throws(() => sandbox.vNextPortalNormalizeMemberNames_(['A', 'B', 'C', 'D', 'E', 'F'], true), /5名以内/);
  assert.throws(() => sandbox.vNextPortalNormalizeMemberNames_(['山田 太郎', '山田　太郎'], true), /同じ関与メンバー/);
  const now = new Date('2026-08-12T00:00:00.000Z');
  assert.equal(sandbox.vNextPortalNormalizeCreationFiscalYear_(2026, now), 2026);
  assert.equal(sandbox.vNextPortalNormalizeCreationFiscalYear_(2036, now), 2036);
  assert.throws(() => sandbox.vNextPortalNormalizeCreationFiscalYear_(2037, now), /一覧から/);
  assert.throws(() => sandbox.vNextPortalNormalizeCreationFiscalYear_(2025, now), /一覧から/);
}

function testCatalogCacheAndStrictSelection() {
  cacheValues.clear();
  const originalRead = sandbox.vNextPortalReadTable_;
  let reads = 0;
  sandbox.vNextPortalReadTable_ = () => {
    reads++;
    return [
      { catalog_key: 'ZAC-2', client_name: 'ベータ製薬', is_active: false, catalog_version: 'CAT-1', synced_at: '2026-08-12T00:00:00.000Z' },
      { catalog_key: 'ZAC-1', client_name: 'アルファ製薬', is_active: true, catalog_version: 'CAT-1', synced_at: '2026-08-12T00:00:00.000Z' }
    ];
  };
  try {
    const first = sandbox.vNextPortalReadClientCatalog_({}, true);
    assert.deepEqual(JSON.parse(JSON.stringify(first.clients)), [{ key: 'ZAC-1', name: 'アルファ製薬' }]);
    assert.equal(reads, 1);
    sandbox.vNextPortalReadClientCatalog_({}, true);
    assert.equal(reads, 1, 'second model load should hit the document cache');
    sandbox.vNextPortalReadClientCatalog_({}, false);
    assert.equal(reads, 2, 'submit path must bypass the catalog cache');

    const selected = sandbox.vNextPortalResolveCreationInput_(
      { clientKey: 'ZAC-1', fiscalYear: 2027, relatedMemberNames: ['佐藤 花子'] },
      {}, 'creator@example.com', false
    );
    assert.equal(selected.clientName, 'アルファ製薬');
    assert.equal(selected.catalogKey, 'ZAC-1');
    assert.throws(() => sandbox.vNextPortalResolveCreationInput_(
      { clientKey: 'BROWSER-TAMPER', fiscalYear: 2027, relatedMemberNames: ['佐藤 花子'] },
      {}, 'creator@example.com', false
    ), /現在のZAC一覧にありません/);
  } finally {
    sandbox.vNextPortalReadTable_ = originalRead;
    cacheValues.clear();
  }
}

function testDuplicateCandidates() {
  const originalDirectory = sandbox.vNextPortalReadDirectory_;
  const originalRequests = sandbox.vNextPortalReadRequestModels_;
  sandbox.vNextPortalReadDirectory_ = () => [{
    source: 'DIRECTORY', directoryKey: 'D1', fiscalYear: 2027, clientId: 'ZAC-C-001',
    clientName: '株式会社テスト製薬', state: 'INPUT_OPEN', requestId: 'REQ-D1',
    url: 'https://docs.google.com/spreadsheets/d/12345678901234567890/edit', updatedAt: '2026-08-10T00:00:00Z'
  }];
  sandbox.vNextPortalReadRequestModels_ = () => [{
    source: 'REQUEST', requestId: 'REQ-FAILED', fiscalYear: 2027, clientId: 'ZAC-C-001',
    clientName: '株式会社テスト製薬', status: 'FAILED', statusLabel: '作成できませんでした', url: ''
  }];
  try {
    const exact = sandbox.vNextPortalBuildDuplicateCheck_({}, {
      fiscalYear: 2027, catalogKey: 'ZAC-C-001', clientName: 'テスト製薬株式会社',
      requestedBy: 'creator@example.com', relatedMemberNames: ['山田 太郎']
    });
    assert.equal(exact.hasExact, true);
    assert.equal(exact.candidates[0].reason, 'クライアントIDが一致');
    assert.equal(exact.candidates.some(item => item.requestId === 'REQ-FAILED'), false);
    assert.match(exact.hash, /^[a-f0-9]{64}$/);
  } finally {
    sandbox.vNextPortalReadDirectory_ = originalDirectory;
    sandbox.vNextPortalReadRequestModels_ = originalRequests;
  }
}

function testAppendRequestContract() {
  const originals = {
    resolve: sandbox.vNextPortalResolveCreationInput_, duplicate: sandbox.vNextPortalBuildDuplicateCheck_,
    lock: sandbox.vNextPortalWithDocumentLock_, append: sandbox.vNextPortalAppendRow_
  };
  let appended = null;
  sandbox.vNextPortalResolveCreationInput_ = () => ({
    catalogKey: 'ZAC-C-001', clientName: '新規クライアント', fiscalYear: 2027,
    relatedMemberNames: ['山田 太郎'], requestedBy: 'creator@example.com'
  });
  sandbox.vNextPortalBuildDuplicateCheck_ = () => ({ hash: 'a'.repeat(64), hasExact: false, hasSimilar: false, candidates: [] });
  sandbox.vNextPortalWithDocumentLock_ = operation => operation();
  sandbox.vNextPortalAppendRow_ = (spreadsheet, name, headers, record) => { appended = { name, headers, record }; };
  try {
    const result = sandbox.vNextPortalSubmitCreationRequest({
      clientKey: 'ZAC-C-001', confirmSimilarDuplicates: false, duplicateCheckHash: 'a'.repeat(64),
      fiscalYear: 2027, relatedMemberNames: ['山田 太郎', '', '', '', '']
    });
    assert.equal(result.status, 'PENDING');
    assert.equal(appended.name, 'VN_PORTAL_REQUEST');
    assert.equal(appended.record.event_type, 'REQUESTED');
    assert.equal(appended.record.catalog_key, 'ZAC-C-001');
    assert.equal(appended.record.forecast_owner_email, 'creator@example.com');
    assert.equal(appended.record.related_member_emails_json, '[]');
    assert.equal(appended.record.related_member_names_json, '["山田 太郎"]');
    assert.equal(appended.record.request_hash, sandbox.vNextPortalSha256Hex_(appended.record.request_json));
    const payload = JSON.parse(appended.record.request_json);
    assert.deepEqual(Object.keys(payload).sort(), [...sandbox.VNEXT_PORTAL.PAYLOAD_KEYS].sort());
    assert.equal(payload.requestedBy, 'creator@example.com');
    assert.equal(Object.hasOwn(payload, 'forecastOwnerEmail'), false);
  } finally {
    sandbox.vNextPortalResolveCreationInput_ = originals.resolve;
    sandbox.vNextPortalBuildDuplicateCheck_ = originals.duplicate;
    sandbox.vNextPortalWithDocumentLock_ = originals.lock;
    sandbox.vNextPortalAppendRow_ = originals.append;
  }
}

function testRequestRowProjection() {
  const payload = v2Payload();
  const current = {
    fiscal_year: 2027, client_id: 'ZAC-C-001', client_name: 'テスト株式会社',
    forecast_owner_email: 'creator@example.com', related_member_emails_json: '[]',
    requested_at: payload.requestedAt, requested_by: payload.requestedBy,
    catalog_key: 'ZAC-C-001', related_member_names_json: '["山田 太郎"]'
  };
  assert.equal(sandbox.vNextPortalAssertRequestRowProjection_(current, payload), true);
  assert.throws(() => sandbox.vNextPortalAssertRequestRowProjection_(
    { ...current, client_name: '改ざん値' }, payload
  ), /projection mismatch/);
  assert.throws(() => sandbox.vNextPortalAssertRequestRowProjection_(
    { ...current, related_member_names_json: '["X"]' }, payload
  ), /projection mismatch/);

  const legacy = legacyPayload();
  const legacyRow = {
    fiscal_year: 2027, client_id: 'CLIENT-1', client_name: 'テスト株式会社',
    forecast_owner_email: 'owner@example.com', related_member_emails_json: '["member@example.com"]',
    requested_at: legacy.requestedAt, requested_by: legacy.requestedBy,
    catalog_key: '', related_member_names_json: ''
  };
  assert.equal(sandbox.vNextPortalAssertRequestRowProjection_(legacyRow, legacy), true);
}

function testStatusEventValidation() {
  const hash = 'b'.repeat(64);
  const base = {
    request_id: 'PORTAL-REQ-12345678-test', request_hash: hash,
    event_type: 'CREATION_STARTED', status: 'CREATING', request_json: '', related_book_url: '',
    fiscal_year: 2027, client_id: 'ZAC-C-001', client_name: 'テスト株式会社',
    forecast_owner_email: 'creator@example.com', related_member_emails_json: '[]',
    requested_at: '2026-08-12T00:00:00.000Z', requested_by: 'creator@example.com',
    catalog_key: 'ZAC-C-001', related_member_names_json: '["山田 太郎"]'
  };
  const payload = v2Payload();
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_(base, base.request_id, hash, payload), true);
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_({ ...base, status: 'COMPLETED' }, base.request_id, hash, payload), false);
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_({ ...base, request_hash: 'c'.repeat(64) }, base.request_id, hash, payload), false);
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_({ ...base, request_json: '{}' }, base.request_id, hash, payload), false);
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_({ ...base, catalog_key: 'TAMPERED' }, base.request_id, hash, payload), false);
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_({ ...base, related_member_names_json: '["X"]' }, base.request_id, hash, payload), false);

  const legacy = legacyPayload();
  const legacyStatus = {
    ...base, client_id: 'CLIENT-1', forecast_owner_email: 'owner@example.com',
    related_member_emails_json: '["member@example.com"]', catalog_key: '', related_member_names_json: ''
  };
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_(legacyStatus, legacyStatus.request_id, hash, legacy), true);
}

function testCreateModel() {
  const originalCatalog = sandbox.vNextPortalReadClientCatalog_;
  sandbox.vNextPortalReadClientCatalog_ = () => ({
    version: 'CAT-1', syncedAt: '2026-08-12T00:00:00.000Z', clients: [{ key: 'ZAC-1', name: 'テスト' }]
  });
  try {
    const model = sandbox.vNextPortalGetCreateModel();
    assert.equal(model.fiscalYears.length, 11);
    assert.equal(model.defaultFiscalYear, model.fiscalYears[0] + 1);
    assert.equal(model.fiscalYears[10], model.fiscalYears[0] + 10);
    assert.equal(model.requesterEmail, 'creator@example.com');
    assert.equal(model.runtimeVersion, 'vnext-portal-1.7.3');
  } finally {
    sandbox.vNextPortalReadClientCatalog_ = originalCatalog;
  }
}

function testRequestProgress() {
  const pending = sandbox.vNextPortalRequestProgressModel_({
    requestId: 'PORTAL-REQ-12345678-test', clientName: 'テスト株式会社', fiscalYear: 2027,
    status: 'PENDING', requestedAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
    detailMessage: '受付済み', url: ''
  });
  assert.equal(pending.phase, 1);
  assert.equal(pending.isTerminal, false);
  assert.equal(pending.isStale, false);
  assert.match(pending.waitMessage, /5分ごと/);

  const completed = sandbox.vNextPortalRequestProgressModel_({
    requestId: 'PORTAL-REQ-12345678-test', clientName: 'テスト株式会社', fiscalYear: 2027,
    status: 'COMPLETED', updatedAt: new Date().toISOString(),
    url: 'https://docs.google.com/spreadsheets/d/12345678901234567890/edit'
  });
  assert.equal(completed.phase, 4);
  assert.equal(completed.isTerminal, true);
  assert.equal(completed.isSuccess, true);
  assert.match(completed.nextAction, /「開く」/);

  const failed = sandbox.vNextPortalRequestProgressModel_({
    requestId: 'PORTAL-REQ-12345678-test', clientName: 'テスト株式会社', fiscalYear: 2027,
    status: 'FAILED', updatedAt: new Date().toISOString(), detailMessage: '管理担当者へ連絡してください。', url: ''
  });
  assert.equal(failed.isTerminal, true);
  assert.equal(failed.isSuccess, false);
  assert.equal(failed.phase, 3);
  assert.equal(failed.nextAction, '管理担当者へ連絡してください。');
  const stale = sandbox.vNextPortalRequestProgressModel_({
    requestId: 'PORTAL-REQ-12345678-test', clientName: 'テスト株式会社', fiscalYear: 2027,
    status: 'PENDING', updatedAt: new Date(Date.now() - 16 * 60 * 1000).toISOString(), url: ''
  });
  assert.equal(stale.isStale, true);
  assert.match(stale.nextAction, /15分以上/);
  assert.throws(() => sandbox.vNextPortalNormalizeRequestId_('invalid'), /受付番号が不正/);

  const originalRead = sandbox.vNextPortalReadRequestModels_;
  sandbox.vNextPortalReadRequestModels_ = () => [{
    requestId: 'PORTAL-REQ-12345678-test', clientName: 'テスト株式会社', fiscalYear: 2027,
    status: 'VALIDATING', updatedAt: new Date().toISOString(), detailMessage: '', url: ''
  }];
  try {
    const endpoint = sandbox.vNextPortalGetRequestProgress('PORTAL-REQ-12345678-test');
    assert.equal(endpoint.status, 'VALIDATING');
    assert.equal(endpoint.phase, 2);
  } finally {
    sandbox.vNextPortalReadRequestModels_ = originalRead;
  }
}

function testEntryModel() {
  const model = sandbox.vNextPortalBuildEntryModel_({
    directory: [{
      directoryKey: '2027|az', clientName: 'アストラゼネカ', fiscalYear: 2027,
      state: 'SUBMITTED', nextAction: '管理者の承認待ちです。',
      url: 'https://docs.google.com/spreadsheets/d/1234567890123456789012/edit'
    }],
    requests: [{
      clientName: '新規株式会社', fiscalYear: 2027, status: 'CREATING',
      detailMessage: '', url: ''
    }]
  }, {
    portalUrl: 'https://docs.google.com/spreadsheets/d/portalportalportalportalpo/edit',
    adminHubUrl: 'https://docs.google.com/spreadsheets/d/hubhubhubhubhubhubhubhubhu/edit'
  });
  assert.equal(model.books.length, 2);
  assert.equal(model.years.join(','), '2027');
  assert.equal(model.books[0].clientName, 'アストラゼネカ');
  assert.equal(model.books[0].tone, 'warn');
  assert.equal(model.books[1].stateLabel, 'ブック作成中');
  assert.equal(model.adminHubUrl.startsWith('https://docs.google.com/spreadsheets/d/'), true);
  assert.equal('actorEmail' in model, false);
  assert.equal(sandbox.vNextPortalSafeSpreadsheetUrl_('https://example.com/x'), '');
}

async function testStaticUxContracts() {
  const core = await readFile(path.join(sourceDir, 'Portal_Core.js'), 'utf8');
  const ux = await readFile(path.join(sourceDir, 'Portal_UX.js'), 'utf8');
  const html = await readFile(path.join(sourceDir, 'Portal_CreateSidebar.html'), 'utf8');
  assert.match(ux, /function onOpen\(event\)/);
  assert.match(ux, /addItem\('案内を開く'/);
  const onOpenBody = ux.slice(ux.indexOf('function vNextPortalOnOpen_'), ux.indexOf('function vNextPortalGoHomeAndShowGuidance'));
  assert.doesNotMatch(onOpenBody, /vNextPortalRefreshViews_|showSidebar/);
  assert.match(ux, /function vNextPortalInstalledGuidanceOnOpen/);
  const openSidebarBody = ux.slice(ux.indexOf('function vNextPortalOpenCreateSidebar'), ux.indexOf('function vNextPortalOpenHelp'));
  assert.doesNotMatch(openSidebarBody, /vNextPortalEnsureStructure_/);
  assert.doesNotMatch(core, /getEffectiveUser/);
  assert.match(html, /id="clientKey"/);
  assert.doesNotMatch(html, /id="clientName"|id="clientId"|id="forecastOwnerEmail"/);
  assert.equal((html.match(/id="memberName[1-5]"/g) || []).length, 5);
  assert.match(html, /関与メンバー/);
  assert.match(html, /重複候補を確認/);
  assert.match(core, /CLIENT_CATALOG_SHEET: 'VN_PORTAL_CLIENT_CATALOG'/);
  assert.doesNotMatch(ux, /\.merge\(/,
    'Portal view rendering must not create merged cells');
  assert.doesNotMatch(ux, /\.setNote\(/,
    'Portal views must not leave hidden note popups or stale memo markers');
  assert.match(ux, /var headers = \['状態', 'クライアント', '年度', '次の案内', '更新', '開く'\]/,
    'Home must expose the compact six-column employee view');
  assert.match(ux, /var headers = \['状態', 'クライアント', '中心見込み', '採用予測', '最終予算', '担当・関与', '次の対応', '更新日', '開く'\]/,
    'FY tabs must expose the compact nine-column planning view');
  assert.match(ux, /dataRange\.getMergedRanges\(\)/);
  assert.doesNotMatch(ux, /sheet\.clear\(\)/,
    'View refresh must clear only the authored range, not the entire sheet');
  assert.match(html, /vNextPortalShowRequestOnHome\(result\.requestId\)/,
    'A successful request must refresh Home without asking for another button click');
  assert.match(core, /function vNextPortalGetRequestProgress\(requestId\)/,
    'Sidebar polling must use a lightweight request-only endpoint');
  assert.match(ux, /function vNextPortalShowRequestOnHome\(requestId\)/,
    'Successful submission must project the accepted request onto Home immediately');
  assert.match(html, /POLL_DELAYS_MS = \[20000, 70000, 210000\]/,
    'Automatic status checks must be sparse and bounded');
  assert.match(html, /if \(busy \|\| submitted/,
    'Client-side submission must synchronously guard double clicks');
  assert.match(html, /id="clientSearch" type="search"/,
    'Client selection must support local catalog filtering');
  const filterBody = html.slice(html.indexOf('function filterClients()'), html.indexOf('function normalizeSearchText'));
  assert.doesNotMatch(filterBody, /google\.script\.run/,
    'Client search must stay local and must not create RPC traffic');
  assert.match(html, /meta name="viewport"/);
  assert.match(html, /aria-live="polite"/);
  assert.match(html, /prefers-reduced-motion/);
  assert.equal((html.match(/id="memberRow[1-5]"/g) || []).length, 5);
  assert.match(ux, /vNextPortalDisplayPlanningValue_\(entry\.centerForecast, actualShortage \? '実績不足' : '未算出'\)/,
    'Empty and insufficient forecast values must be explained, not shown as zero');
  assert.match(ux, /function doGet\(/);
  assert.match(ux, /createHtmlOutputFromFile\('Portal_Entry'\)/);
  assert.match(core, /function vNextPortalGetEntryModel\(/);
  assert.match(core, /function vNextPortalBuildEntryModel_/);
  const entry = await readFile(path.join(sourceDir, 'Portal_Entry.html'), 'utf8');
  assert.match(entry, /新しい個別シートを用意する/);
  assert.match(entry, /vNextPortalGetEntryModel\(\)/);
  assert.match(entry, /vNextPortalPrepareOpenExperience\(\)/);
  assert.match(entry, /data-year/);
  assert.doesNotMatch(entry, /クライアント名で探す|クライアントレイヤー|運用担当|ログイン中|管理者用ハブ/);
  assert.match(entry, /管理者用の画面を開く/);
  assert.match(entry, /class="bubble"/);
  assert.match(entry, /\.bubble \{[\s\S]*?border:1px solid/);
  assert.doesNotMatch(entry, /font-weight:800/);
  assert.equal((entry.match(/class="bubble"/g) || []).length, 3);
  assert.equal((entry.match(/class="who"/g) || []).length, 3);
  assert.match(entry, /新しいシートを用意する方向け/);
  assert.match(entry, /予算を策定する方向け/);
  assert.match(entry, /管理者の方向け/);
  assert.doesNotMatch(entry, /管理者専用|opacity="\.28"|ellipse cx="30"/);
  assert.ok(entry.indexOf('class="choice-head"') < entry.indexOf('class="cast"'),
    'Numbered section headers must appear above each bot and bubble');
  assert.match(entry, /\.cast \{ display:flex; align-items:center;/);
  assert.doesNotMatch(entry, /data-speech|guideSpeech|mouseenter|word-break:keep-all/);
  assert.doesNotMatch(entry, /<br\s*\/?>/);
  assert.doesNotMatch(core, /cached\.actorEmail/);
}
