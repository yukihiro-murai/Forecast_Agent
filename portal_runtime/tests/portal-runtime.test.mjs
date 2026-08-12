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
const sandbox = {
  console,
  Logger: { log() {} },
  SpreadsheetApp: { getActiveSpreadsheet: () => ({}) },
  Session: {
    getActiveUser: () => ({ getEmail: () => 'creator@example.com' }),
    getEffectiveUser: () => ({ getEmail: () => 'creator@example.com' })
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
testExactSchemas();
testNormalization();
testDuplicateCandidates();
testStatusEventValidation();
testAppendRequestContract();
await testStaticUxContracts();
process.stdout.write('PASS portal runtime behavior tests (7)\n');

function testCanonicalJson() {
  assert.equal(sandbox.vNextPortalCanonicalJson_({ b: 2, a: 1 }), '{"a":1,"b":2}');
  assert.equal(sandbox.vNextPortalCanonicalJson_({ z: undefined, a: [1, undefined] }), '{"a":[1,null]}');
  assert.equal(sandbox.vNextPortalSha256Hex_('{"a":1}'), createHash('sha256').update('{"a":1}').digest('hex'));
}

function testExactSchemas() {
  const payload = {
    clientId: 'CLIENT-1', clientName: 'テスト株式会社', fiscalYear: 2027,
    forecastOwnerEmail: 'owner@example.com', relatedMemberEmails: ['member@example.com'],
    requestId: 'PORTAL-REQ-12345678-test', requestType: 'CREATE_CLIENT_FY_BOOK',
    requestedAt: '2026-08-12T00:00:00.000Z', requestedBy: 'creator@example.com',
    schemaVersion: 'vnext-portal-request-1'
  };
  assert.equal(sandbox.vNextPortalValidateRequestPayload_(payload), payload);
  assert.throws(() => sandbox.vNextPortalValidateRequestPayload_({ ...payload, unexpected: true }), /項目が契約と一致/);
  assert.throws(() => sandbox.vNextPortalAssertExactKeys_({ clientName: 'x' }, sandbox.VNEXT_PORTAL.PREVIEW_INPUT_KEYS), /項目が契約と一致/);
}

function testNormalization() {
  const value = sandbox.vNextPortalNormalizeCreationInput_({
    fiscalYear: '2027', clientName: '  テスト株式会社  ', clientId: ' C-001 ',
    forecastOwnerEmail: 'OWNER@EXAMPLE.COM',
    relatedMembersText: 'member@example.com, OWNER@example.com\nmember@example.com\nother@example.com'
  });
  assert.equal(value.fiscalYear, 2027);
  assert.equal(value.clientName, 'テスト株式会社');
  assert.equal(value.forecastOwnerEmail, 'owner@example.com');
  assert.deepEqual([...value.relatedMemberEmails], ['member@example.com', 'other@example.com']);
  assert.equal(sandbox.vNextPortalNormalizeClientName_('株式会社 テスト製薬'), 'テスト製薬');
}

function testDuplicateCandidates() {
  const originalDirectory = sandbox.vNextPortalReadDirectory_;
  const originalRequests = sandbox.vNextPortalReadRequestModels_;
  sandbox.vNextPortalReadDirectory_ = () => [{
    source: 'DIRECTORY', directoryKey: 'D1', fiscalYear: 2027, clientId: 'C-001',
    clientName: '株式会社テスト製薬', state: 'INPUT_OPEN', requestId: 'REQ-D1',
    url: 'https://docs.google.com/spreadsheets/d/12345678901234567890/edit', updatedAt: '2026-08-10T00:00:00Z'
  }, {
    source: 'DIRECTORY', directoryKey: 'D2', fiscalYear: 2027, clientId: '',
    clientName: 'テスト製薬 東日本', state: 'DRAFT_READY', requestId: '', url: '', updatedAt: '2026-08-09T00:00:00Z'
  }];
  sandbox.vNextPortalReadRequestModels_ = () => [{
    source: 'REQUEST', requestId: 'REQ-PENDING', fiscalYear: 2027, clientId: '',
    clientName: '別会社', status: 'PENDING', statusLabel: '受付済み', url: '', updatedAt: '2026-08-11T00:00:00Z'
  }, {
    source: 'REQUEST', requestId: 'REQ-FAILED', fiscalYear: 2027, clientId: '',
    clientName: '株式会社テスト製薬', status: 'FAILED', statusLabel: '作成できませんでした', url: '', updatedAt: '2026-08-11T00:00:00Z'
  }];
  try {
    const exact = sandbox.vNextPortalBuildDuplicateCheck_({}, {
      fiscalYear: 2027, clientId: 'c001', clientName: 'テスト製薬株式会社',
      forecastOwnerEmail: 'owner@example.com', relatedMemberEmails: []
    });
    assert.equal(exact.hasExact, true);
    assert.equal(exact.candidates[0].reason, 'クライアントIDが一致');
    assert.equal(exact.candidates.some((item) => item.requestId === 'REQ-FAILED'), false, 'failed requests must not block retry');
    const pending = sandbox.vNextPortalBuildDuplicateCheck_({}, {
      fiscalYear: 2027, clientId: '', clientName: '別会社',
      forecastOwnerEmail: 'owner@example.com', relatedMemberEmails: []
    });
    assert.equal(pending.hasExact, true, 'pending local requests must prevent a concurrent duplicate');
    assert.match(pending.hash, /^[a-f0-9]{64}$/);
  } finally {
    sandbox.vNextPortalReadDirectory_ = originalDirectory;
    sandbox.vNextPortalReadRequestModels_ = originalRequests;
  }
}

function testAppendRequestContract() {
  const originals = {
    ensure: sandbox.vNextPortalEnsureStructure_, duplicate: sandbox.vNextPortalBuildDuplicateCheck_,
    lock: sandbox.vNextPortalWithDocumentLock_, append: sandbox.vNextPortalAppendRow_
  };
  let appended = null;
  sandbox.vNextPortalEnsureStructure_ = () => ({});
  sandbox.vNextPortalBuildDuplicateCheck_ = () => ({ hash: 'a'.repeat(64), hasExact: false, hasSimilar: false, candidates: [] });
  sandbox.vNextPortalWithDocumentLock_ = (operation) => operation();
  sandbox.vNextPortalAppendRow_ = (spreadsheet, name, headers, record) => { appended = { name, headers, record }; };
  try {
    const result = sandbox.vNextPortalSubmitCreationRequest({
      clientId: '', clientName: '新規クライアント', confirmSimilarDuplicates: false,
      duplicateCheckHash: 'a'.repeat(64), fiscalYear: 2027,
      forecastOwnerEmail: 'owner@example.com', relatedMembersText: 'member@example.com'
    });
    assert.equal(result.status, 'PENDING');
    assert.equal(appended.name, 'VN_PORTAL_REQUEST');
    assert.equal(appended.record.event_type, 'REQUESTED');
    assert.equal(appended.record.request_hash, sandbox.vNextPortalSha256Hex_(appended.record.request_json));
    const payload = JSON.parse(appended.record.request_json);
    assert.deepEqual(Object.keys(payload).sort(), [...sandbox.VNEXT_PORTAL.PAYLOAD_KEYS].sort());
    assert.equal(payload.requestType, 'CREATE_CLIENT_FY_BOOK');
    assert.equal(payload.requestedBy, 'creator@example.com');
  } finally {
    sandbox.vNextPortalEnsureStructure_ = originals.ensure;
    sandbox.vNextPortalBuildDuplicateCheck_ = originals.duplicate;
    sandbox.vNextPortalWithDocumentLock_ = originals.lock;
    sandbox.vNextPortalAppendRow_ = originals.append;
  }
}

function testStatusEventValidation() {
  const hash = 'b'.repeat(64);
  const base = {
    request_id: 'PORTAL-REQ-12345678-test', request_hash: hash,
    event_type: 'CREATION_STARTED', status: 'CREATING', request_json: '', related_book_url: ''
  };
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_(base, base.request_id, hash), true);
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_({ ...base, status: 'COMPLETED' }, base.request_id, hash), false);
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_({ ...base, request_hash: 'c'.repeat(64) }, base.request_id, hash), false);
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_({ ...base, request_json: '{}' }, base.request_id, hash), false);
  assert.equal(sandbox.vNextPortalIsValidStatusEvent_({
    ...base, event_type: 'COMPLETED', status: 'COMPLETED',
    related_book_url: 'https://example.com/not-a-sheet'
  }, base.request_id, hash), false);
}

async function testStaticUxContracts() {
  const ux = await readFile(path.join(sourceDir, 'Portal_UX.js'), 'utf8');
  const html = await readFile(path.join(sourceDir, 'Portal_CreateSidebar.html'), 'utf8');
  assert.match(ux, /function onOpen\(event\)/);
  assert.match(ux, /addItem\('ホームに戻る'/);
  assert.match(ux, /addItem\('新しい年度計画を作る'/);
  assert.match(ux, /addItem\('使い方・困ったとき'/);
  assert.match(ux, /vNextPortalEnsureFiscalYearSheet_/);
  assert.match(html, /重複候補を確認/);
  assert.match(html, /このクライアントの作成を依頼/);
}
