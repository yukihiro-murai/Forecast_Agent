#!/usr/bin/env node

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const bundle = JSON.parse(await readFile(path.join(root, 'client_runtime', 'generated', 'client-runtime-bundle.json'), 'utf8'));
const sourceId = 'SOURCE_script_id_1234567890ABCDE';
const targetId = 'TARGET_script_id_1234567890ABCDE';
const sandbox = {
  console,
  Logger: { log() {} },
  Session: {
    getActiveUser: () => ({ getEmail: () => 'admin@example.com' }),
    getEffectiveUser: () => ({ getEmail: () => 'admin@example.com' })
  },
  Utilities: {
    Charset: { UTF_8: 'UTF_8' },
    DigestAlgorithm: { SHA_256: 'SHA_256' },
    computeDigest(_algorithm, value) {
      return [...createHash('sha256').update(String(value), 'utf8').digest()];
    },
    sleep() {}
  }
};
vm.createContext(sandbox);
vm.runInContext(
  await readFile(path.join(root, 'VNext_PortalRuntimeBundle.js'), 'utf8'),
  sandbox,
  { filename: 'VNext_PortalRuntimeBundle.js' }
);
vm.runInContext(
  await readFile(path.join(root, 'VNext_ClientRuntimeProvisioning.js'), 'utf8'),
  sandbox,
  { filename: 'VNext_ClientRuntimeProvisioning.js' }
);

assert.match(
  await readFile(path.join(root, 'VNext_ClientRuntimeProvisioning.js'), 'utf8'),
  /function vNextClientRuntimeEnableRequiredAppsScriptApi_\(\)[\s\S]*SERVICE_DISABLED[\s\S]*serviceusage\.googleapis\.com\/v1\/projects\/[\s\S]*services\/script\.googleapis\.com:enable/,
  'API bootstrap must enable only the verified script.googleapis.com service from a SERVICE_DISABLED response'
);

testValidCopy();
testPortalManifestContract();
testPortalLegacyExistingFiles();
testPinnedHistoricalCopy();
testIdentityGuards();
testContentGuards();
process.stdout.write('PASS vNext client runtime copy tests (6 groups)\n');

function testPortalLegacyExistingFiles() {
  const verified = sandbox.vNextPortalRuntimeVerifiedBundle_();
  const current = sandbox.vNextPortalRuntimeValidateExistingFiles_(verified.files);
  assert.equal(current.length, 5);
  const legacy = verified.files.filter(file => file.name !== 'Portal_Entry');
  assert.equal(legacy.length, 4);
  const accepted = sandbox.vNextPortalRuntimeValidateExistingFiles_(legacy);
  assert.equal(accepted.map(file => file.name).join(','),
    'Portal_Core,Portal_CreateSidebar,Portal_UX,appsscript');
  assert.throws(
    () => sandbox.vNextPortalRuntimeValidateFiles_(legacy),
    /Portal runtime file count does not match the allowlist/
  );
}

function testPortalManifestContract() {
  const verified = sandbox.vNextPortalRuntimeVerifiedBundle_();
  assert.equal(verified.files.length, 5);
  assert.match(verified.version, /^vnext-portal-\d+\.\d+\.\d+$/);
}

function testValidCopy() {
  const calls = [];
  sandbox.vNextClientRuntimeApiRequest_ = (apiPath, method, body) => {
    calls.push({ apiPath, method, body });
    if (method === 'get') {
      return { scriptId: sourceId, files: clone(bundle.files).reverse() };
    }
    assert.equal(method, 'put');
    assert.equal(apiPath, `/projects/${encodeURIComponent(targetId)}/content`);
    return { scriptId: targetId, files: clone(body.files).reverse() };
  };
  const result = sandbox.vNextClientRuntimeCopyScriptContent_(sourceId, targetId, bundle.sha256.toUpperCase());
  assert.equal(result.ok, true);
  assert.equal(result.sourceScriptId, sourceId);
  assert.equal(result.targetScriptId, targetId);
  assert.equal(result.bundleSha256, bundle.sha256);
  assert.equal(result.fileCount, 10);
  assert.equal(result.updateResult.verificationSource, 'UPDATE_RESPONSE');
  assert.equal(calls.length, 2);
  assert.equal(calls[0].apiPath, `/projects/${encodeURIComponent(sourceId)}/content`);
  assert.equal(calls[0].method, 'get');
  assert.equal(
    JSON.stringify(calls[1].body.files.map(file => file.name)),
    JSON.stringify(['Client_Bridge', 'Client_Core', 'Client_Entry', 'VNext_GuidanceSidebar', 'VNext_HelpSidebar', 'VNext_InputSidebar', 'VNext_PlanSidebar', 'VNext_ReviewSidebar', 'VNext_UX', 'appsscript'])
  );
}

function testPinnedHistoricalCopy() {
  const legacyFiles = clone(bundle.files).filter(file => file.name !== 'VNext_GuidanceSidebar');
  const legacyManifestFile = legacyFiles.find(file => file.name === 'appsscript');
  const legacyManifest = JSON.parse(legacyManifestFile.source);
  legacyManifest.oauthScopes = legacyManifest.oauthScopes.filter(scope =>
    scope !== 'https://www.googleapis.com/auth/script.scriptapp');
  legacyManifestFile.source = JSON.stringify(legacyManifest, null, 2) + '\n';
  const legacySha = filesSha256(legacyFiles);
  assert.throws(
    () => sandbox.vNextClientRuntimeVerifyScriptContent_(
      { scriptId: sourceId, files: legacyFiles }, sourceId, legacySha
    ),
    /exactly 9 client files/,
    'A legacy runtime must never satisfy the fresh-release contract.'
  );
  const pinned = sandbox.vNextClientRuntimeVerifyPinnedScriptContent_(
    { scriptId: sourceId, files: legacyFiles }, sourceId, legacySha
  );
  assert.equal(pinned.historicalContract, true);
  assert.equal(pinned.files.length, 9);

  const calls = [];
  sandbox.vNextClientRuntimeApiRequest_ = (apiPath, method, body) => {
    calls.push({ apiPath, method, body });
    if (method === 'get') return { scriptId: sourceId, files: clone(legacyFiles).reverse() };
    return { scriptId: targetId, files: clone(body.files).reverse() };
  };
  const result = sandbox.vNextClientRuntimeCopyScriptContent_(sourceId, targetId, legacySha);
  assert.equal(result.bundleSha256, legacySha);
  assert.equal(result.fileCount, 9);
  assert.equal(calls.length, 2);

  const previousTenFiles = clone(bundle.files);
  const previousManifestFile = previousTenFiles.find(file => file.name === 'appsscript');
  const previousManifest = JSON.parse(previousManifestFile.source);
  previousManifest.oauthScopes = previousManifest.oauthScopes.filter(scope =>
    scope !== 'https://www.googleapis.com/auth/script.scriptapp');
  previousManifestFile.source = JSON.stringify(previousManifest, null, 2) + '\n';
  const previousSha = filesSha256(previousTenFiles);
  const previousPinned = sandbox.vNextClientRuntimeVerifyPinnedScriptContent_(
    { scriptId: sourceId, files: previousTenFiles }, sourceId, previousSha
  );
  assert.equal(previousPinned.historicalContract, true,
    'the exact pre-trigger ten-file release remains readable only by its stored SHA');

  assert.throws(
    () => sandbox.vNextClientRuntimeVerifyPinnedScriptContent_(
      { scriptId: sourceId, files: legacyFiles }, sourceId, '0'.repeat(64)
    ),
    /hash does not match expectedSha256/
  );
  const altered = clone(legacyFiles);
  altered.push({ name: 'Unexpected', type: 'SERVER_JS', source: '' });
  assert.throws(
    () => sandbox.vNextClientRuntimeVerifyPinnedScriptContent_(
      { scriptId: sourceId, files: altered }, sourceId, filesSha256(altered)
    ),
    /file allowlist mismatch/
  );
}

function testIdentityGuards() {
  let calls = 0;
  sandbox.vNextClientRuntimeApiRequest_ = () => { calls++; return {}; };
  assert.throws(
    () => sandbox.vNextClientRuntimeCopyScriptContent_(' short', targetId, bundle.sha256),
    /sourceScriptId is invalid/
  );
  assert.throws(
    () => sandbox.vNextClientRuntimeCopyScriptContent_(sourceId, sourceId, bundle.sha256),
    /must be different/
  );
  assert.throws(
    () => sandbox.vNextClientRuntimeCopyScriptContent_(sourceId, targetId, 'abc'),
    /exactly 64 hexadecimal/
  );
  assert.equal(calls, 0, 'Identity failures must happen before any API request.');
}

function testContentGuards() {
  assertSourceRejected([...clone(bundle.files), { name: 'Unexpected', type: 'SERVER_JS', source: '' }], bundle.sha256, /exactly 9 client files/);

  const wrongType = clone(bundle.files);
  wrongType.find(file => file.name === 'Client_Core').type = 'HTML';
  assertSourceRejected(wrongType, bundle.sha256, /file type mismatch/);

  const broadManifest = clone(bundle.files);
  const manifestFile = broadManifest.find(file => file.name === 'appsscript');
  const manifest = JSON.parse(manifestFile.source);
  manifest.oauthScopes.push('https://www.googleapis.com/auth/drive');
  manifestFile.source = JSON.stringify(manifest);
  assertSourceRejected(broadManifest, bundle.sha256, /OAuth scopes/);

  const forbidden = clone(bundle.files);
  forbidden.find(file => file.name === 'Client_Core').source += '\nUrlFetchApp.fetch("https://example.com");';
  assertSourceRejected(forbidden, bundle.sha256, /forbidden capability/);

  assertSourceRejected(clone(bundle.files), '0'.repeat(64), /does not match expectedSha256/);

  sandbox.vNextClientRuntimeApiRequest_ = (_apiPath, method, body) => method === 'get'
    ? { scriptId: sourceId, files: clone(bundle.files) }
    : { scriptId: sourceId, files: clone(body.files) };
  assert.throws(
    () => sandbox.vNextClientRuntimeCopyScriptContent_(sourceId, targetId, bundle.sha256),
    /update response scriptId does not match targetScriptId/
  );
}

function assertSourceRejected(files, expectedSha, pattern) {
  let putCalled = false;
  sandbox.vNextClientRuntimeApiRequest_ = (_apiPath, method) => {
    if (method === 'put') putCalled = true;
    return { scriptId: sourceId, files };
  };
  assert.throws(
    () => sandbox.vNextClientRuntimeCopyScriptContent_(sourceId, targetId, expectedSha),
    pattern
  );
  assert.equal(putCalled, false, 'Unverified source content must never be written.');
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function filesSha256(files) {
  return createHash('sha256').update(files.map(file =>
    `${file.name}\u0000${file.type}\u0000${file.source}`
  ).join('\u0000'), 'utf8').digest('hex');
}
