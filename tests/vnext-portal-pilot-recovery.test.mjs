#!/usr/bin/env node

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const source = await readFile(path.join(root, 'VNext_PortalPilotRecovery.js'), 'utf8');
const sandbox = {
  Logger: { log() {} },
  vNextAdminJsonSafe_: value => JSON.parse(JSON.stringify(value))
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox, { filename: 'VNext_PortalPilotRecovery.js' });

testPhaseContract();
testJournalBeforeCheckpoint();
testOrphanQuarantineIsRecoverable();
testPublicResultIsMinimal();
testLegacyV2BridgeContract();
testLegacyV2BridgeBehavior();
testStaticSafetyContract();
process.stdout.write('PASS vNext Portal Pilot recovery tests (7 groups)\n');

function testPhaseContract() {
  const phases = [
    'INITIALIZE', 'TEMPLATE_CREATE', 'TEMPLATE_INITIALIZE', 'TEMPLATE_COPY',
    'TEMPLATE_REGISTER', 'MODEL_REGISTER', 'PAIR_ACTIVATE', 'PORTAL_CREATE',
    'PORTAL_REGISTER', 'VERIFY', 'COMPLETED'
  ];
  phases.forEach((phase, index) => {
    assert.equal(sandbox.vNextAdminPortalPilotRecoveryPhaseIndex_(phase), index);
    assert.ok(sandbox.vNextAdminPortalPilotRecoveryPhaseLabel_(phase));
  });
  assert.equal(sandbox.vNextAdminPortalPilotRecoveryPhaseIndex_('unknown'), -1);
}

function testJournalBeforeCheckpoint() {
  const events = [];
  sandbox.vNextAdminAppendTemplateJournal_ = (_hub, record) => {
    events.push(`journal:${record.phase}`);
    return record;
  };
  sandbox.vNextAdminPortalPilotRecoverySave_ = state => {
    events.push(`save:${state.phase}`);
    return state;
  };
  const state = {
    phase: 'TEMPLATE_CREATE', operationId: 'OP', releaseId: 'REL', modelReleaseId: 'MODEL',
    initialReleaseId: 'OLD_REL', initialModelReleaseId: 'OLD_MODEL', templateSpreadsheetId: ''
  };
  const output = sandbox.vNextAdminPortalPilotRecoveryAdvance_(
    {}, state, 'TEMPLATE_INITIALIZE', 'PILOT_TEMPLATE_CONTAINER_READY', { spreadsheetId: 'SHEET' }
  );
  assert.deepEqual(events, [
    'journal:PILOT_TEMPLATE_CONTAINER_READY', 'save:TEMPLATE_INITIALIZE'
  ], 'Hub journal is durable before the Script Property phase checkpoint advances');
  assert.equal(output.state.phase, 'TEMPLATE_INITIALIZE');
  assert.throws(
    () => sandbox.vNextAdminPortalPilotRecoveryAdvance_({}, output.state, 'TEMPLATE_CREATE', 'BACKWARD', {}),
    /cannot move backward/
  );
}

function testOrphanQuarantineIsRecoverable() {
  const renamed = [];
  const privateIds = [];
  const audits = [];
  const files = ['FILE_A', 'FILE_B'].map(id => ({
    getId: () => id,
    setName: value => renamed.push([id, value])
  }));
  let cursor = 0;
  sandbox.DriveApp = {
    getFolderById: id => ({
      id,
      getFilesByName: () => ({
        hasNext: () => cursor < files.length,
        next: () => files[cursor++]
      })
    })
  };
  sandbox.vNextAdminEnforcePrivateFileAcl_ = file => privateIds.push(file.getId());
  sandbox.vNextAdminWriteAudit_ = (_hub, action, kind, id, status, detail) => {
    audits.push({ action, kind, id, status, detail });
  };
  const count = sandbox.vNextAdminPortalPilotRecoveryQuarantineUnknownFiles_(
    {}, 'FOLDER', 'Staging title', ['admin@example.com'], 'OP', 'TEMPLATE'
  );
  assert.equal(count, 2);
  assert.deepEqual(privateIds, ['FILE_A', 'FILE_B']);
  assert.equal(renamed.length, 2);
  assert.ok(renamed.every(([, name]) => name.startsWith('[UNVERIFIED ORPHAN TEMPLATE]')));
  assert.ok(audits.every(item => item.detail.policy === 'PRESERVED_PRIVATE_NOT_REUSED_WITHOUT_SCRIPT_ID'));
}

function testPublicResultIsMinimal() {
  const result = sandbox.vNextAdminPortalPilotRecoveryPublicResult_({
    operationId: 'OP', phase: 'PORTAL_CREATE', releaseId: 'REL', modelReleaseId: 'MODEL',
    templateSpreadsheetId: 'TEMPLATE', portalId: 'PORTAL', portalSpreadsheetId: '',
    errorCount: 0, lastError: '', updatedAt: '2026-08-12T00:00:00.000Z', completedAt: '',
    evidenceArtifact: 'must-not-leak', adminEmails: ['admin@example.com'], parameters: { secret: 'no' }
  }, {});
  assert.equal(result.phaseIndex, 7);
  assert.equal(result.phaseCount, 10);
  assert.equal(result.needsAnotherRun, true);
  assert.equal(Object.hasOwn(result, 'evidenceArtifact'), false);
  assert.equal(Object.hasOwn(result, 'adminEmails'), false);
  assert.equal(Object.hasOwn(result, 'parameters'), false);
}

function testLegacyV2BridgeContract() {
  const initializeStart = source.indexOf('function vNextAdminPortalPilotRecoveryInitialize_(');
  const resolverStart = source.indexOf('function vNextAdminPortalPilotRecoveryResolveTemplateSource_(');
  const resolverEnd = source.indexOf('function vNextAdminPortalPilotRecoveryTemplateCreate_(', resolverStart);
  assert.ok(initializeStart >= 0 && resolverStart > initializeStart && resolverEnd > resolverStart);
  const initialize = source.slice(initializeStart, resolverStart);
  const resolver = source.slice(resolverStart, resolverEnd);
  assert.doesNotMatch(initialize, /vNextAdminResolveTemplateUiSource_\s*\(/,
    'empty legacy source path must not enter the V2 full-grid verifier');
  assert.match(initialize, /vNextAdminPortalPilotRecoveryResolveTemplateSource_\(hub, req\)/);
  assert.match(resolver, /if \(draftId\)[\s\S]*vNextAdminResolveTemplateUiSource_\(hub, draftId\)/,
    'an explicitly selected mutable draft continues to use the strict existing resolver');
  assert.match(resolver, /VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2/);
  assert.match(resolver, /req\.attestationConfirmed !== true/);
  assert.match(resolver, /activeRelease\.template_content_sha256/);
  assert.match(resolver, /ACTIVE Template BOOK_REGISTRY identity/);
  assert.match(resolver, /vNextAdminActiveReleasePairMirrorsExact_/);
  assert.match(resolver, /ACTIVE Template local routing identity/);
  assert.match(resolver, /vNextClientRuntimeAssertBoundParent_/);
  assert.match(resolver, /vNextAdminAssertPrivateAdminFile_/);
  assert.match(resolver, /const v3ManifestSha256 = vNextAdminTemplateUiManifestHash_\(spreadsheet\)/);
  assert.doesNotMatch(resolver, /vNextAdminTemplateUiManifestHashV2_/);
  assert.doesNotMatch(resolver, /vNextAdminAssertReleaseTemplateManifest_/);
}

function testLegacyV2BridgeBehavior() {
  const shaV2 = 'a'.repeat(64);
  const shaV3 = 'b'.repeat(64);
  const runtimeSha = 'c'.repeat(64);
  const spreadsheet = { getId: () => 'TEMPLATE_SHEET' };
  const hub = { getId: () => 'HUB_SHEET' };
  const release = {
    release_id: 'REL_V2', status: 'ACTIVE', template_spreadsheet_id: 'TEMPLATE_SHEET',
    template_script_id: 'SCRIPT_ID', template_manifest_schema: 'V2',
    template_content_sha256: shaV2, client_runtime_version: 'CLIENT_V1',
    client_runtime_sha256: runtimeSha, schema_version: 'CORE_V1', engine_version: 'ENGINE_V1'
  };
  const registry = {
    mode: 'TEMPLATE', status: 'ACTIVE', book_id: 'BOOK_ID', spreadsheet_id: 'TEMPLATE_SHEET',
    template_release_id: 'REL_V2', client_script_id: 'SCRIPT_ID',
    client_runtime_version: 'CLIENT_V1', client_runtime_sha256: runtimeSha
  };
  sandbox.VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA = 'V3';
  sandbox.VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2 = 'V2';
  sandbox.VN_ADMIN_SHEETS = { RELEASES: 'RELEASES' };
  sandbox.VN_ADMIN_BOOK_CONFIG_SHEET = 'BOOK_CONFIG';
  sandbox.VN_ADMIN_SYSTEM_CONFIG_SHEET = 'SYSTEM_CONFIG';
  sandbox.vNextAdminText_ = value => String(value || '').trim();
  sandbox.vNextAdminRequiredText_ = (value, name) => {
    const text = String(value || '').trim();
    if (!text) throw new Error(`${name} required`);
    return text;
  };
  sandbox.vNextAdminReadActiveReleasePair_ = () => ({
    releaseId: 'REL_V2', modelReleaseId: 'MODEL_V1', templateSpreadsheetId: 'TEMPLATE_SHEET'
  });
  sandbox.vNextAdminReadTable_ = () => ({ rows: [release] });
  sandbox.vNextAdminLatestModelRelease_ = () => ({
    model_release_id: 'MODEL_V1', status: 'ACTIVE', template_version: 'REL_V2'
  });
  sandbox.vNextAdminAssertModelReleaseChecksPassed_ = () => true;
  sandbox.vNextAdminAssertModelTemplateCompatibility_ = () => true;
  sandbox.vNextAdminActiveReleasePairMirrorsExact_ = () => true;
  sandbox.vNextAdminFindRegistryRow_ = (_unused, predicate) => predicate(registry) ? registry : null;
  sandbox.SpreadsheetApp = { openById: id => {
    assert.equal(id, 'TEMPLATE_SHEET');
    return spreadsheet;
  } };
  sandbox.vNextAdminReadKeyValueSheet_ = (target, sheetName) => {
    if (target === hub) return { admin_emails: 'admin@example.com' };
    if (sheetName === 'BOOK_CONFIG') return {
      book_id: 'BOOK_ID', version: 'REL_V2', client_runtime_version: 'CLIENT_V1',
      client_runtime_bundle_sha256: runtimeSha
    };
    return { mode: 'TEMPLATE', book_id: 'BOOK_ID', active_release_id: 'REL_V2' };
  };
  sandbox.vNextDetectBookMode_ = () => 'TEMPLATE';
  sandbox.vNextClientRuntimeAssertBoundParent_ = () => true;
  sandbox.vNextGetRuntimeConfig_ = () => ({ VNEXT_ADMIN_EMAILS: 'admin@example.com' });
  sandbox.vNextAdminActor_ = () => 'admin@example.com';
  sandbox.vNextAdminMergeEmails_ = () => ['admin@example.com'];
  sandbox.DriveApp = { getFileById: id => ({ id }) };
  sandbox.vNextAdminAssertPrivateAdminFile_ = () => true;
  let v3HashCalls = 0;
  sandbox.vNextAdminTemplateUiManifestHash_ = () => { v3HashCalls++; return shaV3; };
  sandbox.vNextAdminResolveTemplateUiSource_ = () => {
    throw new Error('legacy empty path entered strict V2 full-grid resolver');
  };

  const bridged = sandbox.vNextAdminPortalPilotRecoveryResolveTemplateSource_(hub, {
    attestationConfirmed: true, evidenceArtifact: 'reviewed-tests'
  });
  assert.equal(bridged.legacyV2BridgeUsed, true);
  assert.equal(bridged.sourceStoredManifestSha256, shaV2);
  assert.equal(bridged.manifestSha256, shaV3);
  assert.equal(v3HashCalls, 1, 'the V3 used-envelope manifest is still computed exactly once');

  assert.throws(
    () => sandbox.vNextAdminPortalPilotRecoveryResolveTemplateSource_(hub, {}),
    /explicit reviewed Admin attestation/
  );
  release.template_manifest_schema = 'LEGACY_V1';
  assert.throws(
    () => sandbox.vNextAdminPortalPilotRecoveryResolveTemplateSource_(hub, {
      attestationConfirmed: true, evidenceArtifact: 'reviewed-tests'
    }),
    /only V2 or V3/
  );
  release.template_manifest_schema = 'V2';
}

function testStaticSafetyContract() {
  assert.match(source, /function vNextAdminContinueEmployeePortalPilotRecovery\(request\)/);
  assert.match(source, /function vNextAdminGetEmployeePortalPilotRecoveryStatus\(\)/);
  assert.match(source, /function vNextAdminContinueEmployeePortalPilotRecoveryForManualTest\(\)/);
  assert.match(source, /durable before the external create call/g);
  assert.match(source, /PRESERVED_PRIVATE_NOT_REUSED_WITHOUT_SCRIPT_ID/);
  assert.doesNotMatch(source, /\.setTrashed\s*\(/);
  assert.doesNotMatch(source, /\.deleteProperty\s*\(/);
  assert.doesNotMatch(source, /DriveApp\.removeFile/);
}
