#!/usr/bin/env node

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const source = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
const naming = await readFile(path.join(root, '0_VNext_Naming.js'), 'utf8');

checkStaticSafetyContract();
checkPartialHubMetaAppendRecovery();
checkClientOnlyMetaFailsClosed();
checkRetryJournalSafety();
process.stdout.write('PASS vNext empty Pilot Client upgrade tests (4 groups)\n');

function checkStaticSafetyContract() {
  assert.match(source, /const VN_ADMIN_MIGRATION_APPLY_ENABLED = false;/,
    'generic migration APPLY must stay disabled');
  assert.match(source, /function vNextAdminUpgradeEmptyPilotClient\(request\)/);
  assert.match(source, /function vNextAdminRecoverEmptyPilotClientUpgrade\(request\)/);
  assert.match(source, /const dryRun = req\.dryRun !== false;/,
    'the dedicated API must default to read-only dry-run');
  const upgrade = functionSource('vNextAdminUpgradeEmptyPilotClient', 'vNextAdminRecoverEmptyPilotClientUpgrade');
  assert.ok(upgrade.indexOf('vNextAdminReadActiveReleasePair_') <
    upgrade.indexOf('vNextAdminEmptyPilotRelease_(hub, targetReleaseId'),
  'target release must be selected from the canonical pair');
  assert.ok(upgrade.includes("String(targetRelease.template_spreadsheet_id || '') !== pair.templateSpreadsheetId"));
  const eligibility = functionSource('vNextAdminAssertEmptyPilotNoBusinessData_',
    'vNextAdminAssertEmptyPilotPinnedRelease_');
  for (const token of [
    'EVIDENCE_EVENT', 'FORECAST_RUN', 'PLAN_VERSION', 'EVALUATION',
    'STATE_EVENT', 'VN_ADMIN_CLIENT_REQUEST_SHEET', 'VN_ADMIN_SHEETS.APPROVALS',
    'VN_ADMIN_SHEETS.OFFICIAL', "['QUEUED', 'RUNNING']"
  ]) assert.ok(eligibility.includes(token), `eligibility guard missing ${token}`);
  const apply = functionSource('vNextAdminApplyEmptyPilotRelease_',
    'vNextAdminAppendEmptyPilotRepairMeta_');
  const runtimeWrite = apply.indexOf('vNextClientRuntimeCopyScriptContent_');
  const uiWrite = apply.indexOf('vNextAdminCopyTemplateUiToClient_');
  const metaWrite = apply.indexOf('vNextAdminAppendEmptyPilotRepairMeta_');
  const registryWrite = apply.indexOf('vNextAdminPatchRegistryByBookId_');
  assert.ok(runtimeWrite >= 0 && uiWrite > runtimeWrite && metaWrite > uiWrite && registryWrite > metaWrite,
    'registry must be committed after runtime, UI/config and append-only meta');
  assert.ok(apply.includes('vNextAdminAssertReleaseTemplateManifest_(release, client)'));
  const pins = functionSource('vNextAdminAssertEmptyPilotPinnedRelease_', 'vNextAdminEmptyPilotRelease_');
  assert.ok(pins.includes('vNextClientRuntimeAssertBoundParent_') &&
    pins.includes('vNextClientRuntimeVerifyPinnedScriptContent_'));
  const assets = functionSource('vNextAdminAssertEmptyPilotReleaseAssets_',
    'vNextAdminFreezeEmptyPilotClient_');
  assert.ok(assets.includes('vNextClientRuntimeVerifyPinnedScriptContent_'),
    'registered source Template releases must allow only exact hash-pinned historical runtimes');
}

function checkPartialHubMetaAppendRecovery() {
  const { sandbox, hub, client } = repairSandbox({
    hubRows: [sourceMeta(), targetMeta()],
    clientRows: [sourceMeta()]
  });
  const repaired = sandbox.vNextAdminAppendEmptyPilotRepairMeta_(
    hub, client, plan(), sourceRelease(), sourceModel(), 'SOURCE'
  );
  assert.equal(repaired.template_version, 'release-source');
  assert.equal(repaired.supersedes_record_id, 'meta-target');
  assert.deepEqual(client.rows.map(row => row.record_id),
    ['meta-source', 'meta-target', repaired.record_id],
    'trusted Hub target meta is mirrored before rollback meta is appended');
  assert.equal(hub.rows.at(-1).record_id, client.rows.at(-1).record_id);
}

function checkClientOnlyMetaFailsClosed() {
  const { sandbox, hub, client } = repairSandbox({
    hubRows: [sourceMeta()],
    clientRows: [sourceMeta(), { ...targetMeta(), record_id: 'untrusted-client-only' }]
  });
  assert.throws(() => sandbox.vNextAdminAppendEmptyPilotRepairMeta_(
    hub, client, plan(), sourceRelease(), sourceModel(), 'SOURCE'
  ), /Client-only BOOK_META/);
  assert.equal(hub.rows.length, 1, 'untrusted Client metadata is never promoted to Hub');
}

function checkRetryJournalSafety() {
  const upgrade = functionSource('vNextAdminUpgradeEmptyPilotClient', 'vNextAdminRecoverEmptyPilotClientUpgrade');
  assert.match(upgrade, /attemptId:\s*Utilities\.getUuid\(\)/,
    'each retry must receive a distinct migration identity');
  const patcher = functionSource('vNextAdminPatchLatestMigration_', 'vNextAdminApplyFileAcl_');
  assert.match(patcher, /\.filter\([\s\S]*?\.slice\(-1\)\[0\]/,
    'migration phase updates must target the latest matching journal row');
  assert.match(source, /unfinishedEmptyUpgradesByBook/,
    'the Admin model must surface unfinished upgrades even after registry commit');
  assert.match(source, /function vNextAdminUpgradeOnlyEligibleEmptyPilotForManualTest\(\)/,
    'a no-UI, uniquely eligible Pilot fallback must remain available for editor execution');
  assert.match(source, /function vNextAdminRecoverOnlyEmptyPilotForManualTest\(\)/,
    'an interrupted Pilot update must have a no-UI recovery fallback');
}

function repairSandbox(input) {
  const hub = { rows: structuredClone(input.hubRows) };
  const client = { rows: structuredClone(input.clientRows) };
  const sandbox = {
    console,
    Logger: { log() {} },
    Utilities: { getUuid: () => 'uuid' }
  };
  vm.createContext(sandbox);
  vm.runInContext(naming, sandbox, { filename: '0_VNext_Naming.js' });
  vm.runInContext(source, sandbox, { filename: 'VNext_Admin.js' });
  sandbox.vNextAdminReadCoreRows_ = (store, sheetName) => {
    assert.equal(sheetName, 'BOOK_META');
    return store.rows.map(row => ({ ...row }));
  };
  sandbox.vNextAdminAppendMissingCoreRows_ = (store, sheetName, idField, records) => {
    assert.equal(sheetName, 'BOOK_META');
    const ids = new Set(store.rows.map(row => String(row[idField] || '')));
    const missing = records.filter(row => !ids.has(String(row[idField] || ''))).map(row => ({ ...row }));
    store.rows.push(...missing);
    return { appended: missing.length };
  };
  sandbox.vNextAdminCanonicalJson_ = value => JSON.stringify(value);
  sandbox.vNextAdminSha256_ = value => createHash('sha256').update(String(value), 'utf8').digest('hex');
  sandbox.vNextAdminActor_ = () => 'admin@example.com';
  return { sandbox, hub, client };
}

function sourceMeta() {
  return {
    record_id: 'meta-source', book_id: 'book-1', client_id: 'client-1',
    client_name: 'Client', fiscal_year: 2027, state: 'INPUT_OPEN',
    template_version: 'release-source', schema_version: 'schema-1',
    model_release_id: 'model-source', supersedes_record_id: ''
  };
}

function targetMeta() {
  return {
    ...sourceMeta(), record_id: 'meta-target', template_version: 'release-target',
    model_release_id: 'model-target', supersedes_record_id: 'meta-source'
  };
}

function plan() {
  return { bookId: 'book-1', sourceMetaRecordId: 'meta-source' };
}

function sourceRelease() {
  return { release_id: 'release-source', schema_version: 'schema-1' };
}

function sourceModel() {
  return { model_release_id: 'model-source' };
}

function functionSource(name, nextName) {
  const start = source.indexOf(`function ${name}(`);
  const end = source.indexOf(`function ${nextName}(`, start + 1);
  assert.ok(start >= 0 && end > start, `could not isolate ${name}`);
  return source.slice(start, end);
}
