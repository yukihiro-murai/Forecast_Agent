#!/usr/bin/env node

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const admin = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');

const upgrade = isolate('vNextAdminUpgradeDraftReadyPilotUx',
  'vNextAdminResolveDraftReadyPilotUxUpgrade_');
const resolve = isolate('vNextAdminResolveDraftReadyPilotUxUpgrade_',
  'vNextAdminAssertDraftReadyPilotUxBoundary_');
const boundary = isolate('vNextAdminAssertDraftReadyPilotUxBoundary_',
  'vNextAdminUpgradeOnlyDraftReadyPilotUxForManualTest');
const manual = isolate('vNextAdminUpgradeOnlyDraftReadyPilotUxForManualTest',
  'vNextAdminRecoverDraftReadyPilotUxUpgrade');
const recover = isolate('vNextAdminRecoverDraftReadyPilotUxUpgrade',
  'vNextAdminRecoverOnlyDraftReadyPilotUxForManualTest');
const pinnedRelease = isolate('vNextAdminAssertEmptyPilotPinnedRelease_',
  'vNextAdminEmptyPilotRelease_');

assert.match(upgrade, /DRAFT_READY_PILOT_UX_UPGRADE_V1/);
assert.match(upgrade, /preservedState:\s*'DRAFT_READY'/);
assert.match(upgrade, /vNextAdminApplyEmptyPilotRelease_\([\s\S]*resolved\.targetRelease/);
assert.match(upgrade, /resolved\.sourceRelease[\s\S]*'SOURCE'/,
  'upgrade failure must restore the exact source release');
assert.ok(upgrade.indexOf('vNextAdminPatchLatestMigration_') >
  upgrade.indexOf('vNextAdminApplyEmptyPilotRelease_'),
  'journal may succeed only after runtime/UI/config/registry/health apply');

assert.match(resolve, /vNextAdminAssertEmptyPilotPinnedRelease_[\s\S]*'DRAFT_READY'/);
assert.match(resolve, /String\(sourceRelease\.schema_version/);
assert.match(resolve, /vNextAdminAssertEmptyPilotReleaseAssets_\(sourceRelease\)/);
assert.match(resolve, /vNextAdminAssertEmptyPilotReleaseAssets_\(targetRelease\)/);

assert.match(boundary, /\['PLAN_VERSION', 'EVALUATION'\]/);
assert.match(boundary, /VN_ADMIN_SHEETS\.APPROVALS/);
assert.match(boundary, /VN_ADMIN_SHEETS\.OFFICIAL/);
assert.match(boundary, /\['QUEUED', 'RUNNING'\]/);
assert.match(boundary, /related_run_id/);
assert.match(boundary, /String\(hubRun\.status[\s\S]*'SUCCESS'/);
assert.match(boundary, /input_data_hash/);
assert.match(boundary, /Number\(hubRun\.p50/);

assert.match(manual, /access_policy[\s\S]*INTERNAL_OPEN/,
  'the zero-input live helper must exclude legacy PRIVATE DRAFT_READY books');
assert.match(manual, /template_release_id[\s\S]*pair\.releaseId/);

assert.match(pinnedRelease, /VN_ADMIN_CLIENT_STATES\.indexOf/,
  'BOOK_META stores a release snapshot; current workflow state must come from STATE_EVENT');
assert.match(pinnedRelease,
  /vNextAdminLatestClientState_\(hub[\s\S]*expectedState[\s\S]*vNextAdminLatestClientState_\(client/,
  'the latest Hub and Client STATE_EVENT must still equal the requested workflow state');

assert.match(recover, /DRAFT_READY_PILOT_UX_UPGRADE_V1/);
assert.match(recover, /registry\.template_release_id[\s\S]*plan\.targetReleaseId/);
assert.match(recover, /direction === 'TARGET'/);
assert.match(recover, /vNextAdminApplyEmptyPilotRelease_/);

process.stdout.write('PASS vNext DRAFT_READY Pilot UX upgrade contracts\n');

function isolate(name, nextName) {
  const start = admin.indexOf(`function ${name}(`);
  const end = admin.indexOf(`function ${nextName}(`, start + 1);
  assert.ok(start >= 0 && end > start, `could not isolate ${name}`);
  return admin.slice(start, end);
}
