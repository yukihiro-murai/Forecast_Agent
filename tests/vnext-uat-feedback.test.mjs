#!/usr/bin/env node

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const admin = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
const ux = await readFile(path.join(root, 'VNext_UX.js'), 'utf8');
const input = await readFile(path.join(root, 'VNext_InputSidebar.html'), 'utf8');

checkEvidenceMonthNormalization();
checkSafeLivePilotUpgrade();
checkEmployeeInteractionContract();
process.stdout.write('PASS vNext live UAT feedback tests (3 groups)\n');

function checkEvidenceMonthNormalization() {
  const sandbox = {
    console,
    Logger: { log() {} },
    Utilities: {
      formatDate(value) {
        const date = new Date(value);
        return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
      }
    },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' }
  };
  vm.createContext(sandbox);
  vm.runInContext(admin, sandbox, { filename: 'VNext_Admin.js' });
  assert.equal(sandbox.vNextAdminNormalizeEvidenceMonth_('2027-04', 'target_start_month'), '2027-04');
  const gasDate = vm.runInContext("new Date('2027-04-01T00:00:00Z')", sandbox);
  assert.equal(sandbox.vNextAdminNormalizeEvidenceMonth_(gasDate,
    'target_start_month'), '2027-04');
  const accepted = sandbox.vNextAdminCanonicalEvidenceForComparison_({
    evidence_id: 'ev-1', target_start_month: gasDate, target_end_month: '2028-03'
  });
  const submitted = sandbox.vNextAdminCanonicalEvidenceForComparison_({
    evidence_id: 'ev-1', target_start_month: '2027-04', target_end_month: '2028-03'
  });
  assert.equal(JSON.stringify(accepted), JSON.stringify(submitted),
    'Hub Date values and Client YYYY-MM strings must compare as identical evidence');
  assert.equal(sandbox.vNextAdminIsKnownEvidencePreflightFailure_(
    'invalid evidence month: target_start_month'), true);
  assert.equal(sandbox.vNextAdminIsKnownEvidencePreflightFailure_(
    'Client evidence differs from the already accepted Hub record: c2597e43-a91d-41b4-9439-953b0a73f6bd'), true);
  assert.equal(sandbox.vNextAdminIsKnownEvidencePreflightFailure_(
    'Client evidence differs from the already accepted Hub record: bad id!'), false);
  assert.throws(() => sandbox.vNextAdminNormalizeEvidenceMonth_('April 2027',
    'target_start_month'), /invalid evidence month/);
  const validator = functionSource(admin, 'vNextAdminValidateClientEvidenceRows_',
    'vNextAdminNormalizeEvidenceMonth_');
  assert.ok(validator.indexOf('vNextAdminNormalizeEvidenceMonth_') <
    validator.indexOf("/^\\d{4}-\\d{2}$/"),
  'Sheets Date coercion must be normalized before the strict evidence-month check');
  assert.match(validator, /vNextAdminCanonicalEvidenceForComparison_\(existingById\.get\(evidenceId\)\)/,
    'already accepted Hub evidence must receive the same lossless month normalization');
}

function checkSafeLivePilotUpgrade() {
  assert.match(admin, /function vNextAdminUpgradeFailedPreflightPilotClient\(request\)/);
  assert.match(admin, /function vNextAdminRecoverFailedPreflightPilotClientUpgrade\(request\)/);
  assert.match(admin, /function vNextAdminUpgradeOnlyKnownFailedPreflightPilotForManualTest\(\)/);
  const boundary = functionSource(admin, 'vNextAdminAssertFailedPreflightBusinessBoundary_',
    'vNextAdminUpgradeOnlyKnownFailedPreflightPilotForManualTest');
  for (const token of [
    'READY_TO_RUN', 'FORECAST_RUN', 'PLAN_VERSION', 'EVALUATION',
    'VN_ADMIN_SHEETS.APPROVALS', 'VN_ADMIN_SHEETS.OFFICIAL',
    "status === 'RUNNING'",
    "latest.event_type || '').toUpperCase() !== 'FAILED'"
  ]) assert.ok(boundary.includes(token), `failed-preflight boundary missing ${token}`);
  assert.match(boundary, /vNextAdminIsKnownEvidencePreflightFailure_\(row\.error\)/,
    'migration eligibility and requeue must share the same exact preflight failure predicate');
  assert.match(boundary, /vNextAdminIsKnownEvidenceRequeuedJob_\(hub, row\)/,
    'migration may coexist only with the exact unlocked queued retry it preserves');
  const editorFallback = functionSource(admin,
    'vNextAdminUpgradeOnlyKnownFailedPreflightPilotForManualTest',
    'vNextAdminRecoverFailedPreflightPilotClientUpgrade');
  assert.equal(editorFallback.includes('dryRun: true'), false,
    'the editor fallback must not release a lock between slow dry-run and apply');
  const apply = functionSource(admin, 'vNextAdminApplyEmptyPilotRelease_',
    'vNextAdminAppendEmptyPilotRepairMeta_');
  assert.match(apply, /const preservedState = String\(plan\.preservedState \|\| 'INPUT_OPEN'\)/);
  assert.ok(apply.indexOf('vNextClientRuntimeCopyScriptContent_') <
    apply.indexOf('vNextAdminPatchRegistryByBookId_'),
  'runtime/UI/config/meta must be durable before the registry commit marker');
  const sourcePins = functionSource(admin, 'vNextAdminAssertEmptyPilotPinnedRelease_',
    'vNextAdminEmptyPilotRelease_');
  assert.match(sourcePins, /vNextClientRuntimeVerifyPinnedScriptContent_/,
    'the known failed Pilot may migrate only from an exact SHA-pinned historical runtime');
}

function checkEmployeeInteractionContract() {
  assert.equal(/insertImage\s*\(/.test(ux), false,
    'Home must not use the transparent over-grid image that rendered as a red bar');
  assert.match(ux, /key:\s*'small'[\s\S]*?base \* 0\.005[\s\S]*?base \* 0\.02/);
  assert.match(ux, /key:\s*'medium'[\s\S]*?base \* 0\.02[\s\S]*?base \* 0\.05/);
  assert.match(ux, /key:\s*'large'[\s\S]*?base \* 0\.05[\s\S]*?base \* 0\.10/);
  assert.equal(/id=["']previewButton["']/.test(input), false,
    'The employee must not need a separate confirm click before saving');
  assert.equal((input.match(/id=["']saveButton["']/g) || []).length, 1);
  assert.match(ux, /vNextUxAutoOpenGuidance_/,
    'A state-aware sidebar must open automatically for passive employees');
}

function functionSource(source, name, nextName) {
  const start = source.indexOf(`function ${name}(`);
  const end = source.indexOf(`function ${nextName}(`, start + 1);
  assert.ok(start >= 0 && end > start, `could not isolate ${name}`);
  return source.slice(start, end);
}
