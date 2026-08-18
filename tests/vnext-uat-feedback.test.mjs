#!/usr/bin/env node

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const admin = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
const engine = await readFile(path.join(root, 'VNext_Engine.js'), 'utf8');
const ux = await readFile(path.join(root, 'VNext_UX.js'), 'utf8');
const input = await readFile(path.join(root, 'VNext_InputSidebar.html'), 'utf8');
const plan = await readFile(path.join(root, 'VNext_PlanSidebar.html'), 'utf8');
const review = await readFile(path.join(root, 'VNext_ReviewSidebar.html'), 'utf8');
const guidance = await readFile(path.join(root, 'VNext_GuidanceSidebar.html'), 'utf8');

checkEvidenceMonthNormalization();
checkSafeLivePilotUpgrade();
checkEmployeeInteractionContract();
checkStructuredDashboardContract();
process.stdout.write('PASS vNext live UAT feedback tests (4 groups)\n');

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
    'Append-only integrity mismatch EVIDENCE_EVENT id=c2597e43-a91d-41b4-9439-953b0a73f6bd'), true);
  assert.equal(sandbox.vNextAdminCanRetryKnownEvidencePreflightJob_({
    error: 'Append-only integrity mismatch EVIDENCE_EVENT id=c2597e43-a91d-41b4-9439-953b0a73f6bd',
    attempts: 3
  }), true, 'the exact no-run code defect receives one audited corrective retry');
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
  const appendMissing = functionSource(admin, 'vNextAdminAppendMissingCoreRows_',
    'vNextAdminCanonicalCoreRowForIntegrity_');
  assert.match(appendMissing, /vNextAdminCanonicalCoreRowForIntegrity_\(sheetName, row\)/,
    'append-only sync must compare the same lossless canonical evidence representation');
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
  assert.match(boundary, /vNextAdminCanRetryKnownEvidencePreflightJob_\(row\)/,
    'migration eligibility and requeue must share the same exact preflight failure predicate');
  assert.match(boundary, /vNextAdminIsKnownEvidenceRequeuedJob_\(hub, row\)/,
    'migration may coexist only with the exact unlocked queued retry it preserves');
  const editorFallback = functionSource(admin,
    'vNextAdminUpgradeOnlyKnownFailedPreflightPilotForManualTest',
    'vNextAdminRecoverFailedPreflightPilotClientUpgrade');
  assert.equal(editorFallback.includes('dryRun: true'), false,
    'the editor fallback must not release a lock between slow dry-run and apply');
  const stateValidator = functionSource(admin, 'vNextAdminValidateClientStateRows_',
    'vNextAdminVerifyHistoricClientStatePrefix_');
  assert.match(stateValidator, /vNextAdminVerifyHistoricClientStatePrefix_/,
    'early Pilot Client state prefixes must be verified before health checks');
  const recoveryFallback = functionSource(admin,
    'vNextAdminRecoverOnlyFailedPreflightPilotForManualTest',
    'vNextAdminAssertEmptyPilotUpgradeEligibility_');
  assert.match(recoveryFallback, /direction:\s*'TARGET'/,
    'manual recovery must finish the requested current-release upgrade after a verified rollback');
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
  assert.equal((input.match(/vNextSaveEvidence\(payload\(\)\)/g) || []).length, 1,
    'The progressive input must still perform one server save only');
  assert.match(input, /data-step="response"/);
  assert.match(input, /data-step="detail"/);
  assert.match(input, /data-step="impact"/);
  assert.match(input, /data-step="evidence"/);
  assert.match(input, /function showStep\(key\)/,
    'Evidence questions must be disclosed one step at a time');
  assert.match(ux, /vNextUxAutoOpenGuidance_/,
    'A state-aware sidebar must open automatically for passive employees');
  const openMenu = functionSource(ux, 'vNextBuildClientMenu_', 'vNextSetupClientExperience_');
  assert.doesNotMatch(openMenu, /vNextUxGetBookContext_|vNextRefreshEmployeeViews/,
    'Client onOpen must not block guidance on identity or full view rendering');
  assert.doesNotMatch(openMenu, /vNextUxOpenGuidanceShellQuietly_|showSidebar/,
    'Simple onOpen must not call authorized Ui.showSidebar');
  assert.doesNotMatch(openMenu, /vNextUxActivateSheet_/,
    'Simple onOpen must not activate sheets before the menu can appear');
  assert.match(openMenu, /addItem\('案内を開く'/);
  assert.doesNotMatch(openMenu, /自分の情報を入力|予測ダッシュボード|使い方・困ったとき/,
    'Daily employee actions must live in the guidance sidebar, not the top menu');
  const hubOpen = functionSource(ux, 'vNextHandleOnOpen_', 'vNextBuildClientMenu_');
  assert.match(hubOpen, /vNextAdminLooksLikeHub_/);
  assert.doesNotMatch(hubOpen, /vNextIsAdminHub_/,
    'Hub onOpen must not wait on registered-hub verification before the menu appears');
  assert.match(ux, /function vNextInstalledGuidanceOnOpen\(e\)[\s\S]*vNextUxOpenGuidanceShellQuietly_/,
    'Automatic guidance must run from an installable open trigger');
  assert.match(ux, /ScriptApp\.getProjectTriggers\(\)/,
    'Automatic guidance must be a project-level open trigger');
  assert.match(ux, /ScriptApp\.newTrigger\(handler\)\.forSpreadsheet\(spreadsheet\)\.onOpen\(\)\.create\(\)/,
    'The first authorized sidebar open must enable automatic guidance once');
  assert.match(guidance, /vNextRefreshEmployeeViews\(\)/,
    'Guidance must refresh the visible sheets asynchronously after it is rendered');
  const scaleResolver = functionSource(admin, 'vNextAdminResolveClientAnnualSalesScale_',
    'vNextAdminRefreshClientAnnualSalesScale_');
  assert.match(scaleResolver, /asOf:\s*asOf \|\| new Date\(\)/,
    'relative bands must validate actuals against the forecast as-of, not treat cutoff as a new as-of');
  assert.match(scaleResolver, /cutoff:\s*cutoff \|\| vNextAdminCutoffFromAsOf_/);
}

function checkStructuredDashboardContract() {
  assert.match(ux, /ANALYTICS_SHEET:\s*'VN_ANALYTICS_FACT'/);
  assert.match(ux, /var VNEXT_UX_ANALYTICS_HEADERS_/);
  for (const field of ['record_type', 'period_type', 'amount_yen', 'p10_yen', 'p50_yen', 'p90_yen', 'source_table']) {
    assert.match(ux, new RegExp(`['"]${field}['"]`), `analytics projection missing ${field}`);
  }
  assert.match(ux, /function vNextOpenForecastDashboard\(\)[\s\S]*showModelessDialog/,
    'forecast dashboard must use a wide modeless surface rather than a cell chart');
  assert.match(guidance, /data-panel="dashboard"/);
  assert.match(guidance, /data-panel="triangulation"/);
  assert.match(guidance, /独立した予測アプローチ/,
    'Triangulation must compare independent methods rather than cumulative layers');
  for (const method of ['直近3年度の加重平均', '線形回帰トレンド', '減衰CAGR', '統合シミュレーション']) {
    assert.match(engine, new RegExp(method), `Engine must expose ${method}`);
  }
  assert.doesNotMatch(guidance, /shortMoney/,
    'All employee-facing currency must use full yen notation');
  assert.match(guidance, /grid-template-columns:1fr/,
    'Dashboard information groups must use a single reading column');
  assert.match(guidance, /\.metric\.primary\{[^}]*background:#f1f5f8[^}]*color:var\(--ink\)/,
    'Primary forecast card must use a light high-contrast surface');
  assert.match(guidance, /\.primary-action\{/,
    'Footer button styling must use a class that cannot collide with the forecast card');
  assert.doesNotMatch(guidance, /(^|})\.primary\{/m,
    'A generic primary class must not recolor unrelated dashboard components');
  assert.match(guidance, /body\{[^}]*font:15px\/1\.6/,
    'Dashboard body copy must remain readable without oversized typography');
  assert.doesNotMatch(guidance, /repeat\(2|minmax\(0,1fr\).*minmax\(0,1fr\)/,
    'Layer, timing and AI cards must not use dense two-column grids');
  assert.match(guidance, /preserveAspectRatio="xMidYMid meet"/,
    'Monthly graph must preserve its geometry instead of stretching');
  assert.match(guidance, /data-panel="timing"/);
  assert.match(guidance, /id="openWideButton"/,
    'The compact automatic guidance must offer the detailed dashboard as a separate view');
  assert.match(guidance, /id="monthChart"/);
  for (const html of [input, plan, review, guidance]) {
    assert.doesNotMatch(html, /setTimeout\s*\([^)]*google\.script\.host\.close/,
      'successful employee operations must not close their surface automatically');
  }
  assert.match(input, /Math\.trunc/);
  assert.match(plan, /Math\.trunc/);
  assert.match(review, /Math\.trunc/);

  const sandbox = { console, Logger: { log() {} } };
  vm.createContext(sandbox);
  vm.runInContext(ux, sandbox, { filename: 'VNext_UX.js' });
  const forecast = sandbox.vNextUxPublicForecast_({
    fiscalYear: 2027,
    p10: 80.9,
    p50: 100.9,
    p90: 120.9,
    layers: { historyBaseline: 90.8, objectiveForecast: 95.7, systemRecommended: 100.9 },
    months: Array.from({ length: 12 }, (_, index) => ({ month: index + 1, p10: 6.9, p50: 8.9, p90: 10.9 }))
  });
  assert.equal(forecast.center, 100);
  assert.equal(forecast.p10, 80);
  assert.equal(forecast.p90, 120);
  assert.equal(forecast.months.reduce((sum, item) => sum + item.p50, 0), 100,
    'whole-yen monthly values must reconcile to the whole-yen annual forecast');
  assert.equal(forecast.quarters.reduce((sum, item) => sum + item.p50, 0), 100,
    'whole-yen quarters must reconcile to the whole-yen annual forecast');
}

function functionSource(source, name, nextName) {
  const start = source.indexOf(`function ${name}(`);
  const end = source.indexOf(`function ${nextName}(`, start + 1);
  assert.ok(start >= 0 && end > start, `could not isolate ${name}`);
  return source.slice(start, end);
}
