#!/usr/bin/env node

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const source = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
const sidebar = await readFile(path.join(root, 'VNext_AdminSidebar.html'), 'utf8');
const sandbox = { console, Logger: { log() {} } };
vm.createContext(sandbox);
vm.runInContext(await readFile(path.join(root, '0_VNext_Naming.js'), 'utf8'), sandbox, { filename: '0_VNext_Naming.js' });
vm.runInContext(source, sandbox, { filename: 'VNext_Admin.js' });

const safePortalRetry = {
  job_type: 'PORTAL_PROVISION_CLIENT', status: 'FAILED', attempts: 1,
  error: 'Requested release is not ACTIVE: vnext-pilot-20260812'
};
assert.equal(sandbox.vNextAdminIsKnownSafeRetryCandidate_(safePortalRetry), true);
assert.equal(sandbox.vNextAdminIsKnownSafeRetryCandidate_({ ...safePortalRetry, attempts: 3 }), false);
assert.equal(sandbox.vNextAdminIsKnownSafeRetryCandidate_({
  job_type: 'FORECAST_REQUEST', status: 'FAILED', attempts: 1,
  error: 'At least 5 fiscal years of confirmed actual history are required; found 2.'
}), false);

const mismatchedException = sandbox.vNextAdminExceptionForSidebar_({
  exception_type: 'JOB_FAILED', source_ref: 'CURRENT-UNSAFE', book_id: 'BOOK-1'
}, {
  'BOOK-1': { spreadsheet_url: 'https://docs.google.com/spreadsheets/d/BOOK_1/edit' }
}, [
  { ...safePortalRetry, job_id: 'OLDER-SAFE', target_book_id: 'BOOK-1' },
  { job_id: 'CURRENT-UNSAFE', target_book_id: 'BOOK-1', status: 'FAILED', attempts: 3,
    job_type: 'PORTAL_PROVISION_CLIENT', error: 'Unknown failure' }
]);
assert.equal(mismatchedException.actionType, 'OPEN_BOOK',
  'An exception must not inherit a retry action from another job for the same book');
const portalException = sandbox.vNextAdminExceptionForSidebar_({
  exception_type: 'PORTAL_PROVISION_FAILED', source_ref: 'PORTAL-JOB', book_id: 'REQ-1'
}, {}, [{ ...safePortalRetry, job_id: 'PORTAL-JOB', target_book_id: 'REQ-1',
  request_json: '{"clientName":"AstraZeneca","fiscalYear":2027}' }]);
assert.equal(portalException.clientName, 'AstraZeneca');
assert.equal(portalException.fiscalYear, 2027);
assert.equal(portalException.actionType, 'RUN_NOW');

sandbox.vNextAdminResolvePortalForRead_ = () => ({
  spreadsheet: { getUrl: () => 'https://docs.google.com/spreadsheets/d/PORTAL_1/edit' }
});
sandbox.vNextAdminReadTable_ = () => ({ rows: [
  { request_id: 'REQ-1', event_type: 'REQUESTED', status: 'PENDING', client_name: 'Client', fiscal_year: 2027 },
  { request_id: 'REQ-1', event_type: 'REQUESTED', status: 'COMPLETED', client_name: 'Tampered', fiscal_year: 2099 },
  { request_id: 'REQ-2', event_type: 'CREATION_STARTED', status: 'CREATING', client_name: 'Client 2', fiscal_year: 2027 }
] });
const portalProjection = sandbox.vNextAdminPortalRequestsForSidebar_({});
assert.deepEqual(
  { waiting: portalProjection.counts.waiting, processing: portalProjection.counts.processing,
    failed: portalProjection.counts.failed, completed: portalProjection.counts.completed },
  { waiting: 1, processing: 1, failed: 0, completed: 0 }
);
assert.equal(portalProjection.attention[1].clientName, 'Client');

assert.equal(sandbox.vNextAdminAttentionSummary_({
  automationInstalled: true,
  counts: { exceptions: 0, pendingApprovals: 0, portalAttention: 0, queuedJobs: 1, runningJobs: 0 },
  operations: { schedulerStale: false }, portalRequests: { counts: { waiting: 0, processing: 0 } }
}, 0).status, 'PROCESSING');
assert.equal(sandbox.vNextAdminAttentionSummary_({
  automationInstalled: true,
  counts: { exceptions: 0, pendingApprovals: 1, portalAttention: 0, queuedJobs: 0, runningJobs: 0 },
  operations: { schedulerStale: false }, portalRequests: { counts: { waiting: 0, processing: 0 } }
}, 0).status, 'ATTENTION');
assert.equal(sandbox.vNextAdminAttentionSummary_({
  automationInstalled: true,
  counts: { exceptions: 0, pendingApprovals: 0, portalAttention: 0, queuedJobs: 1, runningJobs: 0 },
  operations: { schedulerStale: false, queueStale: true }, portalRequests: { counts: {} }
}, 0).status, 'ERROR');
assert.equal(sandbox.vNextAdminAttentionSummary_({
  automationInstalled: true,
  counts: { exceptions: 0, pendingApprovals: 0, portalAttention: 0, queuedJobs: 0, runningJobs: 0 },
  operations: { schedulerStale: false, queueStale: false }, portalRequests: { unavailable: true, counts: {} }
}, 0).status, 'ATTENTION');

const menuStart = source.indexOf('function vNextBuildAdminMenu_()');
const menuEnd = source.indexOf('function vNextBuildLegacySetupMenu_()', menuStart);
const menuSource = source.slice(menuStart, menuEnd);
assert.equal((menuSource.match(/\.addItem\(/g) || []).length, 5,
  'The normal Admin menu is 案内を開く + two deploy shortcuts + nested irregular ops');
assert.match(menuSource, /addSubMenu/);
assert.doesNotMatch(menuSource, /vNextDetectBookMode_|vNextAdminIsRegisteredHub_|vNextAdminHydrateLocalRuntime_/,
  'Hub menu construction must not wait on config, registry, or property hydration');
assert.match(menuSource, /vNextAdminLooksLikeHub_/);
assert.ok(source.includes('function vNextAdminGetSidebarDetailModel()') &&
  sidebar.includes('vNextAdminGetSidebarDetailModel'),
  'Hub sidebar first paint must not wait on the deferred detail RPC');
assert.doesNotMatch(sidebar, /id="adminPanel" class="hidden"/);
assert.equal(sandbox.vNextAdminIsRegisteredHubFromRows_({ getId: () => 'HUB' }, {
  mode: 'ADMIN', book_id: 'B1'
}, [{ book_id: 'B1', mode: 'ADMIN', spreadsheet_id: 'HUB' }]), true);
assert.equal(sandbox.vNextAdminIsRegisteredHubFromRows_({ getId: () => 'HUB' }, {
  mode: 'ADMIN', book_id: 'B1'
}, [{ book_id: 'B1', mode: 'CLIENT', spreadsheet_id: 'HUB' }]), false);
assert.equal(/Pilot|runtime|Release|\u5f85\u6a5fjob/.test(menuSource), false);
assert.ok(source.includes('function vNextAdminRunOperationalCycle()') &&
  source.includes("vNextAdminWithScriptLock_('admin-run-operational-cycle'") &&
  source.includes('vNextAdminRequeueKnownPilotFailures_(hub)') &&
  source.includes("vNextAdminWriteAudit_(hub, 'RUN_OPERATIONAL_CYCLE'"));

const scriptMatch = sidebar.match(/<script>([\s\S]*?)<\/script>/i);
assert.ok(scriptMatch);
new vm.Script(scriptMatch[1], { filename: 'VNext_AdminSidebar.html' });
const ids = [...sidebar.matchAll(/\bid="([^"]+)"/g)].map(match => match[1]);
assert.deepEqual(ids.filter((id, index) => ids.indexOf(id) !== index), []);
assert.ok(sidebar.includes('<html lang="ja">') && sidebar.includes('id="attentionOverview"') && sidebar.includes('id="portalRequests"') &&
  sidebar.includes('id="recentJobs"') && sidebar.includes('window.confirm(confirmation)'));
assert.equal(/<div class="metric">\u5f85\u6a5fjob/.test(sidebar), false);
assert.equal(sidebar.includes('enqueueMigration(false)'), false,
  'The disabled migration APPLY action must not be exposed in the Admin sidebar');
assert.ok(sidebar.includes('dataset.recoveryRequired') &&
  sidebar.includes('中断した更新を復旧'),
  'an interrupted empty-Pilot upgrade must remain recoverable from the normal Admin UI');

// google.script.run silently drops private (underscore-suffixed) server
// functions: neither handler fires and the UI hangs forever. Ban them from
// every HTML surface so this whole bug class cannot come back.
const { readdir } = await import('node:fs/promises');
const htmlSurfaces = [];
for (const dir of ['.', 'portal_runtime/src', 'client_runtime/src']) {
  for (const name of await readdir(path.join(root, dir))) {
    if (name.endsWith('.html')) htmlSurfaces.push(path.join(dir, name));
  }
}
assert.ok(htmlSurfaces.length >= 8, 'expected to scan the sidebar/entry HTML surfaces');
for (const file of htmlSurfaces) {
  const html = await readFile(path.join(root, file), 'utf8');
  const privateCalls = [...html.matchAll(/\.\s*(vNext\w*_)\s*\(/g)].map(match => match[1]);
  assert.deepEqual(privateCalls, [],
    file + ' must not call private (underscore-suffixed) server functions from google.script.run');
}

process.stdout.write('PASS vNext Admin decision UX contracts\n');
