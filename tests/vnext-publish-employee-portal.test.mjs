#!/usr/bin/env node

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const publish = await readFile(path.join(root, 'scripts/publish_employee_portal.mjs'), 'utf8');
const login = await readFile(path.join(root, 'scripts/gas_agent_login.mjs'), 'utf8');
const oauth = await readFile(path.join(root, 'scripts/lib/gas_oauth.mjs'), 'utf8');
const admin = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
const ignore = await readFile(path.join(root, '.claspignore'), 'utf8');
const gitignore = await readFile(path.join(root, '.gitignore'), 'utf8');
const targets = JSON.parse(await readFile(path.join(root, 'deploy/targets.json'), 'utf8'));

assert.ok(ignore.includes('scripts/**') && ignore.includes('deploy/**'),
  'clasp must not push agent publish scripts into the Admin project');
assert.ok(gitignore.includes('deploy/.gas-oauth.json'),
  'Agent OAuth tokens must stay out of git');
assert.equal(gitignore.includes('client_secret'), false);
assert.ok(targets.hubScriptId && targets.hubSpreadsheetId && targets.centralScriptId &&
  targets.portalSpreadsheetId && targets.quotaProjectId,
  'deploy/targets.json must name Hub, central, portal spreadsheet, and quota project');
assert.equal('portalScriptId' in targets, true,
  'portal script id may be filled from Hub at runtime, but the key must exist');

assert.ok(oauth.includes('https://www.googleapis.com/auth/spreadsheets') &&
  oauth.includes('https://www.googleapis.com/auth/script.webapp.deploy'),
  'Agent login must request Sheets plus web app deploy scopes');
assert.equal(login.includes('forecast') && login.includes('承認'), true);
assert.ok(login.includes('Does not grant forecast approval') ||
  login.includes('計画の承認権限は含まれません'),
  'Login must not be described as forecast approval');

assert.ok(publish.includes("function: 'vNextAdminUpdateSharedPortalRuntimeFromSource'"),
  'Cursor publish must call the Hub function that opens Hub by ID');
assert.ok(publish.includes('confirmRuntimeDrift'),
  'Cursor publish must pass through pin-drift confirmation');
assert.ok(publish.includes('/deployments/') && publish.includes("'PUT'"),
  'Direct fallback must update the existing /exec deployment');
assert.equal(publish.includes("projects/${scriptId}/deployments', {"), false,
  'Direct fallback must not POST a new Web App deployment');
assert.ok(publish.includes('skip-source-sync') && publish.includes('ADMIN_SOURCE_FILES'),
  'Cursor publish must copy the verified 18 Admin files unless skipped');
assert.equal(publish.includes('vNextAdminApprove'), false,
  'Portal publish must not call forecast approval');
assert.equal(publish.includes('RESET_GENERATED_CLIENTS'), false,
  'Portal publish must not reset generated clients');

const fromSourceStart = admin.indexOf('function vNextAdminUpdateSharedPortalRuntimeFromSource(');
const fromSourceEnd = admin.indexOf('function vNextAdminUpdateSharedPortalRuntimeInHub_', fromSourceStart);
const fromSource = admin.slice(fromSourceStart, fromSourceEnd);
assert.ok(fromSourceStart >= 0 && fromSourceEnd > fromSourceStart);
assert.ok(fromSource.includes('SpreadsheetApp.openById(hubId)'),
  'Agent entry must not depend on SpreadsheetApp.getActiveSpreadsheet()');
assert.ok(fromSource.includes('vNextAdminAssertHubAdmin_(hub, true)'),
  'Agent entry must authorize the effective user, not a browser session');
assert.ok(fromSource.includes('vNextAdminUpdateSharedPortalRuntimeInHub_(hub, req)'),
  'Agent entry must reuse the same verified Hub migration');

process.stdout.write('vnext-publish-employee-portal.test.mjs PASS\n');
