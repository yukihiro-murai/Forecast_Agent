#!/usr/bin/env node

import { readFile, readdir } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const runtimeRoot = path.resolve(scriptDir, '..');
const targetArg = readArg('--dir') || 'src';
const targetDir = path.resolve(runtimeRoot, targetArg);
const expected = ['Portal_Core.js', 'Portal_CreateSidebar.html', 'Portal_UX.js', 'appsscript.json'];
const actual = (await readdir(targetDir)).filter((name) => /\.(?:js|html|json)$/.test(name)).sort();

if (JSON.stringify(actual) !== JSON.stringify(expected)) {
  throw new Error(`Portal runtime allowlist mismatch: ${actual.join(', ')}`);
}

const sources = {};
for (const name of actual) sources[name] = await readFile(path.join(targetDir, name), 'utf8');
const serverSource = sources['Portal_Core.js'] + '\n' + sources['Portal_UX.js'];
const allSource = Object.values(sources).join('\n');

new vm.Script(sources['Portal_Core.js'], { filename: 'Portal_Core.js' });
new vm.Script(sources['Portal_UX.js'], { filename: 'Portal_UX.js' });
const inlineScript = sources['Portal_CreateSidebar.html'].match(/<script>([\s\S]*?)<\/script>/);
if (!inlineScript) throw new Error('Portal creation sidebar inline script is missing.');
new vm.Script(inlineScript[1], { filename: 'Portal_CreateSidebar.inline.js' });

const forbidden = [
  /\bDriveApp\b/, /\bUrlFetchApp\b/, /\bPropertiesService\b/,
  /\bopenById\b/, /script\.external_request/, /cloud-platform/, /script\.projects/,
  /getScriptProperties\s*\(/, /getUserProperties\s*\(/
];
for (const pattern of forbidden) {
  if (pattern.test(allSource)) throw new Error(`Forbidden portal capability found: ${pattern}`);
}

const allowedScriptAppMethods = new Set([
  'getUserTriggers', 'getProjectTriggers', 'newTrigger', 'deleteTrigger', 'EventType'
]);
for (const match of allSource.matchAll(/ScriptApp\.([A-Za-z_$][\w$]*)/g)) {
  if (!allowedScriptAppMethods.has(match[1])) {
    throw new Error(`Portal runtimeで許可されていないScriptApp APIです: ${match[1]}`);
  }
}

const manifest = JSON.parse(sources['appsscript.json']);
const expectedScopes = [
  'https://www.googleapis.com/auth/script.container.ui',
  'https://www.googleapis.com/auth/script.scriptapp',
  'https://www.googleapis.com/auth/spreadsheets.currentonly',
  'https://www.googleapis.com/auth/userinfo.email'
];
if (manifest.runtimeVersion !== 'V8') throw new Error('Portal runtime must use V8.');
if (JSON.stringify((manifest.oauthScopes || []).slice().sort()) !== JSON.stringify(expectedScopes.slice().sort())) {
  throw new Error('Portal runtime OAuth scopes are not the exact minimal allowlist.');
}

const menuItems = [...sources['Portal_UX.js'].matchAll(/\.addItem\('([^']+)',\s*'([^']+)'\)/g)]
  .map((match) => [match[1], match[2]]);
const expectedMenu = [
  ['案内を開く', 'vNextPortalGoHomeAndShowGuidance']
];
if (JSON.stringify(menuItems) !== JSON.stringify(expectedMenu)) {
  throw new Error(`Portal menu must contain only the recovery action: ${JSON.stringify(menuItems)}`);
}

const calledFunctions = [...sources['Portal_CreateSidebar.html'].matchAll(/\.([A-Za-z_$][\w$]*)\s*\(/g)]
  .map((match) => match[1])
  .filter((name) => /^vNextPortal/.test(name));
for (const functionName of new Set(calledFunctions)) {
  if (!new RegExp(`function\\s+${functionName}\\s*\\(`).test(serverSource)) {
    throw new Error(`Sidebar calls undefined server function: ${functionName}`);
  }
}

for (const required of [
  'VN_PORTAL_REQUEST', 'PORTAL_DIRECTORY', 'vNextPortalCanonicalJson_',
  'vNextPortalValidateRequestPayload_', 'vNextPortalBuildDuplicateCheck_',
  "HOME_SHEET: 'ホーム'", "REQUEST_TYPE: 'CREATE_CLIENT_FY_BOOK'"
]) {
  if (!serverSource.includes(required)) throw new Error(`Portal contract marker is missing: ${required}`);
}

if (!/Object\.keys\(value\)\.sort\(\)/.test(sources['Portal_Core.js'])) {
  throw new Error('Canonical JSON must sort object keys.');
}
if (!/computeDigest\(Utilities\.DigestAlgorithm\.SHA_256/.test(sources['Portal_Core.js'])) {
  throw new Error('Portal request hash must use SHA-256.');
}
if (/isForecastOwner|isTeamMember|allowedEmails|emailAllowlist/i.test(serverSource)) {
  throw new Error('Portal runtime must not gate creation by client role or an email allowlist.');
}

process.stdout.write(`PASS portal runtime verification (${actual.length} files, 4 scopes, 1 menu item)\n`);

function readArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : '';
}
