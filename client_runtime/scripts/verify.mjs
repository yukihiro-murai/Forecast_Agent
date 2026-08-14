#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { readdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const runtimeRoot = path.resolve(scriptDir, '..');
const directoryArg = readArg('--dir');
const targetDir = directoryArg ? path.resolve(process.cwd(), directoryArg) : path.join(runtimeRoot, 'dist');

const expectedFiles = [
  'Client_Entry.js', 'Client_Core.js', 'Client_Bridge.js', 'VNext_UX.js',
  'VNext_InputSidebar.html', 'VNext_GuidanceSidebar.html', 'VNext_HelpSidebar.html', 'VNext_PlanSidebar.html',
  'VNext_ReviewSidebar.html', 'appsscript.json'
].sort();

const actualFiles = (await readdir(targetDir)).filter((name) => !name.startsWith('.')).sort();
assertEqual(actualFiles, expectedFiles, 'runtime file set');

const contents = {};
for (const name of actualFiles) contents[name] = await readFile(path.join(targetDir, name), 'utf8');
const combined = actualFiles.map((name) => `/* FILE:${name} */\n${contents[name]}`).join('\n');

const forbidden = [
  ['legacy implementation', /Forecast_Agent\.js|initializeAllSheets_|A-1～A-10/],
  ['Admin implementation', /VNext_Admin|vNextBuildAdminMenu_|vNextAdminBootstrap|vNextIsAdminHub_/],
  ['forecast engine', /VNext_Engine|vNextRunForecast_|vNextSimulateForecast_/],
  ['source-system configuration', /FORECAST_SOURCE_SPREADSHEET_ID|VNEXT_ZAC_SOURCE_SPREADSHEET_ID/],
  ['AI runtime configuration', /VERTEX_[A-Z_]+|VertexAI|aiplatform/],
  ['cross-file access API', /DriveApp|UrlFetchApp|SpreadsheetApp\.openById|PropertiesService/],
  ['broad cloud scope', /auth\/cloud-platform|auth\/drive(?:["/])|auth\/script\.projects|auth\/script\.external_request/]
];
for (const [label, pattern] of forbidden) {
  if (pattern.test(combined)) throw new Error(`禁止された${label}がclient runtimeに含まれています: ${pattern}`);
}

const allowedScriptAppMethods = new Set(['getUserTriggers', 'newTrigger', 'EventType']);
for (const match of combined.matchAll(/ScriptApp\.([A-Za-z_$][\w$]*)/g)) {
  if (!allowedScriptAppMethods.has(match[1])) {
    throw new Error(`client runtimeで許可されていないScriptApp APIです: ${match[1]}`);
  }
}

const manifest = JSON.parse(contents['appsscript.json']);
const expectedScopes = [
  'https://www.googleapis.com/auth/script.container.ui',
  'https://www.googleapis.com/auth/script.scriptapp',
  'https://www.googleapis.com/auth/spreadsheets.currentonly',
  'https://www.googleapis.com/auth/userinfo.email'
].sort();
assertEqual((manifest.oauthScopes || []).slice().sort(), expectedScopes, 'minimal OAuth scopes');
if (manifest.runtimeVersion !== 'V8') throw new Error('runtimeVersion must be V8.');

const serverFunctions = new Set();
for (const [name, source] of Object.entries(contents)) {
  if (!name.endsWith('.js')) continue;
  try { new vm.Script(source, { filename: name }); }
  catch (error) { throw new Error(`${name} syntax error: ${error.message}`); }
  for (const match of source.matchAll(/function\s+([A-Za-z_$][\w$]*)\s*\(/g)) serverFunctions.add(match[1]);
}
const requiredFunctions = [
  'onOpen', 'vNextInstalledGuidanceOnOpen', 'vNextInstallAutomaticGuidance',
  'vNextGetClientViewModel', 'vNextPreviewEvidence', 'vNextSaveEvidence',
  'vNextCloseInputAndProceed', 'vNextGetPlanEditorModel', 'vNextPreviewPlan',
  'vNextSubmitPlan', 'vNextGetReviewEditorModel', 'vNextPreviewReview',
  'vNextSaveReview', 'vNextQueueClientForecastRequest'
];
const missingFunctions = requiredFunctions.filter((name) => !serverFunctions.has(name));
if (missingFunctions.length) throw new Error(`必要なserver functionがありません: ${missingFunctions.join(', ')}`);
if (/insertImage\(|assignScript\(/.test(contents['VNext_UX.js'])) {
  throw new Error('ホームに壊れやすい画像ボタンを置かないでください。');
}
if (!/vNextGoHomeAndShowGuidance/.test(contents['VNext_UX.js']) || !contents['VNext_GuidanceSidebar.html']) {
  throw new Error('ホームから再表示できる状態別案内サイドバーがありません。');
}
if (!/officialVintageId:\s*String\(evaluation\.official_vintage_id/.test(contents['VNext_UX.js']) ||
    !/evaluationId:\s*String\(evaluation\.evaluation_id/.test(contents['VNext_UX.js'])) {
  throw new Error('振り返りにcurrent official/evaluation linkageがありません。');
}
if (!/evidenceType:\s*canonical\.responseType[^\n]+COMMITMENT/.test(contents['VNext_UX.js']) ||
    !/\['COMMITMENT',\s*'HUMAN_CHANGE'\]/.test(contents['Client_Core.js'])) {
  throw new Error('契約情報をcommitment lensへ渡すallowlistがありません。');
}

for (const htmlName of actualFiles.filter((name) => name.endsWith('.html'))) {
  const calls = [...contents[htmlName].matchAll(/\.((?:vNext)[A-Za-z0-9_]+)\s*\(/g)].map((match) => match[1]);
  const missing = [...new Set(calls)].filter((name) => !serverFunctions.has(name));
  if (missing.length) throw new Error(`${htmlName} が未定義server functionを呼んでいます: ${missing.join(', ')}`);
  const scripts = [...contents[htmlName].matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)].map((match) => match[1]);
  scripts.forEach((source, index) => {
    try { new vm.Script(source, { filename: `${htmlName}#script${index + 1}` }); }
    catch (error) { throw new Error(`${htmlName} script syntax error: ${error.message}`); }
  });
}

const bundleHash = sha256(expectedFiles.map((name) => `${name}\0${contents[name]}`).join('\0'));
process.stdout.write(`PASS client runtime verification (${expectedFiles.length} files, sha256=${bundleHash})\n`);

function readArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : '';
}

function assertEqual(actual, expected, label) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`${label} mismatch\nactual=${JSON.stringify(actual)}\nexpected=${JSON.stringify(expected)}`);
  }
}

function sha256(value) {
  return createHash('sha256').update(value, 'utf8').digest('hex');
}
