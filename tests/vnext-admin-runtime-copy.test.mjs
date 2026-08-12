#!/usr/bin/env node

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const contract = [
  ['Forecast_Agent', 'SERVER_JS', 'Forecast_Agent.js'],
  ['VNext_AI', 'SERVER_JS', 'VNext_AI.js'],
  ['VNext_Admin', 'SERVER_JS', 'VNext_Admin.js'],
  ['VNext_AdminSidebar', 'HTML', 'VNext_AdminSidebar.html'],
  ['VNext_ClientRuntimeBundle', 'SERVER_JS', 'VNext_ClientRuntimeBundle.js'],
  ['VNext_ClientRuntimeProvisioning', 'SERVER_JS', 'VNext_ClientRuntimeProvisioning.js'],
  ['VNext_Core', 'SERVER_JS', 'VNext_Core.js'],
  ['VNext_Engine', 'SERVER_JS', 'VNext_Engine.js'],
  ['VNext_HelpSidebar', 'HTML', 'VNext_HelpSidebar.html'],
  ['VNext_InputSidebar', 'HTML', 'VNext_InputSidebar.html'],
  ['VNext_PlanSidebar', 'HTML', 'VNext_PlanSidebar.html'],
  ['VNext_PortalPilotRecovery', 'SERVER_JS', 'VNext_PortalPilotRecovery.js'],
  ['VNext_PortalRuntimeBundle', 'SERVER_JS', 'VNext_PortalRuntimeBundle.js'],
  ['VNext_ReviewSidebar', 'HTML', 'VNext_ReviewSidebar.html'],
  ['VNext_Tests', 'SERVER_JS', 'VNext_Tests.js'],
  ['VNext_UX', 'SERVER_JS', 'VNext_UX.js'],
  ['appsscript', 'JSON', 'appsscript.json']
];
const files = await Promise.all(contract.map(async ([name, type, filename]) => ({
  name, type, source: await readFile(path.join(root, filename), 'utf8')
})));
const sourceId = 'ADMIN_source_script_1234567890ABCDE';
const targetId = 'ADMIN_target_script_1234567890ABCDE';
const spreadsheetId = 'ADMIN_target_sheet_1234567890ABCDE';
const folderId = 'ADMIN_target_folder_1234567890ABCDE';
const expectedSha = canonicalSha(files);
let renamed = '';
let movedFolder = '';
const sandbox = {
  console,
  Logger: { log() {} },
  Session: {
    getActiveUser: () => ({ getEmail: () => 'admin@example.com' }),
    getEffectiveUser: () => ({ getEmail: () => 'admin@example.com' })
  },
  ScriptApp: {
    getScriptId: () => sourceId,
    getOAuthToken: () => 'unused-test-token'
  },
  SpreadsheetApp: {
    create: title => ({
      getId: () => spreadsheetId,
      getUrl: () => `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
      title
    })
  },
  DriveApp: {
    getFileById: id => ({
      moveTo(folder) { movedFolder = folder.id; },
      setName(value) { renamed = value; },
      id
    }),
    getFolderById: id => ({ id })
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
  await readFile(path.join(root, 'VNext_ClientRuntimeProvisioning.js'), 'utf8'),
  sandbox,
  { filename: 'VNext_ClientRuntimeProvisioning.js' }
);

testAdminCopy();
testAdminGuards();
testAdminCreate();
process.stdout.write('PASS vNext Admin runtime copy tests (3 groups)\n');

function testAdminCopy() {
  const calls = [];
  sandbox.vNextClientRuntimeApiRequest_ = (apiPath, method, body) => {
    calls.push({ apiPath, method, body });
    if (apiPath === `/projects/${encodeURIComponent(targetId)}` && method === 'get') {
      return { scriptId: targetId, parentId: spreadsheetId, title: 'Known Hub' };
    }
    if (apiPath === `/projects/${encodeURIComponent(sourceId)}/content` && method === 'get') {
      return { scriptId: sourceId, files: clone(files).reverse() };
    }
    if (apiPath === `/projects/${encodeURIComponent(targetId)}/content` && method === 'put') {
      return { scriptId: targetId, files: clone(body.files).reverse() };
    }
    throw new Error(`Unexpected API call ${method} ${apiPath}`);
  };
  const result = sandbox.vNextAdminRuntimeCopyScriptContent_(sourceId, targetId, spreadsheetId);
  assert.equal(result.ok, true);
  assert.equal(result.adminRuntimeSha256, expectedSha);
  assert.equal(result.fileCount, 17);
  assert.equal(result.targetProject.parentId, spreadsheetId);
  assert.equal(result.updateResult.verificationSource, 'UPDATE_RESPONSE');
  assert.equal(calls[0].apiPath, `/projects/${encodeURIComponent(targetId)}`, 'parent binding is checked before source content is read');
  assert.equal(
    JSON.stringify(calls[2].body.files.map(file => file.name)),
    JSON.stringify(files.map(file => file.name).sort()),
    'PUT files use canonical name order'
  );
}

function testAdminGuards() {
  let calls = 0;
  sandbox.vNextClientRuntimeApiRequest_ = () => { calls++; return {}; };
  assert.throws(
    () => sandbox.vNextAdminRuntimeCopyScriptContent_(sourceId, sourceId, spreadsheetId),
    /must be different/
  );
  assert.throws(
    () => sandbox.vNextAdminRuntimeCopyScriptContent_(sourceId, targetId, ' short'),
    /expectedTargetSpreadsheetId is invalid/
  );
  assert.equal(calls, 0, 'identity failures occur before API access');

  let putCalled = false;
  sandbox.vNextClientRuntimeApiRequest_ = (apiPath, method) => {
    if (method === 'put') putCalled = true;
    if (apiPath === `/projects/${encodeURIComponent(targetId)}`) {
      return { scriptId: targetId, parentId: 'DIFFERENT_sheet_1234567890ABCDE' };
    }
    return { scriptId: sourceId, files: clone(files) };
  };
  assert.throws(
    () => sandbox.vNextAdminRuntimeCopyScriptContent_(sourceId, targetId, spreadsheetId),
    /not bound to expectedTargetSpreadsheetId/
  );
  assert.equal(putCalled, false);

  assert.throws(
    () => sandbox.vNextAdminRuntimeValidateFiles_([...clone(files), { name: 'Extra', type: 'SERVER_JS', source: '' }]),
    /exactly the 17 clasp-target files/
  );
  const wrongType = clone(files);
  wrongType.find(file => file.name === 'VNext_Admin').type = 'HTML';
  assert.throws(() => sandbox.vNextAdminRuntimeValidateFiles_(wrongType), /file type mismatch/);
  const wrongManifest = clone(files);
  const manifestFile = wrongManifest.find(file => file.name === 'appsscript');
  const manifest = JSON.parse(manifestFile.source);
  manifest.runtimeVersion = 'DEPRECATED_ES5';
  manifestFile.source = JSON.stringify(manifest);
  assert.throws(() => sandbox.vNextAdminRuntimeValidateFiles_(wrongManifest), /must use V8/);
}

function testAdminCreate() {
  const calls = [];
  movedFolder = '';
  renamed = '';
  sandbox.vNextClientRuntimeApiRequest_ = (apiPath, method, body) => {
    calls.push({ apiPath, method, body });
    if (apiPath === '/projects' && method === 'post') {
      assert.equal(body.parentId, spreadsheetId);
      return { scriptId: targetId, parentId: spreadsheetId };
    }
    if (apiPath === `/projects/${encodeURIComponent(targetId)}` && method === 'get') {
      return { scriptId: targetId, parentId: spreadsheetId, title: 'New Hub Admin Runtime' };
    }
    if (apiPath === `/projects/${encodeURIComponent(sourceId)}/content` && method === 'get') {
      return { scriptId: sourceId, files: clone(files) };
    }
    if (apiPath === `/projects/${encodeURIComponent(targetId)}/content` && method === 'put') {
      return { scriptId: targetId, files: clone(body.files) };
    }
    throw new Error(`Unexpected API call ${method} ${apiPath}`);
  };
  const result = sandbox.vNextAdminRuntimeCreateBoundSpreadsheet_({ title: 'New Empty Hub', folderId });
  assert.equal(result.ok, true);
  assert.equal(result.spreadsheetId, spreadsheetId);
  assert.equal(result.scriptId, targetId);
  assert.equal(result.sourceScriptId, sourceId, 'omitted sourceScriptId uses the current full Admin project');
  assert.equal(result.adminRuntimeSha256, expectedSha);
  assert.equal(result.fileCount, 17);
  assert.equal(movedFolder, folderId);
  assert.equal(renamed, '');
  assert.equal(calls[0].apiPath, '/projects');
}

function canonicalSha(inputFiles) {
  const joined = [...inputFiles].sort((left, right) => left.name < right.name ? -1 : left.name > right.name ? 1 : 0)
    .map(file => `${file.name}\0${file.type}\0${file.source}`).join('\0');
  return createHash('sha256').update(joined, 'utf8').digest('hex');
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}
