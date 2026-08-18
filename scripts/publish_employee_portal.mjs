#!/usr/bin/env node
/**
 * Cursor entrypoint for the employee portal update.
 *
 * 1. Copies the verified 18 Admin files to central + Hub.
 * 2. Runs Hub's vNextAdminUpdateSharedPortalRuntimeFromSource so SHA pin,
 *    rollback, audit, and same-URL /exec republish stay on the Hub path.
 * 3. If Execution API is blocked, copies the portal bundle and pins /exec
 *    directly, then updates Hub/Portal key-value pins.
 *
 * Does not approve, return, or lock a forecast.
 */
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import {
  ADMIN_SOURCE_FILES, REPO_ROOT,
  getAgentAccessToken, googleApi, loadCache, loadTargets, saveCache
} from './lib/gas_oauth.mjs';

const args = new Set(process.argv.slice(2));
const confirmRuntimeDrift = args.has('--confirm-drift');
const skipSourceSync = args.has('--skip-source-sync');
const reasonIndex = process.argv.indexOf('--reason');
const reason = reasonIndex >= 0
  ? String(process.argv[reasonIndex + 1] || '').trim()
  : 'Cursor agent published employee portal runtime';
const portalScriptArgIndex = process.argv.indexOf('--portal-script-id');
const portalScriptArg = portalScriptArgIndex >= 0
  ? String(process.argv[portalScriptArgIndex + 1] || '').trim()
  : '';

if (args.has('--help')) {
  process.stdout.write(`Usage: node scripts/publish_employee_portal.mjs [--confirm-drift] [--skip-source-sync] [--sync-only] [--portal-script-id ID] [--reason TEXT]
--sync-only copies the 18 Admin files to central and Hub without touching /exec.
One-time login: node scripts/gas_agent_login.mjs
`);
  process.exit(0);
}

const targets = await loadTargets();
const token = await getAgentAccessToken();
token.quotaProjectId = token.quotaProjectId || targets.quotaProjectId;

if (!skipSourceSync) {
  await syncAdminProject(targets.hubScriptId, 'hub');
  await syncAdminProject(targets.centralScriptId, 'central');
}
if (args.has('--sync-only')) process.exit(0);

const run = await runHubPortalUpdate(targets, token, reason, confirmRuntimeDrift);
if (run.ok) {
  printResult('execution-api', run.result);
  process.exit(0);
}
if (isPinDrift(run.error) && !confirmRuntimeDrift) {
  process.stderr.write(run.error + '\n');
  process.stderr.write('Pin drift requires a second confirmed run: node scripts/publish_employee_portal.mjs --confirm-drift\n');
  process.exit(2);
}
if (!run.fallbackAllowed) {
  process.stderr.write(run.error + '\n');
  process.exit(1);
}

process.stdout.write('Execution API unavailable; using verified file copy + same-URL /exec pin.\n');
process.stdout.write(run.error + '\n');
const direct = await publishPortalDirect(targets, token, reason, confirmRuntimeDrift);
printResult('direct', direct);

async function syncAdminProject(scriptId, label) {
  const { status, payload } = await googleApi(
    token, 'GET', `https://script.googleapis.com/v1/projects/${scriptId}/content`);
  if (status !== 200) throw new Error(`${label} content get failed: ${formatApiError(status, payload)}`);
  const local = {};
  for (const [name, type, filename] of ADMIN_SOURCE_FILES) {
    local[name] = {
      name,
      type,
      source: await readFile(path.join(REPO_ROOT, filename), 'utf8')
    };
  }
  const changed = [];
  const files = (payload.files || []).map((file) => {
    const next = { name: file.name, type: file.type, source: file.source };
    if (local[file.name] && local[file.name].source !== file.source) {
      next.source = local[file.name].source;
      next.type = local[file.name].type;
      changed.push(file.name);
    }
    return next;
  });
  const present = new Set(files.map((file) => file.name));
  for (const [name, type] of ADMIN_SOURCE_FILES) {
    if (!present.has(name)) {
      files.push({ name, type, source: local[name].source });
      changed.push(name);
    }
  }
  process.stdout.write(`${label} changed ${changed.length ? changed.join(',') : '(none)'}\n`);
  if (!changed.length) return;
  const put = await googleApi(
    token, 'PUT', `https://script.googleapis.com/v1/projects/${scriptId}/content`, { files });
  if (put.status !== 200) throw new Error(`${label} content put failed: ${formatApiError(put.status, put.payload)}`);
}

async function runHubPortalUpdate(targets, token, reason, confirmRuntimeDrift) {
  const url = `https://script.googleapis.com/v1/scripts/${targets.hubScriptId}:run`;
  const { status, payload } = await googleApi(token, 'POST', url, {
    function: 'vNextAdminUpdateSharedPortalRuntimeFromSource',
    devMode: true,
    parameters: [{
      hubSpreadsheetId: targets.hubSpreadsheetId,
      reason,
      confirmRuntimeDrift
    }]
  }, 360_000);
  if (status === 200 && payload && payload.error) {
    const message = executionErrorMessage(payload);
    return { ok: false, fallbackAllowed: !isPinDrift(message), error: message };
  }
  if (status === 200 && payload && (payload.response || payload.done)) {
    const result = (payload.response && payload.response.result) || payload.response || payload;
    return { ok: true, result };
  }
  if (status === 200) return { ok: true, result: payload };
  const message = formatApiError(status, payload);
  const fallbackAllowed = status === 403 || status === 404 || /permission|quota project|not found/i.test(message);
  return { ok: false, fallbackAllowed, error: message };
}

async function publishPortalDirect(targets, token, reason, confirmRuntimeDrift) {
  let config = {};
  try {
    config = await readKeyValueSheet(targets.hubSpreadsheetId, 'VN_SYSTEM_CONFIG');
  } catch (error) {
    process.stdout.write('Hub VN_SYSTEM_CONFIG is not readable with clasp scopes; using local portal script id.\n');
  }
  const portalScriptId = String(
    portalScriptArg || process.env.PORTAL_SCRIPT_ID || targets.portalScriptId ||
    config.portal_script_id || (await loadCache()).portalScriptId || ''
  ).trim();
  if (!portalScriptId) {
    throw new Error(
      'portal_script_id is required. Copy it from Hub VN_SYSTEM_CONFIG or ポータル → 拡張機能 → Apps Script → プロジェクトの設定, then re-run with --portal-script-id ID'
    );
  }
  await saveCache({ ...(await loadCache()), portalScriptId });
  const bundle = JSON.parse(await readFile(
    path.join(REPO_ROOT, 'portal_runtime/generated/portal-runtime-bundle.json'), 'utf8'));
  const targetFiles = bundle.files.map((file) => ({
    name: file.name, type: file.type, source: file.source
  }));
  const targetSha = filesSha256(targetFiles);
  if (targetSha !== bundle.sha256) throw new Error('Local portal bundle hash mismatch.');

  const current = await googleApi(
    token, 'GET', `https://script.googleapis.com/v1/projects/${portalScriptId}/content`);
  if (current.status !== 200) {
    throw new Error('Portal content get failed: ' + formatApiError(current.status, current.payload));
  }
  const portalNames = new Set(targetFiles.map((file) => file.name));
  const currentFiles = (current.payload.files || [])
    .filter((file) => portalNames.has(file.name))
    .map((file) => ({ name: file.name, type: file.type, source: file.source }));
  const currentSha = filesSha256(currentFiles);
  const pinnedSha = String(config.portal_runtime_sha256 || '');
  if (pinnedSha && currentSha !== pinnedSha && !confirmRuntimeDrift) {
    throw new Error(
      'Current Portal runtime does not match its stored SHA-256 pin.' +
      ' pinned=' + pinnedSha + ' current=' + currentSha
    );
  }

  const keep = (current.payload.files || [])
    .filter((file) => !targetFiles.some((target) => target.name === file.name))
    .map((file) => ({ name: file.name, type: file.type, source: file.source }));
  const put = await googleApi(
    token, 'PUT', `https://script.googleapis.com/v1/projects/${portalScriptId}/content`,
    { files: targetFiles.concat(keep) }
  );
  if (put.status !== 200) {
    throw new Error('Portal content put failed: ' + formatApiError(put.status, put.payload));
  }
  const webApp = await republishWebApp(portalScriptId, bundle.version);
  try {
    await upsertKeyValue(targets.hubSpreadsheetId, 'VN_SYSTEM_CONFIG', {
      portal_runtime_version: bundle.version,
      portal_runtime_sha256: targetSha,
      portal_runtime_updated_at: new Date().toISOString(),
      portal_runtime_updated_by: 'cursor-agent'
    });
    if (targets.portalSpreadsheetId) {
      await upsertKeyValue(targets.portalSpreadsheetId, 'VN_PORTAL_CONFIG', {
        runtime_version: bundle.version,
        runtime_sha256: targetSha,
        updated_at: new Date().toISOString(),
        updated_by: 'cursor-agent'
      });
    }
  } catch (error) {
    process.stdout.write('Hub/Portal pin sheets were not updated (clasp scopes omit Sheets). Files and /exec were published.\n');
  }
  return {
    ok: true,
    reused: currentSha === targetSha,
    runtimeVersion: bundle.version,
    runtimeSha256: targetSha,
    webAppUrl: webApp.webAppUrl,
    webAppVersion: webApp.versionNumber,
    reason,
    message: '社員ポータルを最新版へ更新し、社員向けWeb入口も同じURLのまま公開しました。入口をハード再読み込みしてください。'
  };
}

async function republishWebApp(scriptId, versionLabel) {
  const created = await googleApi(
    token, 'POST',
    `https://script.googleapis.com/v1/projects/${scriptId}/versions`,
    { description: versionLabel + ' employee web entry' }
  );
  if (created.status !== 200 || !Number(created.payload.versionNumber)) {
    throw new Error('Portal version create failed: ' + formatApiError(created.status, created.payload));
  }
  const listed = await googleApi(
    token, 'GET', `https://script.googleapis.com/v1/projects/${scriptId}/deployments`);
  if (listed.status !== 200) {
    throw new Error('Portal deployments list failed: ' + formatApiError(listed.status, listed.payload));
  }
  const selected = selectPortalWebAppDeployment(listed.payload.deployments || []);
  const updated = await googleApi(
    token, 'PUT',
    `https://script.googleapis.com/v1/projects/${scriptId}/deployments/${encodeURIComponent(selected.deploymentId)}`,
    {
      deploymentConfig: {
        versionNumber: Number(created.payload.versionNumber),
        manifestFileName: (selected.deploymentConfig && selected.deploymentConfig.manifestFileName) || 'appsscript',
        description: versionLabel
      }
    }
  );
  if (updated.status !== 200) {
    throw new Error('Portal deployment update failed: ' + formatApiError(updated.status, updated.payload));
  }
  const webAppUrl = webAppUrlFromDeployment(updated.payload) || webAppUrlFromDeployment(selected);
  if (!webAppUrl) throw new Error('Portal web app URL was missing after republish.');
  return {
    versionNumber: Number(created.payload.versionNumber),
    webAppUrl,
    deploymentId: selected.deploymentId
  };
}

function selectPortalWebAppDeployment(deployments) {
  const webApps = (deployments || []).map((deployment) => {
    const url = webAppUrlFromDeployment(deployment);
    return url ? { deployment, url } : null;
  }).filter(Boolean);
  if (!webApps.length) {
    throw new Error('Portal web app deployment was not found. Publish /exec once from the Apps Script editor, then retry.');
  }
  const versioned = webApps.filter((item) =>
    Number(item.deployment.deploymentConfig && item.deployment.deploymentConfig.versionNumber || 0) > 0);
  const pool = versioned.length ? versioned : webApps;
  if (pool.length === 1) return pool[0].deployment;
  const domain = pool.filter((item) => (item.deployment.entryPoints || []).some((entry) =>
    String(((entry.webApp || {}).entryPointConfig || {}).access || '').toUpperCase() === 'DOMAIN'));
  if (domain.length === 1) return domain[0].deployment;
  throw new Error('Multiple Portal web app deployments exist: ' + pool.map((item) => item.url).join(', '));
}

function webAppUrlFromDeployment(deployment) {
  for (const entry of (deployment && deployment.entryPoints) || []) {
    if (String(entry.entryPointType || '') !== 'WEB_APP') continue;
    const url = String((entry.webApp || {}).url || '');
    if (url) return url;
  }
  return '';
}

async function readKeyValueSheet(spreadsheetId, sheetName) {
  const { status, payload } = await googleApi(
    token, 'GET',
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}`);
  if (status !== 200) throw new Error(`${sheetName} read failed: ${formatApiError(status, payload)}`);
  const rows = payload.values || [];
  if (!rows.length) return {};
  const header = rows[0].map((cell) => String(cell || '').trim());
  const keyIdx = header.indexOf('key');
  const valueIdx = header.indexOf('value');
  if (keyIdx < 0 || valueIdx < 0) return {};
  const out = {};
  for (const row of rows.slice(1)) {
    const key = String(row[keyIdx] || '').trim();
    if (key) out[key] = row[valueIdx];
  }
  return out;
}

async function upsertKeyValue(spreadsheetId, sheetName, values) {
  const { status, payload } = await googleApi(
    token, 'GET',
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}`);
  if (status !== 200) throw new Error(`${sheetName} read failed: ${formatApiError(status, payload)}`);
  const rows = payload.values || [];
  const header = (rows[0] || ['key', 'value', 'updated_at']).map((cell) => String(cell || '').trim());
  const keyIdx = Math.max(0, header.indexOf('key'));
  const valueIdx = header.indexOf('value') >= 0 ? header.indexOf('value') : 1;
  const updatedIdx = header.indexOf('updated_at');
  const indexByKey = new Map();
  rows.slice(1).forEach((row, offset) => {
    const key = String(row[keyIdx] || '').trim();
    if (key) indexByKey.set(key, offset + 2);
  });
  const data = [];
  for (const [key, value] of Object.entries(values)) {
    const rowNumber = indexByKey.get(key);
    if (!rowNumber) {
      throw new Error(`${sheetName} is missing key ${key}; refusing to append from the agent path.`);
    }
    const row = [...(rows[rowNumber - 1] || [])];
    while (row.length < header.length) row.push('');
    row[valueIdx] = value;
    if (updatedIdx >= 0) row[updatedIdx] = new Date().toISOString();
    data.push({
      range: `${sheetName}!A${rowNumber}`,
      values: [row]
    });
  }
  const written = await googleApi(
    token, 'POST',
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`,
    { valueInputOption: 'RAW', data }
  );
  if (written.status !== 200) {
    throw new Error(`${sheetName} write failed: ${formatApiError(written.status, written.payload)}`);
  }
}

function filesSha256(files) {
  const joined = (files || []).map((file) =>
    `${file.name}\0${file.type}\0${file.source}`).join('\0');
  return createHash('sha256').update(joined, 'utf8').digest('hex');
}

function executionErrorMessage(payload) {
  const details = payload.error && payload.error.details || [];
  for (const detail of details) {
    if (detail.errorMessage) return String(detail.errorMessage);
    if (Array.isArray(detail) && detail[0] && detail[0].errorMessage) return String(detail[0].errorMessage);
  }
  return String(payload.error.message || JSON.stringify(payload.error));
}

function isPinDrift(message) {
  return /stored SHA-256 pin/.test(String(message || ''));
}

function formatApiError(status, payload) {
  if (payload && typeof payload === 'object') {
    const message = payload.error && payload.error.message || payload.message;
    if (message) return `status=${status}; ${message}`;
  }
  return `status=${status}; ${typeof payload === 'string' ? payload.slice(0, 400) : JSON.stringify(payload).slice(0, 400)}`;
}

function printResult(pathName, result) {
  process.stdout.write(`path=${pathName}\n`);
  process.stdout.write(`ok=${Boolean(result && result.ok !== false)}\n`);
  if (result && result.runtimeVersion) process.stdout.write(`runtime=${result.runtimeVersion}\n`);
  if (result && result.webAppUrl) process.stdout.write(`webAppUrl=${result.webAppUrl}\n`);
  if (result && result.webAppVersion) process.stdout.write(`webAppVersion=${result.webAppVersion}\n`);
  if (result && result.message) process.stdout.write(`${result.message}\n`);
}
