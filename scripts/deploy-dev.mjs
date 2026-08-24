#!/usr/bin/env node

import { readFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(scriptDir, '..');
const args = process.argv.slice(2);
const pushOnly = args.includes('--push-only');
const skipRemote = args.includes('--skip-remote');
const reason = readArg('--reason') || 'Cursor deploy:dev';

const config = {
  hubSpreadsheetId: process.env.VNEXT_HUB_SPREADSHEET_ID || '1baEZe6xYQ9KWyMMBk7kzH50v4dTtBPk9kWHK3qT7ID8',
  hubScriptId: process.env.VNEXT_HUB_SCRIPT_ID || '1ANVtOBzPo90DveLcmTYokpEycr4iOphMaKZrhyjvzUizEQUjGqHRScxw',
  centralScriptId: process.env.VNEXT_CENTRAL_SCRIPT_ID || '',
  hubDeploymentId: process.env.VNEXT_HUB_DEPLOYMENT_ID || '',
  centralDeploymentId: process.env.VNEXT_CENTRAL_DEPLOYMENT_ID || ''
};

main().catch(error => {
  process.stderr.write(String(error && error.stack || error) + '\n');
  process.exit(1);
});

async function main() {
  if (!pushOnly) {
    run(process.execPath, ['tests/vnext-integration.test.mjs'], root, 'integration tests');
    run('npm', ['run', 'build'], path.join(root, 'client_runtime'), 'client runtime build');
    run('npm', ['test'], path.join(root, 'client_runtime'), 'client runtime tests');
    run('npm', ['run', 'build'], path.join(root, 'portal_runtime'), 'portal runtime build');
    run('npm', ['test'], path.join(root, 'portal_runtime'), 'portal runtime tests');
  }

  run('clasp', ['push'], root, 'clasp push');
  if (skipRemote) {
    process.stdout.write('\nLocal build/push complete.\n');
    process.stdout.write('管理ハブ上部メニュー「年度予算策定」→「Web入口を最新版にする」→ Web入口 Cmd+Shift+R\n');
    return;
  }

  const token = await (async () => {
    const clasprc = JSON.parse(await readFile(path.join(os.homedir(), '.clasprc.json'), 'utf8'));
    const entry = clasprc.tokens?.default || clasprc.token || clasprc;
    let access = entry.access_token || entry.accessToken;
    const expiry = Number(entry.expiry_date || 0);
    if (expiry && Date.now() > expiry - 60000 && entry.refresh_token) {
      const body = new URLSearchParams({
        client_id: entry.client_id,
        client_secret: entry.client_secret,
        refresh_token: entry.refresh_token,
        grant_type: 'refresh_token'
      });
      const res = await fetch('https://oauth2.googleapis.com/token', { method: 'POST', body });
      const json = await res.json();
      if (!res.ok) throw new Error('token refresh failed');
      access = json.access_token;
    }
    if (!access) throw new Error('clasp OAuth token not found in ~/.clasprc.json');
    return access;
  })();

  if (!config.centralScriptId) {
    config.centralScriptId = JSON.parse(await readFile(path.join(root, '.clasp.json'), 'utf8')).scriptId;
  }

  process.stdout.write('\nPhase A: syncing Hub runtime from central source…\n');
  const phaseA = await runScript({
    token,
    scriptId: config.centralScriptId,
    deploymentId: config.centralDeploymentId,
    functionName: 'vNextAdminDeployVerifiedEmployeeUxReleaseFromSource',
    parameters: [{ hubSpreadsheetId: config.hubSpreadsheetId, reason }]
  });
  process.stdout.write(`Phase A OK: ${summarize(phaseA)}\n`);

  await sleep(4000);

  // Portal first (same order as Hub smart deploy). Client only if still needed later.
  process.stdout.write('Phase B: Portal first…\n');
  try {
    const portal = await runScript({
      token,
      scriptId: config.hubScriptId,
      deploymentId: config.hubDeploymentId,
      functionName: 'vNextAdminDeployVerifiedEmployeeUxPortal_',
      parameters: [{ reason, fastDeploy: true }]
    });
    process.stdout.write(`Portal OK: ${summarize(portal)}\n`);
    await sleep(2000);
    process.stdout.write('Finalize…\n');
    const phaseB = await runScript({
      token,
      scriptId: config.hubScriptId,
      deploymentId: config.hubDeploymentId,
      functionName: 'vNextAdminDeployVerifiedEmployeeUxFinalize_',
      parameters: [{ reason, upgradeEmptyPilots: false, portal }]
    });
    process.stdout.write(`Finalize OK: ${summarize(phaseB)}\n`);
    process.stdout.write('\nDeploy complete. Hard refresh Web入口 /exec (Cmd+Shift+R).\n');
  } catch (error) {
    process.stdout.write('\nHub/Portal API 反映はできませんでした。\n');
    process.stdout.write(`${String(error.message || error)}\n`);
    process.stdout.write('管理ハブ上部メニュー「年度予算策定」→「Web入口を最新版にする」を実行してください。\n');
  }
}

async function runScript({ token, scriptId, deploymentId, functionName, parameters }) {
  const url = deploymentId
    ? `https://script.googleapis.com/v1/scripts/${scriptId}:run`
    : `https://script.googleapis.com/v1/scripts/${scriptId}:run`;
  const body = { function: functionName, parameters, devMode: !deploymentId };
  if (deploymentId) body.deploymentId = deploymentId;

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(`Script API ${response.status}: ${JSON.stringify(payload).slice(0, 500)}`);
  }
  if (payload.error) {
    throw new Error(JSON.stringify(payload.error));
  }
  return payload.response?.result ?? payload.response ?? payload;
}

function summarize(value) {
  if (!value || typeof value !== 'object') return String(value);
  const parts = [];
  if (value.phase) parts.push(`phase=${value.phase}`);
  if (value.clientRuntimeVersion) parts.push(`client=${value.clientRuntimeVersion}`);
  if (value.portalRuntimeVersion) parts.push(`portal=${value.portalRuntimeVersion}`);
  if (value.adminRuntimeSha256) parts.push(`hubSha=${String(value.adminRuntimeSha256).slice(0, 12)}`);
  if (value.verification && value.verification.summary) parts.push(String(value.verification.summary));
  if (value.verification && value.verification.failedCount) parts.push(`failed=${value.verification.failedCount}`);
  if (value.message) parts.push(String(value.message));
  return parts.join(' | ') || JSON.stringify(value).slice(0, 200);
}

function run(command, args, cwd, label) {
  process.stdout.write(`\n> ${label}\n`);
  const result = spawnSync(command, args, { cwd, stdio: 'inherit' });
  if (result.status !== 0) throw new Error(`${label} failed with status ${result.status}`);
}

function readArg(name) {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : '';
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}
