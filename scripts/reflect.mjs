#!/usr/bin/env node
/**
 * Cursor 対話用の最短反映。
 * - clasp push までエージェントがやる
 * - Hub / Portal は可能な範囲で API 実行を試し、ダメならユーザー1操作に落とす
 */
import { readFile, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(scriptDir, '..');
const args = process.argv.slice(2);
const skipTests = args.includes('--skip-tests');
const reason = readArg('--reason') || '開発反映';

const HUB_SHEET =
  process.env.VNEXT_HUB_SPREADSHEET_ID || '1baEZe6xYQ9KWyMMBk7kzH50v4dTtBPk9kWHK3qT7ID8';
const HUB_SCRIPT =
  process.env.VNEXT_HUB_SCRIPT_ID || '1ANVtOBzPo90DveLcmTYokpEycr4iOphMaKZrhyjvzUizEQUjGqHRScxw';
const HUB_URL = `https://docs.google.com/spreadsheets/d/${HUB_SHEET}/edit`;

const USER_NEXT = [
  '',
  '━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
  '【あなた（いまの画面で見つかる手順）】',
  '管理ハブ Spreadsheet で:',
  '  1. 上部メニュー「年度予算策定」→「案内を開く」',
  '  2. 右パネルの「反映」タブ → 緑ボタン「反映する」を1回押す',
  '     （Hub同期後は自動で続行し、Web入口更新まで進みます）',
  '  3. Web入口を Cmd+Shift+R',
  '',
  '※ メニューに「Web入口を最新版にする」が見えるならそれでOK。',
  '　 見えない場合は Hub メニューが古いだけなので、上記 1→2 を使ってください。',
  '',
  `管理ハブ: ${HUB_URL}`,
  '━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
  ''
].join('\n');

main().catch(error => {
  process.stderr.write(String(error && error.stack || error) + '\n');
  process.stdout.write(USER_NEXT);
  process.exit(1);
});

async function main() {
  const totalStartedAt = Date.now();
  process.on('exit', () => {
    process.stdout.write(`\n合計: ${((Date.now() - totalStartedAt) / 1000).toFixed(1)}s\n`);
  });
  if (!skipTests) {
    run(process.execPath, ['tests/vnext-integration.test.mjs'], root, 'contract tests');
  }

  // Portal / Client ソースを触っているときだけ build（失敗しても clasp は続行しない）
  maybeBuildRuntimes();

  run('clasp', ['push'], root, 'clasp push');
  process.stdout.write('\nclasp push 完了（中央ソース）。\n');

  const token = await getAccessToken();
  const centralScriptId = JSON.parse(await readFile(path.join(root, '.clasp.json'), 'utf8')).scriptId;

  let hubSynced = false;
  try {
    process.stdout.write('Hub 同期を試行（中央 → 管理ハブ）…\n');
    const phaseA = await runScript(token, centralScriptId, 'vNextAdminDeployVerifiedEmployeeUxReleaseFromSource', [
      { hubSpreadsheetId: HUB_SHEET, reason }
    ]);
    hubSynced = true;
    process.stdout.write(`Hub 同期 OK: ${summarize(phaseA)}\n`);
  } catch (error) {
    process.stdout.write(`Hub 同期は API からできませんでした（想定内のことが多い）: ${shortErr(error)}\n`);
  }

  if (hubSynced) {
    await sleep(2500);
    try {
      process.stdout.write('Portal 更新を試行…\n');
      const portal = await runScript(token, HUB_SCRIPT, 'vNextAdminRpcDeployPortalStep', [
        { reason, fastDeploy: true }
      ]);
      process.stdout.write(`Portal OK: ${summarize(portal)}\n`);
      const fin = await runScript(token, HUB_SCRIPT, 'vNextAdminRpcDeployFinalizeStep', [
        { reason, upgradeEmptyPilots: false, portal }
      ]);
      process.stdout.write(`確認: ${summarize(fin)}\n`);
      if (fin && fin.verification && fin.verification.ok) {
        process.stdout.write('\n自動反映まで完了。Web入口を Cmd+Shift+R し、右下の「版」表示で確認してください。\n');
        process.stdout.write(`管理ハブ: ${HUB_URL}\n`);
        return;
      }
    } catch (error) {
      process.stdout.write(`Portal API 反映はスキップ: ${shortErr(error)}\n`);
    }
  }

  process.stdout.write(USER_NEXT);
}

function maybeBuildRuntimes() {
  const portalDirty = gitTouched('portal_runtime/');
  const clientDirty = gitTouched('client_runtime/');
  if (portalDirty) {
    run('npm', ['run', 'build'], path.join(root, 'portal_runtime'), 'portal build');
  }
  if (clientDirty) {
    run('npm', ['run', 'build'], path.join(root, 'client_runtime'), 'client build');
  }
}

function gitTouched(prefix) {
  const result = spawnSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' });
  if (result.status !== 0) return false;
  return String(result.stdout || '').split('\n').some(line => line.slice(3).startsWith(prefix));
}

async function getAccessToken() {
  const clasprcPath = path.join(os.homedir(), '.clasprc.json');
  const clasprc = JSON.parse(await readFile(clasprcPath, 'utf8'));
  const entry = clasprc.tokens?.default || clasprc.token || clasprc;
  let access = entry.access_token || entry.accessToken || (typeof entry === 'string' ? entry : '');
  const expiry = Number(entry.expiry_date || 0);
  if (!access) throw new Error('clasp OAuth token missing (~/.clasprc.json)');
  if (expiry && Date.now() > expiry - 60000 && entry.refresh_token) {
    const body = new URLSearchParams({
      client_id: entry.client_id,
      client_secret: entry.client_secret,
      refresh_token: entry.refresh_token,
      grant_type: 'refresh_token'
    });
    const res = await fetch('https://oauth2.googleapis.com/token', { method: 'POST', body });
    const json = await res.json();
    if (!res.ok) throw new Error('token refresh failed: ' + JSON.stringify(json));
    access = json.access_token;
    entry.access_token = access;
    entry.expiry_date = Date.now() + Number(json.expires_in || 3600) * 1000;
    clasprc.tokens = clasprc.tokens || {};
    clasprc.tokens.default = entry;
    await writeFile(clasprcPath, JSON.stringify(clasprc, null, 2));
    process.stdout.write('OAuth token refreshed.\n');
  }
  return access;
}

async function runScript(token, scriptId, functionName, parameters) {
  // Prefer deployment ID as :run target when provided via env.
  const runId = process.env.VNEXT_SCRIPT_RUN_ID || scriptId;
  const res = await fetch(`https://script.googleapis.com/v1/scripts/${runId}:run`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ function: functionName, parameters, devMode: !process.env.VNEXT_SCRIPT_RUN_ID })
  });
  const payload = await res.json();
  if (!res.ok) throw new Error(`Script API ${res.status}: ${JSON.stringify(payload).slice(0, 400)}`);
  if (payload.error) throw new Error(JSON.stringify(payload.error));
  if (payload.response?.error) throw new Error(JSON.stringify(payload.response.error).slice(0, 400));
  return payload.response?.result ?? payload.response ?? payload;
}

function summarize(value) {
  if (!value || typeof value !== 'object') return String(value);
  const parts = [];
  if (value.runtimeVersion) parts.push(`portal=${value.runtimeVersion}`);
  if (value.portalRuntimeVersion) parts.push(`portal=${value.portalRuntimeVersion}`);
  if (value.adminRuntimeSha256) parts.push(`hub=${String(value.adminRuntimeSha256).slice(0, 12)}`);
  if (value.verification?.summary) parts.push(value.verification.summary);
  if (value.message) parts.push(String(value.message).slice(0, 120));
  return parts.join(' | ') || JSON.stringify(value).slice(0, 160);
}

function shortErr(error) {
  return String(error && error.message || error).slice(0, 180);
}

function run(command, argsList, cwd, label) {
  process.stdout.write(`\n> ${label}\n`);
  const startedAt = Date.now();
  const result = spawnSync(command, argsList, { cwd, stdio: 'inherit' });
  process.stdout.write(`  (${label}: ${((Date.now() - startedAt) / 1000).toFixed(1)}s)\n`);
  if (result.status !== 0) throw new Error(`${label} failed (${result.status})`);
}

function readArg(name) {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : '';
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}
