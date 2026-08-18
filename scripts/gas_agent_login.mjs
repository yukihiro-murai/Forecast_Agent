#!/usr/bin/env node
/**
 * Restore clasp login with the default Apps Script CLI scopes.
 * Workspace blocks extra Drive/Sheets consent, so this does not request them.
 * Does not grant forecast approval / 差戻し / 公式化.
 */
import { spawn } from 'node:child_process';
import {
  REPO_ROOT, loadClaspTokenRecord, loadTargets, saveAgentToken
} from './lib/gas_oauth.mjs';

process.stdout.write('社内アカウントは Drive/Sheets の追加許可をブロックするため、clasp の標準スコープだけでログインします。計画の承認権限は含まれません。\n');

const code = await new Promise((resolve, reject) => {
  const child = spawn('clasp', [
    'login',
    '--redirect-port', String(process.env.GAS_AGENT_OAUTH_PORT || 8086)
  ], { cwd: REPO_ROOT, stdio: 'inherit' });
  child.on('error', reject);
  child.on('exit', (exitCode) => resolve(exitCode ?? 1));
});
if (code !== 0) process.exit(code);

const targets = await loadTargets();
const record = await loadClaspTokenRecord();
if (!record) throw new Error('clasp login did not save a refresh token.');
if (!record.scope) {
  const info = await fetch('https://www.googleapis.com/oauth2/v3/tokeninfo?access_token=' +
    encodeURIComponent(record.access_token));
  const payload = await info.json();
  record.scope = payload.scope || '';
}
record.quotaProjectId = targets.quotaProjectId;
await saveAgentToken(record);
process.stdout.write('Saved deploy/.gas-oauth.json. Next: node scripts/publish_employee_portal.mjs\n');
