#!/usr/bin/env node
/**
 * One-time Google login so Cursor can run Hub's verified employee-portal
 * update. Uses clasp (Apps Script CLI), not the blocked gcloud helper app.
 * Does not grant forecast approval / 差戻し / 公式化.
 */
import { spawn } from 'node:child_process';
import {
  REPO_ROOT, loadClaspTokenRecord, loadTargets, saveAgentToken
} from './lib/gas_oauth.mjs';

process.stdout.write('gcloud 用アプリは社内アカウントでブロックされるため、clasp（Apps Script CLI）で許可します。計画の承認権限は含まれません。\n');

const logout = await new Promise((resolve, reject) => {
  const child = spawn('clasp', ['logout'], { cwd: REPO_ROOT, stdio: 'inherit' });
  child.on('error', reject);
  child.on('exit', (exitCode) => resolve(exitCode ?? 1));
});
if (logout !== 0) process.exit(logout);

const code = await new Promise((resolve, reject) => {
  const child = spawn('clasp', [
    'login',
    '--use-project-scopes',
    '--include-clasp-scopes',
    '--redirect-port', String(process.env.GAS_AGENT_OAUTH_PORT || 8085)
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
