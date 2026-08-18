#!/usr/bin/env node
/**
 * One-time Google login so Cursor can run Hub's verified employee-portal
 * update. Does not grant forecast approval / 差戻し / 公式化.
 */
import http from 'node:http';
import { spawn } from 'node:child_process';
import {
  AGENT_SCOPES, loadCloudSdkClient, loadTargets, saveAgentToken
} from './lib/gas_oauth.mjs';

const PORT = Number(process.env.GAS_AGENT_OAUTH_PORT || 8085);
const REDIRECT = `http://localhost:${PORT}/`;

const sdk = await loadCloudSdkClient();
const targets = await loadTargets();
const state = Math.random().toString(36).slice(2);
const authUrl = 'https://accounts.google.com/o/oauth2/v2/auth?' + new URLSearchParams({
  client_id: sdk.client_id,
  redirect_uri: REDIRECT,
  response_type: 'code',
  access_type: 'offline',
  prompt: 'consent',
  include_granted_scopes: 'true',
  state,
  scope: AGENT_SCOPES.join(' ')
});

const record = await new Promise((resolve, reject) => {
  const server = http.createServer(async (req, res) => {
    try {
      const url = new URL(req.url, `http://localhost:${PORT}`);
      if (!url.searchParams.get('code') && !url.searchParams.get('error')) {
        res.writeHead(404);
        res.end();
        return;
      }
      if (url.searchParams.get('state') !== state) {
        throw new Error('OAuth state mismatch.');
      }
      const error = url.searchParams.get('error');
      if (error) throw new Error('OAuth denied: ' + error);
      const code = url.searchParams.get('code');
      if (!code) throw new Error('OAuth code was missing.');
      const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: sdk.client_id,
          client_secret: sdk.client_secret,
          code,
          grant_type: 'authorization_code',
          redirect_uri: REDIRECT
        })
      });
      const payload = await tokenRes.json();
      if (!tokenRes.ok || !payload.refresh_token) {
        throw new Error('OAuth token exchange failed: ' +
          (payload.error_description || payload.error || tokenRes.status));
      }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end('<p>Cursor 用の Google 許可が完了しました。このタブを閉じてください。</p>');
      server.close();
      resolve({
        quotaProjectId: targets.quotaProjectId,
        client_id: sdk.client_id,
        client_secret: sdk.client_secret,
        refresh_token: payload.refresh_token,
        access_token: payload.access_token,
        expiry_date: Date.now() + Number(payload.expires_in || 3600) * 1000,
        scope: payload.scope || AGENT_SCOPES.join(' ')
      });
    } catch (error) {
      res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('OAuth failed.');
      server.close();
      reject(error);
    }
  });
  server.on('error', reject);
  server.listen(PORT, '127.0.0.1', () => {
    process.stdout.write('ブラウザで Google アカウントを許可してください。計画の承認権限は含まれません。\n');
    process.stdout.write(authUrl + '\n');
    spawn('open', [authUrl], { stdio: 'ignore', detached: true }).unref();
  });
});

await saveAgentToken(record);
process.stdout.write('Saved deploy/.gas-oauth.json. Next: node scripts/publish_employee_portal.mjs\n');
