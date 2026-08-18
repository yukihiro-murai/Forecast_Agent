import { readFile, writeFile, mkdir } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
export const TOKEN_PATH = path.join(REPO_ROOT, 'deploy/.gas-oauth.json');
export const TARGETS_PATH = path.join(REPO_ROOT, 'deploy/targets.json');
export const CACHE_PATH = path.join(REPO_ROOT, 'deploy/cache.json');

export const AGENT_SCOPES = Object.freeze([
  'https://www.googleapis.com/auth/script.deployments',
  'https://www.googleapis.com/auth/script.projects',
  'https://www.googleapis.com/auth/script.webapp.deploy'
]);

export const ADMIN_SOURCE_FILES = Object.freeze([
  ['Forecast_Agent', 'SERVER_JS', 'Forecast_Agent.js'],
  ['VNext_AI', 'SERVER_JS', 'VNext_AI.js'],
  ['VNext_Admin', 'SERVER_JS', 'VNext_Admin.js'],
  ['VNext_AdminSidebar', 'HTML', 'VNext_AdminSidebar.html'],
  ['VNext_ClientRuntimeBundle', 'SERVER_JS', 'VNext_ClientRuntimeBundle.js'],
  ['VNext_ClientRuntimeProvisioning', 'SERVER_JS', 'VNext_ClientRuntimeProvisioning.js'],
  ['VNext_Core', 'SERVER_JS', 'VNext_Core.js'],
  ['VNext_Engine', 'SERVER_JS', 'VNext_Engine.js'],
  ['VNext_GuidanceSidebar', 'HTML', 'VNext_GuidanceSidebar.html'],
  ['VNext_HelpSidebar', 'HTML', 'VNext_HelpSidebar.html'],
  ['VNext_InputSidebar', 'HTML', 'VNext_InputSidebar.html'],
  ['VNext_PlanSidebar', 'HTML', 'VNext_PlanSidebar.html'],
  ['VNext_PortalPilotRecovery', 'SERVER_JS', 'VNext_PortalPilotRecovery.js'],
  ['VNext_PortalRuntimeBundle', 'SERVER_JS', 'VNext_PortalRuntimeBundle.js'],
  ['VNext_ReviewSidebar', 'HTML', 'VNext_ReviewSidebar.html'],
  ['VNext_Tests', 'SERVER_JS', 'VNext_Tests.js'],
  ['VNext_UX', 'SERVER_JS', 'VNext_UX.js'],
  ['appsscript', 'JSON', 'appsscript.json']
]);

export async function loadTargets() {
  return JSON.parse(await readFile(TARGETS_PATH, 'utf8'));
}

export async function loadCache() {
  try {
    return JSON.parse(await readFile(CACHE_PATH, 'utf8'));
  } catch (error) {
    if (error && error.code === 'ENOENT') return {};
    throw error;
  }
}

export async function saveCache(cache) {
  await mkdir(path.dirname(CACHE_PATH), { recursive: true });
  await writeFile(CACHE_PATH, `${JSON.stringify(cache, null, 2)}\n`);
}

export async function loadClaspTokenRecord() {
  const clasprc = JSON.parse(await readFile(path.join(os.homedir(), '.clasprc.json'), 'utf8'));
  const token = clasprc.tokens && clasprc.tokens.default;
  if (!token || !token.refresh_token) return null;
  return {
    quotaProjectId: (await loadTargets()).quotaProjectId,
    client_id: token.client_id,
    client_secret: token.client_secret,
    refresh_token: token.refresh_token,
    access_token: token.access_token,
    expiry_date: token.expiry_date || 0,
    scope: token.scope || ''
  };
}

async function refreshAccessToken(record) {
  const body = new URLSearchParams({
    client_id: record.client_id,
    client_secret: record.client_secret,
    refresh_token: record.refresh_token,
    grant_type: 'refresh_token'
  });
  const response = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });
  const payload = await response.json();
  if (!response.ok || !payload.access_token) {
    throw new Error('OAuth refresh failed: ' + (payload.error_description || payload.error || response.status));
  }
  return {
    ...record,
    access_token: payload.access_token,
    expiry_date: Date.now() + Number(payload.expires_in || 3600) * 1000
  };
}

export async function saveAgentToken(record) {
  await mkdir(path.dirname(TOKEN_PATH), { recursive: true });
  await writeFile(TOKEN_PATH, `${JSON.stringify(record, null, 2)}\n`, { mode: 0o600 });
}

export async function loadAgentTokenRecord() {
  try {
    return JSON.parse(await readFile(TOKEN_PATH, 'utf8'));
  } catch (error) {
    if (error && error.code === 'ENOENT') return null;
    throw error;
  }
}

export function missingAgentScopes(scopeText) {
  const granted = new Set(String(scopeText || '').split(/[ ,]+/).filter(Boolean));
  return AGENT_SCOPES.filter((scope) => !granted.has(scope));
}

export async function getAgentAccessToken() {
  const saved = await loadAgentTokenRecord();
  const clasp = await loadClaspTokenRecord().catch(() => null);
  let record = saved && saved.refresh_token ? saved : clasp;
  if (!record || !record.refresh_token) {
    throw new Error('Agent Google login is missing. Run: node scripts/gas_agent_login.mjs');
  }
  if (!record.access_token || Number(record.expiry_date || 0) < Date.now() + 60_000) {
    if (!record.client_secret) {
      throw new Error('Agent Google login is missing a client secret. Run: node scripts/gas_agent_login.mjs');
    }
    record = await refreshAccessToken(record);
    await saveAgentToken(record);
  }
  if (!record.scope) {
    const info = await fetch('https://www.googleapis.com/oauth2/v3/tokeninfo?access_token=' +
      encodeURIComponent(record.access_token));
    const payload = await info.json();
    record = { ...record, scope: payload.scope || '' };
    await saveAgentToken(record);
  }
  const missing = missingAgentScopes(record.scope);
  if (missing.length) {
    throw new Error('Agent Google login is missing scopes: ' + missing.join(', ') +
      '. Run: node scripts/gas_agent_login.mjs');
  }
  const targets = await loadTargets();
  return { ...record, quotaProjectId: record.quotaProjectId || targets.quotaProjectId };
}

export async function googleApi(record, method, url, body, timeoutMs = 60_000) {
  const headers = {
    Authorization: 'Bearer ' + record.access_token
  };
  if (/sheets\.googleapis\.com/.test(url) && record.quotaProjectId) {
    headers['x-goog-user-project'] = record.quotaProjectId;
  }
  const options = { method, headers, signal: AbortSignal.timeout(timeoutMs) };
  if (body !== undefined) {
    headers['Content-Type'] = 'application/json';
    options.body = typeof body === 'string' ? body : JSON.stringify(body);
  }
  const response = await fetch(url, options);
  const text = await response.text();
  let payload = text;
  try { payload = text ? JSON.parse(text) : {}; } catch { /* keep text */ }
  return { status: response.status, payload, text };
}
