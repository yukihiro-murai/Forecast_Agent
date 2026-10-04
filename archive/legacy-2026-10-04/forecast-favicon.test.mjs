#!/usr/bin/env node
/*
 * forecast-favicon.test.mjs — Webアプリのタブのアイコン（案内キャラ「よみ」）の契約テスト。
 * Forecast_WebApp.js の「タブのアイコン」区間（assets/characters/src/build.py が生成）と doGet / webSetFavicon_ を vm で検証する。
 *
 *   node tests/forecast-favicon.test.mjs
 */

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (p, enc) => readFile(path.join(root, p), enc);
const webAppSrc = await read('Forecast_WebApp.js', 'utf8');
const block = /\/\/ ===== タブのアイコン（自動生成[^\n]*\n(?:\/\/[^\n]*\n)*(const FORECAST_FAVICON_URL = '[^']*';)\n\/\/ ===== \/タブのアイコン =====/.exec(webAppSrc);
assert.ok(block, 'Forecast_WebApp.js に「タブのアイコン」区間がある');
const faviconSrc = block[1];
const uiHtml = await read('Forecast_WebAppUI.html', 'utf8');
const assetPng = await read('assets/characters/favicon/yomi_favicon_64.png');
const assetSvg = (await read('assets/characters/favicon/yomi_favicon.svg', 'utf8')).trim();

/** `function name(` から対応する閉じ括弧までのソースを切り出す。 */
function extractFunction(src, name) {
  const start = src.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} not found`);
  let depth = 0;
  for (let p = src.indexOf('{', start); p < src.length; p++) {
    if (src[p] === '{') depth++;
    else if (src[p] === '}' && --depth === 0) return src.slice(start, p + 1);
  }
  throw new Error(`${name} body not closed`);
}

const ctx = vm.createContext({});
vm.runInContext(faviconSrc + '\nglobalThis.__url = FORECAST_FAVICON_URL;', ctx);
const url = ctx.__url;

// ---- 1. URL の形: PNG の data URI で、末尾が #favicon.png（Apps Script は拡張子の無い URL を黙って捨てる） ----
{
  assert.ok(url.startsWith('data:image/png;base64,'), 'PNG の data URI');
  assert.ok(url.endsWith('#favicon.png'), '末尾は #favicon.png');
  assert.equal(new URL(url).hash, '#favicon.png', '#favicon.png は URL の断片');
}

// ---- 2. 中身: 64x64 の RGBA PNG で、生成物 (assets/characters/favicon/yomi_favicon_64.png) と同じ ----
{
  const b64 = url.slice('data:image/png;base64,'.length, -'#favicon.png'.length);
  const png = Buffer.from(b64, 'base64');
  assert.equal(png.subarray(0, 8).toString('hex'), '89504e470d0a1a0a', 'PNG の署名');
  assert.equal(png.toString('ascii', 12, 16), 'IHDR');
  assert.equal(png.readUInt32BE(16), 64, '幅 64');
  assert.equal(png.readUInt32BE(20), 64, '高さ 64');
  assert.equal(png[24], 8, '8 ビット');
  assert.equal(png[25], 6, 'RGBA (透過)');
  assert.ok(png.equals(assetPng), 'Forecast_WebApp.js のアイコンが古い（cd assets/characters/src && python3 build.py）');
  // 標準の data URL の処理（ブラウザと同じ規則）でも、断片を除いた PNG がそのまま読める
  const fetched = Buffer.from(await (await fetch(url)).arrayBuffer());
  assert.ok(fetched.equals(png), 'fetch(data URL) で同じ PNG が読める');
}

// ---- 3. doGet が setFaviconUrl に渡す。失敗しても画面は返す ----
function runDoGet(setFaviconUrl) {
  const calls = { favicon: [], log: [] };
  const out = {
    setTitle() { return out; },
    addMetaTag() { return out; },
    setXFrameOptionsMode() { return out; },
    setFaviconUrl(u) { calls.favicon.push(u); if (setFaviconUrl) setFaviconUrl(u); return out; },
  };
  const sandbox = {
    HtmlService: {
      createTemplateFromFile(name) { assert.equal(name, 'Forecast_WebAppUI'); return { evaluate: () => out }; },
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
    },
    Logger: { log: (m) => calls.log.push(String(m)) },
    webJson_: () => '{}',
    webGetBootstrap_: () => ({}),
  };
  const c = vm.createContext(sandbox);
  vm.runInContext(faviconSrc + '\n' + extractFunction(webAppSrc, 'webSetFavicon_') + '\n' + extractFunction(webAppSrc, 'doGet')
    + '\nglobalThis.__result = doGet({});', c);
  return { result: c.__result, out, calls };
}
{
  const { result, out, calls } = runDoGet();
  assert.equal(result, out, 'doGet は HtmlOutput を返す');
  assert.deepEqual(calls.favicon, [url], 'FORECAST_FAVICON_URL をそのまま渡す');
}
{
  const { result, out, calls } = runDoGet(() => { throw new Error('boom'); });
  assert.equal(result, out, 'setFaviconUrl が失敗しても画面は返す');
  assert.ok(calls.log.some((m) => m.includes('webSetFavicon_') && m.includes('boom')), '失敗は Logger.log に残す');
}

// ---- 4. 管理ハブ runtime の許可リスト (21 ファイル) を崩さない: アイコン用の別ファイルを作らない ----
{
  const names = (await import('node:fs')).readdirSync(root).filter((f) => /^Forecast_Favicon/.test(f));
  assert.deepEqual(names, [], 'Forecast_Favicon.* を作らない（VNEXT_ADMIN_RUNTIME_FILE_TYPES_ に無いファイルは管理ハブの複製で拒否される）');
}

// ---- 5. ページ内の <link rel="icon"> も同じよみの絵（Apps Script は無視するが、単体で開いたときの表示を揃える） ----
{
  const m = /<link rel="icon" href="data:image\/svg\+xml,([^"]*)">/.exec(uiHtml);
  assert.ok(m, 'Forecast_WebAppUI.html に link rel=icon がある');
  assert.equal(m[1], assetSvg.replace(/"/g, "'").replace(/#/g, '%23'), 'link rel=icon が yomi_favicon.svg と同じ');
}

process.stdout.write('PASS forecast-favicon contract tests\n');
