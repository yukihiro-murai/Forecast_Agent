#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const runtimeRoot = path.resolve(scriptDir, '..');
const projectRoot = path.resolve(runtimeRoot, '..');
const sourceDir = path.join(runtimeRoot, 'src');
const lockPath = path.join(runtimeRoot, 'vendor-lock.json');
const update = process.argv.includes('--update');

if (!update) {
  throw new Error('UIの更新は明示操作です。`node scripts/vendor-ui.mjs --update` を使用してください。');
}

const files = [
  ['VNext_UX.js', 'VNext_UX.js'],
  ['VNext_InputSidebar.html', 'VNext_InputSidebar.html'],
  ['VNext_GuidanceSidebar.html', 'VNext_GuidanceSidebar.html'],
  ['VNext_HelpSidebar.html', 'VNext_HelpSidebar.html'],
  ['VNext_PlanSidebar.html', 'VNext_PlanSidebar.html'],
  ['VNext_ReviewSidebar.html', 'VNext_ReviewSidebar.html']
];

await mkdir(sourceDir, { recursive: true });
const lock = { version: 1, sources: {} };

for (const [sourceName, targetName] of files) {
  const sourcePath = path.join(projectRoot, sourceName);
  const targetPath = path.join(sourceDir, targetName);
  const original = await readFile(sourcePath, 'utf8');
  const generated = sourceName === 'VNext_UX.js' ? clientOnlyUx(original) : original;
  await writeFile(targetPath, generated, 'utf8');
  lock.sources[sourceName] = {
    sourceSha256: sha256(original),
    vendoredFile: `src/${targetName}`,
    vendoredSha256: sha256(generated)
  };
}

await writeFile(lockPath, `${JSON.stringify(lock, null, 2)}\n`, 'utf8');
process.stdout.write(`Vendored ${files.length} employee UI files.\n`);

function clientOnlyUx(source) {
  let output = replaceBetween(
    source,
    'function vNextHandleOnOpen_()',
    '/** Client FY Bookに従業員向け4項目だけを表示する。 */',
    `function vNextHandleOnOpen_() {
  try {
    return vNextBuildClientMenu_();
  } catch (error) {
    Logger.log('vNextHandleOnOpen_ error: ' + vNextUxErrorText_(error));
    return false;
  }
}

`
  );

  const start = "    if (typeof vNextEngineRunForecast === 'function') {";
  const end = '    vNextRefreshEmployeeViews();';
  const startIndex = output.indexOf(start, output.indexOf('function vNextRequestForecast()'));
  const endIndex = output.indexOf(end, startIndex);
  if (startIndex < 0 || endIndex < 0) throw new Error('vNextRequestForecastのservice bridge範囲を特定できません。');
  output = output.slice(0, startIndex) + '    vNextEngineRunForecast(request);\n' + output.slice(endIndex);
  return output;
}

function replaceBetween(source, startMarker, endMarker, replacement) {
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start);
  if (start < 0 || end < 0 || end <= start) throw new Error(`置換範囲を特定できません: ${startMarker}`);
  return source.slice(0, start) + replacement + source.slice(end);
}

function sha256(value) {
  return createHash('sha256').update(value, 'utf8').digest('hex');
}
