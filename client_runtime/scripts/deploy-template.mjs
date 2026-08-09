#!/usr/bin/env node

import { mkdir, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const runtimeRoot = path.resolve(scriptDir, '..');
const distDir = path.join(runtimeRoot, 'dist');
const scriptId = readArg('--script-id');
const apply = process.argv.includes('--apply');

if (!scriptId || !/^[A-Za-z0-9_-]{20,}$/.test(scriptId)) throw new Error('--script-id にTemplateのbound script IDを指定してください。');

run(process.execPath, [path.join(scriptDir, 'build.mjs')]);
await mkdir(distDir, { recursive: true });
const projectPath = path.join(distDir, '.clasp.json');
await writeFile(projectPath, `${JSON.stringify({ scriptId, rootDir: '.' }, null, 2)}\n`, { mode: 0o600 });

try {
  run('clasp', ['-P', projectPath, 'status']);
  if (!apply) {
    process.stdout.write('DRY RUN: remoteは変更していません。反映する場合は --apply を追加してください。\n');
    process.exit(0);
  }
  run('clasp', ['-P', projectPath, 'push', '--force']);
  process.stdout.write('Template bound scriptへclient-only runtimeを反映しました。\n');
} finally {
  await rm(projectPath, { force: true });
}

function readArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : '';
}

function run(command, args) {
  const result = spawnSync(command, args, { cwd: runtimeRoot, stdio: 'inherit' });
  if (result.status !== 0) throw new Error(`${command} ${args.join(' ')} failed with status ${result.status}`);
}
