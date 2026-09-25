#!/usr/bin/env node

import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const scriptId = readArg('--script-id');
if (!scriptId || !/^[A-Za-z0-9_-]{20,}$/.test(scriptId)) throw new Error('--script-id が必要です。');

const temporary = await mkdtemp(path.join(os.tmpdir(), 'forecast-client-runtime-'));
try {
  const projectPath = path.join(temporary, '.clasp.json');
  await writeFile(projectPath, `${JSON.stringify({ scriptId, rootDir: '.' }, null, 2)}\n`, { mode: 0o600 });
  run('clasp', ['-P', projectPath, 'pull'], temporary);
  run(process.execPath, [path.join(scriptDir, 'verify.mjs'), '--dir', temporary], temporary);
  process.stdout.write('PASS remote Template contains only the verified client runtime.\n');
} finally {
  await rm(temporary, { recursive: true, force: true });
}

function readArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : '';
}

function run(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, stdio: 'inherit' });
  if (result.status !== 0) throw new Error(`${command} ${args.join(' ')} failed with status ${result.status}`);
}
