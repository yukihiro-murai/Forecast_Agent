#!/usr/bin/env node
/*
 * app-run-log-summary.test.mjs — 実行ログの CSV の集計ツール（tools/run-log-summary.mjs）のテスト。
 * 合成の行だけを使う（実データではない）。見出しが Schema.js の RUN と同じこと、段を二重に数えないこと、
 * JST の日の境目、Web の読み込みをトリガーの枠に入れないこと、行の中身・ID を出さないことを確かめる。
 *
 *   node app/tests/app-run-log-summary.test.mjs
 */

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { RUN_COLUMNS, parseCsv, summarize, contextOf } from '../tools/run-log-summary.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));

// ==== 1. 見出しは Schema.js の RUN と同じ ====
{
  const schema = await readFile(path.join(here, '..', 'src', 'Schema.js'), 'utf8');
  const m = schema.match(/RUN: \[([^\]]*)\]/);
  assert.ok(m, 'Schema.js に RUN の列がある');
  assert.deepEqual(m[1].split(',').map(s => s.trim().replace(/'/g, '')), RUN_COLUMNS);
}

// ==== 2. CSV の引用符・カンマ・改行 ====
{
  const rows = parseCsv('﻿a,b\r\n"x,1","he said ""hi""\nnext"\n');
  assert.deepEqual(rows, [['a', 'b'], ['x,1', 'he said "hi"\nnext']]);
}

const head = RUN_COLUMNS.join(',');
const line = (id, kind, fin, ms, status, detail) =>
  [id, 'req-' + id, kind, '', fin, ms, status, detail || '', ''].map(v => /[",\n]/.test(String(v)) ? '"' + String(v).replace(/"/g, '""') + '"' : v).join(',');

// ==== 3. 段ごとの行を足す・重複を除く・JST の日・Web を除く ====
{
  const csv = [head,
    line('a1', 'JOB:FORECAST.RUN', '2026-10-05T01:00:00.000Z', 110000, 'DONE', '{"jobId":"j1","requestedBy":"someone@example.com"}'),
    line('a2', 'JOB:FORECAST.RUN_CALC', '2026-10-05T01:01:30.000Z', 60000, 'DONE'),
    line('a3', 'JOB:FORECAST.RUN_SAVE', '2026-10-05T01:02:00.000Z', 30000, 'DONE'),
    line('a3', 'JOB:FORECAST.RUN_SAVE', '2026-10-05T01:02:00.000Z', 30000, 'DONE'),       // 同じ行を 2 回
    line('b1', 'BACKUP', '2026-10-05T14:59:00.000Z', 20000, 'OK'),                          // JST 23:59 → 10/05
    line('b2', 'HOUSEKEEPING', '2026-10-05T15:00:30.000Z', 40000, 'OK'),                    // JST 00:00 → 10/06
    line('w1', 'PLAN.VIEW', '2026-10-05T02:00:00.000Z', 25000, 'SLOW'),
    line('f1', 'JOB:PLAN.EDIT', '2026-10-05T03:00:00.000Z', 5000, 'FAILED'),
    line('s1', 'SCHEMA.ENSURE', '2026-10-05T03:00:00.000Z', '', 'OK')
  ].join('\n');
  const res = summarize([csv]);
  assert.equal(res.rows, 9);
  assert.equal(res.counted, 8);
  const d5 = res.days.find(d => d.day === '2026-10-05');
  const d6 = res.days.find(d => d.day === '2026-10-06');
  // 110 + 60 + 30 + 20 + 5 = 225 秒 = 3.75 分（PLAN.VIEW は入らない）
  assert.equal(d5.triggerRuns, 5);
  assert.equal(d5.triggerMin, 3.75);
  assert.equal(d5.webSec, 25);
  assert.equal(d5.workspacePct, 1.04);   // 225 秒 / 360 分 = 1.04%
  assert.equal(d5.consumerPct, 4.17);    // 225 秒 / 90 分 = 4.17%
  assert.equal(d6.triggerMin, Math.round(40000 / 600) / 100);
  assert.equal(res.peak.day, '2026-10-05');
  const save = res.kinds.find(k => k.kind === 'JOB:FORECAST.RUN_SAVE');
  assert.equal(save.n, 1);
  const edit = res.kinds.find(k => k.kind === 'JOB:PLAN.EDIT');
  assert.equal(edit.failed, 1);
  assert.ok(res.warnings.some(w => w.includes('run_id')));
  assert.ok(res.warnings.some(w => w.includes('時間の無い行が 1 件')));
  // 行の中身・ID・メールは出さない
  const s = JSON.stringify(res);
  ['example.com', 'req-a1', 'j1', 'a1'].forEach(x => assert.ok(!s.includes('"' + x) && !s.includes(x + '"') && !s.includes('example.com'), x));
}

// ==== 4. 見出しが違えば止める（監査ログなどを渡した） ====
{
  assert.throws(() => summarize(['event_id,occurred_at\n1,2']), /見出しが実行ログ/);
}

// ==== 5. 文脈の分け方 ====
{
  assert.equal(contextOf('JOB:FORECAST.RUN'), 'trigger');
  assert.equal(contextOf('MAINTENANCE.NOTIFY'), 'trigger');
  assert.equal(contextOf('PLAN.VIEW'), 'web');
  assert.equal(contextOf('JOURNAL.RECOVER'), 'other');
}

console.log('app-run-log-summary: ok');
