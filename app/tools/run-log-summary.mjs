#!/usr/bin/env node
/**
 * run-log-summary.mjs — 実行ログ（ログのスプレッドシートの RUN_yyyy_MM シート）の CSV を手元で集計し、
 * トリガーの 1 日の実行時間の上限に対する「業務の段の時間の合計（下限）」を日ごとに出す。
 *
 * CSV の取り方（所有者）: マイドライブ「売上予測アプリ（システム）」の「売上予測アプリ ログ FYxxxx」を開き、
 * RUN_yyyy_MM のシートを選んで「ファイル → ダウンロード → カンマ区切り形式（.csv）」。
 * 原ログ（メール・ID・中身）は公開リポジトリに置かない。このツールも集計した数字だけを出す（行の中身・ID・エラーの文は出さない）。
 *
 *   node app/tools/run-log-summary.mjs <RUN_2026_10.csv> [<RUN_2026_09.csv> ...]
 *   node app/tools/run-log-summary.mjs --json <files...>
 *
 * 注意:
 * - 1 つの処理の段（JOB:FORECAST.RUN → RUN_CALC → RUN_SAVE）はそれぞれ 1 行で、親のまとめの行は無い。段の行を足しても二重にならない。
 *   同じ実行の中で続けて動いた段（inline）も、段ごとに 1 行。
 * - 合計はトリガーの実行時間の「下限」。トリガーの起動・ロック待ち・待ち行列の確認・ログを書く時間は入らない。
 *   上限の判定には Apps Script の「実行数」の画面の時間と、同じアカウントの他のアプリの時間も要る。
 * - PLAN.VIEW（画面の読み込みが遅かった印）は Web アプリの実行で、トリガーの枠には数えない。
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const RUN_COLUMNS = ['run_id', 'request_id', 'kind', 'started_at', 'finished_at', 'duration_ms', 'status', 'detail_json', 'error'];
const TZ = 'Asia/Tokyo';
// トリガーの合計実行時間の 1 日の上限（分）。https://developers.google.com/apps-script/guides/services/quotas（2026-10-05 確認）
export const QUOTA_MIN = { workspace: 360, consumer: 90 };

/** RFC 4180 の CSV を行の配列に（引用符の中のカンマ・改行・"" を扱う） */
export function parseCsv(text) {
  const rows = [];
  let row = [], field = '', q = false;
  const s = String(text).replace(/^﻿/, '');
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (q) {
      if (c === '"') { if (s[i + 1] === '"') { field += '"'; i++; } else q = false; }
      else field += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && s[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows.filter(r => !(r.length === 1 && r[0] === ''));
}

/** 処理の種類を、トリガーで動くもの・Web アプリで動くもの・どちらとも言えないものに分ける */
export function contextOf(kind) {
  if (/^JOB:/.test(kind) || kind === 'BACKUP' || kind === 'HOUSEKEEPING' || kind === 'MAINTENANCE.NOTIFY') return 'trigger';
  if (kind === 'PLAN.VIEW') return 'web';
  return 'other';   // JOURNAL.RECOVER（画面からも処理からも動く）・SCHEMA.ENSURE（時間なし）など
}

function jstDay(iso) {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return null;
  return new Intl.DateTimeFormat('en-CA', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
}

function quantile(sorted, p) {
  if (!sorted.length) return 0;
  const idx = Math.min(sorted.length - 1, Math.ceil(p * sorted.length) - 1);
  return sorted[Math.max(0, idx)];
}

const sec = ms => Math.round(ms / 100) / 10;

/** CSV の本文の配列を集計する。中身・ID・メールは返さない */
export function summarize(csvTexts) {
  const warnings = [];
  const seen = new Set();
  let total = 0, duplicates = 0, badDuration = 0, badDate = 0;
  const byDay = {}, byKind = {};
  csvTexts.forEach((text, fi) => {
    const rows = parseCsv(text);
    if (!rows.length) { warnings.push(`ファイル ${fi + 1}: 空です`); return; }
    const head = rows[0].map(h => h.trim());
    if (RUN_COLUMNS.some((c, j) => head[j] !== c)) {
      throw new Error(`ファイル ${fi + 1}: 見出しが実行ログ（RUN_yyyy_MM）と違います。RUN のシートを CSV で落としてください`);
    }
    rows.slice(1).forEach(r => {
      total++;
      const o = {};
      RUN_COLUMNS.forEach((c, j) => { o[c] = r[j] === undefined ? '' : r[j]; });
      if (o.run_id && seen.has(o.run_id)) { duplicates++; return; }
      if (o.run_id) seen.add(o.run_id);
      const day = jstDay(o.finished_at);
      if (!day) { badDate++; return; }
      const ms = Number(o.duration_ms);
      const hasMs = o.duration_ms !== '' && isFinite(ms) && ms >= 0;
      if (!hasMs) badDuration++;
      const ctx = contextOf(o.kind);
      const k = byKind[o.kind] || (byKind[o.kind] = { kind: o.kind, context: ctx, n: 0, ok: 0, failed: 0, other: 0, ms: [] });
      k.n++;
      if (/^(OK|DONE)$/.test(o.status)) k.ok++; else if (o.status === 'FAILED') k.failed++; else k.other++;
      if (hasMs && ms > 0) k.ms.push(ms);
      const d = byDay[day] || (byDay[day] = { day, triggerMs: 0, triggerRuns: 0, webMs: 0, otherMs: 0, jobsByKind: {} });
      if (!hasMs) return;
      if (ctx === 'trigger') {
        d.triggerMs += ms; d.triggerRuns++;
        d.jobsByKind[o.kind] = (d.jobsByKind[o.kind] || 0) + 1;
      } else if (ctx === 'web') d.webMs += ms;
      else d.otherMs += ms;
    });
  });
  if (duplicates) warnings.push(`同じ run_id の行が ${duplicates} 件あり、1 件として数えました（同じ月を 2 回渡していないか確かめてください）`);
  if (badDate) warnings.push(`終わりの時刻が読めない行が ${badDate} 件あり、除きました`);
  if (badDuration) warnings.push(`時間の無い行が ${badDuration} 件あり、時間の合計に入れていません`);
  const days = Object.values(byDay).sort((a, b) => (a.day < b.day ? -1 : 1)).map(d => ({
    day: d.day, triggerRuns: d.triggerRuns, triggerMin: Math.round(d.triggerMs / 600) / 100,
    workspacePct: Math.round(d.triggerMs / (QUOTA_MIN.workspace * 600) * 100) / 100,
    consumerPct: Math.round(d.triggerMs / (QUOTA_MIN.consumer * 600) * 100) / 100,
    webSec: sec(d.webMs), otherSec: sec(d.otherMs), jobsByKind: d.jobsByKind
  }));
  const peak = days.reduce((m, d) => (!m || d.triggerMin > m.triggerMin ? d : m), null);
  const kinds = Object.values(byKind).sort((a, b) => (a.kind < b.kind ? -1 : 1)).map(k => {
    const s = k.ms.slice().sort((a, b) => a - b);
    return { kind: k.kind, context: k.context, n: k.n, ok: k.ok, failed: k.failed, other: k.other,
      medianSec: sec(quantile(s, 0.5)), p90Sec: sec(quantile(s, 0.9)), maxSec: sec(s.length ? s[s.length - 1] : 0) };
  });
  return { rows: total, counted: total - duplicates - badDate, days, peak, kinds, warnings,
    note: 'トリガーの時間は業務の段の合計（下限）。起動・ロック待ち・他アプリの時間は入らない。上限の判定には Apps Script の実行数の画面の時間も要る' };
}

function print(res) {
  const out = [];
  out.push(`行 ${res.rows}（数えた ${res.counted}）`);
  out.push('', '日ごと（JST）: 日付 / トリガーの段の数 / 合計分（下限） / 360 分枠の % / 90 分枠の % / Web の遅い読み込み 秒');
  res.days.forEach(d => out.push(`${d.day}  ${d.triggerRuns}  ${d.triggerMin}  ${d.workspacePct}%  ${d.consumerPct}%  ${d.webSec}`));
  if (res.peak) out.push('', `ピーク: ${res.peak.day} ${res.peak.triggerMin} 分（360 分枠の ${res.peak.workspacePct}%・下限）`);
  out.push('', '種類ごと: 種類 / 文脈 / 件数 / 成功 / 失敗 / 他 / 中央値秒 / 90% 秒 / 最大秒');
  res.kinds.forEach(k => out.push(`${k.kind}  ${k.context}  ${k.n}  ${k.ok}  ${k.failed}  ${k.other}  ${k.medianSec}  ${k.p90Sec}  ${k.maxSec}`));
  if (res.warnings.length) out.push('', '注意:', ...res.warnings.map(w => '- ' + w));
  out.push('', res.note);
  return out.join('\n');
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2);
  const json = args.includes('--json');
  const files = args.filter(a => a !== '--json');
  if (!files.length) { console.error('使い方: node app/tools/run-log-summary.mjs [--json] <RUN_yyyy_MM.csv> ...'); process.exit(2); }
  const texts = await Promise.all(files.map(f => readFile(f, 'utf8')));
  const res = summarize(texts);
  console.log(json ? JSON.stringify(res, null, 2) : print(res));
}
