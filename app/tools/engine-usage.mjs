#!/usr/bin/env node
/**
 * engine-usage.mjs — 旧来のコード（Forecast_Agent.js・Forecast_WebApp.js）のうち、新アプリが使う関数と使わない関数を分ける。
 *
 * 新アプリが呼ぶ関数（build-engine.mjs の ENGINE_EXPORTS）と、トップレベルの文から、名前で参照をたどる（文字列の中の名前も
 * 参照として数える＝広めに「使う」とする）。たどり着かない関数は、新アプリでは動かない（旧ブックのメニュー・ダイアログ・
 * サイドバー・旧 Web アプリの入口・管理ハブの集約など）。
 *
 *   node app/tools/engine-usage.mjs          … 一覧を出す
 *   node app/tools/engine-usage.mjs --json   … JSON で出す
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { ENGINE_EXPORTS } from './build-engine.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const files = ['Forecast_Agent.js', 'Forecast_WebApp.js'];

/**
 * JS を文字ごとに読み、文字列・テンプレート文字列・コメント・正規表現の中を除いた「コード」と、トップレベルの関数の範囲を返す。
 * コードでは、文字列などの中身を空白にする（位置は元のまま）。テンプレート文字列の ${...} の中はコードとして残す
 */
function scan(src) {
  const out = src.split('');
  const blank = (a, b) => { for (let i = a; i < b; i++) if (out[i] !== '\n') out[i] = ' '; };
  const fns = [];
  let depth = 0, i = 0, prev = '';   // prev: 直前の意味のある文字（正規表現か割り算かを決める）
  const tmplStack = [];               // ${ の中の深さ
  let fnStart = null;
  const wordBefore = (k) => { let j = k - 1; while (j >= 0 && /\s/.test(src[j])) j--; let e = j + 1; while (j >= 0 && /[A-Za-z0-9_$]/.test(src[j])) j--; return src.slice(j + 1, e); };
  while (i < src.length) {
    const c = src[i], n = src[i + 1];
    if (c === '/' && n === '/') { const e = src.indexOf('\n', i); const end = e < 0 ? src.length : e; blank(i, end); i = end; continue; }
    if (c === '/' && n === '*') { const e = src.indexOf('*/', i + 2); const end = e < 0 ? src.length : e + 2; blank(i, end); i = end; continue; }
    if (c === "'" || c === '"') { let j = i + 1; while (j < src.length && src[j] !== c && src[j] !== '\n') { if (src[j] === '\\') j++; j++; } blank(i + 1, j); i = j + 1; prev = 'a'; continue; }
    if (c === '`') {
      let j = i + 1;
      while (j < src.length && src[j] !== '`') {
        if (src[j] === '\\') { j += 2; continue; }
        if (src[j] === '$' && src[j + 1] === '{') break;
        j++;
      }
      blank(i + 1, j);
      if (src[j] === '$') { tmplStack.push(depth); depth++; i = j + 2; prev = '{'; continue; }
      i = j + 1; prev = 'a'; continue;
    }
    if (c === '}' && tmplStack.length && tmplStack[tmplStack.length - 1] === depth - 1) {
      // テンプレート文字列の ${ ... } の終わり: 続きの文字列を読む
      tmplStack.pop(); depth--;
      let j = i + 1;
      while (j < src.length && src[j] !== '`') {
        if (src[j] === '\\') { j += 2; continue; }
        if (src[j] === '$' && src[j + 1] === '{') break;
        j++;
      }
      blank(i + 1, j);
      if (src[j] === '$') { tmplStack.push(depth); depth++; i = j + 2; prev = '{'; continue; }
      i = j + 1; prev = 'a'; continue;
    }
    if (c === '/') {
      const w = wordBefore(i);
      const isRegex = !prev || '(,=:[!&|?{};+-*%<>~^'.indexOf(prev) >= 0 || /^(return|typeof|case|in|of|void|delete|throw|new)$/.test(w);
      if (isRegex) {
        let j = i + 1, cls = false;
        while (j < src.length && src[j] !== '\n') {
          if (src[j] === '\\') { j += 2; continue; }
          if (src[j] === '[') cls = true; else if (src[j] === ']') cls = false; else if (src[j] === '/' && !cls) break;
          j++;
        }
        blank(i + 1, j); i = j + 1; while (/[a-z]/.test(src[i] || '')) i++; prev = 'a'; continue;
      }
    }
    if (depth === 0 && c === 'f' && src.startsWith('function', i) && (i === 0 || /[\s;}]/.test(src[i - 1]))) {
      const m = /^function\s+([A-Za-z0-9_$]+)\s*\(/.exec(src.slice(i, i + 200));
      if (m) fnStart = { name: m[1], start: i };
    }
    if (c === '{') { depth++; }
    else if (c === '}') {
      depth--;
      if (depth === 0 && fnStart) { fns.push(Object.assign(fnStart, { end: i + 1 })); fnStart = null; }
    }
    if (!/\s/.test(c)) prev = /[A-Za-z0-9_$)\]]/.test(c) ? 'a' : c;
    i++;
  }
  return { code: out.join(''), fns: fns };
}

/** トップレベルの関数（本体はコードだけ）と、関数の外の文 */
function parse(src, file) {
  const sc = scan(src);
  const fns = sc.fns.map((f) => ({ name: f.name, file, start: f.start, end: f.end, body: sc.code.slice(f.start, f.end), chars: f.end - f.start,
    line: src.slice(0, f.start).split('\n').length }));
  let outside = '', last = 0;
  fns.forEach((f) => { outside += sc.code.slice(last, f.start); last = f.end; });
  outside += sc.code.slice(last);
  return { fns, outside };
}

export async function analyze() {
  const all = [];
  let outside = '';
  for (const f of files) {
    const src = await readFile(path.join(root, f), 'utf8');
    const p = parse(src, f);
    all.push(...p.fns);
    outside += '\n' + p.outside;
  }
  const byName = new Map(all.map((x) => [x.name, x]));
  const names = [...byName.keys()];
  const idRe = /[A-Za-z_$][A-Za-z0-9_$]*/g;
  // 文字列とコメントの中の名前は数えない（メニューの項目・トリガーの名前など、旧ブックでだけ動く呼び出し）。
  // テンプレート文字列の ${...} の中はコードなので残す
  const refs = (text) => new Set((text.match(idRe) || []).filter((w) => byName.has(w)));
  const seen = new Set();
  const queue = [...ENGINE_EXPORTS.filter((n) => byName.has(n)), ...refs(outside)];
  while (queue.length) {
    const n = queue.pop();
    if (seen.has(n)) continue;
    seen.add(n);
    refs(byName.get(n).body).forEach((r) => { if (!seen.has(r)) queue.push(r); });
  }
  const unused = all.filter((x) => !seen.has(x.name));
  const size = (xs) => xs.reduce((s, x) => s + x.chars, 0);
  return { total: all.length, used: all.length - unused.length, unused, unusedChars: size(unused), totalChars: size(all), names };
}

if (import.meta.url === 'file://' + process.argv[1]) {
  const r = await analyze();
  if (process.argv.includes('--json')) {
    console.log(JSON.stringify({ total: r.total, used: r.used, unusedChars: r.unusedChars, totalChars: r.totalChars,
      unused: r.unused.map((x) => ({ name: x.name, file: x.file, line: x.line, chars: x.chars })) }, null, 2));
  } else {
    console.log(`関数 ${r.total} のうち、新アプリから届くもの ${r.used}・届かないもの ${r.unused.length}（${r.unusedChars} / ${r.totalChars} 文字）`);
    r.unused.forEach((x) => console.log(`${x.file}:${x.line}\t${x.name}\t${x.chars}`));
  }
}
