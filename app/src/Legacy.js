/**
 * Legacy.js — 計算用ブックのシートと、データ本体の計算用の表（ENG_*）の相互変換。
 * 旧来の計算は、セルの値の型（数値・文字列・日付）と表示形式（文字列が日付や数値に自動で変わるか）に依存する。
 * そこで値は型ごと、表示形式は列ごとの並びとして持ち、計算用ブックに保存したときと同じ状態を組み立て直せるようにする。
 * 表の形のシートは見出しと同じ列名の表に 1 行 = 1 件、表の形でないシートと表の外のセルは ENG_ROWS に行ごとに持つ。
 */

// ---- セル 1 つ ----

function appIsDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]';
}

/**
 * セルの値を「型 1 文字」と文字列にする。e=空 s=文字列 n=数値 b=真偽 d=日時（UTC の ISO） f=数式
 * 数値は String() で書くと Number() で元の値に戻る（-0 だけ別に書く）
 */
function appCellEncode_(v, formula) {
  if (formula) return ['f', String(formula)];
  if (v === '' || v === null || v === undefined) return ['e', ''];
  if (typeof v === 'number') return ['n', Object.is(v, -0) ? '-0' : String(v)];
  if (typeof v === 'boolean') return ['b', v ? 'TRUE' : 'FALSE'];
  if (appIsDate_(v)) {
    if (isNaN(v.getTime())) throw new Error('日時として読めないセルがあります。');
    return ['d', v.toISOString()];
  }
  return ['s', String(v)];
}

function appCellDecode_(t, text) {
  switch (t) {
    case 'e': return '';
    case 's': return String(text);
    case 'n': return text === '-0' ? -0 : Number(text);
    case 'b': return text === 'TRUE';
    case 'd': return new Date(text);
    default: throw new Error('未知のセルの型: ' + t);
  }
}

/** 2 つのセルの値が同じか（型も比べる。日時は時刻の値で比べる） */
function appCellSame_(a, b) {
  const ea = a === '' || a === null || a === undefined;
  const eb = b === '' || b === null || b === undefined;
  if (ea || eb) return ea && eb;
  if (appIsDate_(a) || appIsDate_(b)) return appIsDate_(a) && appIsDate_(b) && a.getTime() === b.getTime();
  return typeof a === typeof b && a === b;
}

/** 行・列（0 始まり）を A1 表記に */
function appA1_(r, c) {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s + (r + 1);
}

// ---- シート 1 枚 ----

/** シートを読む（値・数式・表示形式・大きさ）。表示形式は全行 × fmtCols 列 */
function appSheetSnapshot_(sh, minCols) {
  const maxRows = sh.getMaxRows();
  const maxCols = sh.getMaxColumns();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const has = lastRow > 0 && lastCol > 0;
  const range = has ? sh.getRange(1, 1, lastRow, lastCol) : null;
  const fmtCols = Math.min(maxCols, Math.max(lastCol, minCols || 0));
  return {
    name: sh.getName(), maxRows: maxRows, maxCols: maxCols, lastRow: has ? lastRow : 0, lastCol: has ? lastCol : 0, fmtCols: fmtCols,
    values: has ? range.getValues() : [],
    formulas: has ? range.getFormulas() : [],
    formats: fmtCols ? sh.getRange(1, 1, maxRows, fmtCols).getNumberFormats() : []
  };
}

/**
 * 読んだシートを、計画 planId の行（ENG_シート名 / ENG_ROWS / ENG_FORMATS / ENG_SHEETS）にする。
 * 見出しが旧来の定義と違うシートは、値を失わないよう行ごとに持つ（warnings に残す）。
 */
function appEngEncodeSheet_(planId, snap) {
  const reg = APP_ENGINE_SHEETS[snap.name];
  if (!reg) throw new Error('移す対象でないシートです: ' + snap.name);
  const warnings = [];
  let mode = reg.mode;
  const header = mode === 'table' ? reg.header : null;
  if (mode === 'table') {
    const row1 = snap.values[0] || [];
    if (!(snap.lastRow >= 1 && snap.lastCol >= header.length && header.every((h, j) => row1[j] === h))) {
      mode = 'rows';
      warnings.push('見出しが旧来の定義と違うので、行ごとに持ちます');
    }
  }
  const cell = (r, c) => (c < snap.lastCol ? appCellEncode_(snap.values[r][c], snap.formulas[r] && snap.formulas[r][c]) : ['e', '']);
  const tableRows = [];
  const rowSegs = [];
  let formulaCount = 0;
  for (let r = 0; r < snap.lastRow; r++) {
    if (mode === 'table' && r >= 1) {
      const o = { plan_id: planId, seq: String(r + 1) };
      let types = '';
      header.forEach((h, j) => {
        const e = cell(r, j);
        o[h] = e[1];
        types += e[0];
        if (e[0] === 'f') formulaCount++;
      });
      o._types = types;
      tableRows.push(o);
    }
    const from = mode === 'table' ? header.length : 0;
    const cells = [];
    for (let c = from; c < snap.lastCol; c++) {
      const e = cell(r, c);
      if (e[0] === 'f') formulaCount++;
      cells.push(e[0] + e[1]);
    }
    while (cells.length && cells[cells.length - 1] === 'e') cells.pop();
    if (cells.length) {
      const json = JSON.stringify(cells);
      if (json.length > 49000) throw new Error(snap.name + ' の ' + (r + 1) + ' 行目が長すぎて保存できません。');
      rowSegs.push({ plan_id: planId, sheet: snap.name, row_no: String(r + 1), col_from: String(from + 1), cells_json: json });
    }
  }
  const formatRows = [];
  for (let c = 0; c < snap.fmtCols; c++) {
    const runs = [];
    for (let r = 0; r < snap.maxRows; r++) {
      const f = String((snap.formats[r] || [])[c] || '');
      const last = runs[runs.length - 1];
      if (last && last[2] === f) last[1] = r + 1; else runs.push([r + 1, r + 1, f]);
    }
    const json = JSON.stringify(runs);
    if (json.length > 49000) throw new Error(snap.name + ' の ' + (c + 1) + ' 列目の表示形式が複雑すぎて保存できません。');
    formatRows.push({ plan_id: planId, sheet: snap.name, col: String(c + 1), runs_json: json });
  }
  const strip = o => { const x = Object.assign({}, o); delete x.plan_id; return x; };
  const content = JSON.stringify([snap.name, mode, snap.maxRows, snap.maxCols, snap.lastRow, snap.lastCol, snap.fmtCols,
    tableRows.map(strip), rowSegs.map(strip), formatRows.map(strip)]);
  const sheetRow = {
    plan_id: planId, sheet: snap.name, mode: mode, max_rows: String(snap.maxRows), max_columns: String(snap.maxCols),
    last_row: String(snap.lastRow), last_column: String(snap.lastCol), fmt_columns: String(snap.fmtCols),
    content_hash: appSha256Hex_(content), import_batch_id: '', updated_at: '', updated_by: ''
  };
  return { sheetRow: sheetRow, tableRows: tableRows, rowSegs: rowSegs, formatRows: formatRows, warnings: warnings, formulaCount: formulaCount };
}

/** 保存した行から、シートの中身（値・数式・表示形式・大きさ）を組み立てる */
function appEngDecodeSheet_(sheetRow, tableRows, rowSegs, formatRows) {
  const name = sheetRow.sheet;
  const mode = sheetRow.mode;
  const maxRows = Number(sheetRow.max_rows);
  const maxCols = Number(sheetRow.max_columns);
  const lastRow = Number(sheetRow.last_row);
  const lastCol = Number(sheetRow.last_column);
  const fmtCols = Number(sheetRow.fmt_columns);
  const values = Array.from({ length: lastRow }, () => Array(lastCol).fill(''));
  const formulas = Array.from({ length: lastRow }, () => Array(lastCol).fill(''));
  const put = (r, c, t, text) => {
    if (r >= lastRow || c >= lastCol) throw new Error(name + ' の保存した行が範囲の外にあります（' + appA1_(r, c) + '）。');
    if (t === 'f') formulas[r][c] = text; else values[r][c] = appCellDecode_(t, text);
  };
  if (mode === 'table') {
    const header = APP_ENGINE_SHEETS[name].header;
    header.forEach((h, j) => put(0, j, 's', h));
    tableRows.forEach(o => {
      const r = Number(o.seq) - 1;
      header.forEach((h, j) => put(r, j, o._types.charAt(j), o[h]));
    });
  }
  rowSegs.forEach(seg => {
    const r = Number(seg.row_no) - 1;
    const c0 = Number(seg.col_from) - 1;
    JSON.parse(seg.cells_json).forEach((x, k) => put(r, c0 + k, x.charAt(0), x.slice(1)));
  });
  const formats = Array.from({ length: maxRows }, () => Array(fmtCols).fill(''));
  formatRows.forEach(fr => {
    const c = Number(fr.col) - 1;
    JSON.parse(fr.runs_json).forEach(run => { for (let r = run[0]; r <= run[1]; r++) formats[r - 1][c] = run[2]; });
  });
  return { name: name, mode: mode, maxRows: maxRows, maxCols: maxCols, lastRow: lastRow, lastCol: lastCol, fmtCols: fmtCols,
    values: values, formulas: formulas, formats: formats };
}

/** データ本体から、計画 1 つ分の旧来のシートをすべて（only を渡すとそのシートだけ）組み立てる（{ シート名: 組み立てた中身 }） */
function appEngLoadPlanSheets_(planId, only) {
  const sheets = appReadPlanTable_('ENG_SHEETS', planId).filter(r => !only || only.indexOf(r.sheet) >= 0);
  const group = (name) => {
    const by = {};
    appReadPlanTable_(name, planId).forEach(r => { (by[r.sheet] = by[r.sheet] || []).push(r); });
    return by;
  };
  const segs = group('ENG_ROWS');
  const fmts = group('ENG_FORMATS');
  const out = {};
  sheets.forEach(s => {
    const rows = s.mode === 'table' ? appReadPlanTable_('ENG_' + s.sheet, planId) : [];
    out[s.sheet] = appEngDecodeSheet_(s, rows, segs[s.sheet] || [], fmts[s.sheet] || []);
  });
  return out;
}

// ---- 計算用ブックに書く ----

/** シートの行数・列数をそろえる */
function appFitGrid_(sh, rows, cols) {
  const mr = sh.getMaxRows();
  if (mr < rows) sh.insertRowsAfter(mr, rows - mr); else if (mr > rows) sh.deleteRows(rows + 1, mr - rows);
  const mc = sh.getMaxColumns();
  if (mc < cols) sh.insertColumnsAfter(mc, cols - mc); else if (mc > cols) sh.deleteColumns(cols + 1, mc - cols);
}

/** Sheets が「自動」（形式を付けていない）セルの表示形式として返す文字列 */
const APP_AUTO_FORMAT = '0.###############';

/**
 * 組み立てたシートを計算用ブックに書く。表示形式 → 値 → 数式 の順（保存したときと同じ自動変換になるように）。
 * 表示形式は、保存したときに形式が付いていたセル（「自動」でも空でもないもの）にだけ付け、ほかは新しいシートのままにする。
 * 形式の付いていないセルに日付を書くと Sheets は空の形式にし（旧来の DASHBOARD!B17 など）、「自動」を明示すると
 * 0.############### になるため（2026-10-02 に本物で確認）。
 * 書いた後に値と表示形式を読み戻し、元と違うセルは書き方を変えて直す:
 *   値) 1. 表示形式をいったん外して値を書き、元の表示形式に戻す（旧来が「書いてから形式を付けた」セル）
 *       2. それでも違う文字列のセルは、書式なしテキストにして書く（forcedText に数える）
 *   形式) 「自動」は「自動」を明示する。空は形式を消して書き直す → 形式を付けていないセルの形式を写して書き直す → 空の形式を付けて書き直す、の順に試す
 * 返り値: { mismatches, repaired, forcedText, samples, formatMismatches, formatFixed, formatSamples, blankMethod }
 */
function appEngWriteSheet_(ss, dec) {
  let sh = ss.getSheetByName(dec.name);
  if (!sh) sh = ss.insertSheet(dec.name);
  appFitGrid_(sh, Math.max(1, dec.maxRows), Math.max(1, dec.maxCols));
  const explicit = f => f !== '' && f !== APP_AUTO_FORMAT;
  // 1) 形式の付いていたセルにだけ形式を付ける（列ごとの続いた範囲で）
  for (let c = 0; c < dec.fmtCols; c++) {
    let r0 = 0;
    for (let r = 1; r <= dec.maxRows; r++) {
      if (r < dec.maxRows && dec.formats[r][c] === dec.formats[r0][c]) continue;
      const f = dec.formats[r0][c];
      if (explicit(f)) sh.getRange(r0 + 1, c + 1, r - r0, 1).setNumberFormat(f);
      r0 = r;
    }
  }
  let bad = [];
  let repaired = 0;
  let forced = 0;
  const forcedCells = {};   // 文字列に固定したセル（表示形式は '@' になるので、形式の照合から除く）
  const rewrite = run => sh.getRange(run.r0 + 1, run.c + 1, run.r1 - run.r0 + 1, 1)
    .setValues(dec.values.slice(run.r0, run.r1 + 1).map(row => [run.c < dec.lastCol ? row[run.c] : '']));
  if (dec.lastRow && dec.lastCol) {
    sh.getRange(1, 1, dec.lastRow, dec.lastCol).setValues(dec.values);
    for (let r = 0; r < dec.lastRow; r++) {
      for (let c = 0; c < dec.lastCol; c++) if (dec.formulas[r][c]) sh.getRange(r + 1, c + 1).setFormula(dec.formulas[r][c]);
    }
    const diff = () => {
      const back = sh.getRange(1, 1, dec.lastRow, dec.lastCol).getValues();
      const out = [];
      for (let r = 0; r < dec.lastRow; r++) {
        for (let c = 0; c < dec.lastCol; c++) {
          // 数式のセルは比べない（保存したのは数式で、値は計算でできる）
          if (dec.formulas[r][c]) continue;
          if (!appCellSame_(dec.values[r][c], back[r][c])) out.push([r, c]);
        }
      }
      return out;
    };
    bad = diff();
    if (bad.length) {
      const before = bad.length;
      // 値 1) 列ごとの続いた範囲に分けて、形式を外して書き、元の形式に戻す
      appCellRuns_(bad.filter(x => !dec.formulas[x[0]][x[1]])).forEach(run => {
        const rg = sh.getRange(run.r0 + 1, run.c + 1, run.r1 - run.r0 + 1, 1);
        rg.clearFormat();
        rewrite(run);
        if (run.c < dec.fmtCols) {
          const fs = dec.formats.slice(run.r0, run.r1 + 1).map(row => row[run.c]);
          if (fs.every(explicit)) rg.setNumberFormats(fs.map(f => [f]));
        }
      });
      bad = diff();
      repaired = before - bad.length;
      // 値 2) 残った文字列のセルは書式なしテキストにして書く
      const strs = bad.filter(x => typeof dec.values[x[0]][x[1]] === 'string' && !dec.formulas[x[0]][x[1]]);
      appCellRuns_(strs).forEach(run => {
        sh.getRange(run.r0 + 1, run.c + 1, run.r1 - run.r0 + 1, 1).setNumberFormat('@');
        rewrite(run);
        forced += run.r1 - run.r0 + 1;
        for (let r = run.r0; r <= run.r1; r++) forcedCells[r + ':' + run.c] = true;
      });
      if (strs.length) bad = diff();
    }
  }
  // 2) 表示形式を読み戻して確かめる（文字列に固定したセルは除く）
  const fdiff = () => {
    if (!dec.fmtCols) return [];
    const back = sh.getRange(1, 1, dec.maxRows, dec.fmtCols).getNumberFormats();
    const out = [];
    for (let r = 0; r < dec.maxRows; r++) {
      for (let c = 0; c < dec.fmtCols; c++) {
        const got = String((back[r] || [])[c] || '');
        if (got !== dec.formats[r][c] && !(got === '@' && forcedCells[r + ':' + c])) out.push([r, c, got]);
      }
    }
    return out;
  };
  let fbad = fdiff();
  let formatFixed = 0;
  let blankMethod = '';
  if (fbad.length) {
    const before = fbad.length;
    const runsOf = (pred) => appCellRuns_(fbad.filter(x => pred(dec.formats[x[0]][x[1]])));
    // 「自動」・形式の付いたセル: その形式を付け直す
    runsOf(f => f !== '').forEach(run => {
      const f = dec.formats[run.r0][run.c];
      const fs = dec.formats.slice(run.r0, run.r1 + 1).map(row => row[run.c]);
      sh.getRange(run.r0 + 1, run.c + 1, run.r1 - run.r0 + 1, 1).setNumberFormats(fs.map(x => [x === APP_AUTO_FORMAT ? 'General' : x || f]));
    });
    // 空の形式のセル: 書き方を順に試す（最初に元どおりになったものを残りにも使う）
    let blanks = runsOf(f => f === '');
    let pristine = null;
    const methods = [
      ['形式を消して書き直す', run => { sh.getRange(run.r0 + 1, run.c + 1, run.r1 - run.r0 + 1, 1).clearFormat(); rewrite(run); }],
      ['形式のないセルの形式を写して書き直す', run => {
        if (!pristine) pristine = ss.insertSheet('_FMT_' + Utilities.getUuid().slice(0, 8));
        pristine.getRange(1, 1).copyTo(sh.getRange(run.r0 + 1, run.c + 1, run.r1 - run.r0 + 1, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        rewrite(run);
      }],
      ['空の形式を付けて書き直す', run => { sh.getRange(run.r0 + 1, run.c + 1, run.r1 - run.r0 + 1, 1).setNumberFormat(''); rewrite(run); }]
    ];
    for (let i = 0; i < methods.length && blanks.length; i++) {
      try { blanks.forEach(methods[i][1]); } catch (e) { Logger.log('空の形式の直し方「' + methods[i][0] + '」: ' + (e && e.message ? e.message : e)); continue; }
      fbad = fdiff();
      const left = runsOf(f => f === '');
      if (left.length < blanks.length || !left.length) blankMethod = methods[i][0];
      blanks = left;
    }
    if (pristine) ss.deleteSheet(pristine);
    fbad = fdiff();
    formatFixed = before - fbad.length;
  }
  return { sheet: sh, mismatches: bad.length, repaired: repaired, forcedText: forced, samples: bad.slice(0, 5).map(x => appA1_(x[0], x[1])),
    formatMismatches: fbad.length, formatFixed: formatFixed, blankMethod: blankMethod,
    formatSamples: fbad.slice(0, 5).map(x => appA1_(x[0], x[1]) + ' 「' + dec.formats[x[0]][x[1]] + '」/「' + x[2] + '」') };
}

/** セルの一覧を、列ごとに続いた行の範囲にまとめる */
function appCellRuns_(cells) {
  const byCol = {};
  cells.forEach(x => { (byCol[x[1]] = byCol[x[1]] || []).push(x[0]); });
  const runs = [];
  Object.keys(byCol).forEach(k => {
    const c = Number(k);
    const rows = byCol[k].sort((a, b) => a - b);
    let r0 = rows[0];
    let prev = rows[0];
    rows.slice(1).concat([null]).forEach(r => {
      if (r !== null && r === prev + 1) { prev = r; return; }
      runs.push({ c: c, r0: r0, r1: prev });
      r0 = r; prev = r;
    });
  });
  return runs;
}
