/**
 * Migrate.js — 旧ブック（クライアント別売上予測）からデータ本体への取り込み（段階2）。所有者だけが実行する。
 * 1) 試し（appMigrationDryRun_）: 旧ブックを読み、保存の形にして戻すと元と同じになるか、計算用ブックに組み立てて読み戻すと元と同じになるかを確かめる。データ本体には書かない。
 * 2) 取り込み（appMigrationImport_）: 試しのときと内容が同じこと（内容のハッシュ）を確かめてから、計画 1 つ分の ENG_* を入れ替える。
 * 旧ブックは読むだけで、何も変えない。
 */

/** URL か ID からスプレッドシートの ID を取り出す */
function appParseBookId_(input) {
  const s = String(input || '').trim();
  const m = /\/spreadsheets\/d\/([a-zA-Z0-9_-]{20,})/.exec(s);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  throw new Error('旧ブック（クライアント別売上予測）の URL を入れてください。');
}

/** 旧ブックを読む（書き込まない） */
function appLegacyRead_(bookId) {
  let ss = null;
  try { ss = SpreadsheetApp.openById(bookId); } catch (e) { throw new Error('旧ブックを開けません。URL と共有を確かめてください。'); }
  const config = ss.getSheetByName('CONFIG');
  if (!config) throw new Error('CONFIG シートがありません。旧ブック（クライアント別売上予測）の URL か確かめてください。');
  const client = String(config.getRange('B2').getValue() === null ? '' : config.getRange('B2').getValue()).trim();
  const fy = Number(config.getRange('B3').getValue());
  const people = String(config.getRange('B4').getValue() || '');
  if (!client) throw new Error('旧ブックの CONFIG にクライアント名（B2）がありません。');
  if (!(fy > 2000)) throw new Error('旧ブックの CONFIG の予測年度（B3）が数値ではありません。');
  const snaps = [];
  const missing = [];
  Object.keys(APP_ENGINE_SHEETS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { missing.push(name); return; }
    const reg = APP_ENGINE_SHEETS[name];
    snaps.push(appSheetSnapshot_(sh, reg.header ? reg.header.length : 0));
  });
  const names = ss.getSheets().map(s => s.getName());
  return {
    bookId: bookId, title: ss.getName(), locale: ss.getSpreadsheetLocale(), timeZone: ss.getSpreadsheetTimeZone(),
    client: client, fy: fy, people: people, snaps: snaps, missing: missing,
    notMigrated: names.filter(n => APP_ENGINE_NOT_MIGRATED.indexOf(n) >= 0),
    unknown: names.filter(n => !APP_ENGINE_SHEETS[n] && APP_ENGINE_NOT_MIGRATED.indexOf(n) < 0)
  };
}

/** 旧ブックの内容全体のハッシュ（試しと取り込みで同じ内容かを確かめる） */
function appLegacyContentHash_(src, encoded) {
  return appSha256Hex_(JSON.stringify([src.bookId, src.locale, src.timeZone, src.client, src.fy, src.people, src.missing,
    encoded.map(e => [e.sheetRow.sheet, e.sheetRow.content_hash])]));
}

/** この旧ブックから作った計画（無ければ null） */
function appPlanOfBook_(bookId) {
  return appReadTable_('PLANS').filter(p => p.source_book_id === bookId)[0] || null;
}

function appMigrationDryRun_(ctx, input) {
  const t0 = new Date().getTime();
  const bookId = appParseBookId_(input && input.bookUrl);
  const src = appLegacyRead_(bookId);
  const t1 = new Date().getTime();
  const encoded = src.snaps.map(snap => appEngEncodeSheet_('PL-DRYRUN', snap));
  const t2 = new Date().getTime();
  return appWithLock_(() => {
    const scratch = appScratchBook_();
    try { scratch.setSpreadsheetTimeZone(src.timeZone); scratch.setSpreadsheetLocale(src.locale); } catch (e) {
      Logger.log('計算用ブックの地域設定: ' + (e && e.message ? e.message : e));
    }
    const placeholder = appScratchReset_(scratch);
    const sheets = src.snaps.map((snap, i) => {
      const enc = encoded[i];
      const dec = appEngDecodeSheet_(enc.sheetRow, enc.tableRows, enc.rowSegs, enc.formatRows);
      const lossless = appEngCompare_(snap, dec);
      const w = appEngWriteSheet_(scratch, dec, snap.values);
      return {
        sheet: snap.name, mode: enc.sheetRow.mode, rows: snap.lastRow, cols: snap.lastCol, records: enc.tableRows.length,
        rowSegments: enc.rowSegs.length, formulas: enc.formulaCount, warnings: enc.warnings,
        losslessMismatch: lossless.length, losslessSamples: lossless.slice(0, 5),
        mismatch: w.mismatches, repaired: w.repaired, forcedText: w.forcedText, samples: w.samples,
        formatMismatches: w.formatMismatches, formatFixed: w.formatFixed, formatSamples: w.formatSamples
      };
    });
    if (scratch.getSheets().length > 1) scratch.deleteSheet(placeholder);
    const plan = appIsSetUp_() ? (() => { try { return appPlanOfBook_(bookId); } catch (e) { return null; } })() : null;
    const lastBatch = plan ? appLastImport_(plan.plan_id) : null;
    const contentHash = appLegacyContentHash_(src, encoded);
    const t3 = new Date().getTime();
    return {
      timing: { readMs: t1 - t0, encodeMs: t2 - t1, scratchMs: t3 - t2, totalMs: t3 - t0 },
      book: { title: src.title, locale: src.locale, timeZone: src.timeZone },
      client: src.client, fy: src.fy, peopleCount: src.people.split(',').map(x => x.trim()).filter(Boolean).length,
      sheets: sheets, missing: src.missing, notMigrated: src.notMigrated, unknown: src.unknown,
      lossless: sheets.every(x => x.losslessMismatch === 0),
      faithful: sheets.every(x => x.mismatch === 0 && x.forcedText === 0 && x.formatMismatches === 0),
      contentHash: contentHash,
      existingPlan: plan ? { planId: plan.plan_id, lastImportedAt: lastBatch ? lastBatch.finished_at : '', unchanged: !!(lastBatch && lastBatch.content_hash === contentHash),
        runsSinceImport: appRunsSinceImport_(plan.plan_id, lastBatch) } : null,
      audit: { entityId: plan ? plan.plan_id : '' }
    };
  });
}

/** 最後の取り込みの後に、新アプリで実行した予測の数（取り込み直すと、計算用の表はそれを旧ブックの内容で置き換える） */
function appRunsSinceImport_(planId, lastBatch) {
  try {
    const after = lastBatch ? lastBatch.finished_at : '';
    return appReadTable_('FORECAST_RUNS').filter(r => r.plan_id === planId && r.finished_at >= after).length;   // 時刻は秒までなので同じ秒も数える
  } catch (e) {
    return 0;   // まだ表がない（版 3 の前）
  }
}

/** 計画の最後の取り込み（成功したもの） */
function appLastImport_(planId) {
  const rows = appReadTable_('IMPORT_BATCHES').filter(b => b.plan_id === planId && b.status === 'OK');
  return rows.length ? rows[rows.length - 1] : null;
}

function appMigrationImport_(ctx, input) {
  const bookId = appParseBookId_(input && input.bookUrl);
  const expected = String(input && input.contentHash || '');
  if (!expected) throw new Error('先に「試しに読む」を実行してください。');
  appEnsureTables_(ctx);
  const started = appNowIso_();
  const src = appLegacyRead_(bookId);
  // 内容が試しと同じか・保存の形から元に戻せるかを、書き込む前に確かめる（計画の ID は内容のハッシュに入らない）
  const encoded = src.snaps.map(snap => appEngEncodeSheet_('PL-TEMP', snap));
  const contentHash = appLegacyContentHash_(src, encoded);
  if (contentHash !== expected) throw new Error('旧ブックの内容が「試しに読む」のときから変わりました。もう一度「試しに読む」から行ってください。');
  if (encoded.some((enc, i) => appEngCompare_(src.snaps[i], appEngDecodeSheet_(enc.sheetRow, enc.tableRows, enc.rowSegs, enc.formatRows)).length)) {
    throw new Error('保存の形にすると元に戻らないシートがあります。取り込みを止めました。');
  }
  return appWithLock_(() => {
    appJournalRecover_(ctx);   // 前の保存が途中で止まっていれば、先に書き終える
    const now = appNowIso_();
    const ops = [];
    // 計画とクライアント（無ければ作る。前に取り込んだ計画があれば、そのクライアントを使う）
    let plan = appPlanOfBook_(bookId);
    const clients = appReadTable_('CLIENTS');
    const normalized = appNormalizeName_(src.client);
    let client = (plan && clients.filter(c => c.client_id === plan.client_id)[0]) || clients.filter(c => c.normalized_name === normalized)[0];
    if (!client) {
      client = { client_id: appId_('CL'), client_name: src.client.slice(0, 100), zac_code: '', normalized_name: normalized, aliases_json: '[]',
        is_active: true, note: '旧ブックから取り込み', created_at: now, created_by: ctx.actor, updated_at: now, updated_by: ctx.actor, row_version: 1 };
      ops.push({ table: 'CLIENTS', mode: 'ensure', rows: [client] });
    }
    const planId = plan ? plan.plan_id : appId_('PL');
    encoded.forEach(e => { [e.sheetRow].concat(e.tableRows, e.rowSegs, e.formatRows).forEach(o => { o.plan_id = planId; }); });
    const last = plan ? appLastImport_(planId) : null;
    if (last && last.content_hash === contentHash) {
      return { planId: planId, unchanged: true, lastImportedAt: last.finished_at, audit: { entityId: planId, clientId: client.client_id } };
    }
    const batchId = appId_('IB');
    const written = {};
    // 表の形のシート（旧ブックに無いシートの行も消す）
    Object.keys(APP_ENGINE_SHEETS).filter(n => APP_ENGINE_SHEETS[n].mode === 'table').forEach(name => {
      const enc = encoded.filter(e => e.sheetRow.sheet === name)[0];
      const rows = enc && enc.sheetRow.mode === 'table' ? enc.tableRows : [];
      ops.push(appOpReplacePlan_('ENG_' + name, planId, null, rows));
      written['ENG_' + name] = rows.length;
    });
    const segs = [].concat.apply([], encoded.map(e => e.rowSegs));
    const fmts = [].concat.apply([], encoded.map(e => e.formatRows));
    const sheetRows = encoded.map(e => Object.assign({}, e.sheetRow, { import_batch_id: batchId, updated_at: now, updated_by: ctx.actor }));
    ops.push(appOpReplacePlan_('ENG_ROWS', planId, null, segs));
    ops.push(appOpReplacePlan_('ENG_FORMATS', planId, null, fmts));
    ops.push(appOpReplacePlan_('ENG_SHEETS', planId, null, sheetRows));
    written.ENG_ROWS = segs.length;
    written.ENG_FORMATS = fmts.length;
    written.ENG_SHEETS = sheetRows.length;
    const planPatch = { client_id: client.client_id, fy: String(src.fy), client_label: src.client, people_csv: src.people, source_book_id: bookId,
      locale: src.locale, time_zone: src.timeZone, state: 'ACTIVE', note: '' };
    const before = plan ? appStripRow_(plan) : null;
    if (plan) {
      ops.push({ table: 'PLANS', mode: 'patch', key: { plan_id: planId }, patch: planPatch, actor: ctx.actor });
    } else {
      ops.push({ table: 'PLANS', mode: 'ensure', rows: [Object.assign({ plan_id: planId, created_at: now, created_by: ctx.actor, updated_at: now, updated_by: ctx.actor,
        row_version: 1 }, planPatch)] });
    }
    appJournalRun_(ctx, '旧ブックの取り込み（' + batchId + '）', planId, ops);
    plan = appPlanOf_(planId);
    // 書いた後にデータ本体から読み戻して、旧ブックと同じに戻るかを確かめる（本物のシートでしか分からない差を拾う）
    const loaded = appEngLoadPlanSheets_(planId);
    const verify = src.snaps.map(snap => {
      const dec = loaded[snap.name];
      const bad = dec ? appEngCompare_(snap, dec) : ['シートが保存されていない'];
      return { sheet: snap.name, mismatch: bad.length, samples: bad.slice(0, 5) };
    }).filter(v => v.mismatch);
    const status = verify.length ? 'VERIFY_FAILED' : 'OK';
    const summary = { title: src.title, sheets: encoded.length, missing: src.missing, written: written, verify: verify };
    appInsertRows_('IMPORT_BATCHES', [{ batch_id: batchId, plan_id: planId, source_type: 'LEGACY_BOOK', source_id: bookId, content_hash: contentHash,
      summary_json: summary, status: status, started_at: started, finished_at: appNowIso_(), actor_email: ctx.actor }]);
    return { planId: planId, batchId: batchId, unchanged: false, written: written, verified: !verify.length, verify: verify, before: before,
      audit: { entityId: planId, clientId: client.client_id } };
  });
}

// ---- 取り込んだ旧ブックの片付け（旧ブックは使わない。2026-10-04 村井さん）。消さない: アーカイブのフォルダの「旧ブック」へ移す ----

function appLegacyArchiveFolder_() {
  const archiveId = appProps_().getProperty(APP_PROP.archiveFolderId);
  if (!archiveId) throw new Error('アーカイブのフォルダがありません。「初期設定」をもう一度実行してください。');
  const archive = DriveApp.getFolderById(archiveId);
  const it = archive.getFoldersByName('旧ブック');
  return it.hasNext() ? it.next() : archive.createFolder('旧ブック');
}

/** 取り込んだ旧ブックの一覧（名前・場所・アーカイブに移したか） */
function appLegacyBooks_() {
  const names = appClientNameMap_();
  const archiveId = appProps_().getProperty(APP_PROP.archiveFolderId);
  return appReadTable_('PLANS').filter(p => p.source_book_id).map(p => {
    const o = { planId: p.plan_id, clientName: names[p.client_id] || p.client_label, fy: p.fy, bookId: p.source_book_id };
    try {
      const f = DriveApp.getFileById(p.source_book_id);
      o.name = f.getName(); o.url = f.getUrl(); o.trashed = f.isTrashed();
      const parents = [];
      const it = f.getParents();
      while (it.hasNext()) { const x = it.next(); parents.push({ id: x.getId(), name: x.getName() }); }
      o.folder = parents.map(x => x.name).join('、');
      o.archived = parents.some(x => x.name === '旧ブック') || parents.some(x => x.id === archiveId);
    } catch (e) { o.error = '開けません（' + String(e && e.message ? e.message : e).slice(0, 80) + '）'; }
    return o;
  });
}

/** 取り込んだ旧ブックを、アーカイブのフォルダの「旧ブック」へ移す（消さない。移せないものは理由を返す） */
function appArchiveLegacyBooks_(ctx) {
  const folder = appLegacyArchiveFolder_();
  const res = appLegacyBooks_().map(b => {
    if (b.error || b.archived || b.trashed) return { bookId: b.bookId, name: b.name, skipped: b.error || (b.archived ? 'すでに移した' : 'ゴミ箱にある') };
    try { DriveApp.getFileById(b.bookId).moveTo(folder); return { bookId: b.bookId, name: b.name, moved: true }; }
    catch (e) { return { bookId: b.bookId, name: b.name, skipped: '移せません（' + String(e && e.message ? e.message : e).slice(0, 80) + '）' }; }
  });
  return { books: res, moved: res.filter(x => x.moved).length, audit: { entityId: 'LEGACY_BOOKS' } };
}
