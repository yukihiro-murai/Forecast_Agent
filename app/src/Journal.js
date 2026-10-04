/**
 * Journal.js — データ本体への書き込みの控え（途中で止まっても、続きから書き直せるようにする）。
 *
 * 予測・保存・実行の結果を戻すときは、何枚もの表（ENG_シート名・ENG_ROWS・ENG_FORMATS・ENG_SHEETS・記録の表）を続けて書く。
 * 途中で止まる（1 回の上限 6 分・Google 側の一時的なエラー）と、表どうしが食い違い、その計画を組み立てられなくなる。
 * そこで、書く前に「書き終えた後の表の中身」を控え（システムのフォルダのファイル）に置いてから書き、書き終えたら控えを消す。
 * 控えが残っていれば、次の書き込みの前に控えのとおりに書き直す（何度書き直しても同じ結果になる）。
 *
 * 控えの中身（ops）:
 *   { table, mode: 'replace', rows }       … 表全体をこの行にする（ほかの計画の行も含めた、書き終えた後の全部。前の版の控え）
 *   { table, mode: 'replacePlan', planId, sheets, rows }
 *                                          … 計画 planId の行（sheets を渡せばそのシートの行だけ）を rows に入れ替える。
 *                                            ほかの計画の行は控えに入れない（控えが小さく、書く行も少ない）。元の行があった場所に置く
 *   { table, mode: 'ensure', rows }        … キーの無い行だけ足す（追記だけの表・記録の表）
 *   { table, mode: 'patch', key, patch }   … キーの行の列を、この値にする（同じ値なら書かない）
 */
const APP_JOURNAL_PROP = 'APP_WRITE_JOURNAL';

/** 書きかけの控えがあるか（{ id, fileId, label, planId, at } か null） */
function appJournalPending_() {
  const raw = appProps_().getProperty(APP_JOURNAL_PROP);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return { id: '', fileId: '', label: '（読めない控え）', planId: '', at: '' }; }
}

/**
 * 控えを置いてから ops を書き、書き終えたら控えを消す（ロックの中で呼ぶ）。返り値は表ごとの書いた件数。
 * 途中で例外になったら控えは残る（次の書き込みの前に appJournalRecover_ が書き直す）
 */
function appJournalRun_(ctx, label, planId, ops) {
  appJournalRecover_(ctx);   // 前の書きかけがあれば、先に書き終える
  const id = appId_('JNL');
  const body = JSON.stringify({ id: id, label: label, planId: planId || '', actor: ctx.actor, createdAt: appNowIso_(), ops: ops });
  const folderId = appProps_().getProperty(APP_PROP.folderId);
  if (!folderId) throw new Error('初期設定がまだです。');
  const file = DriveApp.getFolderById(folderId).createFile(APP_FILES.journalPrefix + id + '.json', body, MimeType.PLAIN_TEXT);
  appProps_().setProperty(APP_JOURNAL_PROP, JSON.stringify({ id: id, fileId: file.getId(), label: label, planId: planId || '', at: appNowIso_() }));
  const written = appJournalApply_(ops);
  appJournalFinish_(file.getId());
  return written;
}

/** 書きかけの控えがあれば、控えのとおりに書き直して消す（ロックの中で呼ぶ）。返り値は書き直した控え（無ければ null） */
function appJournalRecover_(ctx) {
  const p = appJournalPending_();
  if (!p) return null;
  let body;
  try {
    body = JSON.parse(DriveApp.getFileById(p.fileId).getBlob().getDataAsString());
  } catch (e) {
    throw new Error('データ本体の保存が途中で止まっていますが、その控え（' + (p.label || '') + '）が読めません。管理者に連絡してください。');
  }
  const t0 = new Date().getTime();
  const written = appJournalApply_(body.ops || []);
  appJournalFinish_(p.fileId);
  appRunLog_({ requestId: (ctx && ctx.requestId) || '', kind: 'JOURNAL.RECOVER', status: 'OK', durationMs: new Date().getTime() - t0,
    detail: { journalId: p.id, label: p.label, planId: p.planId, startedAt: p.at, written: written } });
  return { id: p.id, label: p.label, planId: p.planId, written: written };
}

function appJournalFinish_(fileId) {
  appProps_().deleteProperty(APP_JOURNAL_PROP);
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { Logger.log('保存の控えを消せません: ' + (e && e.message ? e.message : e)); }
}

/** ops を順に書く（何度書いても同じ結果になる） */
function appJournalApply_(ops) {
  const written = {};
  ops.forEach(op => {
    const prev = written[op.table];
    let res;
    if (op.mode === 'replace') res = Object.assign({ replaced: true }, appReplaceWhole_(op.table, op.rows), op.removed !== undefined ? { removed: op.removed } : {});
    else if (op.mode === 'replacePlan') res = Object.assign({ replaced: true }, appReplaceWhole_(op.table, appPlanRowsSwapped_(op)), op.removed !== undefined ? { removed: op.removed } : {});
    else if (op.mode === 'ensure') res = { appended: appEnsureRows_(op.table, op.rows).length };
    else if (op.mode === 'patch') res = { patched: appPatchRow_(op.table, op.key, op.patch, op.actor) ? 1 : 0 };
    else throw new Error('控えの書き方が不明です: ' + op.mode);
    written[op.table] = prev ? [].concat(prev, res) : res;
  });
  return written;
}

// ---- 控えを作る側の道具 ----

/** 計画 planId の行（sheets を渡せばそのシートの行だけ）を、objs に入れ替える op */
function appOpReplacePlan_(name, planId, sheets, objs) {
  const hit = appPlanRowMatcher_(planId, sheets);
  return { table: name, mode: 'replacePlan', planId: planId, sheets: sheets || null, rows: objs, removed: appReadTable_(name).filter(hit).length };
}

function appPlanRowMatcher_(planId, sheets) {
  return r => r.plan_id === planId && (!sheets || sheets.indexOf(r.sheet) >= 0);
}

/**
 * replacePlan の書き終えた後の表の全行。入れ替える行は、元の行があった場所に置く（シートごと。元が無ければ後ろに足す）。
 * 前後の行の位置が変わらないので、書き直す行が少なくて済む（appReplaceWhole_ は同じ行を書かない）
 */
function appPlanRowsSwapped_(op) {
  const hit = appPlanRowMatcher_(op.planId, op.sheets);
  const group = r => (r.sheet === undefined ? '' : String(r.sheet));
  const fresh = {};
  op.rows.forEach(r => { (fresh[group(r)] = fresh[group(r)] || []).push(r); });
  const def = APP_TABLES[op.table];
  const keyOf = o => def.key.map(k => String(o[k] === undefined || o[k] === null ? '' : o[k])).join('\u0001');
  const seen = {};
  const out = [];
  appReadTable_(op.table).forEach(r => {
    if (!hit(r)) {
      // 前の書き込みが途中で止まると、ほかの計画の行が 2 度ある（上へ詰めた後、余りを消す前に止まった）。最初の 1 つだけ残す
      const k = keyOf(r);
      if (!seen[k]) { seen[k] = true; out.push(appStripRow_(r)); }
      return;
    }
    const g = group(r);
    if (fresh[g]) { fresh[g].forEach(x => out.push(x)); delete fresh[g]; }
  });
  op.rows.forEach(r => { const g = group(r); if (fresh[g]) { fresh[g].forEach(x => out.push(x)); delete fresh[g]; } });
  return out;
}

/** 計画 1 つ分の行を入れ替える op（appReplaceOrAppend_ と同じ判断: 足されただけなら足した行だけ） */
function appOpReplaceOrAppend_(name, planId, objs) {
  const def = APP_TABLES[name];
  const key = o => def.columns.map(c => String(o[c] === undefined || o[c] === null ? '' : o[c])).join('\u0001');
  const all = appReadTable_(name);
  const cur = all.filter(r => r.plan_id === planId).map(appStripRow_);
  const next = {};
  objs.forEach(o => { next[o.seq] = o; });
  if (cur.every(o => next[o.seq] && key(next[o.seq]) === key(o))) {
    const have = {};
    cur.forEach(o => { have[o.seq] = true; });
    return { table: name, mode: 'ensure', rows: objs.filter(o => !have[o.seq]) };
  }
  return appOpReplacePlan_(name, planId, null, objs);
}
