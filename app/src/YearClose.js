/**
 * YearClose.js — 年度の締め（決まった年度を、元の行はそのままに読み取り専用にする）。
 *
 * - 年度の控え（YearSnapshot.js が選んで並べた全行のままの文字）は、アーカイブのフォルダに 1 ファイルとして残す（消さない）。
 * - 控えを読み戻して元と完全に一致した年度だけ、YEAR_CLOSURES に 1 行を足して凍結する。元の計画・表の行は 1 つも動かさない・消さない。
 * - 凍結した年度は、書き込み（保存・実行・公式版・計画の作成・学習の書き込み・保存の控えの続き）を止める。読むのはこれまでどおり。元に戻す操作は無い。
 * - 控えの作成・登録の途中で失敗したら、控えのファイルはそのまま残し（ゴミ箱に送らない）、凍結の記録は付けない（もう一度やり直せる）。
 */

const APP_YEAR_FROZEN_MESSAGE = 'この年度は締め済みです（FY{fy}）。元の行は残したまま読み取り専用です。';
function appYearFrozenErr_(fy) { return new Error(String(APP_YEAR_FROZEN_MESSAGE).replace('{fy}', fy)); }

/**
 * 年度 fy の締めの記録（無ければ null）。行の選び方は YearSnapshot.js の appYearStoredClosure_ と同じ（年度 1 行・重複は止まる）。
 * 表・見出し・保存場所の読み取りに失敗したらそのまま投げる（fail-closed。読めないのに書き込みを通さない）。
 * 記録があれば、中身が壊れていても「その年度は凍結されている」と扱う（壊れた記録を信じて開けることはしない）
 */
function appYearClosureOf_(fy) {
  return appYearStoredClosure_(fy);
}

/** 記録が「検証を終えた正常な凍結」か（壊れていれば false。壊れていても凍結は維持する） */
function appYearClosureOk_(rec) {
  return !!rec && rec.state === 'CLOSED' && !!rec.file_id && /^[a-f0-9]{64}$/.test(String(rec.snapshot_sha256 || '')) &&
    Number(rec.plan_count) >= 1 && Number(rec.row_count) >= 1 && Number(rec.bytes) >= 1;
}

/** 年度は凍結されているか（記録があれば壊れていても true） */
function appYearIsFrozen_(fy) {
  return !!appYearClosureOf_(fy);
}

/** 凍結している年度なら止める（書き込みの入口で必ず呼ぶ） */
function appRequireOpenYear_(fy) {
  if (appYearIsFrozen_(fy)) throw appYearFrozenErr_(fy);
}

/** 計画の年度が凍結していなければ返す（計画が無いときは計画のエラーが先） */
function appRequireOpenPlan_(planId) {
  const plan = appPlanOf_(planId);
  appRequireOpenYear_(plan.fy);
  return plan;
}

/** 計画の年度（データ本体か、控えの中の PLANS の行から）。どちらにも無ければ止める（fail-closed。年度が分からないものは書かない） */
function appYearPlanFyOf_(planId, ops) {
  const stored = appYearStoredPlans_().filter(x => x.plan_id === planId)[0];
  if (stored) return Number(stored.fy);
  const row = (ops || []).filter(op => op.table === 'PLANS').map(op => op.rows || [])
    .reduce((a, r) => a.concat(r), []).filter(r => r.plan_id === planId)[0];
  if (row && row.fy !== undefined && row.fy !== '') return Number(row.fy);
  throw new Error('計画（' + planId + '）の年度を確かめられません。書き込みを止めました。');
}

/** 書き込みの控え（ops）が触る計画・年度が凍結していないか。1 つでも触れば止める（fail-closed） */
function appRequireOpenJournalOps_(ops) {
  const check = (planId) => { appRequireOpenYear_(appYearPlanFyOf_(planId, ops)); };
  (ops || []).forEach(op => {
    // 表全体の入れ替えは plan_id の行をまるごと書き直すので、締めた年度の計画が 1 つでもあれば止める
    // （控えに入れた行だけ書き直すと、入れていない締めた年度の行が消えるため）
    if (op && op.mode === 'replace' && APP_TABLES[op.table] && APP_TABLES[op.table].columns.indexOf('plan_id') >= 0 &&
        appYearStoredPlans_().some(p => appYearIsFrozen_(p.fy))) {
      throw new Error('締め済みの年度の計画があります（読み取り専用）。表全体の入れ替えはできません。');
    }
    if (op && op.planId) check(op.planId);
    (op && op.rows || []).forEach(r => {
      if (r && r.plan_id !== undefined && r.plan_id !== '') check(r.plan_id);
    });
    if (op && op.table === 'PLANS') (op.rows || []).forEach(r => {
      if (r && r.fy !== undefined && r.fy !== '') appRequireOpenYear_(Number(r.fy));
    });
    if (op && op.mode === 'patch' && op.key && op.key.plan_id) check(op.key.plan_id);
    // 計画の年度を締めた年度へ書き換えることもできない（open の計画を closed の年度へ動かさない）
    if (op && op.table === 'PLANS' && op.mode === 'patch' && op.patch && op.patch.fy !== undefined && op.patch.fy !== '') {
      appRequireOpenYear_(Number(op.patch.fy));
    }
  });
}

/** 年度の一覧（今の年度・計画のある年度の計画数と状態。管理画面の「状態」に出す） */
function appYearStatus_() {
  const currentFy = appFy_(new Date());
  const count = {};
  appYearStoredPlans_().forEach(p => { const f = String(p.fy); count[f] = (count[f] || 0) + 1; });
  const states = {};
  appYearStoredClosures_().forEach(r => { states[String(r.fy)] = appYearClosureOk_(r) ? 'CLOSED' : 'ERROR'; });
  const fys = Object.keys(Object.assign({}, count, states)).sort();
  return { currentFy: currentFy,
    years: fys.map(f => ({ fy: Number(f), planCount: count[f] || 0, state: states[f] || 'OPEN' })) };
}

/**
 * 締められる年度か（ロックの中で呼ぶ）。過去の年度だけ・計画がある・書きかけの保存なし・ほかの処理なし・
 * 承認待ちの版なし・全部の計画に公式版（APPROVED）がある、の全部を満たす。
 * すでに正常な凍結があるときは投げずに { closed } を返す（冪等）。壊れた記録は投げる（凍結のまま）。
 */
function appYearCheck_(fy, ownJobId) {
  const year = Number(fy);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) throw new Error('年度を 4 桁の数で選んでください。');
  if (year >= appFy_(new Date())) throw new Error('今年度とこれからの年度は締められません。');
  const rec = appYearClosureOf_(year);
  if (rec) {
    if (appYearClosureOk_(rec)) return { fy: year, closed: true, record: rec };
    throw new Error('FY' + year + ' の締めの記録が壊れています。凍結のままです（元の行は動いていません）。管理者に連絡してください。');
  }
  const plans = appYearStoredPlans_(year);
  if (!plans.length) throw new Error('この年度の計画はありません。');
  if (appJournalPending_()) throw new Error('データ本体の保存が途中で止まっています。予測の画面の「保存の続きを書く」を先に行ってください。');
  const busy = appJobList_().filter(j => j.id !== ownJobId && (j.status === 'QUEUED' || (j.status === 'RUNNING' && !appJobIsStale_(j))));
  if (busy.length) throw new Error('ほかの処理（' + appJobSpec_(busy[0].kind).label + '）が終わるまでお待ちください。');
  const planIds = plans.map(x => x.plan_id);
  const versions = appYearStoredVersions_(planIds);
  if (versions.some(v => v.state === 'SUBMITTED')) throw new Error('承認待ちの公式版があります。承認か却下を済ませてください。');
  // 公式版が要るのは予算を立てる計画だけ（測る専用の計画は外す。版 10 の 3-9。V10.js）
  if (appYearPlansNeedingVersion_(plans).some(x => !versions.some(v => v.plan_id === x.plan_id && v.state === 'APPROVED'))) {
    throw new Error('公式版のない計画があります。年度を締めるには、全部の計画に公式版が要ります。');
  }
  // 年度の最後の四半期の当たりを数え終えてから締める（締めた後は書けないため。版 10 の 9 章 10。V10.js）
  const hitsPending = appYearHitsPending_(year, plans);
  if (hitsPending) throw new Error(hitsPending);
  return { fy: year, closed: false, planIds: planIds, plans: plans };
}

/**
 * 年度を締める前の確認（内容を確認）: 締められるなら控えの指紋（SHA-256）を返し、画面の「締める」がその指紋を渡す。
 * 控えのファイルはまだ作らない・何も書かない。締められなければ canClose:false と理由だけ（投げない）。
 */
function appYearPreview_(ctx, input) {
  const fy = Number(input && input.fy);
  return appWithLock_(() => {
    let check;
    try { check = appYearCheck_(fy, ''); }
    catch (e) { return { fy: Number.isInteger(fy) ? fy : (input && input.fy), canClose: false, reason: String(e && e.message ? e.message : e) }; }
    if (check.closed) return { fy: fy, canClose: false, reason: appYearFrozenErr_(fy).message, closed: true };
    try {
      const s = appYearSnapshot_(fy);
      return { fy: fy, canClose: true, inputHash: s.sha256, planCount: s.planCount, rowCount: s.rowCount, bytes: s.bytes, tables: s.tables };
    } catch (e) {
      return { fy: fy, canClose: false, reason: String(e && e.message ? e.message : e) };
    }
  });
}

/**
 * 年度を締める（裏の処理 YEAR.CLOSE から。ロックの中）。
 * 確認した控えの指紋（p.inputHash）と今の控えが一致しなければ止める（画面を開いた後に年度が変わった）。
 * 控えをアーカイブへ置き、読み戻して一致・元データが変わっていないことをもう一度確かめてから、YEAR_CLOSURES に 1 行足す。
 * 途中で失敗しても元の行は変わらず、置いた控えのファイルは残す（もう一度締めればまた検証する。同じ年度を 2 度は書かない）。
 */
function appYearClose_(ctx, p, job) {
  return appWithLock_(() => {
    const rec0 = appYearClosureOf_(p.fy);
    if (rec0) {
      if (!appYearClosureOk_(rec0)) throw new Error('FY' + p.fy + ' の締めの記録が壊れています。凍結のままです（元の行は動いていません）。管理者に連絡してください。');
      return { fy: Number(rec0.fy), state: 'CLOSED', planCount: Number(rec0.plan_count), rowCount: Number(rec0.row_count), bytes: Number(rec0.bytes),
        snapshotSha256: String(rec0.snapshot_sha256), fileId: String(rec0.file_id), already: true, audit: { entityId: 'FY' + rec0.fy } };
    }
    const check = appYearCheck_(p.fy, job && job.id);
    const snapshot = appYearSnapshot_(check.fy);
    if (!/^[a-f0-9]{64}$/.test(String(p.inputHash || '')) || snapshot.sha256 !== String(p.inputHash)) {
      throw new Error('確認したあとに年度の内容が変わりました（または確認していません）。「内容を確認」からやり直してください。');
    }
    const folderId = String(appProps_().getProperty(APP_PROP.archiveFolderId) || '');
    if (!folderId) throw new Error('アーカイブのフォルダが決まっていません。管理者が「初期設定」を実行してください。');
    const folder = DriveApp.getFolderById(folderId);   // 開けなければここで止まる（控えはまだ作っていない）
    const file = folder.createFile(APP_FILES.yearPrefix + check.fy + ' ' + appId_('YC') + '.json', snapshot.text, MimeType.PLAIN_TEXT);
    // 置いた控えを読み戻して確かめる。失敗したら控えは残し、凍結にはしない（元の行は変わっていない）
    appYearVerifySnapshot_(snapshot, file);
    const parents = file.getParents();
    let inArchive = false;
    while (parents.hasNext()) { if (parents.next().getId() === folderId) { inArchive = true; break; } }
    if (!inArchive) throw new Error('年度の控えがアーカイブのフォルダに置けませんでした。凍結しません（置いたファイルは残します）。');
    // 書き込みの隙間で元が変わっていないか、もう一度同じ選び方で確かめる（変わっていたら凍結しない）
    appStoreForget_();
    const again = appYearSnapshot_(check.fy);
    if (again.sha256 !== snapshot.sha256) {
      throw new Error('年度の内容が控えを作っている間に変わりました。凍結しません（置いた控えは残します）。もう一度「内容を確認」からやり直してください。');
    }
    appInsertRows_('YEAR_CLOSURES', [{ fy: String(check.fy), state: 'CLOSED', file_id: file.getId(), snapshot_sha256: snapshot.sha256,
      plan_count: snapshot.planCount, row_count: snapshot.rowCount, bytes: snapshot.bytes, closed_at: appNowIso_(), closed_by: ctx.actor }]);
    return { fy: check.fy, state: 'CLOSED', planCount: snapshot.planCount, rowCount: snapshot.rowCount, bytes: snapshot.bytes,
      snapshotSha256: snapshot.sha256, fileId: file.getId(), already: false, audit: { entityId: 'FY' + check.fy } };
  });
}
