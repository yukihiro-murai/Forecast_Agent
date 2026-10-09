/**
 * InputLog.js — 入力の記録（INPUT_LOG。SCHEMA_PLAN_v10-12_JA.md の 3-1。2026-10-08 村井さん承認「版10はおすすめで」）。
 * - 入力の 4 つの表（ENG_PRODUCT・ENG_CLIENT・ENG_OPINIONS・ENG_DEV_SPOT）が変わる保存と実行のすべて（入力の保存・売上の取り込み・計画を作る ほか）で、
 *   変わった行だけを前と後で残す（Forecast.js の appChangedOps_ から。本体の保存と同じ控えに入れて一緒に書く）。前の行は書く前のデータ本体から読む。
 * - 行の印 row_key = 種類|担当者|製品|月（スポットは案件名）。同じ印が 2 行あれば 2 つ目から #2・#3（上からの順）。
 *   欄の値の「%」「|」「#」は %25・%7C・%23 にして書く（案件名「案件#2」と 2 つ目の「案件」を同じ印にしない。appInputLogKey_）。
 *   入力の保存では、画面が保存ごとに fromRows: true を、行ごとに「開いたときの何行目から来たか」（fromRow）を渡す。そのときは前と後をその印で組む
 *   （外した行・足した行・変えた行が、位置がずれても取り違えない。fromRow の無い行は足した行、どの行の出どころにもならなかった行は外した行。
 *   開いたときの行を全部外して足したときも同じ。fromRows の無い前の画面は、どれかの行に fromRow があるときだけ）。
 *   fromRows も fromRow も無い保存（前の画面）とほかの操作は、同じ印どうしで比べ、印の無くなった行と新しい印の行は、
 *   同じ位置（行の番号）なら「変えた」とみなす（担当者を入れた・月を変えた）。
 * - 同じ印の組（#2・#3…）の行を外すと、下の行の #n が詰まる。出どころで組んだ行の印が #n だけ変わったら、その行の自信と跡は前の印から引き継ぐ:
 *   新しい印に CHANGE（中身は同じ。before_json の _key に前の印。同じ保存で中身も変えた行は、その CHANGE の before_json に _key）を足し、
 *   画面の跡はそこから前の印の跡（その保存より前）をたどる（ほかの行の跡を付けない）。出どころで組んで担当者・製品・月を変えた行も、
 *   CHANGE の before_json の _key に前の印を入れ、跡はその行の前の印の跡をたどる（前にその印を使ったほかの行・外した行の跡を付けない）。
 *   _key の無い CHANGE（同じ位置で組んだ保存）で担当者・製品・月が変わったものは、跡をそこで止める（それより前は、その印にいたほかの行のもの）。
 *   新しい印のもとの行と、中身も跡も同じ行（見分けられない行。例: 手を付けていないひな形）なら足さない（見え方が同じなので）。
 *   付け直しだけの行は画面の跡に出さない（中身は変わっていない。自信はその行に付けたまま）。足した行（ADD）が跡の始まり（前に同じ印だった行の跡を付けない）。
 * - 自信（self_conf: 高い・ふつう・低い・空）は製品とメーカー全体の行だけ。入力の画面で選び、この記録にだけ書く（旧来の表には渡さない。計算に使わない）。
 *   行の中身が同じで自信だけ変えたら CONF。自信を送らない行（画面は、選び直した行だけ送る）とほかの操作で変えた・外した行は、
 *   その行の一番新しい自信を引き継ぐ（足した行は空）。比べるのは、その行の前の印の自信（#n が詰まった行でも、その行自身の自信）。
 * - 見せる範囲（appInputLogView_）: 予算策定担当以上（そのメーカーの担当を含む）は全部。ほかの人は、人のつなぎ（PERSON_LINKS）で本人と分かる行だけで、
 *   その中でも担当者の名前と保存した人は送らない。本人の行かは、保存した日に効くつなぎで決める（今日のつなぎで前の行を決めない）。
 *   担当者が変わった行（CHANGE）の前の中身は、前の担当者も本人のときだけ送る。跡は今の担当者の間の分だけ送る（担当者を変えた行の跡は
 *   前の担当者の記録までたどるが、前の担当者だった本人に、今はほかの人の行の跡と自信を出さない）。ほかの行は種類ごとの件数だけ（V10.js の appLogViewer_・appLogVisible_）。
 * - 版 10 の移行の後に 1 回、今の 4 つの表の行を BASELINE として写す（appInputLogBackfill_。前の保存の履歴は残っていないので作れない）。
 *   した人は仕組み（APP_V10_BACKFILL_ACTOR。画面には人として出さない）。行を確かめられない計画は飛ばし、ほかの計画を写す（書きかけの控えを残さない）。
 * 行は消さず、書き換えない。記録は計算に使わない。
 */

/** 入力の 4 つの表（旧来のシート → 記録の種類） */
const APP_INPUT_LOG_SHEETS = { PRODUCT: 'product', CLIENT: 'client', OPINIONS: 'opinions', DEV_SPOT: 'devspot' };
/** 種類ごとの行の項目（旧来の見出しの順。画面の入力の行と同じ名前。before_json・after_json の中身） */
const APP_INPUT_LOG_FIELDS = {
  product: ['person', 'product', 'ym', 'step', 'reason'],
  client: ['person', 'ym', 'step', 'reason'],
  opinions: ['person', 'ym', 'step', 'conf', 'note'],
  devspot: ['person', 'ym', 'project', 'amount', 'conf']
};
/** 行の印に使う項目（担当者・製品・月・スポットの名前） */
const APP_INPUT_LOG_KEYS = { product: ['person', 'product', 'ym'], client: ['person', 'ym'], opinions: ['person', 'ym'], devspot: ['person', 'ym', 'project'] };
/** 自信の段階（9 章 7。計算には使わない）と、選べる種類（製品とメーカー全体の行だけ） */
const APP_INPUT_CONF = ['高い', 'ふつう', '低い'];
const APP_INPUT_CONF_KINDS = ['product', 'client'];
/** 保存の理由（任意）の上限 */
const APP_INPUT_REASON_MAX = 200;
/** 画面に送る、1 行あたりの変わった跡の数（新しい順） */
const APP_INPUT_LOG_HIST = 5;
/** 一度だけの写しの 1 回の目安（それを過ぎたら続きは次の操作で） */
const APP_INPUT_LOG_BACKFILL_MS = 20000;
/** 操作を渡さない道（控えの番号の頭 → 操作）: 計画を作る（Portfolio.js）・予測（Forecast.js）・学習の事前分布（Learning.js） */
const APP_INPUT_LOG_BATCH = { NEW: 'PLAN.CREATE', RUN: 'FORECAST.RUN', POOL: 'LEARN.POOL' };

// ---- 行の形と印 ----

/** 種類の旧来のシートの名前 */
function appInputLogSheet_(kind) {
  return Object.keys(APP_INPUT_LOG_SHEETS).filter(s => APP_INPUT_LOG_SHEETS[s] === kind)[0];
}

/**
 * ENG_ の表の 1 行（型つきの文字列）を、記録の行の形にする。値が 1 つも無い行は null（旧来の画面の読み取り webParseInputs_ と同じく数えない）。
 * 日付は月（yyyy-MM。月の欄のほかは yyyy-MM-dd）にする。数は数のまま
 */
function appInputLogRow_(kind, o) {
  const header = APP_ENGINE_SHEETS[appInputLogSheet_(kind)].header;
  const types = String(o._types || '');
  const out = {};
  let any = false;
  APP_INPUT_LOG_FIELDS[kind].forEach((f, j) => {
    const t = types.charAt(j) || 'e';
    const raw = o[header[j]] === undefined || o[header[j]] === null ? '' : o[header[j]];
    let v = t === 'f' ? '=' + String(raw).replace(/^=/, '') : appCellDecode_(t, raw);
    if (appIsDate_(v)) v = Utilities.formatDate(v, APP_TZ, f === 'ym' ? 'yyyy-MM' : 'yyyy-MM-dd');
    if (v !== '' && v !== null) any = true;
    out[f] = v;
  });
  return any ? out : null;
}

/**
 * 行の印（同じ印の 2 つ目からの #n は appInputLogList_ が付ける。ここで返すのは #n の無い印）。
 * 欄の値の「%」「|」「#」は %25・%7C・%23 にする（appInputLogKeyPart_）。区切りの | と 2 つ目からの #n を、欄の中の文字と取り違えない
 * （例: 案件名「案件#2」の印は …|案件%232 で、2 つ目の「案件」の …|案件#2 と別。ほかの文字はそのままなので、ふつうの名前の印は読めるまま）
 */
function appInputLogKey_(kind, row) {
  return kind + '|' + APP_INPUT_LOG_KEYS[kind].map(f => appInputLogKeyPart_(row[f])).join('|');
}

/** 行の印の 1 つの欄（前後の空白を除き、% | # だけを %25 %7C %23 にする。% も置き換えるので、もとの値と 1 対 1） */
function appInputLogKeyPart_(v) {
  return String(v === null || v === undefined ? '' : v).trim().replace(/[%|#]/g, c => (c === '%' ? '%25' : c === '|' ? '%7C' : '%23'));
}

/** 印から 2 つ目からの #n を除いた印（欄の中の # は %23 にしてあるので、末尾の「#数」だけが組の番号） */
function appInputLogBaseKey_(key) {
  return String(key === null || key === undefined ? '' : key).replace(/#\d+$/, '');
}

/** ENG_ の表の行（計画 1 つ・1 種類）を、行の番号の順に、印つきの記録の形にする（空の行は除く）。返り値 [{ seq, row, key }] */
function appInputLogList_(kind, objs) {
  const seen = {};
  return (objs || []).slice().sort((a, b) => Number(a.seq) - Number(b.seq)).map(o => ({ seq: Number(o.seq), row: appInputLogRow_(kind, o) }))
    .filter(x => x.row).map(x => {
      const k = appInputLogKey_(kind, x.row);
      seen[k] = (seen[k] || 0) + 1;
      return { seq: x.seq, row: x.row, key: seen[k] > 1 ? k + '#' + seen[k] : k };
    });
}

/**
 * 前と後の行を比べる。返り値 [{ change: 'ADD'|'CHANGE'|'REMOVE'|''（同じ）, a: 後の行, b: 前の行 }]。
 * from（後の行の番号 → 前の行の位置（0 から。画面が開いたときの行の何行目か））を渡したときは、その印で組む（appInputLogDiffFrom_。
 * classOf = 見分けない前の行の組の印）。渡さないときは同じ印どうしで比べる。印の無くなった前の行と新しい印の後の行は、
 * 同じ位置（行の番号）なら CHANGE（担当者を入れた・月を変えた）、ほかは REMOVE と ADD
 */
function appInputLogDiff_(before, after, from, classOf) {
  if (from) return appInputLogDiffFrom_(before, after, from, classOf);
  const bk = {}, ak = {};
  before.forEach(b => { bk[b.key] = b; });
  after.forEach(a => { ak[a.key] = a; });
  const gone = {};
  before.forEach(b => { if (!ak[b.key]) gone[b.seq] = b; });
  const paired = {};
  const out = [];
  after.forEach(a => {
    const b = bk[a.key];
    if (b) { out.push({ change: JSON.stringify(a.row) === JSON.stringify(b.row) ? '' : 'CHANGE', a: a, b: b }); return; }
    const p = gone[a.seq];
    if (p && !paired[p.key]) { paired[p.key] = true; out.push({ change: 'CHANGE', a: a, b: p }); return; }
    out.push({ change: 'ADD', a: a, b: null });
  });
  before.forEach(b => { if (!ak[b.key] && !paired[b.key]) out.push({ change: 'REMOVE', a: null, b: b }); });
  return out;
}

/**
 * 画面が渡した行の出どころで前と後を組む。from[後の行の番号] = 前の行の位置（before の何番目か。画面の開いたときの行と同じ並び）。
 * 出どころのある後の行は、その前の行と比べる（同じ中身なら ''・違えば CHANGE）。
 * 出どころの無い後の行・前の行の数を超える位置・2 つ目からの同じ位置は ADD。どの後の行の出どころにもならなかった前の行は REMOVE（外した行）。
 * 見分けられない行（classOf が同じ: 中身も、その印の跡も同じ。省くと中身だけ）は入れ替えても見え方が同じなので、組み直して印の変わる行を減らす:
 * 同じ組の中で、後の行には同じ印だった前の行を先に当て、残りの前の行を外したことにする
 * （例: 手を付けていない同じ中身の 2 行の 1 行目を外すと、残った行は 1 行目の印のまま、外したのは 2 行目の印。前の画面の組み方と同じ）。
 * 中身か跡の違う行は組み直さない（その行の自信と跡は、印が #n だけ変わっても、その行についていく。appInputLogOps_）
 */
function appInputLogDiffFrom_(before, after, from, classOf) {
  const cls = classOf || (b => JSON.stringify(b.row));
  const used = {};
  const out = [];
  after.forEach(a => {
    const i = Object.prototype.hasOwnProperty.call(from, a.seq) ? from[a.seq] : null;
    const b = i !== null && i >= 0 && i < before.length && !used[i] ? before[i] : null;
    if (!b) { out.push({ change: 'ADD', a: a, b: null }); return; }
    used[i] = true;
    out.push({ change: JSON.stringify(a.row) === JSON.stringify(b.row) ? '' : 'CHANGE', a: a, b: b });
  });
  const gone = before.filter((b, i) => !used[i]);
  if (!gone.length && out.every(d => d.change !== '' || d.a.key === d.b.key)) return out;   // 外した行も印の変わった行も無い（組み直すものが無い）
  // 見分けられない行の組み直し（中身の同じ組 = 同じ中身の後の行と前の行なので、組み直しても '' のまま）
  const groups = {};
  out.forEach(d => {
    if (d.change !== '') return;
    const k = cls(d.b);
    const g = groups[k] = groups[k] || { pairs: [], pool: [] };
    g.pairs.push(d);
    g.pool.push(d.b);
  });
  gone.forEach(b => { const g = groups[cls(b)]; if (g) g.pool.push(b); });
  const removed = gone.filter(b => !groups[cls(b)]);
  Object.keys(groups).forEach(k => {
    const g = groups[k];
    const avail = g.pool.slice();
    const take = b => { avail.splice(avail.indexOf(b), 1); return b; };
    const left = [];
    g.pairs.forEach(d => {
      const same = avail.filter(b => b.key === d.a.key)[0];   // 同じ印だった前の行を先に
      if (same) d.b = take(same); else left.push(d);
    });
    const wait = [];
    left.forEach(d => { if (avail.indexOf(d.b) >= 0) take(d.b); else wait.push(d); });   // 次に、もとの組
    wait.forEach(d => { d.b = take(avail.slice().sort((x, y) => before.indexOf(x) - before.indexOf(y))[0]); });
    avail.forEach(b => removed.push(b));
  });
  before.forEach(b => { if (removed.indexOf(b) >= 0) out.push({ change: 'REMOVE', a: null, b: b }); });
  return out;
}

/** CHANGE の記録が指す、その行の前の印（before_json の _key。#n の付け直しと、出どころで組んで印を変えた行。つなぎが無ければ ''） */
function appInputLogMovedFrom_(r) {
  if (!r || r.change !== 'CHANGE') return '';
  const b = appInputLogJson_(r.before_json);
  const k = b && typeof b === 'object' && typeof b._key === 'string' ? b._key : '';
  return k && k !== r.row_key ? k : '';
}

/** 記録の JSON の列から、記録だけの印（_key）を除いた中身（画面と比べ方に使う） */
function appInputLogContent_(v) {
  const x = appInputLogJson_(v);
  if (!x || typeof x !== 'object') return x;
  const c = Object.assign({}, x);
  delete c._key;
  return c;
}

/**
 * 前の印へのつなぎ（_key）の無い CHANGE で、前の中身の印が、この記録の印（#n を除く）と違うか。
 * 同じ位置で組んだ保存（前の画面・ほかの操作）で、担当者・製品・月が変わった行。その前のこの印の記録は、前にこの印だったほかの行のもの
 */
function appInputLogLeftKey_(r) {
  if (!r || r.change !== 'CHANGE' || !APP_INPUT_LOG_KEYS[r.kind] || appInputLogMovedFrom_(r)) return false;
  const b = appInputLogContent_(r.before_json);
  if (!b || typeof b !== 'object') return false;
  return appInputLogKey_(r.kind, b) !== appInputLogBaseKey_(r.row_key);
}

/** 印の付け直しだけの記録（#n が詰まった。中身は同じ）。画面の跡には出さない（自信と跡を引き継ぐための行） */
function appInputLogRenumbered_(r) {
  return !!appInputLogMovedFrom_(r) && JSON.stringify(appInputLogContent_(r.before_json)) === JSON.stringify(appInputLogContent_(r.after_json));
}

/** 跡の 1 行の比べ方（同じ跡か。記録の番号と付け直しの前の印は入れない） */
function appInputLogRecSig_(r) {
  return [r.change, r.action, r.action_id, r.saved_at, r.actor_email, String(r.self_conf || ''), String(r.reason || ''),
    appInputLogContent_(r.before_json), appInputLogContent_(r.after_json)];
}

/**
 * 記録（appInputLogRead_ の並び）から、行の跡をたどる道具 trail(印) を作る。返り値の trail は、今その印にある行の跡（新しい順。同じ保存の中は
 * appInputLogHistOrder_ の順）。足した行（ADD）で止める（それより前の同じ印の記録は、前にその印だった別の行のもの。初めの姿 BASELINE は
 * 記録の無い印にだけ写すので、いつも一番古い）。
 * 前の印へのつなぎ（before_json の _key。appInputLogMovedFrom_: #n だけ詰まった行と、出どころで組んで担当者・製品・月を変えた行）に来たら、
 * 前の印の跡のうち、その保存より前のものへ続ける（その行自身の前の跡）。つなぎの無い CHANGE で担当者・製品・月が変わったもの
 * （appInputLogLeftKey_。同じ位置で組んだ前の画面・ほかの操作）でも止める（それより前は、前にその印だったほかの行の跡）
 */
function appInputLogTrails_(log) {
  const at = new Map(), byKey = {};
  (log || []).forEach((r, i) => { at.set(r, i); (byKey[r.row_key] = byKey[r.row_key] || []).push(r); });
  const saveOf = r => String(r.action_id || '') + '\u0001' + String(r.saved_at || '');
  const memo = {};
  const trail = (key, bound) => {
    if (!bound && memo[key]) return memo[key];
    const list = appInputLogHistOrder_((byKey[key] || []).filter(r => !bound || (at.get(r) < bound.i && saveOf(r) !== bound.save)).reverse());
    const out = [];
    for (let j = 0; j < list.length; j++) {
      const r = list[j];
      out.push(r);
      if (r.change === 'ADD') break;
      const from = appInputLogMovedFrom_(r);
      if (from) { Array.prototype.push.apply(out, trail(from, { i: at.get(r), save: saveOf(r) })); break; }
      if (appInputLogLeftKey_(r)) break;
    }
    if (!bound) memo[key] = out;
    return out;
  };
  return trail;
}

/** 画面が渡した行の出どころ（fromRow。開いたときの行の位置。0 から）。整数でない・負の値は無いことにする（null） */
function appInputLogFromRow_(v) {
  if (v === undefined || v === null || v === '' || typeof v === 'boolean') return null;
  const n = typeof v === 'number' ? v : Number(v);
  return Number.isInteger(n) && n >= 0 ? n : null;
}

// ---- 書く（本体の保存と同じ控えに入れる）----

/**
 * 入力の保存（INPUT.SAVE）の中身から、自信・行の出どころ・保存の理由を取り出す（旧来の webSaveInputs には渡さない。旧来の表に列を足さない）。
 * info を埋めて、旧来に渡す中身（写し）を返す: info.reason = 保存の理由、info.conf = { kind, bySeq: { 書く行の番号: 自信（送らない行は null） } }、
 * info.from = { kind, bySeq: { 書く行の番号: 開いたときの行の位置 } }（今の画面の保存 = 中身の頭に fromRows: true があるとき。
 * 行に fromRow が 1 つも無くても組む: 全部の行を外して足した保存は、足した行が ADD・開いたときの行が REMOVE。
 * fromRows の無い保存（前の画面）では、どれかの行が fromRow を渡したときだけ）。
 * 書く行の番号は旧来と同じ数え方（値の無い行を飛ばして 2 行目から詰める）。自信は製品とメーカー全体の行だけ。ほかの操作はそのまま返す
 */
function appInputLogTake_(args, info) {
  if (!info || info.action !== 'INPUT.SAVE' || !args || typeof args !== 'object' || Array.isArray(args)) return args;
  const out = Object.assign({}, args);
  const reason = out.reason === undefined || out.reason === null ? '' : String(out.reason).trim();
  if (reason.length > APP_INPUT_REASON_MAX) throw new Error('保存の理由は ' + APP_INPUT_REASON_MAX + ' 字までです。');
  delete out.reason;
  info.reason = reason;
  const fromRows = out.fromRows === true;   // 今の画面: 行の出どころで必ず組む（旧来には渡さない）
  delete out.fromRows;
  if (!Array.isArray(out.rows)) return out;
  const kind = String(out.kind || '');
  const useConf = APP_INPUT_CONF_KINDS.indexOf(kind) >= 0;
  const bySeq = {}, fromBySeq = {};
  let n = 0, anyFrom = false;
  out.rows = out.rows.map(r => {
    if (!r || typeof r !== 'object') return r;
    const x = Object.assign({}, r);
    const c = x.selfConf === undefined || x.selfConf === null ? null : String(x.selfConf).trim();   // 送らない（null）= 今の自信のまま。空 = 外した
    const f = appInputLogFromRow_(x.fromRow);   // 開いたときの何行目から来たか（足した行は無い）
    delete x.selfConf;
    delete x.fromRow;
    if (c && APP_INPUT_CONF.indexOf(c) < 0) throw new Error('自信は「高い」「ふつう」「低い」から選んでください。');
    // 旧来の webSaveInputs と同じ見分け方（どの値も空なら書かない）
    if (Object.keys(x).some(k => String(x[k] || '').trim() !== '')) {
      n++;
      if (useConf) bySeq[n + 1] = c;
      if (f !== null) { fromBySeq[n + 1] = f; anyFrom = true; }
    }
    return x;
  });
  if (useConf) info.conf = { kind: kind, bySeq: bySeq };
  if (anyFrom || fromRows) info.from = { kind: kind, bySeq: fromBySeq };
  return out;
}

/**
 * 入力の記録を足す控えの書き方（ops。appChangedOps_ が本体の ops に足す）。changed = 控えの形にしたシート（appCaptureChanged_ の changed）、
 * batchId = 操作の番号（action_id）、info = { action, reason, conf, from, hashChecked }（appInputLogTake_。無ければ番号の頭から操作を決める）。
 * 行の出どころ（from）は、画面が開いたときのデータ本体と同じだと確かめた保存（hashChecked = 入力のハッシュが一致）でだけ使う。
 * 入力の表が変わっていなくても、自信だけ変えた行は CONF として足す。前の行と、その行の跡（一番新しい自信）は、書く前のデータ本体から読む。
 * 出どころで組んだ行の印が #n だけ変わったら（同じ印の組の上の行を外した）、その行は前の印の行のまま:
 * - 自信は前の印の自信と比べる（送らなければ前の印の自信を引き継ぐ。ほかの行の自信と比べて CONF を足さない）
 * - 中身が同じなら、新しい印に CHANGE（before_json = 前の行と _key = 前の印）を足して、自信と跡を新しい印へ引き継ぐ。
 *   ただし新しい印のもとの行と中身も跡も同じ（見分けられない）なら足さない（見え方が同じ）。中身を変えた行は CHANGE の before_json に _key を入れる
 * 出どころで組んで担当者・製品・月を変えた行（印が変わった CHANGE）も、before_json の _key に前の印を入れる（跡はその行の前の印の跡をたどり、
 * 前にその印を使ったほかの行の跡を付けない）。同じ位置で組んだ CHANGE には入れない（どの行から来たか確かでないので、跡はそこで止まる）
 */
function appInputLogOps_(ctx, planId, changed, batchId, info) {
  const o = info || {};
  const enc = {};
  (changed || []).forEach(e => {
    const k = APP_INPUT_LOG_SHEETS[e.sheetRow.sheet];
    if (k && e.sheetRow.mode === 'table') enc[k] = e;   // 見出しが変わって行ごとに持つシートは比べない（旧来は見出しを変えない）
  });
  const confKind = o.conf && APP_INPUT_CONF_KINDS.indexOf(o.conf.kind) >= 0 ? o.conf.kind : '';
  const fromKind = o.from && o.hashChecked ? String(o.from.kind || '') : '';
  if (!Object.keys(enc).length && !confKind) return [];
  const action = o.action || APP_INPUT_LOG_BATCH[String(batchId || '').split('-')[0]] || '';
  const now = appNowIso_();
  let trails = null;
  const trailOf = key => { if (!trails) trails = appInputLogTrails_(appInputLogRead_(planId)); return trails(key); };
  const confOf = key => { const r = trailOf(key).filter(x => x.change !== 'REMOVE')[0]; return r ? String(r.self_conf || '') : ''; };   // その印の行の一番新しい自信
  const rows = [];
  Object.keys(APP_INPUT_LOG_FIELDS).forEach(kind => {
    const asked = kind === confKind ? o.conf.bySeq : null;   // 画面が渡した自信（書く行の番号 → 自信）
    if (!enc[kind] && !asked) return;
    const stored = appInputLogList_(kind, appReadPlanTable_('ENG_' + appInputLogSheet_(kind), planId));
    const next = enc[kind] ? appInputLogList_(kind, enc[kind].tableRows) : stored;
    const withConf = APP_INPUT_CONF_KINDS.indexOf(kind) >= 0;
    const fromMode = kind === fromKind;
    const held = {};
    stored.forEach(b => { held[b.key] = b; });
    const sig = {};
    // 見分けられない行の組の印: 中身と、画面に出る跡（付け直しだけの記録は見え方を変えないので入れない）
    const classOf = b => sig[b.key] || (sig[b.key] = JSON.stringify(b.row) + '\u0002' +
      JSON.stringify(trailOf(b.key).filter(r => !appInputLogRenumbered_(r)).map(appInputLogRecSig_)));
    const push = (change, x, before, after, conf) => rows.push({ plan_id: planId, log_id: appNewLogId_(APP_LOG_PREFIX.INPUT_LOG), action_id: String(batchId || ''),
      action: action, kind: kind, change: change, row_key: x.key, person: String(x.row.person === null || x.row.person === undefined ? '' : x.row.person),
      before_json: before, after_json: after, self_conf: conf, reason: o.reason || '', signal_id: '', actor_email: ctx.actor, saved_at: now });
    appInputLogDiff_(stored, next, fromMode ? o.from.bySeq : null, fromMode ? classOf : null).forEach(d => {
      const s = withConf && asked && d.a ? asked[d.a.seq] : null;
      const said = s === null || s === undefined ? null : s;   // 画面が選んだ自信（送らなければ null = 今の自信のまま）
      const own = () => (withConf && d.b ? confOf(d.b.key) : '');   // その行の今の自信（前の印で読む）
      const conf = () => (!withConf ? '' : said !== null ? said : own());   // 自信を送らない保存・ほかの操作: 変えた・外した行は今の自信を引き継ぐ（足した行は空）
      // 出どころで組んだ行の印が変わった（担当者・製品・月を変えた。同じ印の組の上の行を外した・足したときは #n だけ）
      const relinked = fromMode && d.a && d.b && d.a.key !== d.b.key;
      const moved = relinked && appInputLogKey_(kind, d.a.row) === appInputLogKey_(kind, d.b.row);   // #n だけ変わった
      const linked = () => Object.assign({}, d.b.row, { _key: d.b.key });   // 前の印へのつなぎ（跡はその行の前の印の跡をたどる）
      if (d.change === '') {
        if (moved && (!held[d.a.key] || classOf(held[d.a.key]) !== classOf(d.b))) push('CHANGE', d.a, linked(), d.a.row, own());   // 印の付け直し: 自信と跡を引き継ぐ
        if (said !== null && said !== own()) push('CONF', d.a, '', d.a.row, said);
        return;
      }
      if (d.change === 'CHANGE') push('CHANGE', d.a, relinked ? linked() : d.b.row, d.a.row, conf());
      else if (d.change === 'ADD') push('ADD', d.a, '', d.a.row, conf());
      else push('REMOVE', d.b, d.b.row, '', conf());
    });
  });
  return appLogOps_('INPUT_LOG', rows);
}

// ---- 読む ----

/**
 * 計画の入力の記録（古い順。同じ秒なら BASELINE が先、その次は表に足した順。表が読めなければ空）。
 * 足すのはロックの中で 1 つずつなので、表の順が保存の順（記録の番号の後ろは乱数なので、同じ秒の 2 回の保存の順には使わない）
 */
function appInputLogRead_(planId) {
  let rows;
  try { rows = appReadPlanTable_('INPUT_LOG', planId); } catch (e) { return []; }
  const k = r => [r.saved_at, r.change === 'BASELINE' ? '0' : '1'].join('\u0001');
  return rows.map((r, i) => ({ r: appStripRow_(r), i: i })).sort((a, b) => (k(a.r) < k(b.r) ? -1 : k(a.r) > k(b.r) ? 1 : a.i - b.i)).map(x => x.r);
}

/**
 * 行の跡（新しい順）のうち、今の担当者の間の分（先頭から、記録の担当者の欄が一番新しい記録と同じ間）。
 * 担当者を変えた CHANGE は変えた後の担当者の記録なので入り、その前の担当者の記録から先は入らない
 */
function appInputLogSamePerson_(path) {
  const who = r => String(r.person === null || r.person === undefined ? '' : r.person).trim();
  const list = path || [];
  if (!list.length) return list;
  let i = 0;
  while (i < list.length && who(list[i]) === who(list[0])) i++;
  return list.slice(0, i);
}

/**
 * 1 つの印の跡（新しい順）を、保存ごとにまとめて並べ直す。1 回の保存で、外した行（REMOVE）と同じ印の行を足す・変えることがある
 * （行の出どころで組んだとき。例: 行を外し、別の行をその印に変えた）。同じ保存の中では外した方を古い方にし、残った行の跡が先に来るようにする。
 * 保存の並び（新しい順）は変えない
 */
function appInputLogHistOrder_(h) {
  const groups = [], at = {};
  (h || []).forEach(r => {
    const k = String(r.action_id || '') + '\u0001' + String(r.saved_at || '');
    if (!at[k]) { at[k] = []; groups.push(at[k]); }
    at[k].push(r);
  });
  return [].concat.apply([], groups.map(g => g.filter(r => r.change !== 'REMOVE').concat(g.filter(r => r.change === 'REMOVE'))));
}

/** 記録の JSON の列を読む（空・読めなければ null） */
function appInputLogJson_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'object') return v;
  try { return JSON.parse(String(v)); } catch (e) { return null; }
}

/**
 * 計画の画面に付ける入力の記録（apiPlanView。見る人ごと）。rows[種類][i] は画面の入力の行（boot.input[種類][i]）と同じ並び
 * （同じ ENG_ の表から、同じく空の行を除いて作る）。行ごとに { conf: 一番新しい自信, n: 記録の数, hist: 新しい順の跡 }。記録の無い行は null。
 * 跡は行ごとにたどる（appInputLogTrails_: 足した行で止まり、前の印へのつなぎ _key（#n の付け直し・出どころで組んで印を変えた行）は前の印の跡へ続き、
 * つなぎの無い CHANGE で印が変わったものでは止まる）。付け直しだけの記録は跡に出さない。
 * 予算策定担当以上でない人には、本人と分かる行だけを送り、その中の担当者の名前と保存した人は送らない。ほかは種類ごとの件数（counts）だけ。
 * その人に送る跡は、今の担当者の間の分だけ（appInputLogSamePerson_。担当者を変えた行の跡は前の担当者の記録までたどるが、
 * 前の担当者だった本人に、今はほかの人の行の跡と自信を出さない）。
 * 入力の記録の表がまだ使えない（appLogReady_('INPUT_LOG'): 表が無い・列を足す前）・読めないときは null（画面はそのまま出す）
 */
function appInputLogView_(ctx, plan) {
  if (!appLogReady_('INPUT_LOG')) return null;
  try {
    const viewer = appLogViewer_(ctx, plan.client_id);
    const log = appInputLogRead_(plan.plan_id);
    const vis = appLogVisible_(viewer, log, { personOf: r => r.person, dateOf: r => r.saved_at, typeOf: r => r.kind });
    const shown = new Set(vis.rows);
    const trail = appInputLogTrails_(log);   // 行の跡は全部の記録でたどり、見せる行だけを送る
    // 記録だけの印（_key）は送らない。予算策定担当以上でない人には担当者の名前も送らない
    const clean = x => {
      if (!x || typeof x !== 'object') return x;
      const c = Object.assign({}, x);
      delete c._key;
      if (!viewer.full) delete c.person;
      return c;
    };
    const hideName = j => clean(appInputLogJson_(j));
    // 本人の行の前の中身: 担当者が変わった行（ほかの人の行を引き継いだ）は、前の担当者もその日に本人のときだけ送る
    const beforeOf = r => {
      const x = appInputLogJson_(r.before_json);
      if (!x || viewer.full) return clean(x);
      const was = String(x.person === null || x.person === undefined ? '' : x.person).trim();
      if (was && was !== String(r.person || '').trim() &&
        appPersonEmailOf_(was, viewer.clientId, appLogDay_(r.saved_at) || appToday_(), viewer.links) !== viewer.email) return null;
      return clean(x);
    };
    const byOf = r => (viewer.full && !appLogSystemActor_(r.actor_email) ? r.actor_email : '');   // 一度だけの写しの行は人ではない
    const rows = {};
    Object.keys(APP_INPUT_LOG_SHEETS).forEach(sheet => {
      const kind = APP_INPUT_LOG_SHEETS[sheet];
      rows[kind] = appInputLogList_(kind, appReadPlanTable_('ENG_' + sheet, plan.plan_id)).map(x => {
        // 予算策定担当以上でない人には、今の担当者の間の跡だけ（担当者を変える前の跡は、前の担当者の行として出さない）
        const path = viewer.full ? trail(x.key) : appInputLogSamePerson_(trail(x.key));
        const all = path.filter(r => shown.has(r));
        if (!all.length) return null;
        const live = all.filter(r => r.change !== 'REMOVE')[0];   // 自信は付け直しの記録からも読む（その行の自信を引き継いだ行）
        const h = all.filter(r => !appInputLogRenumbered_(r));
        return { conf: APP_INPUT_CONF_KINDS.indexOf(kind) >= 0 && live ? String(live.self_conf || '') : '', n: h.length,
          hist: h.slice(0, APP_INPUT_LOG_HIST).map(r => ({ at: r.saved_at, change: r.change, action: r.action, by: byOf(r),
            reason: viewer.full ? r.reason : '', conf: r.self_conf, before: beforeOf(r), after: hideName(r.after_json) })) };
      });
    });
    return { full: viewer.full, rows: rows, counts: vis.counts, hidden: vis.hidden };
  } catch (e) {
    appLogError_('INPUT_LOG.VIEW', e, ctx);
    return null;
  }
}

// ---- 一度だけの写し（版 10 の移行の後。V10.js の appV10BackfillInputLog_ から）----

/**
 * 今の 4 つの入力の表の行を、計画ごとに BASELINE として 1 回だけ写す（締めた年度の計画は飛ばす）。キーは計画・印から決める（appStableLogId_。
 * 2 つの操作が同時に動かしても二重にならない）。BASELINE がもうある計画は飛ばす。1 回は約 20 秒までで、残りは次の操作で（{ more: true }）。
 * 計画ごとにロックの中で、書きかけの保存を先に書き終えてから（appJournalRecover_）読む（途中の表から写さない）。
 * 記録がもうある行（写しより前に保存した行）には BASELINE を足さない（写しの方が新しくなり、その保存の自信を隠すため。その行の初めの姿は前の記録にある）。
 * した人は仕組み（APP_V10_BACKFILL_ACTOR）。写しを動かした操作の人ではない。
 * 行を確かめられない計画（同じキーの 2 行など。appLogOps_ が控えを置く前に止める）は、エラーのログに残して飛ばし、ほかの計画を続ける。
 * 飛ばした計画があれば最後に投げる（済みにしない。appRunBackfills_ が失敗として残し、10 分あけてその計画だけやり直す）
 */
function appInputLogBackfill_(ctx) {
  if (!appLogReady_()) return { more: true, rows: 0 };
  const t0 = new Date().getTime();
  const done = {};
  let log;
  try { log = appReadTable_('INPUT_LOG'); } catch (e) { return { more: true, rows: 0 }; }   // 表がまだ無い（移行の前）: 次の操作でまた
  log.forEach(r => { if (r.change === 'BASELINE') done[r.plan_id] = true; });
  const plans = appReadTable_('PLANS').filter(p => !done[p.plan_id] && !appYearIsFrozen_(p.fy));
  let rows = 0;
  const skipped = [];   // 写す行を確かめられなかった計画（控えを置かずに飛ばす）
  for (let i = 0; i < plans.length; i++) {
    if (new Date().getTime() - t0 > APP_INPUT_LOG_BACKFILL_MS) return { more: true, rows: rows };
    const planId = plans[i].plan_id;
    rows += appWithLock_(() => {
      appJournalRecover_(ctx);   // 書きかけの保存を先に書き終える（その保存の入力の行と記録が入ってから読む）
      const plan = appPlanOf_(planId);
      if (appYearIsFrozen_(plan.fy)) return 0;
      const had = appReadPlanTable_('INPUT_LOG', plan.plan_id);
      if (had.some(r => r.change === 'BASELINE')) return 0;   // ほかの操作が先に写した
      const logged = {};
      had.forEach(r => { logged[r.row_key] = true; });
      const now = appNowIso_();
      const list = [];
      Object.keys(APP_INPUT_LOG_SHEETS).forEach(sheet => {
        const kind = APP_INPUT_LOG_SHEETS[sheet];
        appInputLogList_(kind, appReadPlanTable_('ENG_' + sheet, plan.plan_id)).filter(x => !logged[x.key]).forEach(x => list.push({
          plan_id: plan.plan_id, log_id: appStableLogId_(APP_LOG_PREFIX.INPUT_LOG, ['BASELINE', plan.plan_id, x.key]), action_id: '', action: 'BASELINE',
          kind: kind, change: 'BASELINE', row_key: x.key, person: String(x.row.person === null || x.row.person === undefined ? '' : x.row.person),
          before_json: '', after_json: x.row, self_conf: '', reason: '', signal_id: '', actor_email: APP_V10_BACKFILL_ACTOR, saved_at: now }));
      });
      if (!list.length) return 0;
      let ops;
      try {
        ops = appLogOps_('INPUT_LOG', list);   // 行を確かめる（同じキーの 2 行など）。止まったら控えを置かない
      } catch (e) {
        skipped.push(plan.plan_id);
        appLogError_('INPUT_LOG.BACKFILL', new Error('入力の記録の初めの姿を写せないので、この計画を飛ばしました（' + plan.plan_id + '）: ' +
          String(e && e.message ? e.message : e)), ctx);
        return 0;
      }
      appJournalRun_(ctx, '入力の記録の初めの姿（' + plan.plan_id + '）', plan.plan_id, ops);
      return list.length;
    });
  }
  // 飛ばした計画があれば、済みにしない（一度だけの写しの失敗として残り、10 分あけてその計画だけやり直す。ほかの計画の写しと保存は止めない）
  if (skipped.length) throw new Error('入力の記録の初めの姿を写せない計画を飛ばしました（' + skipped.length + ' 計画: ' + skipped.join(', ') + '）。ほかの計画は写しました。');
  return { rows: rows };
}
