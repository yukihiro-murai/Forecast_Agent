/**
 * Engine.js — 予測の計算を動かす土台（段階2）。
 * 旧来の計算（Forecast_Agent.js）は書き換えず、新アプリの中の「計算用ブック」の上で動かす。
 * 計算用ブックは「売上予測アプリ（システム）」フォルダに 1 つだけ置き、実行のたびに中身を入れ替える（所有者だけ・人は触らない）。
 * 乱数は実行ごとの種（seed）で固定し、同じ入力と同じ種なら同じ P10/P50/P90 になるようにする（FORECAST_RUNS に種を残す）。
 */

/** 文字列から 128 ビットの種を作る（cyrb128） */
function appSeedFrom_(text) {
  const str = String(text);
  let h1 = 1779033703, h2 = 3144134277, h3 = 1013904242, h4 = 2773480762;
  for (let i = 0; i < str.length; i++) {
    const k = str.charCodeAt(i);
    h1 = h2 ^ Math.imul(h1 ^ k, 597399067);
    h2 = h3 ^ Math.imul(h2 ^ k, 2869860233);
    h3 = h4 ^ Math.imul(h3 ^ k, 951274213);
    h4 = h1 ^ Math.imul(h4 ^ k, 2716044179);
  }
  h1 = Math.imul(h3 ^ (h1 >>> 18), 597399067);
  h2 = Math.imul(h4 ^ (h2 >>> 22), 2869860233);
  h3 = Math.imul(h1 ^ (h3 >>> 17), 951274213);
  h4 = Math.imul(h2 ^ (h4 >>> 19), 2716044179);
  h1 ^= (h2 ^ h3 ^ h4); h2 ^= h1; h3 ^= h1; h4 ^= h1;
  return [h1 >>> 0, h2 >>> 0, h3 >>> 0, h4 >>> 0];
}

/** 種つきの乱数（sfc32。0 以上 1 未満） */
function appSeededRandom_(seedText) {
  let [a, b, c, d] = appSeedFrom_(seedText);
  const next = () => {
    a >>>= 0; b >>>= 0; c >>>= 0; d >>>= 0;
    let t = (a + b) | 0;
    a = b ^ (b >>> 9);
    b = (c + (c << 3)) | 0;
    c = (c << 21) | (c >>> 11);
    d = (d + 1) | 0;
    t = (t + d) | 0;
    c = (c + t) | 0;
    return (t >>> 0) / 4294967296;
  };
  for (let i = 0; i < 15; i++) next();  // 種の偏りを捨てる
  return next;
}

/** fn の間だけ Math.random を種つきの乱数にする（入れ子にしない。終われば必ず元に戻す） */
function appWithSeededRandom_(seedText, fn) {
  if (APP_RANDOM_PATCHED_) throw new Error('乱数の固定が入れ子になっています。');
  const original = Math.random;
  Math.random = appSeededRandom_(seedText);
  APP_RANDOM_PATCHED_ = true;
  try {
    return fn();
  } finally {
    Math.random = original;
    APP_RANDOM_PATCHED_ = false;
  }
}
let APP_RANDOM_PATCHED_ = false;

/**
 * 計算用ブック（無ければ作る）。人が開いて触る前提のファイルではないので、名前で示す。
 * 計算は appWithLock_ の中で行い、同時に 2 つの計算が同じブックを使わないようにする。
 */
function appScratchBook_() {
  const props = appProps_();
  const id = props.getProperty(APP_PROP.scratchId);
  if (id) {
    try {
      const ss = SpreadsheetApp.openById(id);
      if (!DriveApp.getFileById(id).isTrashed()) return ss;
    } catch (e) {
      Logger.log('計算用ブックを開けないので作り直します: ' + (e && e.message ? e.message : e));
    }
  }
  const folderId = props.getProperty(APP_PROP.folderId);
  if (!folderId) throw new Error('初期設定がまだです。');
  const ss = SpreadsheetApp.create(APP_FILES.scratch);
  DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(folderId));
  props.setProperty(APP_PROP.scratchId, ss.getId());
  return ss;
}

/**
 * 計算用ブックを空にする（シートを 1 枚の空のシートだけにする）。
 * 旧来の計算は必要なシートを自分で作るか、呼ぶ側が入れたシートだけを読む。
 */
function appScratchReset_(ss) {
  const keep = ss.insertSheet('_EMPTY_' + Utilities.getUuid().slice(0, 8));
  ss.getSheets().forEach(s => { if (s.getSheetId() !== keep.getSheetId()) ss.deleteSheet(s); });
  return keep;
}
