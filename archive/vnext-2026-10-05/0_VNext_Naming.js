/**
 * Forecast vNext — user-facing Japanese labels (正本).
 * Technical IDs, sheet names, and code identifiers are unchanged.
 *
 * User-facing layer numbers match Web入口 top-to-bottom (01→02→03):
 *   第1層 申請入口 → 第2層 クライアント年度ブック → 第3層 管理ハブ
 * Drive folder prefixes (01_管理ハブ 等) are infrastructure only — not shown as layer numbers.
 */
var VNEXT_NAMING = Object.freeze({
  SYSTEM: '年度予算策定',
  MENU: '年度予算策定',
  SHARED_DRIVE: '年度予算策定',
  LEGACY_SHARED_DRIVE: '年度計画',
  ADMIN_HUB: '管理ハブ',
  PORTAL: '申請入口',
  CLIENT_BOOK: 'クライアント年度ブック',
  LAYER1_SHORT: '第1層：申請入口',
  LAYER2_SHORT: '第2層：クライアント年度ブック',
  LAYER3_SHORT: '第3層：管理ハブ',
  PORTAL_DEFAULT_TITLE: '申請入口',
  WEB_ENTRY: '年度予算策定 Web入口',
  FOLDER_ADMIN: '01_管理ハブ',
  FOLDER_PORTAL: '02_申請入口',
  FOLDER_BOOKS: '03_クライアント年度ブック',
  FOLDER_TEMPLATES: '04_テンプレート',
  FOLDER_AUDIT: '監査',
  TEMPLATE_CURRENT: '現行',
  TEMPLATE_DRAFT: '下書き',
  TEMPLATE_HISTORY: '履歴',
  FOLDER_LEGACY: Object.freeze({
    '01_管理ハブ': '03_管理者',
    '02_申請入口': '01_社員ポータル',
    '03_クライアント年度ブック': '02_クライアント年度ブック'
  }),
  FORECAST_OWNER: '予算策定担当',
  FORMAL_BUDGET: '正式予算',
  BUDGET_DRAFT: '予算案',
  ADMIN_CONTACT: '管理ハブ担当者',
  CLIENT_BOOK_ID: 'クライアント年度ブック ID',
  /** @deprecated Use ADMIN_HUB / PORTAL / CLIENT_BOOK. Kept for internal audit notes. */
  LAYER1: '管理ハブ',
  /** @deprecated Use PORTAL. */
  LAYER2: '申請入口',
  /** @deprecated Use CLIENT_BOOK. */
  LAYER3: 'クライアント年度ブック',
  /** @deprecated Use PORTAL_DEFAULT_TITLE. */
  LAYER2_DEFAULT_TITLE: '申請入口'
});

/** @param {string} clientName @param {number|string} fiscalYear */
function vNextFormatClientBookTitle_(clientName, fiscalYear) {
  return String(clientName || 'クライアント') + '｜FY' + String(fiscalYear || '') + ' 年度予算';
}
