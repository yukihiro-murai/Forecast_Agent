# データ契約 — 年度予算策定システム（Forecast vNext）

対象読者: 上位システム（BI / AI / データ統合基盤）との連携を実装する人、将来の保守者。
最終更新: 2026-08-25（portal 1.7.27）

---

## 1. 全体構造と正本

| 層 | 実体 | 役割 | 正本データ |
|---|---|---|---|
| 第1層 | クライアント年度予算の管理表（Spreadsheet + Web入口） | 申請の受付・状況表示・機械連携の入口 | `VN_PORTAL_REQUEST` / `PORTAL_DIRECTORY` |
| 第2層 | 年度予算シート（クライアント1社×1年度ごとの Spreadsheet） | 入力・予測・予算の作業場 | 各ブック内テーブル |
| 第3層 | 管理ハブ（Spreadsheet） | 登録・承認・ジョブ・監査の中枢 | `BOOK_REGISTRY` / `APPROVALS` / `JOBS` / `AUDIT_LOG` ほか |

- 表示名の正は Web入口の定義（`0_VNext_Naming.js` / `VNEXT_PORTAL_NAMING`）。
- Drive ファイル名は反映（deploy finalize）のたびに表示名へ自動整合される。

## 2. 機械連携の入口（推奨順）

### 2.1 JSON スナップショット API

```
GET <Web入口のURL>?format=json
```

- 認証: Web アプリと同じ（ドメイン内 Google アカウント）。
- キャッシュ: サーバ側 60 秒（申請イベント発生時は即時無効化）。
- スキーマ ID: `vnext-portal-snapshot-1`。**同一 ID 内でフィールドは後方互換を維持**し、
  形を変えるときは ID を bump する（利用側は `schema` を必ず確認すること）。

```json
{
  "ok": true,
  "schema": "vnext-portal-snapshot-1",
  "runtimeVersion": "vnext-portal-1.7.27",
  "generatedAt": "2026-08-25T09:00:00.000Z",
  "books": [{
    "requestId": "PORTAL-REQ-…", "clientId": "ZAC-C-001",
    "clientName": "…", "fiscalYear": 2027,
    "state": "SUBMITTED", "stateLabel": "承認待ち",
    "centerForecast": 1000, "adoptedForecast": 1200, "finalBudget": null,
    "nextAction": "…", "url": "https://docs.google.com/…", "updatedAt": "ISO8601"
  }],
  "requests": [{
    "requestId": "PORTAL-REQ-…", "clientName": "…", "fiscalYear": 2027,
    "status": "CREATING", "statusLabel": "年度予算シートを作成中",
    "relatedBookId": "CLIENT-…", "url": "", "requestedAt": "ISO8601", "updatedAt": "ISO8601"
  }]
}
```

- `books` = 作成済み年度予算シートの現在状態（`PORTAL_DIRECTORY` 射影）。
- `requests` = 作成申請の進行状況（イベントログの検証済み射影）。
- 失敗時は `{ "ok": false, "schema": "…", "error": "…" }`。

### 2.2 スプレッドシート表の直接読み取り（読み取り専用）

| シート | 意味 | 備考 |
|---|---|---|
| `VN_PORTAL_REQUEST` | 申請の**追記専用イベントログ**（schema `vnext-portal-request-2`） | 1申請=複数イベント行。`request_hash` は正規化 JSON の SHA-256。改ざん・不正行は読み取り時に無視される |
| `PORTAL_DIRECTORY` | ACTIVE な年度予算シートの射影 | 管理ハブが書き込む。`directory_key` で最新行が有効 |
| `VN_PORTAL_CLIENT_CATALOG` | クライアント候補（ZAC 由来） | `catalog_version` 世代管理 |
| 管理ハブ `BOOK_REGISTRY` | 全ブックの台帳（`book_id` が主キー） | `requests.relatedBookId` と結合可能 |

**書き込みは全て所定の関数経由のみ。外部システムがセルへ直接書き込むことは契約違反。**

### 2.3 人間ビュー（契約外）

`ホーム` / `FY20xx` シートは表示専用で、レイアウトは予告なく変わる。機械連携で参照しないこと。

## 3. 状態機械

- 申請 `status`: `PENDING → VALIDATING → CREATING → COMPLETED`（終端: `COMPLETED` / `FAILED` / `REJECTED`）
- ブック `state`: `INPUT_OPEN → READY_TO_RUN → RUNNING → DRAFT_READY → SUBMITTED →`
  `OFFICIAL_LOCKED → REVIEW_DUE → YEAR_CLOSED`（差戻し: `CHANGES_REQUESTED`）
- 日本語ラベルはコード側の対応表（`vNextPortalDirectoryStateLabel_` / `REQUEST_STATUS_LABELS`）が正。
  機械連携は英字コードを使い、ラベルを解析しないこと。

## 4. データ衛生ルール

1. イベントログは追記専用。訂正は新イベントで表現し、行の書き換え・削除をしない。
2. 時刻は ISO 8601（UTC）。表示整形はビュー層でのみ行う。
3. 金額は生の数値で保持（`¥` 等の書式はビュー層のみ）。
4. 表示文言（説明文・案内）をデータ行に混ぜない。データシートはヘッダー+データのみ
   （恒常ルール: gas-workspace `.ai/memory/decisions/2026-08-25-information-design-principle.md`）。
5. 結合キー: `requestId`（申請）、`book_id`（ブック台帳）、`client_id + fiscal_year`（業務キー）。

## 5. スキーマ変更の手順（保守者向け）

1. 列追加はヘッダー検証（`*_HEADERS`）と読み書き双方の対応、旧版行の互換読み取りをセットで行う。
2. 版数カスケード（毎回同時に更新する5点）: `Portal_Core RUNTIME_VERSION` /
   `portal_runtime/package.json` / `VN_ADMIN_PORTAL_RUNTIME_VERSION` + legacy 一覧 /
   `VN_ADMIN_RUNTIME_BUILD_STAMP` / テストの版数ピン（portal 1 + root 2）。
3. スナップショットの形を変えるときは `SNAPSHOT_SCHEMA` を bump し、本書 §2.1 を更新する。

## 6. 成長と運用上の既知事項

- `VN_PORTAL_REQUEST` は無制限に伸びる。**目安 5,000 行**を超えたら、終端状態かつ
  60 日超のイベントをアーカイブシートへ移す保守処理を実装すること（未実装・計画済み）。
- Web 公開は Apps Script の保存版数上限（約200）を消費する。150 版以降は反映結果に警告が出る。
