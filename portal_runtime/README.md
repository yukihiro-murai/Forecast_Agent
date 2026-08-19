# Forecast vNext Shared Portal Runtime

全社員がクライアント別年度計画を探し、存在しなければ作成依頼を送るための、Spreadsheet-bound最小runtimeです。個別クライアントの予測計算や外部データ取得は行いません。

## 社員画面

- `ホーム`: 作成依頼の受付・作成中・完成・要確認を一覧表示
- `FYyyyy`: 1行を「1クライアント × 1年度」とする年度別一覧
- メニュー: `案内を開く` のみ。日常作業は右側の案内
- Web入口: `doGet` / `Portal_Entry.html`。各手順に単色Botと吹き出し案内。01で新しい個別シート、02で作成済み計画、03で管理者の承認・整備。

専用ブックが完成すると、`PORTAL_DIRECTORY`または完了イベントのURLから「開く」リンクを表示します。社員のGoogleアカウントをクライアント別allowlistでは制限しません。作成担当（Forecast Owner）は入力させず、送信時のログインユーザーを自動設定します。

作成依頼後は、サイドバーに`受付済み → 内容確認 → ブック作成 → 利用可能`の4段階を表示します。状態確認は依頼ログだけを読む軽量APIで、20秒後・約1分半後・約5分後の最大3回に限定します。自動確認後も手動で再確認でき、15分以上状態が変わらない場合は受付番号を添えて管理者へ連絡する案内を表示します。

## Request API

社員向け公開関数:

- `vNextPortalGetCreateModel()`
- `vNextPortalPreviewCreation(input)`
- `vNextPortalSubmitCreationRequest(input)`
- `vNextPortalGetRequestProgress(requestId)`
- `vNextPortalShowRequestOnHome(requestId)`
- `vNextPortalGetEntryModel()`
- `vNextPortalGoHome()`
- `vNextPortalOpenCreateSidebar()`
- `vNextPortalOpenHelp()`

`vNextPortalPreviewCreation`は次のexact keysだけを受け入れます。

```text
clientKey, fiscalYear, relatedMemberNames
```

`vNextPortalSubmitCreationRequest`は次のexact keysだけを受け入れます。

```text
clientKey, confirmSimilarDuplicates, duplicateCheckHash,
fiscalYear, relatedMemberNames
```

`clientKey`は管理側がZACから事前同期した`VN_PORTAL_CLIENT_CATALOG`の選択値です。社員によるクライアント名・IDの自由入力は受け付けません。関与メンバーは氏名を1〜5名入力します。候補確認hashは、同じFYの`PORTAL_DIRECTORY`と処理中のローカル依頼を再読込して生成します。送信時にロック内で再計算し、候補が増減していれば再確認を要求します。同一IDまたは正規化名の完全一致は送信を止め、類似名は社員の明示確認を必須にします。

## `VN_PORTAL_REQUEST` contract

列順は固定です。

```text
request_event_id, request_id, event_type, status, request_hash, request_json,
fiscal_year, client_id, client_name, forecast_owner_email,
related_member_emails_json, requested_at, requested_by, related_book_id,
related_book_url, detail_code, detail_message, created_at, created_by,
catalog_key, related_member_names_json
```

runtimeは`REQUESTED / PENDING`だけをappendします。`request_json`は次のexact payloadをcanonical JSON化したもの、`request_hash`はそのUTF-8 SHA-256です。

```text
catalogKey, clientName, fiscalYear, relatedMemberNames,
requestId, requestType, requestedAt, requestedBy, schemaVersion
```

固定値:

- `requestType = CREATE_CLIENT_FY_BOOK`
- `schemaVersion = vnext-portal-request-2`

既存の`vnext-portal-request-1`行は読み取り・状態更新の後方互換を維持します。

非同期処理側は同じ`request_id`と`request_hash`を使い、同じテーブルへ状態イベントをappendします。推奨ペア:

| event_type | status | 意味 |
|---|---|---|
| `VALIDATION_STARTED` | `VALIDATING` | 内容確認中 |
| `CREATION_STARTED` | `CREATING` | 専用ブック作成中 |
| `COMPLETED` | `COMPLETED` | URL利用可能 |
| `FAILED` | `FAILED` | 再試行可能な失敗 |
| `REJECTED` | `REJECTED` | 重複等の人による確認が必要 |

`COMPLETED`では`related_book_id`と`related_book_url`を設定します。失敗・差戻し理由は`detail_code`と社員向け`detail_message`へ入れます。元の`REQUESTED`行は更新しません。
非`REQUESTED`イベントの`request_json`は空欄にします。runtimeはevent/statusペア、元requestと同じhash、URL形式を検証し、不正な状態行を表示へ採用しません。

## `PORTAL_DIRECTORY` contract

列順は固定です。

```text
directory_event_id, directory_key, fiscal_year, client_id, client_name,
forecast_owner_email, related_member_emails_json, state, center_forecast,
adopted_forecast, final_budget, next_action, client_book_url, request_id,
updated_at, updated_by, related_member_names_json
```

同じ`directory_key`の最終行を現在値として表示するappend projectionです。URLはGoogle SpreadsheetのHTTPS URLだけをリンク化します。

## Build and verification

```bash
cd portal_runtime
npm run check
```

生成物`../VNext_PortalRuntimeBundle.js`は、検証済み4ファイルだけを`VNEXT_PORTAL_RUNTIME_BUNDLE_`として保持します。runtime manifestは`container.ui`、`spreadsheets.currentonly`、`userinfo.email`の3 scopeだけです。
