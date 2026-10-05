# Forecast vNext Client Runtime

Client FY Bookへ入れるコードを、Admin Hub・legacy・予測計算・Vertex AIから物理的に分離するための独立ビルドです。

Client Runtimeに含めるのは次だけです。

- Clientローカルの追記型レコードと状態遷移
- 現場情報、計画案、振り返りの保存
- Admin所有workerが回収する予測依頼キュー
- 3画面、4つの従業員メニュー、4つのサイドバー
- Clientローカルへ返された予測・計画・評価の読取

次は含めません。

- `Forecast_Agent.js`、`VNext_Admin.js`、`VNext_Engine.js`
- ZAC source ID、Admin Hub ID、Vertex設定、AI raw、prompt本文、秘密情報
- `DriveApp`、`UrlFetchApp`、`openById`、Script Properties
- Drive、Cloud Platform、external request、Apps Script project編集のOAuth scope

Admin側syncとの列ずれを防ぐため、追記型テーブルのheader順は中央schemaと同一です。ただしClient側では`source_spreadsheet_id`、AI model/prompt/rule列を常に空にし、`MODEL_RELEASE`はheaderだけで行を配布しません。header名は非表示・保護された互換契約であり、source ID、prompt、係数、設定値そのものはClient Bookに存在しません。

## なぜ空のSpreadsheetからTemplateを作るのか

Spreadsheetをコピーするとbound scriptも丸ごと複製されます。したがって、legacy/Adminコード入りブックをコピーしてシートだけ削除しても物理分離にはなりません。また、コピー後の新しいbound script IDをSpreadsheet APIから取得する安定した方法はありません。

初回だけ次の順でMaster Templateを作ります。

1. Admin側で空のSpreadsheetを新規作成する。
2. Apps Script API `projects.create(parentId=spreadsheetId)` で新しいbound projectを作る。
3. `projects.updateContent` で検証済みClient Runtimeだけを投入する。
4. Admin側からTemplateのシート・空の監査テーブルを初期化する。
5. Client FY Bookは、このclean TemplateだけをDriveでコピーする。

新規の [`VNext_ClientRuntimeProvisioning.js`](../VNext_ClientRuntimeProvisioning.js) が1〜3を実装します。返されたTemplate `scriptId`はAdmin Hubだけへ保存し、Client Bookへ書きません。

## ビルド

Node.jsの外部packageは不要です。

```bash
cd client_runtime
npm run vendor
npm run build
```

`vendor`は従業員UXと4 HTMLを`src/`へ固定し、`vendor-lock.json`へsource/output SHA-256を記録します。通常の`build`はroot UXが後から変わった場合に失敗するため、差分をレビューせず古いClient Runtimeを配布しません。

`build`は以下を自動検査します。

- 配布ファイルがallowlistの9ファイルだけであること
- Admin/legacy/forecast engine/Vertex/source設定/cross-file APIがないこと
- manifestが`currentonly`を含む3 scopeだけであること
- HTMLが呼ぶserver functionが定義済みであること
- bundle hashが再現できること

Admin health scanでは静的検査に加え、全Client Bookで次をfail-closed検査します。

- `BOOK_META.source_spreadsheet_id`の非空行が0件
- `EVIDENCE_EVENT`のAI/prompt/rule列の非空行が0件
- `MODEL_RELEASE`のデータ行が0件
- Client `VN_SYSTEM_CONFIG`にHub/Template/source/script IDがない

生成される`VNext_ClientRuntimeBundle.js`はAdmin側GASへ`clasp push`するための新規root fileです。Clientへ送るcontentだけを文字列として保持し、Admin/秘密値は含みません。

## Admin Provisionerへの接続

Admin側manifestに、Template生成時だけ必要な次のscopeを追加し、対応GCP projectでApps Script APIを有効化します。

```json
"https://www.googleapis.com/auth/script.projects"
```

`script.external_request`はAdmin側manifestに既に必要です。Client側manifestにはどちらもありません。

Template生成箇所では、legacy spreadsheetの`makeCopy`を使用せず次の形にします。

```javascript
var created = vNextClientRuntimeCreateBoundSpreadsheet_({
  title: 'Forecast vNext Master Template',
  folderId: rootFolderId
});
var template = SpreadsheetApp.openById(created.spreadsheetId);
vNextAdminInitializeTemplate_(template, {
  resetCopied: true,
  bookId: templateBookId,
  releaseId: releaseId,
  adminEmails: adminEmails,
  templateSpreadsheetId: created.spreadsheetId,
  now: new Date(),
  actor: actor
});
```

Admin Hubの非表示設定へ`created.scriptId`、`created.bundleSha256`、`created.runtimeVersion`を保存します。Client初期化時の`VN_SYSTEM_CONFIG`へはHub/Template/source/script IDを入れず、`mode`、`book_id`、`active_release_id`、`schema_version`だけを書きます。

AIイベントの`EVIDENCE_EVENT`をHubからClientへ丸ごと同期してはいけません。社員向け引用は、prompt/raw metadataを除いた公開要約を`FORECAST_RUN`へ保存して返します。

## claspによる既存Template更新

Template script IDがAdmin Hubで分かっている場合は、dry-run確認後に独立して反映できます。

```bash
npm run deploy:template -- --script-id SCRIPT_ID
npm run deploy:template -- --script-id SCRIPT_ID --apply
npm run verify:remote -- --script-id SCRIPT_ID
```

`deploy-template`は一時的な`dist/.clasp.json`を権限`0600`で作り、終了時に削除します。rootの`.clasp.json`やAdmin projectへは向きません。

## リリース運用

- stable release作成時にClient Runtimeのbundle SHAを`MODEL_RELEASE/RELEASES`へ記録する。
- 新FYは更新済みclean Templateから作る。
- 公式確定済みClient Bookへ通常更新をかけない。
- 重大修正だけ、対象bookとruntime SHAを固定したversioned migrationで適用する。
- Template copy後も、Client BookのApps Script Editorで配布ファイルallowlistと最小scopeをspot checkする。
