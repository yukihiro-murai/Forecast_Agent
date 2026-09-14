# 管理ハブ サイドバー簡素化 ＋ 更新一発化（設計・実装記録）

対象: `VNext_AdminSidebar.html` / `VNext_Admin.js` / `VNext_ClientRuntimeProvisioning.js`
実装 commit: `34fe40c`（ブランチ `cursor/admin-sidebar-simplify-0efd`、Portal 1.7.38 を merge 済み → `38d2020`）
中央 clasp push: 2026-09-14 17:51 JST（19ファイル）

## 1. 操作の分類

| 操作 | 分類 | Before（サイドバー） | After |
|---|---|---|---|
| 状態カード・件数（承認待ち/要確認/自動処理中/ブック数） | 日常・状態表示 | 常設 | **サイドバー**（そのまま） |
| 承認 / 差戻し / 却下 | 日常 | 常設 | **サイドバー** |
| 要確認事項（正式予算を再反映・申請を今すぐ処理・ブックを開く） | 日常 | 常設 | **サイドバー** |
| 申請・自動処理の状況 ＋「申請を今すぐ処理」 | 日常 | 常設 | **サイドバー** |
| 自動運用を有効化 | 1回きり | 常設（有効化後は非表示） | **サイドバー**（未設定のときだけ表示。既存挙動） |
| テスト前ブックの最新版更新（安全条件確認→更新→復旧） | たまに（条件付き） | 候補があるときだけ表示 | **サイドバー**（そのまま。候補が無ければ非表示） |
| 版・更新の状態（Hub SHA / 申請入口版 / 更新ジョブ） | 状態表示 | 「保守・高度な操作」内に埋没 | **サイドバー**に「版・更新」カードを新設 |
| 全ブックを確認 / 要確認一覧を更新 | たまに・裏側 | サイドバー「状態確認（詳細）」 | **メニュー 保守** |
| ZAC候補を更新 | たまに・裏側（sweep が自動更新） | サイドバー | **メニュー 保守** |
| 登録一覧を開く | たまに | メニュー その他 | **メニュー 保守** |
| 中央配備版へ更新（Hub runtime） | たまに（deploy 時） | 保守・高度な操作（理由必須） | **メニュー 保守 →「最新版に更新」**に統合 |
| 申請入口を最新版へ更新（Portal runtime + /exec） | たまに（deploy 時） | 申請入口の設定 → 既存の申請入口を最新版へ | 同上に統合（**1操作**） |
| 申請入口を準備する | 1回きり（フォーム2項目） | サイドバー | **高度な操作**ビュー（初回のみ、折りたたみ） |
| 初回 Pilot 段階準備 / Canary 承認 | 1回きり・復旧 | サイドバー | **高度な操作**ビュー |
| クライアント年度ブックを個別作成（例外） | たまに（フォーム） | サイドバー | **高度な操作**ビュー |
| Vertex AI 接続 | 1回きり（フォーム） | 保守・高度な操作 | **高度な操作**ビュー |
| Template Draft / STAGED / pair 有効化 | たまに（フォーム） | 保守・高度な操作 | **高度な操作**ビュー |
| モデル版 DRAFT 登録 / 有効化 / ロールバック | たまに（フォーム） | 保守・高度な操作 | **高度な操作**ビュー |
| ブック個別操作（AI取消 / 訂正 / 差戻し / 振り返り / 年度終了） | たまに（フォーム） | 保守・高度な操作 | **高度な操作**ビュー |
| 共有ドライブへ移す | 1回きり | 初回セットアップ ＋ 保守・高度な操作（2か所） | **メニュー 初回・復旧**（確認ダイアログ→実行） |
| 受入試験をゼロからやり直す | 1回きり・危険 | 保守・高度な操作（確認語入力欄） | **メニュー 初回・復旧**（対象一覧 → 確認語 prompt → 削除の3段階） |
| 初回セットアップ手順（テキスト） | 1回きり | サイドバー常設アコーディオン | 削除（README_VNEXT_JA / AI_HANDOFF に集約） |
| LEGACY モード（bootstrap / 復旧 / API有効化 / Pilot生成） | 1回きり | LEGACY 時のみ表示 | 変更なし |

## 2. Before / After

### メニュー「年度計画」
```
Before                      After
├ 案内を開く                ├ 案内を開く
└ その他 ▸                  ├ 保守 ▸
   ├ 全クライアントの状態点検 │  ├ 最新版に更新（管理ハブ＋申請入口）   ← 新規・一発化
   └ 登録一覧を開く          │  ├ ZACクライアント候補を更新
                            │  ├ 全クライアントの状態点検
                            │  ├ 要確認一覧を更新
                            │  ├ 登録一覧を開く
                            │  ├ ───
                            │  └ 高度な操作を開く                     ← フォーム群（同一HTML の advanced ビュー）
                            └ 初回・復旧 ▸
                               ├ 共有ドライブへ整理（初回のみ）        ← 確認ダイアログ
                               └ 受入試験をゼロからやり直す（削除）    ← 一覧→確認語→削除
```

### サイドバー（daily ビュー）
Before: 状態 → 件数 → 承認待ち → 要確認 → 申請・自動処理（＋状態確認詳細）→ 初回セットアップ → 申請入口の設定 → 個別作成 → テスト前更新 → 保守・高度な操作（runtime / 共有ドライブ / Vertex / Release / モデル版 / ブック個別 / リセット）
After: 状態 → 件数 → 承認待ち → 要確認 → 申請・自動処理 → テスト前更新（条件付き）→ **版・更新**

### 実装方式
- `VNext_AdminSidebar.html` は 1 ファイルのまま `view` テンプレート変数（`daily` / `advanced` / `updater`）で出し分け（`vNextAdminSidebarOutput_(view)`、`HtmlService.createTemplateFromFile`）。管理ハブ runtime の 18 ファイル allowlist を変えないため HTML を増やさない（旧コードの allowlist 検証で新版コピーが拒否されるのを避ける）。
- 全 DOM id は 3 ビューで共存し、`render()` の参照先は不変。`applyView()` が表示切替。
- 危険操作の確認: 承認/差戻し/却下は `window.confirm`（既存）、共有ドライブ整理は `ui.alert OK_CANCEL`、リセットは `ui.alert`（対象一覧）→ `ui.prompt`（確認語 `RESET_GENERATED_CLIENTS`）→ 不一致なら何もしない。

## 3. 更新一発化の仕組み

### 問題
「中央配備版へ更新」は **実行中のコード自身**を書き換える。同じ実行で続けて申請入口を更新すると、旧コードが持つ旧 Portal bundle を配ってしまう。そのため従来は「再読み込み → 申請入口を更新」の 2 段だった。

### 解決
```
メニュー 保守 → 最新版に更新 → ダイアログ「更新を開始」
  ├ (1) vNextAdminUpdateAllFromSource        …旧コードで実行
  │     ・vNextAdminUpdateHubRuntimeFromSource({skipIfCurrent:true})
  │         中央 getContent → 検証 → Hub へ PUT → SHA 検証（同一なら省略）
  │     ・中央ファイル群から VNext_PortalRuntimeBundle の version / sha256 を正規表現で読む
  │       （vNextAdminRuntimePortalBundleIdentity_、コードは評価しない）
  │     ・VN_SYSTEM_CONFIG.runtime_update_job_json に
  │       {jobId, phase:ADMIN_UPDATED|DONE, targetPortalSha256, ...} を記録、監査 UPDATE_ALL_RUNTIMES STARTED
  │     ・portal_runtime_sha256 == targetPortalSha256 なら申請入口は不要 → DONE
  └ (2) vNextAdminContinueRuntimeUpdate({jobId})  …5秒ごとに再試行（最大 36 回）
        ・VNEXT_PORTAL_RUNTIME_BUNDLE_.sha256 != targetPortalSha256 → WAITING_FOR_NEW_CODE（まだ旧コード）
        ・一致したら script lock 下で phase=PORTAL_UPDATING に claim（多重実行防止、15分で lease 切れ）
        ・既存の vNextAdminUpdateSharedPortalRuntime（ロールバック付き、/exec ピン更新は scriptId 必須送信を維持）
        ・DONE / FAILED を記録。FAILED は要確認事項 RUNTIME_UPDATE_FAILED に出す
  完了 → 「閉じて案内を開き直す」で新しい案内を表示（再読み込み操作不要）
```
- 旧コードに (2) が無い初回移行では「Script function not found」を WAITING と同様に再試行。
- ダイアログを閉じても **`vNextAdminScheduledSweep`（5分ごと）の `vNextAdminAutoFollowRuntimeUpdate_`** が引き継ぐ:
  - (a) pending ジョブがあれば (2) を実行。
  - (b) ジョブが無くても `portal_runtime_sha256 != 実行中 bundle sha`（例: 旧 UI の「中央配備版へ更新」だけ実行した後）なら自動でジョブを作って申請入口を更新。同じ目標 SHA で失敗済みなら再試行しない。`VN_SYSTEM_CONFIG.runtime_auto_follow = OFF` で停止。
  - (c) 任意 `admin_auto_pull = ON`: 中央 project の `projects.get().updateTime` が記録値と変わっていれば (1) を自動実行（= clasp push を検知して Hub 取込）。既定 OFF。理由: WIP の clasp push も 5 分以内に本番 Hub へ入るため、運用判断が必要。

### 「中央 clasp push 後に自動で走らせる」手段の比較
| 手段 | 実現性 | 採用 |
|---|---|---|
| Hub 側 5分 sweep が中央の `updateTime` を監視して取込（上記 (c)） | 実装済み。追加コストは `projects.get` 1回/5分（内容は取らない） | **opt-in**（`admin_auto_pull=ON`） |
| Hub 側 sweep が Portal を Hub bundle に追従（(b)） | 実装済み。追加 API 呼び出しなし | **既定 ON** |
| 中央 project 側に時間トリガーを置いて Hub へ push | 中央→Hub の Apps Script API 呼び出しは可能だが、中央 project は LEGACY ブック bound で trigger owner の権限依存。(c) と同等効果で二重管理になる | 不採用 |
| `clasp run vNextAdminUpdateAllFromSource` | 中央 project を標準 GCP project に紐付け、API 実行可能 deployment ＋ OAuth クライアント設定が必要。かつ実行先は中央 project であり Hub bound project ではない | 不採用（提案のみ） |
| Apps Script API `projects.updateContent` を CI（GitHub Actions）から直接 Hub へ | Hub bound script ID へ直接 PUT。Hub 側の検証（allowlist / SHA / 監査）を迂回するため既存設計と矛盾 | 不採用 |

## 4. 版
- Admin runtime: 版番号は持たず SHA-256 で識別（中央 = Hub が一致していることをサイドバー「版・更新」で表示）。今回のコード変更で SHA は更新される（ダウングレードなし）。
- Portal runtime: 本ブランチでは Portal source を変更しない。並行ワーカーの Portal 1.7.38（`8e721fb`）を merge して中央へ push したため、中央は **1.7.38**（1.7.37 → 1.7.38、ダウングレードなし）。
- `/exec` ピン: `deploymentConfig.scriptId` 必須送信を維持。

## 5. まだ人手が要る手順
1. **今回1回だけ**: 本番 Hub はまだ旧コード。旧 UI「保守・高度な操作 → 中央配備版へ更新」（理由必須）→ 再読み込み。以後は新メニューが出る。申請入口は sweep が 5 分以内に 1.7.38 へ自動追従（待たない場合はメニュー「保守 → 最新版に更新」で即時）。
2. 初回の OAuth 同意（deployment 更新スコープ）が出た場合の許可。
3. `admin_auto_pull` を ON にするかの判断（VN_SYSTEM_CONFIG に `admin_auto_pull` = `ON` を書く）。
4. Web アプリ初回公開（済）。以降は同じ URL のまま自動で差し替え。
5. GAS 実機での動作確認（ローカルは Node 契約テストのみ）。
