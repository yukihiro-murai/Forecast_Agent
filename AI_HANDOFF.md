# Forecast vNext — AIエージェント引継ぎ

最終更新: 2026-08-18 JST（社員入口3層カード、Portal 1.7.0、Cursorから同じURLで入口を公開）  
対象ブランチ: `codex/vnext-annual-planning`  
この文書の目的: チャット履歴や端末固有メモリを使わず、GitHub上のリポジトリだけから安全に作業を再開できるようにする。

## 最初に行うこと

1. [`AGENTS.md`](./AGENTS.md) とワークスペースルートの `AGENTS.md` / `GIT_SYNC_RULES.md` を読む。
2. `git branch --show-current`、`git status -sb`、`git fetch origin --prune` を実行する。
3. この文書と [`README_VNEXT_JA.md`](./README_VNEXT_JA.md) を読む。
4. `origin/codex/vnext-annual-planning` と同期していることを確認する。分岐、未push、競合があれば勝手にmerge/rebase/stashしない。
5. 実装前に下記の「ライブ状態」と実際のHub状態を照合する。Sheets/GASの状態はGitだけでは更新されないため、記載値を盲信しない。

## 現在のコードと配備状態

IDは認証情報ではないが、公開資料へ転載しない。Git上のruntime版と、Hubで有効なACTIVE pairは一致しないことがある。

| 対象 | 現在値 |
|---|---|
| Git remote | `git@github.com:yukihiro-murai/Forecast_Agent.git` |
| コード上のClient runtime | `vnext-client-1.8.0` |
| コード上のClient bundle SHA-256 | `bc4e6f38e6bfedcd21a1e4d289a56123780887ca530f3b7a2678a04a9f5aa4f3` |
| コード上のPortal runtime | `vnext-portal-1.7.0` |
| コード上のPortal bundle SHA-256 | `331bd88d9c09545ed48c9c03db60fc91fcac27fabae5021681ea7e0c8c23d026` |
| Forecast Engine | `vnext-engine-0.5.0`（変更なし） |
| 中央clasp source Script ID | `1CkHthmMuU5r66ZpWJLw4bXrhNhDzcHCjBb2o1sFdIR1I0p1wNAao_erV` |
| Admin Hub Spreadsheet ID | `1baEZe6xYQ9KWyMMBk7kzH50v4dTtBPk9kWHK3qT7ID8` |
| Admin Hub bound Script ID | `1ANVtOBzPo90DveLcmTYokpEycr4iOphMaKZrhyjvzUizEQUjGqHRScxw` |
| 社員ポータル Spreadsheet ID | `16uiBgEF6Pz6N3UUpPU1DPJ6axelYRDJ3WuXFJUWy1dU` |
| ライブ ACTIVE Template / Model | まだ `vnext-client-1.7.0` / `model-portal-BEFDD31F149CE292A6CD`。Hubで新pair発行とポータル更新が必要 |

## 操作面（2026-08-17）

日常作業は右側の案内だけを辿る。上部メニュー「年度計画」は復旧用の「案内を開く」と、管理者の入れ子「その他」だけ。誰向け・いつ使うかの接頭辞は付けない。

簡易onOpenからは案内を出せない。案内が一度開いたあと、そのブックのproject onOpen triggerで以後は自動表示する。Hub はシート名だけでメニューを出し、案内の承認・例外カードを先に描画してからポータル／catalog／model詳細を後読みする。Apps Script の冷起動そのものは残る。

Hubの日常「申請を今すぐ処理」は案内の中。ポータルの作成フォームも同じ案内の中（最初は次の一歩、ボタンで作成へ）。

社員の共通入口はポータル runtime の Web アプリ（`doGet` / `Portal_Entry.html`）。2026-08-18に社内ドメイン向け `/exec` を公開済み。Portal 1.7.0 は入口を「年度予算の策定 / デジタルソリューション部」とし、部署ポータル・クライアントショートカット・管理者ハブの3層カードで開く。既存計画は年度ボタン→クライアントボタン。Hubの「社員ポータルを最新版へ更新」後に入口へ反映する。共有ドライブ名は AutoAnalysis と同じ並びの「年度計画」。

## ライブUAT対象: アストラゼネカ FY2027

- Spreadsheet: [Forecast アストラゼネカ FY2027](https://docs.google.com/spreadsheets/d/1bJwCUSn2GJ8uKZeBG1vyXMOX5fkl81t3h-hljTMlK7w/edit)
- 2026-08-16時点の状態: `SUBMITTED`、管理者承認待ち。
- 同じURLのままClient runtime 1.7.0へ更新済み。
- 提出済み計画、承認待ち、入力、既存の予測runは保持した。
- 保存済みrun ID: `RUN-68D50704E1BA8E3A73A5B87BB1EA9202`
- 保存済み中心予測: `¥185,642,622`（表示は1円単位、内部ログは元の数値も保持）
- 保存済みP10–P90: `¥137,153,921 ～ ¥261,460,828`

重要: この保存済みrunは旧Engine `vnext-engine-0.3.1`で計算されたimmutable recordである。Client 1.7.0へ画面を更新しても、過去runの分位点を後から書き換えてはいけない。Engine 0.5.0の新しい不確実性校正を実値で確認する場合は、承認前にAdmin Hubで差戻し経路「同じ入力で再予測」を選び、新しいrunとして実行する。ユーザーの明示判断なしに承認・差戻し・公式化を行わない。

## 直近変更の意図

### UI・可読性

- [`VNext_GuidanceSidebar.html`](./VNext_GuidanceSidebar.html) の予測カードを淡い背景＋濃い文字へ変更した。
- 保存ボタン用の汎用`.primary`が`<article class="metric primary">`にも当たっていたため、ボタンを`.primary-action`へ分離した。
- 本文15px、主要見出し26px、主要金額28pxを基準とし、派手な装飾や多カラムを避けた。
- 同じファイルを`client_runtime`へvendorし、[`VNext_ClientRuntimeBundle.js`](./VNext_ClientRuntimeBundle.js)を再生成した。

### 予測幅

[`VNext_Engine.js`](./VNext_Engine.js) で表示上のcapや任意倍率は使わず、生成過程を変更した。

- 年次変動は全期間の外れたregimeを混ぜず、直近5FYのdetrended log residualをMAD＋shrinkageで校正する。
- アストラゼネカの同じ年次実績では、年次sigmaが旧`0.261031...`から新`0.103498...`へ下がる純粋テストを追加した。
- UNKNOWN SPOTは12か月それぞれで大きなtailを引かず、年間総額を1回sampleして月へ配分する。
- AI/provider失敗はAI差分0の参考情報扱いとし、失敗だけを理由に売上分布を機械的に15%拡大しない。
- seed、FY/Q/月の加算整合、分位点、layer、input hash、versionは引き続きFORECAST_RUNへ保存する。

## 主要ファイル

| ファイル/ディレクトリ | 役割 |
|---|---|
| [`VNext_Admin.js`](./VNext_Admin.js) | Hub、provision、job、承認、release、AI調査、公式化 |
| [`VNext_Core.js`](./VNext_Core.js) | append-only schema、context、hash、状態イベント |
| [`VNext_Engine.js`](./VNext_Engine.js) | 実績読込、レイヤー、simulation、評価 |
| [`VNext_AI.js`](./VNext_AI.js) | AI調査・grounding・公開要約 |
| [`VNext_UX.js`](./VNext_UX.js) | Client FY Bookの表示・操作routing |
| [`VNext_GuidanceSidebar.html`](./VNext_GuidanceSidebar.html) | 1カラムの案内・予測dashboard |
| [`client_runtime/`](./client_runtime/) | Client専用最小runtimeの正本・build・tests |
| [`portal_runtime/`](./portal_runtime/) | 社員ポータル専用runtimeの正本・build・tests |
| [`VNext_ClientRuntimeProvisioning.js`](./VNext_ClientRuntimeProvisioning.js) | known-bound project作成・runtime copy/verify |
| [`tests/`](./tests/) | Node/V8契約・runtime copy・UAT回帰テスト |

`Forecast_Agent.js`はlegacy資産であり、vNextの新規employee flowを追加する場所ではない。vNext公開関数は原則`vNext` prefixを維持する。

## ローカル検証

最低限、変更範囲に応じて次を実行する。

```bash
node --check VNext_Admin.js
node --check VNext_Core.js
node --check VNext_Engine.js
node --check VNext_UX.js
node tests/vnext-integration.test.mjs
node tests/vnext-uat-feedback.test.mjs
node tests/vnext-client-runtime-copy.test.mjs
node tests/vnext-admin-runtime-copy.test.mjs
node tests/vnext-portal-pilot-recovery.test.mjs

cd client_runtime
npm run vendor
npm run build
npm run verify
npm test

cd ../portal_runtime
npm run check
```

Client/Portalのsourceを変更した場合はbundle再生成を省略しない。integration testのSHA不一致は、sourceと生成bundleのどちらかが古いことを示す。

## Cursorからの社員ポータル更新

Hub案内の「社員ポータルを最新版へ更新」と同じ検証経路（SHAピン、rollback、監査、同じURLの `/exec` 再公開）を、このマシンの Google ログインで実行する。計画の承認・差戻し・公式化は対象外。

Workspace は gcloud / clasp への Drive・Sheets 追加許可をブロックする。Cursor からの入口更新は clasp 標準スコープ（Apps Script ファイル更新と既存 `/exec` の差し替え）だけを使う。Hub 関数の直接実行は使わない。ポータル bound script ID は一度だけ `deploy/targets.json` の `portalScriptId` か `--portal-script-id` で渡す。

```bash
node scripts/gas_agent_login.mjs
node scripts/publish_employee_portal.mjs --portal-script-id <PortalのスクリプトID>
```

## Git・GAS反映

1. 対象ファイルだけ`git add`する。
2. commit後、`git pull --ff-only`、`git push origin codex/vnext-annual-planning`。
3. `clasp status`で中央projectの対象を確認して`clasp push`。
4. AdminコードをHubへ反映する場合は、Hubの管理画面「中央配備版へ更新」を使うか、緊急bridgeとしてHub bound Script IDへ同一18ファイルを直接pushする。対象Script IDとparent Spreadsheetを再確認する。
5. Client/Engine/UI変更は中央/Hubへpushしただけでは既存Clientへ反映されない。新しいimmutable Template＋Model pairを発行・有効化し、既存bookは状態別の専用same-URL upgradeを使う。

`clasp push`だけで「ライブ反映完了」と判断しない。中央source、Hub bound project、ACTIVE pair、対象Client pinの4層を確認する。

## 変更してはいけない境界

- 保存済みFORECAST_RUN、PLAN_VERSION、OFFICIAL_RUNを上書きしない。
- 採用判断、営業上積み、最終予算をモデル学習へ戻さない。
- Client runtimeへHub ID、ZAC source ID、Vertex設定、raw prompt、Admin処理を入れない。
- Portal runtimeからZAC/API/Hubを直接読まない。Adminがcatalogを投影し、Hub側でrequestを再検証する。
- Client hidden sheetは未信頼staging。Hub正本へ無検証で同期しない。
- P10/P90を見栄え目的でclip/capしない。狭める場合はデータ生成、校正、event構造を変え、理由とtestを残す。
- 既存Clientの状態を変える承認・差戻し・再予測・公式化は、ユーザーの明示判断なしに実行しない。
- `.clasp.json`のtargetを確認せずpushしない。`--force` pushは禁止。

## 社員テストをゼロからやり直す

ユーザーが「管理シートだけの状態に戻す」と明示した場合の専用操作。確認語 `RESET_GENERATED_CLIENTS` と `apply:true` が揃うまで Drive 削除も Hub 書込もしない。

残すもの: Admin Hub、社員ポータル、Template/Model release、ZACクライアント候補。
消すもの: 生成済み Client FY Book（ゴミ箱へ移動）、予測/申請/承認/job の試験データ、`ADMIN_AUDIT_LOG`、Drive監査フォルダのAI JSON等、ポータルの依頼・directory 投影。リセット直後に `FRESH_UAT_RESET` の監査行が1件だけ残る。

実行手順（Admin Hubを再読み込みしてから）:

1. メニュー「年度計画」で案内を開く。
2. 「保守・高度な操作」の最下部「社員テストをゼロからやり直す」。
3. 「対象を確認（変更なし）」で Client 冊数を見る。
4. 確認語 `RESET_GENERATED_CLIENTS` を入力し、「生成済み年度計画を削除して初期状態へ戻す」。

## ユーザーが今やる設定順

リセットを先に行い、検証ブックを共有ドライブへコピーしない。

1. [Admin Hub](https://docs.google.com/spreadsheets/d/1baEZe6xYQ9KWyMMBk7kzH50v4dTtBPk9kWHK3qT7ID8/edit) を開く。案内が出なければメニュー「年度計画 → 案内を開く」。
2. 「保守・高度な操作」→「中央配備版へ更新」（理由必須）→ 再読み込み。
3. 「社員テストをゼロからやり直す」。確認語 `RESET_GENERATED_CLIENTS`。
4. 「共有ドライブ『年度計画』へ移す」（理由必須）→ 再読み込み。
5. 「社員ポータルの設定・更新」→「既存ポータルを最新版へ更新」（理由必須）。
6. ポータル bound script から Web アプリを1回公開（実行者=アクセスしているユーザー、アクセス=ドメイン）。そのURLが社員入口。

Apps Script の実行履歴は Google 側の記録のため、このリセットでは消えない。HubシートとDrive上の検証ファイルは手順3で消える。

## 次の安全な作業候補

1. Admin Hub を再読み込みし、上部メニューと案内の初回表示が以前の約10秒より短いことを確認する。
2. 案内の「社員ポータルを最新版へ更新」を押し、ブックマーク済み Web 入口で年度→クライアントボタンを確認する。
3. Client 1.8.0 の新しい Template/Model pair を発行・有効化してから、ポータルでゼロから年度×クライアントを作る。
4. ユーザーが希望した場合のみ、新しいrunで Engine 0.5.0 の P10/P90 を確認する。
5. `main`へ統合する場合は、現ブランチのライブ配備状態とGitHub差分を独立レビューしてから行う。

## 終了時に更新する項目

作業を引き継ぐ場合、この文書の次を必ず更新してcommit/pushする。

- 最終更新日時、ブランチ、commit
- Central/Hubへclasp反映したか
- ACTIVE Template/Model、Client/Portal/Engine versionとSHA
- ライブUAT対象のstateと、実行した外部操作
- PASS/FAILしたtest
- 未解決事項と、次の具体的な1操作
- Cursor が社員入口を更新する場合は `node scripts/gas_agent_login.mjs` 済みか
