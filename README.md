# Forecast Agent / Forecast vNext

詳細な構成・データ契約・運用手順は [`README_VNEXT_JA.md`](./README_VNEXT_JA.md) を参照してください。

別端末・別AIエージェントで現在の作業を引き継ぐ場合は、最初に
[`AI_HANDOFF.md`](./AI_HANDOFF.md) を参照してください。現在のライブ版、配備先、
アストラゼネカFY2027の状態、テスト・反映手順、変更禁止境界をまとめています。

## 初期版の信頼境界

Forecast vNext の初期版は、Admin Hub とクライアント年度 book を物理的に分離します。Hub は管理者だけに共有し、従業員には担当クライアントの年度 book だけを共有します。Vertex AI の設定、raw prompt、raw AI evidence、他クライアントの情報はクライアント book に保存しません。

ただし、現行の Sheets + bound Apps Script 方式では、クライアント book のスクリプトは実行ユーザー権限で動きます。そのため、従業員の正規入力を許可しながら hidden sheet を暗号学的に改ざん不能にすることはできません。クライアント内の append-only log は「未信頼のステージング」であり、Hub の正本ではありません。

Admin 所有 trigger は Hub へ取り込む前に、次をサーバー側で再検証します。

- book / client / FY と BOOK_REGISTRY の一致
- Forecast Owner、TEAM、入力者、依頼日時、状態遷移の整合
- evidence type の allowlist と、人間入力への AI metadata 混入禁止
- forecast request の hash、時刻、cutoff 再計算
- submitted plan と SUCCESS forecast の対応、金額算術、理由、12か月配分

違反レコードは Hub へ取り込まず、永続的な例外として管理者へ表示します。この検証は誤操作・通常の改ざんリスクを抑えますが、署名付きの信頼境界そのものではありません。

高保証運用へ移行する場合は、`executeAs=USER_DEPLOYING` の Admin 所有 Web App を唯一の書込口にし、OAuth で呼出者を確認したうえで Hub に直接 append してください。その段階ではクライアントからの local fallback を廃止し、通信失敗時は fail-closed とします。

## 初回運用

1. Legacy book の所有者が Forecast vNext の初期設定を開き、ZAC 実績元 Spreadsheet ID を入力します。
2. private な空Spreadsheetへ、中央配備済みのAdmin専用runtimeとClient専用runtimeをそれぞれbindingし、Admin Hubとimmutable Master Templateを生成します。元のLegacy bookやそのbound scriptはコピーしません。
3. 生成した Admin Hub を開き、「自動運用を有効化」を1回実行します。
4. 社員ポータルでFY、ZACクライアント、関与メンバー氏名を指定して年度 book の作成を依頼します。Forecast Ownerは依頼した社内ユーザーへ自動設定されます。

Adminコードを改修した後は、中央projectへclasp反映してからAdmin Hubの「中央配備版へ更新」を実行します。更新対象は検証済み21ファイルだけで、Hubの履歴・設定・正式計画は置換しません。Client UI/MEMOの改修は管理者限定Template Draftから新しいimmutable Template Releaseとして公開され、既存年度bookは原則固定されます。汎用migration APPLYは停止したままです。従業員テスト前で回答・依頼・予測・計画等が完全に0件、source pinsとruntime SHAが一致する空のPilot Clientだけは、read-only事前判定後に同じURLのままcanonical ACTIVE pairへ更新でき、途中停止時は専用journalから復旧します。

初回展開は2～3 Client、明示承認後のcanaryでも最大5 Clientです。6冊目は30冊負荷試験とrelease承認が完了するまでserver-sideで拒否します。Client fileはAdmin管理のprivate root配下にだけ生成し、共有境界を確認できないfolderは使用しません。

クライアント年度 book の通常表示は `1_ホーム` と `2_予測と計画` の2シートです。内部処理、承認履歴、raw AI metadata は従業員向け画面には表示しません。

## 売上予測 Webアプリ（Legacy book の doGet）

「クライアント別売上予測」スプレッドシートの bound script を `executeAs=USER_DEPLOYING` の Webアプリとして公開しています。**2026-10-01 から公開範囲は `access=MYSELF`（所有者のみ）**です（段階0。理由と記録は下記）。Legacy メニュー（A-2 取り込み → A-3 加工 → A-4 AI調査 → A-5〜A-8 主観入力 → A-9 予測 → A-10 予算入力 → B-1〜B-5 検証・自動学習 → C-1〜C-3 四半期レビュー）を、スプレッドシートを開かずにブラウザだけで一巡できる SPA です。

B-5（月次ベイズ自動学習 + Vertexアシスト）により、B-1/B-2 で実績評価が蓄積されるたび補正係数が自動更新され、次回 A-9 予測へ反映されます。設計は [`DESIGN_bayesian_autolearn_vertex_JA.md`](./DESIGN_bayesian_autolearn_vertex_JA.md) を参照。

- 実装: [`Forecast_WebApp.js`](./Forecast_WebApp.js)（`doGet` + `webGetBootstrap` + `webRun*/webSave*` RPC）と [`Forecast_WebAppUI.html`](./Forecast_WebAppUI.html)（SPA）
- 反映: `clasp push` 後に `clasp deploy --deploymentId <id>` で exec URL の版を上げる（`/dev` は Google ログイン必須のため運用には使わない）
- 運用上の注意と検証手順は [`AI_HANDOFF.md`](./AI_HANDOFF.md) の「Webアプリ」を参照
- **公開範囲を所有者のみにした理由（2026-10-01）:** Apps Script の Web アプリでは、名前の末尾が `_` でない関数（このプロジェクトで178個。初期化 `setupForecastBook`、全シート削除 `adminSetupGuideOnly`、vNext の管理機能を含む）を、画面に無くてもブラウザから `google.script.run` で呼べる。`USER_DEPLOYING` かつ社内ドメイン公開のままでは、社内の誰でもデプロイした人の権限でそれらを実行できたため。社内の他の人が使うのは、新しいアプリ（[`DESIGN_data_platform_JA.md`](./DESIGN_data_platform_JA.md)）ができてから。
- **操作の記録:** Web からの書き込み・実行（17関数）は `webAudited_` を通し、年度ごとのログ用スプレッドシート「売上予測 ログ FYyyyy」の月別シート `AUDIT_yyyy_MM` に、開始と終了（OK / NEEDS_CONFIRM / FAILED / DENIED）、操作した人、変更前後の値を残す。各行は前の行のハッシュとつなぐ（改ざん・削除の検知用）。開始を記録できなければ処理しない。ログのファイルの ID は Script Property `FORECAST_LOG_SPREADSHEETS_JSON`、最新のハッシュは `FORECAST_AUDIT_LAST_HASH`。所有者以外の管理者を足すときは `FORECAST_WEB_ADMIN_EMAILS`（カンマ区切り）。画面上部の「操作の記録」から開ける（管理者のみ）。スプレッドシートのメニューからの操作は記録しない。
