# Forecast vNext — 年度売上予測・予算策定システム

## 目的

Forecast vNext は、未来の売上を一点で「当てる」ものではなく、情報締切時点で利用できた確定実績・案件・明示的な前提から、クライアント別FY売上の条件付き分布を作るシステムです。

従業員向けの操作と管理ハブ担当者向けの設定を分離し、次の数値を別々のレコードとして保持します。

1. 過去実績からの基準
2. 未把握の単発売上
3. 確度の高い案件・契約
4. 類似企業の参考値
5. 公開情報から確認できる変化
6. 担当者が把握する変化
7. AI調査による変化
8. システム推奨予測
9. 採用判断差分・採用予測
10. 営業上積み・最終予算

採用判断差分、営業上積み、最終予算はモデル学習へ戻しません。

## 今すぐやること（初期設定）

管理ハブ を開き、案内の「最初にやること」を上から1つずつ実行します。各操作のあと画面を再読み込みしてください。

1. **中央配備版へ更新**（理由必須）
2. **受入試験をゼロからやり直す**（確認語 `RESET_GENERATED_CLIENTS`）。検証用ブックと試験ログを消します
3. **共有ドライブ「年度予算策定」へ移す**（理由必須）
4. **既存ポータルを最新版へ更新**（理由必須）
5. 申請入口 Spreadsheet のスクリプトから **Web アプリを社内向けに公開**。その URL が **年度予算策定 Web入口** です

案内が出ないときは、上部メニュー「年度予算策定 → 案内を開く」を使います。

## 物理構成（3層・フォルダ番号順）

- **共有ドライブ「年度予算策定」**: AutoAnalysis / PulseCycle と同じ社内共有ドライブ。既存の「年度計画」名も移設時に引き継ぎます。
- **01_管理ハブ（第1層）**: 承認・例外・release・監査。配下の `監査` に AI raw と凍結 backup。
- **02_申請入口（第2層）**: 全社員が client×FY を探し、未作成分を申請する Spreadsheet。年度予算策定 Web入口からもここへ進みます。
- **03_クライアント年度ブック / FYyyyy / クライアント名（第3層）**: 1 client × 1 FY の予 작業ブック。
- **04_テンプレート**: `現行` / `下書き` / `履歴`。
- **Legacy book**: 現行運用を維持する参照元。vNext 初期化では変更しません。
- **Master Template**: 実クライアントデータを持たない immutable な クライアント年度ブック 生成 release。
- **クライアント年度ブック**: 1 client × 1 FY。従業員入力、予測・予算、振り返りだけを表示。

管理情報を単に隠すのではなく、クライアント年度ブックへ置かないことが基本方針です。クライアント年度ブックの非表示シートは、同一クライアント内の誤編集と認知負荷を抑えるためだけに使います。

管理ハブとMaster Templateはいずれもlegacy bookのコピーではありません。bootstrap元の中央projectが空のSpreadsheetを作り、Apps Script APIで管理ハブ用の完全runtimeまたは[`client_runtime`](./client_runtime/README.md)の最小runtimeだけをbindingします。生成Hubにはsource/target script IDとcanonical runtime SHA-256を管理ハブ担当者専用設定として記録し、以後は管理ハブサイドバーの1操作で中央clasp配備版へ更新できます。したがってクライアント年度ブックのApps Script projectには、Admin処理、ZAC source ID、Vertex設定、AI raw、prompt、予測エンジンを含めません。Client側のOAuth scopeも、現在のブック、UI、利用者メールだけに限定します。

申請入口も空 Spreadsheet へ [`portal_runtime`](./portal_runtime/README.md) だけを binding します。通常表示は `ホーム` と `FYyyyy` タブで、Client 名、状態、次の対応、クライアント年度ブックリンクを管理ハブ正本から投影します。

公開済みMaster Templateは手編集するworking sheetではなく、immutable release artifactです。クライアント年度ブック画面、MEMO、書式をシート上で改訂する場合は、管理ハブから管理ハブ担当者限定の`TEMPLATE_DRAFT`を作り、その表示3シートだけを編集します。公開操作はdraft（code-only更新時は現ACTIVEの表示3シート）を新しいclean Templateへコピーし、Client runtime・表示内容・版の組を検証してからpointerを切り替えます。表示3シートのvalues/formulas/notes/format/validation等はcanonical hashへ結び付けられ、公開後の直接編集を検出した場合は新規Client生成を停止します。既存Clientはreleaseに固定されます。

Client側に残る追記型イベントは、初期版では**未信頼のステージング**です。管理ハブ所有の定期workerがクライアント年度ブック／client／FY、担当者、許可イベント種別、状態遷移、金額恒等式、時刻とhashを再検証し、受理した記録だけを管理ハブの横断監査正本へ追記します。違反は取り込まず、永続例外として管理ハブ担当者へ表示します。クライアント年度ブック向けに返すAI情報は上位5件までの公開要約・調査観点・人が確認する問い・URL・出典日・適用額だけで、raw responseやprompt metadataは管理ハブ監査 Folderから出しません。

この境界は、通常の誤操作と直接編集を検出・拒否するための初期実装であり、暗号学的な署名境界ではありません。より高い保証が必要になった時点で、管理ハブ担当者権限で実行するWeb Appを唯一の書込口とし、OAuthで呼出者を確認してHubへ直接appendする構成へ移行します。Client側のlocal fallbackは設けず、通信失敗時はfail-closedとします。

## 操作面の設計

日常の作業は、ブックを開いたときに右側へ出る案内に従うだけで完了するようにします。シートは見るための画面であり、セルを直接編集しません。

上部の「年度予算策定」メニューは復旧と、普段使わない操作のための控えです。案内が出ないときだけ「案内を開く」を使います。日常の入力・作成・承認は案内の中で順番に進め、メニューと案内を往復しません。

Google の制約で、簡易な onOpen からは案内を出せません。案内が一度開いたタイミングで、そのブックのプロジェクトに自動表示用の onOpen を1つ付けます。以後は開くだけで案内が出ます。

## クライアント年度ブックの利用フロー

通常表示は `1_ホーム` と `2_予測と計画` だけです。実績評価が始まると `3_振り返り` を表示します。

1. 従業員が「変化あり／確認したが変化なし／わからない」を回答する。
2. 予算策定担当が回答状況を確認し、予測を依頼する。
3. vNextがドラフトを作り、過去実績からAI調査まで7層の積み上げと振れ幅を表示する。
4. 予算策定担当が採用判断と営業上積みを別々に入力して提出する。
5. 管理ハブが管理ハブから承認または差戻しする。
6. 承認runを公式vintageとして凍結する。
7. 実績確定後は公式runだけを評価し、次年度の情報収集とモデルreleaseへ反映する。

ホームには状態に応じた次の操作を1つだけ表示します。操作は右側の案内から行い、シート上のセルを直接編集しません。処理待ち状態では操作を増やさず、通常の待ち時間と、停滞時に管理ハブ担当者へ伝える内容を表示します。

クライアント年度ブックと申請入口は、会社Workspaceドメイン内のリンク共有を選択できます。これは「誰が予算策定担当か」とは別の概念です。社内ユーザーは担当登録がなくても閲覧と情報提供ができ、予算策定担当だけが予測依頼・予算案提出を行います。初期版で関与メンバーとして保存する氏名は、関与状況を見やすくする表示情報であり、メール本人確認や回答分母には使いません。メールで同定できる予算策定担当を必須回答者とし、社内の自由参加者の未回答で処理を停止しません。社員名簿とメールの安全な対応表を導入した段階で、関与メンバーを回答分母へ追加できます。管理ハブ、Template、監査folderは引き続きprivateです。

AI調査は、予測差分を作るだけでなく、担当者が公開情報からクライアントの変化を発見するための補助レイヤーです。調査観点は、業績・投資余力、DXの実行度、製品・市場の勢い、戦略・組織変化、規制・供給・調達、採用・特許・臨床試験・認証・設備投資・提携・公開頻度などの先行シグナルです。各調査結果には「予測へ反映」または「参考情報」を明示し、担当者が次に確認する問いと出典URLを付けます。一次情報または公的登録情報で、出典強度・売上関連性・確度の基準をすべて満たす方向性のある情報だけを金額差分へ反映し、それ以外は0円の参考情報として残します。

Vertex/providerが一時的に失敗した場合や、引用URLを検証できない場合は、失敗を管理ハブ例外とrun監査へ残し、AI差分を0として継続性・案件・現場情報で予測を完了し、情報不足分だけ通常の振れ幅を広げます。grounding metadataの複数形式を解釈し、それでも引用がない場合は設定済みの検索データストアを代替経路として使用します。未引用のAI文章を金額へ反映することはありません。AI取消の比較runは、元runが保存した有効evidence ID集合、非AI入力hash、model/version、as-of、seed、simulation数を再検証します。非AI入力が変わっていればAIだけの反実仮想とは呼ばず、通常の新runを要求します。

## 実績データ契約

- 実績ソースIDは `FORECAST_SOURCE_SPREADSHEET_ID` で管理します。
- クライアント列: AO
- サービスカテゴリ列: AT
- 製品列: AX
- 実績日列: BE
- 金額列: BN
- BDの売上予定日は実績日fallbackとして使用しません。
- `as_of`の前月末を超える行は読み込みません。
- 対象履歴は利用可能な5〜8年度です。

## 版と監査

公式runは、少なくとも次を保存します。

- `run_id`, `issued_at`, `as_of`, `data_cutoff`, `seed`
- `template_version`, `schema_version`, `model_version`, `prompt_version`
- 正規化済み実績・前提・AI event・有効設定のSHA-256
- FY、Q、月次のP10/P50/P90と各差分レイヤー
- 提出者、承認者、状態遷移、変更理由

公式runは上書きしません。訂正はamendmentとして新しい版を追加します。

## 管理ハブの日常運用

- 通常時は「今日、人が判断する必要がある例外」だけを確認します。
- 予測依頼、AI調査、予測計算、Clientへの結果返却は5分triggerで非同期処理します。
- timeoutしたjobにはlease期限と再試行上限を設け、失敗確定時はClientを操作可能な状態へ戻します。
- 承認時はHubにあるSUCCESS runとSUBMITTED planからsnapshotを再構築し、組合せを検証してから公式vintageを凍結します。
- 正式予算の訂正は、現在の公式vintageを参照するamendmentとしてだけ発行します。
- 実績評価は、対象FYの確定実績と現在の公式vintageを検証してから生成します。
- 管理ハブ runtime改修は中央clasp projectへpush後、管理ハブの「中央配備版へ更新」で反映します。source/targetの同一ID、target parent、18ファイルallowlist、V8 manifest、書込後SHA-256を検証します。
- Client runtime/UI改修は現行Templateを上書きしません。管理ハブ担当者限定Draft（code-only更新は現ACTIVE UI）から新しいprivate `STAGED` Templateを作り、そのrelease IDへ厳密に結び付いたPASS済みModel candidateとの組だけを有効化します。canonical pair pointerをCASで切り替え、property cache更新後にだけ旧TemplateをRETIREDへ移します。各phaseは追記型journalへ残るため、中断後も同じoperation IDから再開できます。
- クライアント年度ブックは管理ハブ管理のprivate root配下へだけ生成します。任意の共有folderや、共有境界を証明できない保存先はfail-closedで拒否します。

## 初期releaseの適応範囲と展開ゲート

- 自動更新するのは不確実性、案件確率・月ずれ、季節配分、誤差校正などの状態です。モデル構造や係数releaseはbacktest/canaryのcandidate hashが一致し、管理ハブが有効化したものだけを使います。
- 参照クラスpriorは、比較可能なcohortと十分な完了年度がそろうまで`DISABLED`です。全Clientへ同じ絶対金額priorを流用することは拒否します。初期pilotでは、継続性・案件・明示的な客観/現場変化の三者を意味の異なる根拠として保持し、参照クラスはデータ蓄積後のreleaseで有効化します。
- 最初は2～3 Clientを現行運用と並行runし、操作完了率、所要時間、入力形骸化、区間校正、job所要時間を測定します。管理ハブがcanary開始を明示承認した場合だけ4～5冊へ進め、6冊目は負荷試験releaseまでserver-sideで拒否します。30冊展開には、30 Client×AI＋forecast、quota、timeout、lease crashを含む負荷試験とSLO承認が必要です。
- pilot releaseでは汎用Client migrationのAPPLYをserver-sideで停止し、dry-runだけを許可します。例外として、受入試験前で`ACTIVE / INPUT_OPEN`、回答・依頼・予測・予算案・承認・公式・評価がすべて0件、source release/model/runtime SHAと管理ハブ／クライアント年度ブック metadataが完全一致するクライアント年度ブックに限り、管理ハブサイドバーの「テスト前のクライアント年度ブックを最新版へ更新」を使用できます。この専用処理は最初にread-only判定を行い、同一schemaのcanonical ACTIVE pairへ同じSpreadsheet URLのまま更新し、registryを最後にcommitします。中断時はmigration journalから旧版または新版の整合状態へ復旧します。公式・振り返り中・終了済みクライアント年度ブックのin-place更新は行いません。
- 初期MODEL_RELEASEのbacktest/canary PASSは管理ハブ担当者 attestationです。自動評価済みと称しません。本番releaseではdataset snapshot、candidate/code hash、metric、threshold、実行者・時刻を持つsystem生成artifactへ置き換えます。
- Sheets内のClient stagingは署名境界ではありません。30冊本番または監査保証を強める段階では、管理ハブ所有Web Appを唯一のwrite endpointにする移行をrelease gateとします。

## 開発・反映

このリポジトリはclasp管理です。変更後はリポジトリ規約に従い、GitHubへpushしてから次の順でGASへ反映します。

```text
clasp status
clasp push
```

新年度のクライアント年度ブックは、最新ACTIVEのimmutable Master Template Releaseから管理ハブのプロビジョニング機能で生成します。選択Templateとactive Model ReleaseのEngine/Core/schema互換性が一致しなければfail-closedです。従業員にApps Script Editor操作や初期セットアップを要求しません。

初回bootstrap前に、管理ハブ projectでApps Script APIを有効化し、`FORECAST_SOURCE_SPREADSHEET_ID`、管理ハブ担当者メール、必要なVertex設定をScript Propertiesへ保存します。管理ハブ側manifestはTemplate生成のため`script.projects`を持ちますが、Client runtimeにはこのscopeを配布しません。
