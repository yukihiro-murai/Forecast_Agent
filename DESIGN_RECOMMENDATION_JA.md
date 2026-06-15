# 売上予測スクリプト 設計書 v2.3.4（2.3.30〜2.3.31 同期 / doc-sync）

## 0. 文書の目的
この設計書は、実装（Forecast_Agent.js）の現行挙動を正確に記述する。
旧v7の「三角観測 w1/w2/w3/w4 + 逆sMAPE重み更新」は実装に存在しないため撤去済み。
3目的は不変：(1)予測精度向上 (2)透明化（根拠明示・再現性） (3)学習性（継続改善）。

対象実装：`Forecast_Agent.js`（VERSION='2.3.31-dev' / BUILD_STAGE='raw-migrate-header-match'）
設計書版：v2.3.4（v2.3.3 からのドキュメント改訂。output-display-tidy（2.3.30）は表示専用、
raw-migrate-header-match（2.3.31）は生ログ移送ロジックのみの変更で、いずれも予測の P10/P50/P90 は不変。
本改訂は「同梱の raw-migrate-header-match プロンプトを適用済み」の状態を記述する＝
**コードを先に適用・検証してから本書へ差し替える**）。
DLM_BUILD_STAGE は 'v8-step3c3c-1'（DLMロジック無変更のため据え置き）。

### 0.0 v2.3.3 → v2.3.4 の更新点（この改訂で反映したもの）
本改訂は2系統の実装変更をドキュメントへ同期する。各々は適用済み（raw-migrate-header-match は同梱プロンプトで
適用する）であり、本改訂で新たな別系統のコード変更は行わない。

- **output-display-tidy（2.3.30）**：表示専用の整形。
  ① 初期設定ダイアログ冒頭の黄色 notice（および `.notice` CSS）を撤去（このタイミングでの表示は遅く、他箇所でカバー）。
  ② OUTPUT 6行目の注記ブロックの赤字判定を「実警告（`⚠` / 製品重み警告 / `web_error`・`rag_error`・`structure_error` が
  1以上）があるときのみ赤」へ変更。`vertex_rows=…; web_error=0;…` の all-zero 情報サマリーだけでは赤くしない
  （`step3aHasError` で error≥1 を検出）。
  ③ row10「定量/主観/AI/KnownSpot 100%分解」テキストを E列→F列へ移動（E10は空・NOTEもF10）。
  ④ F9/F10 の長文を折り返し＋行高調整ではみ出し解消。
  ⑤ 21行目 coverage を topic ごとの改行表示＋A〜F幅（6列）へ縮小（行高72）、22行目（degraded / benchmark不足）・
  23行目（中立化バナー）も折り返し・幅縮小で横伸びを抑制。
  **いずれも表示位置/書式のみで数値は不変。** 変更スコープは `showInitialSetupDialog_` / `writeOutputFY_` の2関数のみ。
- **raw-migrate-header-match（2.3.31）**：`migrateLegacyAIResearchRawSheets_` の旧2枚→RAW 移送を
  **位置ベースから「ヘッダ名マッチ＋シート名由来の axis 確定」へ変更**（§10.6・§12）。旧 WEB/EXTERNAL の列順が
  RAW と一致していなくても、`client/as_of_date/topic/evidence/note` は同名ヘッダの列へ入り、`axis` は
  シート名（AI_RESEARCH_WEB→web / AI_RESEARCH_EXTERNAL→rag）から確定する。旧2枚に axis 列が無くても空にならず、
  frozen_flag は旧側に無ければ 0・frozen_at は空で補完。これにより §12 の latent issue（移送時の axis 1列ずれ）を
  **解消**。生ログ移送のみで予測の P10/P50/P90 は不変、冪等性も不変。

### 0.0' v2.3.2 → v2.3.3 の更新点（参考・再掲）
本改訂は **8系統・複数ビルドの実装変更**をドキュメントへ同期した。各々は適用済みであり、新たなコード変更は行っていない。

- **snapshot-vestigial-removal（2.3.12）**：FORECAST_SNAPSHOT の三角測量系 vestigial カラム
  （`linear_pred` / `robust_pred` / `regime_pred` / `simulation_pred` / `w1`〜`w4`）を物理削除。
  現行ヘッダは15列、`final_pred` は J列（index 9）に確定。B-2 の突合は `r[9]` を `final_pred` として読む。
- **config-simplify（2.3.13）**：CONFIG の「[互換] 担当者（B10 = B4参照）」を撤去。担当者は B4 のみが
  単一の真実源。チューニング表から `QUAL_SUBJECTIVE_MAX_SCALE` 行を撤去（cap pass-through で不要）。
- **config-deadknob-removal（2.3.15）**：未使用CONFIG行を撤去
  （`QUAL_SHARE_ALERT/TARGET_*` / `QUARTERLY_REVIEW_PERIOD_MONTHS` / `BIAS_CORRECTION_*` /
  `AI_DIRECTION_HIT_*` / `AI_EFFECT_MIN_MEANINGFUL` / `DLM_FORECAST_HORIZON`）。予測コアは不変。
- **rag-config-defaults（2.3.16）**：CONFIG に RAG 既定値を投入
  （`VERTEX_DATASTORE_ID=fujikeizai-portfolio-2025` / `VERTEX_SEARCH_LOCATION=global` /
  `VERTEX_SERVING_CONFIG=default_search`）。Vertex/RAG 環境行は黄色＋テキスト書式（番地非依存・キー照合）。
  `discoveryEngineHost_` を地域エンドポイント対応化（global は従来同一ホスト）。
- **ai-dx-confidence-diagnostics（2.3.21）**：confidence 欠落で event 候補が重み0となり不採用になった件数を
  `confDrop` として OUTPUT の coverage 行・degraded 警告へ可視化。`Number('')→0` を避けるため `isFinite(conf)`
  での欠落検出を明示化。採点（P10/P50/P90）は不変の純診断追加。
- **output-note-relocate（2.3.22）**：年度合計行と月次表ヘッダの間にあった長い「年度≠月次合算」説明ブロックを
  短い1行注記（`OUTPUT_RANGE_EXPLAIN_PRIMARY_SHORT_TEXT`）に置換し、長文本文（`OUTPUT_RANGE_EXPLAIN_MAIN_TEXT`）は
  混合セクションの Scenario Split の下（目立たない位置）へ移設。表示位置のみの変更で数値は不変。
- **ai-score-robustness（2.3.23）**：event_score の算出式を是正（§10.3）。旧 `sign × |impact-50| × conf` を
  `direction符号 × (impact/100) × 50 × conf` に変更。impact_score は 0〜100 の「影響の大きさ」（50は中立でない）。
  併せて `AI_MISSING_CONFIDENCE_DEFAULT`（既定0.5）を導入し、confidence/relative_confidence 欠落を
  既定値で補完して採用（0なら従来どおり不採用）。`confDefault` 件数を診断計上。
- **ai-score-degroundless（2.3.24）**：AI寄与の「根拠なき中立化／捏造ゼロ」を是正（§10.3・§10.4）。
  ① `AI_TOTAL_NEUTRAL_THRESHOLD` 既定を 0 にし dead-zone を撤去（上限は ±3% キャップのみで担保）。
  ② event_only への構造的 ×0.5 ペナルティを撤去（信号の弱さは coverage 由来 qualityMultiplier のみで反映）。
  ③ benchmark 行は相対位置の根拠があるときだけ書く（50/middle のプレースホルダ禁止）。
  ④ 構造化抽出は temperature=0 に固定（同一web/RAG入力 → 同一構造化結果）。
  ⑤ momentum の lookback を `as_of_date`（調査リフレッシュ）単位で dedup（A-9実行回数では動かない）。
- **calibration-blank-override-fix（2.3.25）**：`applyCalibrationToTuning_` の override 空欄判定を是正。
  `CALIBRATION_STATE` の `ai_weight_override` / `ai_max_abs_effect_override` が空欄のとき `Number('')→0` が
  finite として通り `aiWeight=0`／`aiMaxAbsEffect=0` に化け kAI が常時 1.0 に潰れていた欠陥を修正。
  空欄(=未設定)と明示0(=意図的ゼロ)を区別する（§4.6）。
- **rag-query-frame-align（2.3.26）**：`buildRagQuery_` を AI調査の新フレーム（外部観測可能な「支援需要に
  影響しうる外部環境」）に整合。topic 別の日本語展開語（市場規模/競合/MR/DX 等）で外部環境を狙い、
  英語topic素語の連結や出版社名（弁別力なし）を撤去。A-4側のクエリのみの変更。
- **qual-share-const-removal（2.3.27）**：未使用 const `QUAL_SHARE_ALERT_THRESHOLD` /
  `QUAL_SHARE_TARGET_CENTER/LOW/HIGH` をコードから撤去（デッドコード削除・予測不変）。
  `QUAL_SUBJECTIVE_MONTHLY_CAP` / `QUAL_CALIBRATION_ENABLED` / `SUBJECTIVE_OVERLAY_TARGET_*` は残存。
- **orphan-fn-removal（2.3.28）**：孤立関数 2件（`adminShowCloneGuide` / `importPastSalesToSalesTab`）を
  grep でゼロ依存確認のうえ物理削除。book複製手順の案内は当時 `showInitialSetupDialog_` の notice に残存していたが、
  output-display-tidy（2.3.30）でこの notice は撤去された（§0.0 参照）。
- **ai-research-raw-merge（2.3.29）**：生ログシート `AI_RESEARCH_WEB` / `AI_RESEARCH_EXTERNAL` を
  単一の `AI_RESEARCH_RAW` に統合（§6・§10.5）。web/rag の区別は `axis` 列で持つ。既存bookの旧2枚は
  `migrateLegacyAIResearchRawSheets_` が一度だけ RAW へ移送し旧タブを削除する（冪等）。
  ※移送ロジックは raw-migrate-header-match（2.3.31）でヘッダ名マッチへ是正された（§10.6）。

### 0.0'' v2.3.1 → v2.3.2 の更新点（参考・再掲）
- **objonly-dealias**：`runForecastFYCore_` の `objOnly` を `quantOnly` の独立コピー（`.slice()`）化。
  `closedMonthMode='actual'` の実績上書きが `quantOnly`（KPI診断）/ `opsQuantOnly`（DLM比較の旧Ops参照）を
  巻き込み変異させない配線になった。現行データフローでは actual_closed 月が立たないため挙動中立であり、
  予測コア（OUTPUTのP10/P50/P90）は不変。
- **gem-path-removal**：旧Gem手動貼付経路の関数群を物理削除。A-4 は `runVertexAIResearch`、
  A-9 は `countAIResearchStructuredRows_` で件数把握する経路に確定。

### 0.0''' v2.3 で反映済みの更新点（参考・再掲）
- **AI調査をVertex AI自動実行へ移行**。旧A-4「Gemプロンプト生成→TSV手動貼付→parse」は廃止し、
  A-4は `runVertexAIResearch`（Vertex grounded web検索 + Vertex AI Search RAG + 構造化出力）へ置換済み（§10）。
- **AI調査サマリービュー（AI_RESEARCH）を正典化**。`writeAIResearchSummaryView_` が3段ビューを再描画（§11）。
- **ランタイムシート初期化を正典化**（§1.1・§10.5）。
- **FORECAST_REPORT は撤去済み**（§6）。

---

## 0.5 v7からの主要な乖離（同期のために明記）
- **三角観測は廃止済み**。実装は「単一Opsモデル（線形トレンド×季節指数）＋残差Monte Carlo」が定量土台。
- **逆sMAPE重み更新（w_i）は存在しない**。FORECAST_SNAPSHOT の三角測量系カラム（w1〜w4 等）も**物理削除済み**
  （snapshot-vestigial-removal）。現行ヘッダは15列・`final_pred`はindex 9。
- **DLM（対数空間の状態空間モデル）が追加**され、CONFIGの DLM_ENGINE_MODE（off/shadow/primary）で制御。
- **主観は乗算係数（kProd/kClient/kOpinion/kAI）として反映**し、月次cap（QUAL_SUBJECTIVE_MONTHLY_CAP）でクリップ。
  v6の「主観キャリブレータ（オーバーレイ率ターゲット探索）」は撤去済み（cap pass-through）。
- **AIは予測係数ではなく、benchmark/event blend のスコアとして kAI に限定反映**。
  品質不足（coverage由来）時のみ中立化。**「総合スコアのdead-zone中立化」は撤去**（既定閾値0）。
  上限は ±AI_MAX_ABS_EFFECT（±3%）のみで担保（ai-score-degroundless）。
- **event_score の算出が是正済み**。`direction符号 × (impact/100) × 50 × confidence`。
  impact_score は 0〜100 の影響の大きさ（50は中立ではない）。旧 `|impact-50|` 式は撤去（ai-score-robustness）。
- **AI調査の取得経路が Vertex AI へ移行**。生ログは `AI_RESEARCH_RAW`（旧 WEB/EXTERNAL を統合）に集約。
  旧2枚→RAW の移送は**ヘッダ名マッチ＋シート名由来 axis**（raw-migrate-header-match）。
- **信頼度（SOURCE_RELIABILITY）** が追加され、各ソースの寄与に reliability_r を乗じる。既定ON（空ならno-op）。
- **LMDI分解** が追加（CONFIG LMDI_DECOMPOSITION_ENABLED、既定OFF）。
- **POOL_PRIOR のクライアント横断集約** が追加（中央集約book→各bookへfan-out / 手動実行）。
- **calibration override の空欄/明示0 区別を是正**（calibration-blank-override-fix）。空欄は「未設定」、
  明示0は「意図的ゼロ」として扱う。

---

## 1. 実装スコープ

### 1.1 現行実装済み
- データ取込：外部実績SS → SALES_INPUT（A-2）→ SALES_MONTHLY 48ヶ月横持ち BASE/SPOT/TOTAL（A-3）
- 主観入力：PRODUCT / CLIENT / OPINIONS / DEV_SPOT（A-5〜A-8）
- AI調査（Vertex AI自動）：A-4 `runVertexAIResearch`。grounded web検索 + Vertex AI Search RAG +
  構造化出力 → AI_RESEARCH_STRUCTURED へ記録、AI_RESEARCH（サマリービュー）へ再描画（§10・§11）。
  生ログは AI_RESEARCH_RAW（axis=web/rag）と AI_RESEARCH_TASK_LOG に保存。
- 予測実行：runPhase1Forecast（A-9）。OUTPUT / FORECAST_SNAPSHOT 更新
- ダッシュボード：updatePhase1Dashboard（A-10）
- 検証：実績取込（B-1）→ EVAL_LOG / EVAL_COMPARE_MONTHLY（B-2）→ EVAL_INSIGHTS（B-3）
- 四半期レビュー：runQuarterlyReview（C-1）→ applyQuarterlyProposals（C-2）→ ログ閲覧（C-3）
- 信頼度：SOURCE_RELIABILITY 適用＋C-1のreliability提案＋RELIABILITY_EVIDENCEへの raw hit/n 永続化
- 横断プール：POOL_PRIOR のクライアント横断集約（adminSetupPoolHub / adminAggregatePoolPriorAcrossBooks）
- **シート初期化**：AI_RESEARCH_STRUCTURED は A-1（`setupForecastBook` の order[] と `buildPhase1Sheets_`）で
  先行作成される。Vertex実行時に真に遅延作成されるのは **AI_RESEARCH_TASK_LOG / AI_RESEARCH_RAW の2枚**
  （旧 WEB/EXTERNAL は RAW に統合済み）。`ensureAIResearchRuntimeSheets_` は STRUCTURED/TASK_LOG/RAW を
  冪等に存在確認・ヘッダ整合し、旧2枚があれば `migrateLegacyAIResearchRawSheets_` で RAW へ一度だけ移送する。
  POOL_PRIOR / POOL_REGISTRY / POOL_AGGREGATION_LOG / DLM_STATE / BACKTEST_REPORT / LANDING_FORECAST も
  管理関数や予測実行時に必要に応じて `getOrCreateSheet_` で作成する。

### 1.2 既定OFF/中立のシャドウ機能（検証待ち）
- DLM_ENGINE_MODE = off（shadow/primaryはCONFIGで切替可）
- LMDI_DECOMPOSITION_ENABLED = 0
- FORECAST_CLOSED_MONTH_MODE = actual（**月次表示のみ**を切り替えるトグル。年度合計の算出には影響しない。§9）
- AI_SCORE_BASIS = level（momentumに切替でAIスコアを相対位置の変化として扱う）
- AI_TOTAL_NEUTRAL_THRESHOLD = 0（dead-zone なし＝既定。>0 にすると総合中立化が復活するが、根拠ある場合のみ推奨）
- ※ RELIABILITY_APPLY_ENABLED は既定1（ON）。ただしSOURCE_RELIABILITY空ならno-op（予測不変）。
- ※ AI_RESEARCH_ENABLED は既定1。0でA-4のVertex調査をスキップし、AI_RESEARCH_STRUCTUREDの既存行のみ参照。
- ※ AI_MISSING_CONFIDENCE_DEFAULT は既定0.5（confidence欠落を0.5で補完して採用）。0で従来どおり不採用。

### 1.3 未実装（今後）
- 構造変化検知（CUSUM/Bai-Perron）、分位点回帰の本格適用、学習窓の動的最適化。
- POOL_PRIOR集約のprecision導出の経験ベイズ化（現状はΣhit/Σn加重＋固定precision=SHRINKAGE_K）。
- person系ソースの横断プール（現状は普遍キーのsource_type単位まで）。
- AI_RESEARCH_RAW の `frozen_flag` / `frozen_at` 列は将来のスナップショット凍結用の予約列（現状は未使用 / 既定0・空）。

### 1.4 メニュー構成（実装同期）
- A-1 初期セットアップ（setupForecastBook）
- A-2 売上データを取り込む（importSalesInputMonthly）
- A-3 予測用に売上データを加工（aggregateSalesData）
- A-4 AI調査を取り込む（**runVertexAIResearch** / Vertex AI自動）
- A-5 製品ごとの動向を入力（openProductTrendEntryDialog）
- A-6 クライアント動向を入力（openClientTrendEntryDialog）
- A-7 担当者意見を入力（openOpinionsEntryDialog）
- A-8 開発/スポット要因を入力（openDevEntryDialog）
- A-9 予測を実行（runPhase1Forecast）
- A-10 予測ダッシュボードを更新（updatePhase1Dashboard）
- B-1 検証用に実績データを取り込み（importActualEvalMonthly）
- B-2 検証レポートを更新（updatePhase1EvaluationReport）
- B-3 検証インサイトを更新（updatePhase1LearningInsights）
- C-1 四半期レビューを実行（runQuarterlyReview）
- C-2 承認済み提案を適用（applyQuarterlyProposals）
- C-3 過去の提案履歴を開く（openQuarterlyReviewLog）
- 管理者用（メニュー非掲載 / スクリプトエディタから手動）：
  adminSetupGuideOnly / adminInitDLMAndBacktest / adminSetupPoolHub / adminAggregatePoolPriorAcrossBooks

---

## 2. 予測エンジンの実体（runForecastFYCore_）

### 2.1 定量土台
1. SALES_MONTHLY から baseSeries48（BASE 48ヶ月）を読む。
2. adjustForUnclosedMonths_：未確定月（前月までを確定）を同月トレンド係数で補完。補完後は途中実績を下回らない。
3. fitOpsModelTrendSeason_：OLS線形トレンド＋移動平均ベース季節指数（0.80〜1.20でクリップ）。
4. buildResidualPool_：確定月の残差%プール（MADクリップ＋中央値方向へ収縮）。
5. forecastByResidualQuantiles_ で P10/P50/P90 の定量土台を算出。
6. DLM_ENGINE_MODE=primary かつ実績充足時は、BASEを対数空間DLMの予測へ差し替え（shadowは比較のみ）。

### 2.2 SPOT
- 背景SPOT（未知の再発）：fitSpotRecurringModel_ が月別の期待値・発生確率・severity標本を作る（BASE P50比でcap）。
- 既知SPOT（DEV_SPOT）：金額×確度を月別に固定加算（knownSpot）。背景との二重計上はoffset率で調整。

### 2.3 主観乗算（forecastMonteCarloMixed_）
- kProd = 1 + Σ(製品構成比 × 製品step × reliability)
- kClient = 1 + Σ(client step × reliability)
- kOpinion = 担当者別の最新意見を ±5% jitter込みで合成（× reliability）
- kAI = 1 + clamp(Σ(topicスコア × reliability) × AI_WEIGHT, ±AI_MAX_ABS_EFFECT)。
  **総合中立化（dead-zone）は既定OFF**（AI_TOTAL_NEUTRAL_THRESHOLD=0）。上限は ±AI_MAX_ABS_EFFECT のみで担保。
- Monte Carlo（N_SIM=1000）で各simの total を生成し、月次P10/P50/P90を取る。
- 主観差分は月次cap（QUAL_SUBJECTIVE_MONTHLY_CAP、既定0.20）でクリップ（capHitを診断記録）。
- **kProd全月1.0のフェイルセーフ**：PRODUCTに有効行があるのに kProdが全月1.0（直近12ヶ月closed BASE実績の
  無い製品＝weight=0）の場合、throwせず警告（`productWeightWarning`）に置換し A-9を完走させる。警告は
  weight=0 となった製品名を列挙し、OUTPUT上部の警告ブロックとRUN_LOG note（`prodw=...`）に残す。

### 2.4 信頼度（reliability）
- readReliabilityApplyEnabled_ が真のとき readSourceReliability_(client) を適用。
- getSourceReliability_(map, type, key)：未登録は 1.0（中立＝フェイルセーフ）。
- 適用先：factor_product:person / factor_client:person / opinion:person / ai_topic:topic。
- 既定ON。ただし SOURCE_RELIABILITY が空なら全ソース1.0でno-op（予測不変）。

### 2.5 LMDI分解（診断のみ）
- lmdiDecompose_：Πk-1 を kProd/kClient/kOpinion/kAI に厳密加法分解（全正値はLMDI-I、非正値は線形フォールバックで和保存）。
- 既定OFF。ONのときOUTPUTに寄与シェア／絶対レンジ／相対レンジを出力。

### 2.6 AIスコアの予測反映
- readAIResearchScores_ が AI_RESEARCH_STRUCTURED から topic別（Market/Competitor/Channel/DX）の
  final blended score（benchmark/event blend）を返す。品質不足（coverage由来）topicは中立化（multiplier=0/0.5）。
- AI_SCORE_BASIS=momentum のときは AI_SCORE_HISTORY の過去runとの差分（momentum）に切替（既定はlevel）。
  momentum の lookback は `as_of_date`（調査リフレッシュ）単位で dedup され、A-9実行回数では動かない。
- これがA-4（Vertex）で書かれた行を読む唯一の入口。A-9はVertex APIを呼ばず、シート上の構造化行のみ参照する。

### 2.7 年度合計のセマンティクスと2経路の分離（確定）
本ツールには性質の異なる2つの経路があり、実績の扱いが明確に違う。混同しないこと。

**(a) 予測・提出経路（A-9 / runForecastFYCore_）— 実績非依存**
- 入力は SALES_MONTHLY（過去48ヶ月のBASE/SPOT履歴）と主観入力・AI調査・DEV_SPOT。
- 予測対象は予測FYの12ヶ月（fy/04〜fy+1/03）。出力は OUTPUT と FORECAST_SNAPSHOT。
- **年度合計(P10/P50/P90)は、12ヶ月すべての予測simから算出する**。経過月の実績で固定した着地値ではない
  ＝**通年予測の見通し**（§9で確定した選択肢A）。
- 運用上、この経路は会社公式予測を提出する前に回すため実績が出る前に実行される。加えて構造的にも、
  A-9 は実行直前に `syncSalesFromSalesInput_` で SALES窓を `[fy-4/04, fy/03]` に再整列し、予測窓
  `[fy/04, fy+1/03]` とは同一fyで隣接非重複に確定する。よって全月 `forecast_open` となり実績は混ざらない。

**(b) 検証・学習経路（B-1 → B-2 → B-3、C-1）— 実績を使うが予測は書き換えない**
- B-1：実績を ACTUAL_EVAL_MONTHLY に取り込む。
- B-2：提出済みの FORECAST_SNAPSHOT（`final_pred`=r[9]）と実績を突合し EVAL_LOG / EVAL_COMPARE_MONTHLY に記録。
- B-3：EVAL_INSIGHTS に外れ要因・次アクションを整理。
- C-1：EVAL_LOG と impact履歴（AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY）から reliability を学習。
- 実績は「答え合わせと学習の材料」であり「予測の入力」ではない。提出済み予測を retroactive に書き換えない。

---

## 3. CONFIG パラメータ（読取方式）

### 3.1 読取規約
- readConfigLabelMap_ が CONFIG A:B を読み、configKeyOf_ で「（」または「(」より前をキー化して完全一致マップを作る。
- read*FromConfig_ 系（readModelTuningFromConfig_ / readDlmEngineMode_ / readReliabilityApplyEnabled_ /
  readLmdiDecompositionEnabled_ / readDlmPrimarySpotCapBasis_ / readForecastClosedMonthMode_ /
  readAiScoreBasis_ / readAiMissingConfidenceDefault_ / readVertexConfig_）はすべてこのマップ経由。
- セル番地直読みは廃止。tuneRows に行を挿入しても壊れない。
- ラベル末尾の注記（全角括弧内）は自由に変更可。キー部分が一致すれば読める。

### 3.2 主要キー（抜粋・現行同期）
- AI_WEIGHT / AI_MAX_ABS_EFFECT
- **AI_MISSING_CONFIDENCE_DEFAULT**（confidence/relative_confidence 欠落時の補完既定値 0〜1 / 既定0.5 / 0で不採用）
- AI_TOTAL_NEUTRAL_THRESHOLD（既定**0**＝総合dead-zone中立化なし。>0は根拠ある時のみ）
- AI_QUALITY_NEUTRAL_THRESHOLD / AI_QUALITY_PARTIAL_THRESHOLD
- AI_SCORE_BASIS（level/momentum）/ AI_MOMENTUM_LOOKBACK_QUARTERS / AI_MOMENTUM_MIN_HISTORY
- AI_WEIGHT_PROPOSAL_MIN / AI_WEIGHT_PROPOSAL_MAX
- QUAL_SUBJECTIVE_MONTHLY_CAP / QUAL_CALIBRATION_ENABLED
- SPOT_BG_* / KNOWN_SPOT_* / SEASONAL_*
- DLM_ENGINE_MODE / DLM_PRIMARY_SPOT_CAP_BASIS / DLM_BACKTEST_* / DLM_LOG_EPSILON_RATE
- RELIABILITY_APPLY_ENABLED（既定1）/ RELIABILITY_R_MIN / R_MAX / SHRINKAGE_K / MIN_SAMPLES / MIN_CHANGE
- POOL_MIN_CLIENTS（横断集約の最低クライアント数、既定2）
- LMDI_DECOMPOSITION_ENABLED
- FORECAST_CLOSED_MONTH_MODE（actual/forecast、既定actual。§9参照）
- **Vertex AI調査キー（§10）**：VERTEX_PROJECT_ID / VERTEX_LOCATION / VERTEX_GEMINI_MODEL /
  VERTEX_DATASTORE_ID / VERTEX_SEARCH_LOCATION / **VERTEX_SERVING_CONFIG**（既定 default_search） /
  AI_RESEARCH_ENABLED（既定1）

### 3.3 撤去済みCONFIG行（参考）
config-simplify / config-deadknob-removal / qual-share-const-removal により以下は CONFIG から撤去済み：
- 担当者の B10 互換参照（=B4）、QUAL_SUBJECTIVE_MAX_SCALE
- QUAL_SHARE_ALERT_THRESHOLD / QUAL_SHARE_TARGET_CENTER/LOW/HIGH（const もコードから撤去）
- QUARTERLY_REVIEW_PERIOD_MONTHS / BIAS_CORRECTION_* / AI_DIRECTION_HIT_* / AI_EFFECT_MIN_MEANINGFUL /
  DLM_FORECAST_HORIZON
残す調整キー（AI_WEIGHT / QUAL_SUBJECTIVE_MONTHLY_CAP / RELIABILITY_* / DLM_BACKTEST_MIN_MONTHS 等）は不変。

---

## 4. 学習ループ（四半期レビュー C-1〜C-3）

### 4.1 データ収集（collectQuarterlyReviewData_）
- EVAL_LOG から client一致・scenario=neutral・constraint_relevant_flag=1 の行を抽出（ヘッダidx参照）。
- 直近3ヶ月（last3）が揃わなければ ready:false（提案ゼロで正常終了）。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY / AI_SCORE_HISTORY を last3 で結合。

### 4.2 サンプル水増し対策（dedup / open限定）
背景：GUIDEは「確認→修正→再実行を前提とする」と明記しており、同じ月に対しA-9を複数回押す運用が正規。
旧実装はA-9のたびに履歴を追記（dedupなし）し n が膨張、提案が「高」信頼度に化ける欠陥があった。

対策（現行 computeReliabilityHitStats_）：
- **forecast_source 列**：AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY の末尾に
  `forecast_source`（'forecast_open' / 'actual_closed'）を持つ。記録値は runForecastFYCore_ の
  result.sourceByMonth[i] をそのまま書く。
- **open限定**：C-1集計は `forecast_source==='forecast_open'` の行のみ採用。closed後の actual上書き行
  （surprise=0）を自動除外。空/欠落の旧（汚染）行も自動除外。
- **dedup**：
  - quant側（AI_IMPACT_HISTORY）：month単位で最新 run_at の1件のみ採用（latestQuantByMonth）。
  - subjective側（SUBJECTIVE_IMPACT_HISTORY）：(target_month, source_type, source_key) 単位で
    最新 run_at の1件のみ採用（latestSubjectiveByUnit）。
  これにより n が「A-9実行回数」に依存しなくなる。
- 触らない箇所：evalActualByMonth（EVAL_LOG由来のactual）は month単位 last-write-wins のままでよい。

### 4.3 信頼度提案（generateReliabilityProposals_）
- 各 (source_type, source_key) について、push方向と「実績−定量(quant_only)」のサプライズ方向の的中率hを集計。
- rHat = clamp(2h, R_MIN, R_MAX)。POOL_PRIOR（横断事前）と shrinkage_k で収縮 → rShrunk。
  rShrunk = clamp((n*rHat + k*rPool) / (n+k), R_MIN, R_MAX)。
  k は POOL_PRIOR の precision（無ければ SHRINKAGE_K）。
- |rShrunk − 現在値| >= MIN_CHANGE のとき提案化。confidenceは n と変化量で判定。
- プール入力は「生のhit/n」を母数とする（rShrunk再プールは自己強化になるため不可）。
  各bookのC-1で source_type別 hit/n を RELIABILITY_EVIDENCE に永続化する（§7の集約前提）。

### 4.4 適用（applyQuarterlyProposals）
- QUARTERLY_REVIEW の承認列（承認/却下/保留）に従う。空欄は保留。
- 承認かつ auto_update_enabled=1 のときのみ SOURCE_RELIABILITY を upsert ＋ CALIBRATION_HISTORY 記録。
- 同一 review_id の二重適用は早期終了。

### 4.5 AI_weight 提案（generateQuarterlyProposals_）
- AI方向一致率（hitRate）と mean|kAI-1| から ai_weight_override 提案を生成。
- 現在値の読取は **空欄=未設定→CONFIG既定、明示数値→その値** で扱う（calibration-blank-override-fix と同方針）。
- 提案値が aiWeightProposalMin〜Max の範囲内のときのみ提案化。

### 4.6 calibration override の空欄/明示0 区別（calibration-blank-override-fix）
- `applyCalibrationToTuning_` 内の `calOverrideNum_(v)`：`''`/`null`/`undefined` は null（未設定→無視）、
  数値は isFinite なら採用。これにより CALIBRATION_STATE の `ai_weight_override` /
  `ai_max_abs_effect_override` が空欄でも `Number('')→0` で aiWeight=0 に化けず、kAI が不当に 1.0 へ潰れない。
- 明示的に 0 を入れた場合は 0 が採用され、意図的にAIを止める運用は従来どおり有効。
- この修正の影響は混合セクション（AI効果分 ±3%内）のみ。客観のみ・SPOT・DLM・他の主観係数は不変。

---

## 5. 検証ポリシー（KPI）
- 計画用単一値＝P50。P10/P90は説明帯（hard gateではない）。
- 制約：annual_abs_error_rate <= 10% / half_wape <= 12%（将来目標10%）/ over-forecast rate <= 5%。
- 診断：月次APE、Q差分、定量寄与率、主観オーバーレイ率（AI除く）、AI寄与率、Known Spot寄与率、レンジ逸脱数。
- レンジ逸脱月（actualがP10-P90外）はB-3で追加調査を必須化。
- 1 client = 1 book。

---

## 6. ログ／永続シート
- ユーザ表示シート：GUIDE / CONFIG / SALES_INPUT / SALES_MONTHLY / AI_RESEARCH（サマリービュー）/
  PRODUCT / CLIENT / OPINIONS / DEV_SPOT / OUTPUT / DASHBOARD（hideNonUserSheets_ で制御）。
- 内部管理（原則非表示）：RUN_LOG / FORECAST_SNAPSHOT / PROCESS_STATUS / AI_SCORE_HISTORY /
  AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY / CALIBRATION_STATE / CALIBRATION_HISTORY /
  QUARTERLY_REVIEW / QUARTERLY_REVIEW_LOG / DLM_STATE / BACKTEST_REPORT / SOURCE_RELIABILITY /
  RELIABILITY_EVIDENCE / POOL_PRIOR / POOL_REGISTRY / POOL_AGGREGATION_LOG / LANDING_FORECAST /
  AI_RESEARCH_STRUCTURED / **AI_RESEARCH_TASK_LOG / AI_RESEARCH_RAW**。
- **FORECAST_REPORT は撤去済み**。
- **FORECAST_SNAPSHOT は15列**（snapshot_id / run_date / client / target_month / scenario / base_pred /
  subjective_adj / ai_adj / deterministic_adj / **final_pred(index 9)** / confidence_interval_lower /
  confidence_interval_upper / key_factors_json / subjective_input_date / calibration_applied_json）。
  三角測量系 vestigial カラム（linear/robust/regime/simulation_pred / w1〜w4）は**物理削除済み**。
- **AI_RESEARCH_RAW（旧 WEB + EXTERNAL の統合）**：headers は
  `client / as_of_date / axis / topic / direction / magnitude / uncertainty / relative_position /
  evidence / frozen_flag / frozen_at / note`。`axis`=web/rag で生応答の出所を区別する。
  実際の生応答・citation・prompt（web）／summary・documents（rag）は `note`（JSON）に保存。
  `frozen_flag`/`frozen_at` は将来用の予約列（現状未使用 / 既定0・空）。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY 末尾に forecast_source（§4.2）。

---

## 7. POOL_PRIOR 横断集約（3c-3c）
### 7.1 構成
- ハブbook初期化：adminSetupPoolHub（メニュー非掲載 / スクリプトエディタから手動）。
  POOL_REGISTRY / POOL_AGGREGATION_LOG / POOL_PRIOR を作成。
- 集約実行：adminAggregatePoolPriorAcrossBooks（同・手動）。REGISTRYの enabled=1 book を openById で開き、
  各bookの RELIABILITY_EVIDENCE から source_type別の生 hit/n を読み、集約してPOOL_PRIORへ書き込み＋fan-out。

### 7.2 集約ロジック
- 集約入力：生 hit/n をプール（rShrunk再プール禁止）。
- 粒度：reliability:{source_type}（factor_product/factor_client/opinion/ai_topic）。person系キーは横断不可。
- pooled_value = clamp(2 * Σhit/Σn, R_MIN, R_MAX)。precision = SHRINKAGE_K（固定）。
- フェイルセーフ：
  - n_clients < POOL_MIN_CLIENTS → written=false / reason=min_clients（POOL_PRIORに書かない＝空のまま）。
  - Σn < MIN_SAMPLES → written=false / reason=min_samples。
  - 単一book読取失敗（権限/不存在）は status=excluded でログ記録し、全体は止めない。
- 監査：POOL_AGGREGATION_LOG に per-book 行と per-scope 行を1 run_id で記録。

### 7.3 予測への波及
- POOL_PRIOR は C-1 の提案収縮（rShrunk）に効くのみ。予測値そのものは変えない（集約直後のA-9はno-op）。
- POOL_MIN_CLIENTS未満で空のまま据え置く設計は auditability（1.0で埋めない）を保つため。

---

## 8. バージョン整合と適用順序
- VERSION / BUILD_STAGE / 設計書版 / 手動チェックリストは各リリースで同期する。
- 現行コードは VERSION='2.3.31-dev' / BUILD_STAGE='raw-migrate-header-match'。
  DLM_BUILD_STAGE は 'v8-step3c3c-1'（DLMロジック無変更のため据え置き）。
- 本改訂（設計書 v2.3.4）は output-display-tidy（2.3.30 / 表示専用）と raw-migrate-header-match
  （2.3.31 / 生ログ移送のヘッダ名マッチ化）をドキュメントへ反映する doc sync。raw-migrate-header-match は
  同梱の Codex プロンプトで適用する前提のため、**コードを先に適用・検証してから本書へ差し替える**こと。
- 設計書ドリフトは既知の再発リスク。リリースごとに doc sync を独立タスクとして扱う。

---

## 9. 通年予測モード（FORECAST_CLOSED_MONTH_MODE）【確定：A=通年予測 / コード変更なし】
年間総計のセマンティクスは **「通年（12ヶ月）すべて予測の見通し」（選択肢A）** で確定。

### 9.1 確定した仕様
- **年度合計(P10/P50/P90)は、FORECAST_CLOSED_MONTH_MODE の値にかかわらず、常に12ヶ月すべての予測simから算出する。**
  年度合計は sim配列（mixedは `totalCalibratedSimByMonth`、客観は quant側sim）の年次集約（`aggregateAnnualSim_`）に
  percentile を取ったもので、closed月上書きは sim配列に触れず月次表示用配列にしか作用しない。
- よって本トグルは **「月次表示の見せ方」だけを切り替える**もので、計画の主数値（年度合計）の意味は変えない。

### 9.2 モード別の月次表示
- `actual`（既定）：closed月は実績で上書き表示（objOnly/mixed/regTotal）。
- `forecast`：closed月も予測のまま表示（通年予測モード）。経過月の実績は ActualClosed 列に参考併記。

### 9.3 closed月上書きの発火条件と、実運用での非発火
- actualモードの上書きは `sourceByMonth[i]==='actual_closed'` の月のみに作用する。
- A-9 経路では、timing（実績前に回す）と structure（syncSalesFromSalesInput_ による窓非重複）の二重で
  発火条件に到達しない（全月 forecast_open）。よって上書き対象の closed月が存在しない。

### 9.4 回帰確認（1回のみ）
- OUTPUT「（参考）内訳とメモ（P50比較）」表の `ForecastSource` 列が、A-9 実行後に全行 `forecast_open` で
  あることを1回確認すれば、§9.3 の非発火はその場で確認できる。

### 9.5 表示位置と OUTPUT 表示整形（output-note-relocate / output-display-tidy）
- 年度合計行と月次表ヘッダの間は短い1行注記のみ（混合は `..._PRIMARY_SHORT_TEXT`、客観は従来の1行注記）。
- 「年度≠月次合算」の長文説明は混合セクションの Scenario Split の下（目立たない灰色小フォント）に配置。
  これは表示位置のみの変更で数値は不変（output-note-relocate）。
- **output-display-tidy（2.3.30）の表示整形**（いずれも `writeOutputFY_` のみ・数値不変）：
  - OUTPUT 6行目の注記ブロックは、実警告（`⚠` / 製品重み警告 / `web_error`・`rag_error`・`structure_error` が1以上）が
    あるときのみ赤字。all-zero の情報サマリーだけでは黒字（`step3aHasError` が error≥1 を検出）。
  - row10「定量/主観/AI/KnownSpot 100%分解」テキストは F列（E10は空・NOTEもF10）。F9/F10 は折り返し表示。
  - 21行目 coverage は topic ごとの改行表示＋A〜F幅（6列）。22行目（degraded / benchmark不足）・
    23行目（中立化バナー）も折り返し・幅縮小で横伸びを抑える。
- 初期設定ダイアログの黄色 notice は output-display-tidy で撤去（複製手順は GUIDE 等でカバー）。

---

## 10. AI調査パイプライン（Vertex AI / A-4 = runVertexAIResearch）

### 10.1 全体像
A-4はメーカー名（CONFIG!B2）について、4 topic（Market/Competitor/Channel/DX）ごとに以下を実行する：
1. **grounded web検索**（callVertexGeminiGrounded_）：Gemini + googleSearch（不可時 googleSearchRetrieval へ
   フォールバック）で最新Web情報を取得。groundingメタからcitationを抽出。
2. **RAG検索**（callVertexSearchRAG_）：Vertex AI Search データストアを検索し summary + citations + documents を取得。
   VERTEX_DATASTORE_ID/SEARCH_LOCATION 未設定時はskip（web-only）。`VERTEX_SERVING_CONFIG`（既定 default_search）で
   サービング構成を指定。`discoveryEngineHost_` は global/地域ロケーションに応じてエンドポイントを生成。
3. **構造化**（callVertexGeminiStructured_）：web結果をevent行、RAG結果をbenchmark行として
   responseMimeType=application/json で構造化（buildVertexStructuredRows_）。**temperature=0 固定**（再現性）。

### 10.2 readiness 判定
- readVertexConfig_ が geminiReady（projectId+location+geminiModel）と ragReady（datastoreId+searchLocation）を分離。
- A-4の入口判定は geminiReady。RAG未設置でも web-only で実行できる。
- AI_RESEARCH_ENABLED=0 のときは「無効」メッセージで何もせず終了。
- geminiReady=false のときは「必須設定が未入力」エラーで停止（手動経路には落ちない）。

### 10.3 スコア化（event / benchmark）【ai-score-robustness / degroundless で是正済み】
- **event_score**：`direction符号(up=+1/down=-1/neutral=0) × (impact_score/100) × 50 × confidence` を ±50 にclamp。
  - impact_score は 0〜100 の「影響の大きさ」（0=なし、100=最大）。**50は中立ではない**。向きは direction が担う。
  - 旧式 `sign × |impact-50| × conf` は撤去（impact=0.9 を過大評価する不具合を解消）。
  - confidence 欠落時：direction か impact があり AI_MISSING_CONFIDENCE_DEFAULT>0 なら既定値で補完して採用
    （confDefault 計上）。補完OFF(=0)なら不採用（confDrop 計上）。
  - 採点での再構成は readAIResearchScores_ 側でも行う（A-4が空 event_score を書いても A-9 が補完式で再計算）。
- **benchmark_score**：`(relative_percentile-50) × relative_confidence × quality倍率(high1/medium0.75/low0.5)` を ±50 にclamp。
  - relative_confidence 欠落時は relative_percentile があれば既定値で補完（benchDefault 計上）。
  - **benchmark 行は相対位置の根拠があるときだけ書く**。根拠が無いトピックは relative_percentile/label/peer_* を
    空（null）にし benchmark を一切埋めない（50/middle のプレースホルダ禁止）。
- これらを readAIResearchScores_ が topic別に blend（Market/Competitor/Channel=0.65:0.35、DX=0.50:0.50）。
- AI_SCORE_HISTORY に topic別 blended_score（level）を記録（momentum算出の履歴源）。

### 10.4 品質中立化と「根拠なき中立化／捏造ゼロ」の是正（ai-score-degroundless）
- **総合 dead-zone 中立化は撤去**（AI_TOTAL_NEUTRAL_THRESHOLD 既定0）。総合スコアが小さくても kAI は dead-zone で
  1.0 に潰れず、±AI_MAX_ABS_EFFECT（±3%）内で反映される。上限はこのキャップのみで担保。
- **event_only への構造的 ×0.5 ペナルティを撤去**。信号の弱さは coverage 由来 qualityMultiplier のみで反映。
  - qualityScore = clamp((benchCount×2 + eventCount)/4, 0, 1)。
  - < AI_QUALITY_NEUTRAL_THRESHOLD(0.25) → 0倍（中立化）／< AI_QUALITY_PARTIAL_THRESHOLD(0.50) → 0.5倍。
- **honest-zero**：web が真に neutral（動きなし）かつ benchmark 根拠なしのトピックは依然 0 になりうる。
  これは捏造しない設計どおりの正常動作（強制非ゼロにはしない）。
- OUTPUT 行23の「信頼度不足で中立化」バナーは coverage 由来の中立化時のみ表示（総合dead-zone由来は出ない）。

### 10.5 失敗時の扱い（フェイルセーフ）
- web/rag/structure の各段で失敗してもrun全体は止めず、AI_RESEARCH_TASK_LOG に per-topic per-aspect 行を残す。
- 全topicで構造化行が0件なら、AI_RESEARCH_STRUCTURED は**上書きせず既存行を保持**し、エラー通知して終了。
- 1件以上得られたら AI_RESEARCH_STRUCTURED を全置換し、サマリービュー（§11）を再描画。

### 10.6 ランタイムシート（ai-research-raw-merge で統合 / raw-migrate-header-match で移送是正）
- AI_RESEARCH_STRUCTURED：A-1で先行作成される構造化結果シート（event/benchmark行）。A-9が読む唯一の入口。
- A-1では作らず、Vertex実行時に遅延作成されるのは以下の**2枚**：
  - AI_RESEARCH_TASK_LOG：run_id/topic/aspect(web/rag/structure)/status/duration/usage/citations/error を記録。
  - **AI_RESEARCH_RAW**：旧 AI_RESEARCH_WEB + AI_RESEARCH_EXTERNAL を統合した単一の生ログシート。
    `axis`=web/rag で出所を区別し、生応答（web は prompt/text/citations、rag は query/summary/documents）を
    `note`（JSON）に保存。append 行は `appendAIResearchRawRow_` が axis を明示付与する。
- **既存bookの旧2枚→RAW 移送（migrateLegacyAIResearchRawSheets_）は raw-migrate-header-match（2.3.31）で
  ヘッダ名マッチ化された**：
  - 旧2枚の各行を、行ヘッダ名と RAW ヘッダ名の一致で写す（位置ベースは撤去）。旧2枚の列順が RAW と
    異なっても `client/as_of_date/topic/evidence/note` は同名ヘッダの列へ入る。
  - `axis` は**旧シート名から確定**する（AI_RESEARCH_WEB→web / AI_RESEARCH_EXTERNAL→rag）。旧2枚に axis 列が
    無くても空にならず、列ずれも起きない。
  - `frozen_flag` は旧側に無ければ 0、`frozen_at` は空で補完。
  - 移送は一度だけ（旧タブを移送後に削除＝冪等。再実行で二重移送しない）。A-1（全上書き）経路では旧2枚は
    移送せず削除される。

### 10.7 CONFIGキー（既定値）
- VERTEX_PROJECT_ID = forecast-agent-498907
- VERTEX_LOCATION = global（grounding推奨）
- VERTEX_GEMINI_MODEL = gemini-3.1-pro-preview（またはgemini-3.5-flash）
- VERTEX_DATASTORE_ID = fujikeizai-portfolio-2025
- VERTEX_SEARCH_LOCATION = global
- **VERTEX_SERVING_CONFIG = default_search**（アプリにより default_config の場合あり）
- AI_RESEARCH_ENABLED = 1
- 必要OAuthスコープ：cloud-platform / script.external_request（appsscript.jsonに設定済み）

### 10.8 RAGクエリのフレーム整合（rag-query-frame-align）
- `buildRagQuery_` は AI調査の新フレーム（外部観測可能な「支援需要に影響しうる外部環境」）に整合。
- topic 別の日本語展開語（Market=市場規模/需要/成長率、Competitor=競合/上市/シェア、Channel=MR/チャネル/販促、
  DX=DX/デジタル/投資 等）で外部環境を狙う。英語topic素語の連結や出版社名（単一コーパスで弁別力なし）は撤去。
- 本変更は A-4 側のクエリのみ。A-4 を再実行しない限り A-9 の OUTPUT は不変。再実行後は benchmark スコアが
  変わりうる（=意図した改善）。客観のみ（quantOnly/objOnly）は不変。

---

## 11. AI調査サマリービュー（AI_RESEARCH シート）
A-4実行時に writeAIResearchSummaryView_ が AI_RESEARCH シートを再描画する（ユーザ表示シート）。

### 11.1 構成（3段）
- ① topic別サマリー（要約文）：構造化結果の report_text を topic単位で表示（splitAIResearchReportTextByTopic_ で
  【Market】等の見出しで分割）。全topicで要約文が空ならフェイルセーフ文を表示。
- ② AIスコア サマリー（4軸）：topic別 Final Score / event_score / benchmark_score / 最新as_of / 備考。
  Final Score は readAIResearchScores_ と一致（同一runのOUTPUT 4軸スコアと整合）。
- ③ スコア根拠（event / benchmark 明細）：行種別ごとに direction/impact/confidence/relative_*/quality/evidence 等。

### 11.2 再描画の不変条件
- A-4を複数回実行しても追記でなく再描画（重複しない）。
- A-4失敗（outRows空）時は上書きせず前回内容を保持。
- 本ビューは表示専用で、予測コア（A-9）には影響しない。

---

## 12. 既知の latent issue（非ブロッキング / 別スコープで対応）
本設計書では現状を記述するに留め、修正は別プロンプトで扱う（1スコープ1変更の原則）。

- **client名マッチの不整合**：`isSameClient_` と `normalizeClientName_` 経由の直接比較が混在。
  両者とも正規化を通すため実害は低いが、表記の統一は別スコープ。

解消済みメモ：
- **AI_RESEARCH_RAW 旧データ移送の列整合**は raw-migrate-header-match（2.3.31）で**解消**。位置ベース移送を
  ヘッダ名マッチへ変更し、`axis` をシート名（WEB→web / EXTERNAL→rag）から確定するようにした。旧2枚の列順差による
  axis 1列ずれが構造的に起きなくなった（旧データを持つbookでのHOW TO TEST実機確認は維持を推奨だが、
  列ずれリスク自体は除去済み）。
- **FORECAST_SNAPSHOT の三角測量系 vestigial カラム**（w1〜w4 / linear/robust/regime/simulation_pred）は
  snapshot-vestigial-removal で**物理削除済み**。
- LANDING_FORECAST / EVAL_LOG の再実行追記増殖は複合キー upsert 化で解消済み。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY は設計どおり追記のままで、重複排除は C-1 読取側で行う。
- shadow mode 表示の `opsQuantOnly` エイリアス問題は objOnly 独立コピー化で解消済み。
- 旧Gem手動貼付経路の到達不能関数群、孤立関数（adminShowCloneGuide / importPastSalesToSalesTab）は物理削除済み。
- calibration override の空欄/明示0 取り違えは calibration-blank-override-fix で解消済み。

---

この v2.3.4 は、annual-forecast-mode（FORECAST_CLOSED_MONTH_MODE）の方針を「A=通年予測」とする正典を維持しつつ、
output-display-tidy（2.3.30 / 表示専用）と raw-migrate-header-match（2.3.31 / 生ログ移送のヘッダ名マッチ化）を
現行実装へ同期したドキュメント改訂である。前者は表示位置/書式のみ、後者は生ログ移送ロジックのみで、いずれも
予測の P10/P50/P90 は不変。年度合計は常に12ヶ月すべての予測simから算出する通年予測であり、実績は検証・学習経路
（B/C）でのみ使用する。残存 latent issue は §12 の1件（client名マッチ統一）のみとなり、別スコープで対応する。
