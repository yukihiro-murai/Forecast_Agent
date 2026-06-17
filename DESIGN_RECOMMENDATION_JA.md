# 売上予測スクリプト 設計書 v2.3.6（2.3.34〜2.3.37 同期 / overlay-cap-raise）

## 0. 文書の目的
この設計書は、実装（Forecast_Agent.js）の現行挙動を正確に記述する。
旧v7の「三角観測 w1/w2/w3/w4 + 逆sMAPE重み更新」は実装に存在しないため撤去済み。
3目的は不変：(1)予測精度向上 (2)透明化（根拠明示・再現性） (3)学習性（継続改善）。

対象実装：`Forecast_Agent.js`（VERSION='2.3.37-dev' / BUILD_STAGE='overlay-cap-raise'）
設計書版：v2.3.6（v2.3.5 からのドキュメント改訂）。
DLM_BUILD_STAGE は 'v8-step3c3c-1'（DLMロジック無変更のため据え置き）。

本改訂は、設計書 v2.3.5 が記述する `2.3.33-dev / a9-client-normalize` から現行コード `2.3.37-dev /
overlay-cap-raise` までの差分を同期する。コードはこの間に4増分（2.3.34〜2.3.37）進んでいるが、
最終コードに残る実質的な挙動差分は **主観オーバーレイ月次cap と AI係数上限の2キャップ引き上げのみ**であり、
これを overlay-cap-raise として記述する（中間stageが最終コードに痕跡を残していない場合は本書では追わない）。

### 0.0 v2.3.5 → v2.3.6 の更新点（この改訂で反映したもの）— overlay-cap-raise（2.3.37）
**従来の doc sync（表示・構造のみで予測不変）とは異なり、本改訂は予測値が変わる数値変更である。**
2系統のキャップ定数を引き上げた。いずれも「主観/AIが定量土台へ上乗せできる量の上限」を緩める変更で、
主観入力・AI調査の効いているbookでは P10/P50/P90 が変化しうる（意図した変更であり回帰ではない）。

- **主観オーバーレイ月次cap の引き上げ**：`QUAL_SUBJECTIVE_MONTHLY_CAP` を **0.20 → 0.40** に引き上げ。
  Monte Carlo 内で主観連続差分（kProd/kClient/kOpinion/kAI 由来）を `±(quantOpsBase × cap)` でクリップする
  唯一の制御点（`applySubjectiveCap_`）。cap=0.40 により、主観差分が定量土台の ±40% まで通るようになった
  （従来は ±20%）。`QUAL_CALIBRATION_ENABLED=0` のときは従来どおり cap 無効（無制限）。
- **AI係数上限の引き上げ**：`AI_MAX_ABS_EFFECT` を **0.03（±3%）→ 0.05（±5%）** に引き上げ。
  `kAI = 1 + clamp(Σ(topicスコア × reliability) × AI_WEIGHT, ±AI_MAX_ABS_EFFECT)` の絶対上限。
  `readModelTuningFromConfig_` の上限clamp（`Math.min(0.05, …)`）と OUTPUT 注記（`AI_MAX_ABS_EFFECT=±5%`）も
  ±5% へ整合済み。AI_TOPIC_SCORE_ABS_CAP（各軸±50 / 4軸合計±200）と AI_WEIGHT_DEFAULT（0.0008）は不変。

**発効条件（既存bookの前提）**：両キャップは CONFIG のチューニング表セルから読まれる（番地非依存・キー照合）。
コードの const 既定値（0.40 / 0.05）は CONFIG セルが欠落/不読のときのフォールバックに過ぎない。
clasp push しただけでは既存bookの CONFIG セルは旧値（0.20 / 0.03）のまま残るため、**A-1初期セットアップで
CONFIG を再生成するか、当該セルを 0.40 / 0.05 へ手動修正しない限り発効しない**（ai-score-degroundless と同じ前提）。
新規book（A-1済み）は既定で 0.40 / 0.05 になる。

**検証の扱い**：本変更は数値変更のため、主観/AIオーバーレイのあるbookでは A-9 の P10/P50/P90 が変わる。
B-2/B-3 の精度KPI（annual_abs_error_rate / half_wape / over-forecast rate）を再確認すること。
A-9 実行前の K_TOTAL 警告（warn ±30% / block ±50%）と Step/DEV_SPOT 警告は不変で、過大入力のガードは維持される。

### 0.0' v2.3.4 → v2.3.5 の更新点（参考・再掲）
- **client-match-unify（2.3.32）**：クライアント名マッチの比較方式を統一。`isSameClient_` を
  `normalizeClientName_(a) === normalizeClientName_(b)` に統一し、読取側のクライアント突合
  （`readSourceReliability_` / `readCalibrationState_` / `writeCalibrationState_` /
  `collectQuarterlyReviewData_` / `computeProductWeightsFromSalesInputClosed12_` /
  `syncSalesFromSalesInput_` / `fetchClientMonthlyRecords_` / `readAiScoreHistoryByTopic_` /
  `readAIReportTextForClient_` / `updatePhase1LearningInsights` / `writeDlmState_` ほか）を**すべて
  `isSameClient_` に集約**。inline の直接比較を撤去。読取は両側を正規化してから比較するため、書込側に
  表記ゆれがあっても突合が頑健。§12 の latent issue を解消。予測の P10/P50/P90 は不変。
- **a9-client-normalize（2.3.33）**：A-9（`runPhase1Forecast`）の入口で CONFIG!B2 を
  `normalizeClientName_` で正規化。以降の**永続クライアント値**（`result.clientName` /
  FORECAST_SNAPSHOT / AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY / LANDING_FORECAST /
  OUTPUTタイトル）が正規化済みになる。ACTUAL_EVAL_MONTHLY は取込時に既に正規化済みのため、両者の表記が
  一致し、**B-2 の生文字列複合キー突合**（`[client, target_month]`）が未正規化表記でも外れない。予測値は不変。

### 0.0'' v2.3.3 → v2.3.4 の更新点（参考・再掲）
- **output-display-tidy（2.3.30）**：表示専用の整形。
  ① 初期設定ダイアログ冒頭の黄色 notice（および `.notice` CSS）を撤去。
  ② OUTPUT 6行目の注記ブロックの赤字判定を「実警告（`⚠` / 製品重み警告 / `web_error`・`rag_error`・
  `structure_error` が1以上）があるときのみ赤」へ変更。all-zero 情報サマリーだけでは赤くしない。
  ③ row10「定量/主観/AI/KnownSpot 100%分解」テキストを E列→F列へ移動（E10は空・NOTEもF10）。
  ④ F9/F10 の長文を折り返し＋行高調整ではみ出し解消。
  ⑤ 21行目 coverage を topic ごとの改行表示＋A〜F幅へ縮小（行高72）、22・23行目も折り返し・幅縮小。
  いずれも表示位置/書式のみで数値は不変。変更スコープは `showInitialSetupDialog_` / `writeOutputFY_` のみ。
- **raw-migrate-header-match（2.3.31）**：`migrateLegacyAIResearchRawSheets_` の旧2枚→RAW 移送を
  **位置ベースから「ヘッダ名マッチ＋シート名由来の axis 確定」へ変更**（§10.6・§12）。旧 WEB/EXTERNAL の列順が
  RAW と一致していなくても、`client/as_of_date/topic/evidence/note` は同名ヘッダ列へ入り、`axis` は
  シート名（AI_RESEARCH_WEB→web / AI_RESEARCH_EXTERNAL→rag）から確定する。旧2枚に axis 列が無くても空にならず、
  frozen_flag は旧側に無ければ 0・frozen_at は空で補完。生ログ移送のみで予測の P10/P50/P90 は不変、冪等性も不変。

### 0.0''' v2.3.2 → v2.3.3 の更新点（参考・再掲 / 抜粋）
- snapshot-vestigial-removal（2.3.12）：FORECAST_SNAPSHOT を15列化、`final_pred` は J列（index 9）。
- config-simplify（2.3.13）/ config-deadknob-removal（2.3.15）：未使用CONFIG行・互換参照を撤去。
- rag-config-defaults（2.3.16）：CONFIG に RAG 既定値（VERTEX_DATASTORE_ID / VERTEX_SEARCH_LOCATION /
  VERTEX_SERVING_CONFIG）を A-1 で自動投入。`discoveryEngineHost_` を地域エンドポイント対応化。
- ai-dx-confidence-diagnostics（2.3.21）：confidence 欠落で event 不採用となった件数を `confDrop` 可視化。
- output-note-relocate（2.3.22）：年度≠月次合算の長文説明を Scenario Split の下へ移設。
- ai-score-robustness（2.3.23）：event_score を `direction符号 × (impact/100) × 50 × conf` へ是正。
  `AI_MISSING_CONFIDENCE_DEFAULT`（既定0.5）導入。
- ai-score-degroundless（2.3.24）：根拠なき中立化／捏造ゼロを是正（dead-zone撤去・event_only×0.5撤去・
  benchmark プレースホルダ禁止・構造化 temperature=0・momentum lookback を as_of_date 単位で dedup）。
- calibration-blank-override-fix（2.3.25）：override 空欄判定を是正（`Number('')→0` の罠を回避）。
- rag-query-frame-align（2.3.26）：`buildRagQuery_` を AI調査の新フレームへ整合。
- qual-share-const-removal（2.3.27）/ orphan-fn-removal（2.3.28）：デッドコード・孤立関数を物理削除。
- ai-research-raw-merge（2.3.29）：生ログを `AI_RESEARCH_RAW`（旧 WEB/EXTERNAL 統合）に集約。axis 列で区別。

---

## 0.5 v7からの主要な乖離（同期のために明記）
- **三角観測は廃止済み**。実装は「単一Opsモデル（線形トレンド×季節指数）＋残差Monte Carlo」が定量土台。
- **逆sMAPE重み更新（w_i）は存在しない**。FORECAST_SNAPSHOT の三角測量系カラムも物理削除済み。現行ヘッダは15列・`final_pred`はindex 9。
- **DLM（対数空間の状態空間モデル）が追加**され、CONFIGの DLM_ENGINE_MODE（off/shadow/primary）で制御。
- **主観は乗算係数（kProd/kClient/kOpinion/kAI）として反映**し、月次cap（QUAL_SUBJECTIVE_MONTHLY_CAP、既定0.40）でクリップ。
- **AIは予測係数ではなく、benchmark/event blend のスコアとして kAI に限定反映**。総合中立化（dead-zone）は撤去（既定閾値0）。上限は ±AI_MAX_ABS_EFFECT（既定±5%）のみで担保。
- **event_score の算出が是正済み**（`direction符号 × (impact/100) × 50 × confidence`）。
- **AI調査の取得経路が Vertex AI へ移行**。生ログは `AI_RESEARCH_RAW`（旧 WEB/EXTERNAL を統合）。旧2枚→RAW の移送はヘッダ名マッチ＋シート名由来 axis。
- **信頼度（SOURCE_RELIABILITY）** が追加され、各ソースの寄与に reliability_r を乗じる。既定ON（空ならno-op）。
- **LMDI分解** が追加（CONFIG LMDI_DECOMPOSITION_ENABLED、既定OFF）。
- **POOL_PRIOR のクライアント横断集約** が追加（中央集約book→各bookへfan-out / 手動実行）。
- **calibration override の空欄/明示0 区別を是正**。
- **クライアント名マッチを統一・正規化**（client-match-unify / a9-client-normalize）。読取側の突合は
  `isSameClient_`（両側を `normalizeClientName_` で正規化して比較）に集約。A-9 入口で永続クライアントを
  正規化し、FORECAST_SNAPSHOT 等と ACTUAL_EVAL_MONTHLY の表記が一致する。`normalizeClientName_` の正規化対象は
  現状 ｳﾞｨｱﾄﾘｽ系の表記統合のみ。

---

## 1. 実装スコープ

### 1.1 現行実装済み
- データ取込：外部実績SS → SALES_INPUT（A-2）→ SALES_MONTHLY 48ヶ月横持ち BASE/SPOT/TOTAL（A-3）
- 主観入力：PRODUCT / CLIENT / OPINIONS / DEV_SPOT（A-5〜A-8）
- AI調査（Vertex AI自動）：A-4 `runVertexAIResearch`。grounded web検索 + Vertex AI Search RAG +
  構造化出力 → AI_RESEARCH_STRUCTURED へ記録、AI_RESEARCH（サマリービュー）へ再描画（§10・§11）。
  生ログは AI_RESEARCH_RAW（axis=web/rag）と AI_RESEARCH_TASK_LOG に保存。
- 予測実行：runPhase1Forecast（A-9）。OUTPUT / FORECAST_SNAPSHOT 更新。
  入口で CONFIG!B2 を `normalizeClientName_` で正規化し、以降の永続クライアントを正規化済みに統一（a9-client-normalize）。
- ダッシュボード：updatePhase1Dashboard（A-10）
- 検証：実績取込（B-1）→ EVAL_LOG / EVAL_COMPARE_MONTHLY（B-2）→ EVAL_INSIGHTS（B-3）
- 四半期レビュー：runQuarterlyReview（C-1）→ applyQuarterlyProposals（C-2）→ ログ閲覧（C-3）
- 信頼度：SOURCE_RELIABILITY 適用＋C-1のreliability提案＋RELIABILITY_EVIDENCEへの raw hit/n 永続化
- 横断プール：POOL_PRIOR のクライアント横断集約（adminSetupPoolHub / adminAggregatePoolPriorAcrossBooks）
- シート初期化：AI_RESEARCH_STRUCTURED は A-1 で先行作成。Vertex実行時に遅延作成されるのは
  AI_RESEARCH_TASK_LOG / AI_RESEARCH_RAW の2枚。`ensureAIResearchRuntimeSheets_` が STRUCTURED/TASK_LOG/RAW を
  冪等に整合し、旧2枚があれば `migrateLegacyAIResearchRawSheets_` で RAW へ一度だけ移送する。

### 1.2 既定OFF/中立のシャドウ機能（検証待ち）
- DLM_ENGINE_MODE = off（shadow/primaryはCONFIGで切替可）
- LMDI_DECOMPOSITION_ENABLED = 0
- FORECAST_CLOSED_MONTH_MODE = actual（**月次表示のみ**を切り替えるトグル。年度合計の算出には影響しない。§9）
- AI_SCORE_BASIS = level（momentumに切替でAIスコアを相対位置の変化として扱う）
- AI_TOTAL_NEUTRAL_THRESHOLD = 0（dead-zone なし＝既定）
- ※ RELIABILITY_APPLY_ENABLED は既定1（ON）。SOURCE_RELIABILITY空ならno-op（予測不変）。
- ※ AI_RESEARCH_ENABLED は既定1。0でA-4のVertex調査をスキップ。
- ※ AI_MISSING_CONFIDENCE_DEFAULT は既定0.5（0で従来どおり不採用）。

### 1.3 未実装（今後）
- 構造変化検知（CUSUM/Bai-Perron）、分位点回帰の本格適用、学習窓の動的最適化。
- POOL_PRIOR集約のprecision導出の経験ベイズ化（現状はΣhit/Σn加重＋固定precision=SHRINKAGE_K）。
- person系ソースの横断プール（現状は普遍キーのsource_type単位まで）。
- AI_RESEARCH_RAW の `frozen_flag` / `frozen_at` 列は将来のスナップショット凍結用の予約列（現状は未使用 / 既定0・空）。
- `normalizeClientName_` の正規化辞書拡張（現状は ｳﾞｨｱﾄﾘｽ系のみハードコード）。新メーカーで表記ゆれが
  発生した場合の汎用正規化（全角/半角・法人格表記の一般則化）は別スコープ。
- キャップ値（主観月次cap / AI上限）の経験ベイズ的な自動調整。現状は CONFIG の固定値で運用し、
  検証KPIに基づき手動で調整する方針（overlay-cap-raise も手動の固定値引き上げ）。

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
  総合中立化（dead-zone）は既定OFF。上限は ±AI_MAX_ABS_EFFECT（既定±5%）のみで担保。
- Monte Carlo（N_SIM=1000）で各simの total を生成し、月次P10/P50/P90を取る。
- 主観差分は月次cap（QUAL_SUBJECTIVE_MONTHLY_CAP、**既定0.40**）でクリップ（capHitを診断記録）。
  cap は `±(quantOpsBase × cap)` を限度として主観連続差分をクリップする唯一の制御点（`applySubjectiveCap_`）。
  `QUAL_CALIBRATION_ENABLED=0` のときは cap 無効（無制限）。
- **kProd全月1.0のフェイルセーフ**：PRODUCTに有効行があるのに kProdが全月1.0（直近12ヶ月closed BASE実績の
  無い製品＝weight=0）の場合、throwせず警告（`productWeightWarning`）に置換し A-9を完走させる。

### 2.4 信頼度（reliability）
- readReliabilityApplyEnabled_ が真のとき readSourceReliability_(client) を適用。
- getSourceReliability_(map, type, key)：未登録は 1.0（中立＝フェイルセーフ）。
- 適用先：factor_product:person / factor_client:person / opinion:person / ai_topic:topic。
- 既定ON。ただし SOURCE_RELIABILITY が空なら全ソース1.0でno-op（予測不変）。
- SOURCE_RELIABILITY のクライアント絞り込みは `isSameClient_` 経由（client-match-unify）。

### 2.5 LMDI分解（診断のみ）
- lmdiDecompose_：Πk-1 を kProd/kClient/kOpinion/kAI に厳密加法分解。既定OFF。

### 2.6 AIスコアの予測反映
- readAIResearchScores_ が AI_RESEARCH_STRUCTURED から topic別 final blended score を返す。品質不足topicは中立化。
- 各軸 final score は ±AI_TOPIC_SCORE_ABS_CAP（=50）でクリップ（4軸合計±200 / OUTPUT 17-18行の表示レンジと一致）。
- AI_SCORE_BASIS=momentum のときは AI_SCORE_HISTORY の過去runとの差分（momentum）に切替（既定はlevel）。
  momentum の lookback は `as_of_date` 単位で dedup され、A-9実行回数では動かない。
- AI_SCORE_HISTORY のクライアント絞り込みは `isSameClient_` 経由（client-match-unify）。

### 2.7 年度合計のセマンティクスと2経路の分離（確定）
**(a) 予測・提出経路（A-9 / runForecastFYCore_）— 実績非依存**
- 年度合計(P10/P50/P90)は、12ヶ月すべての予測simから算出する（経過月の実績で固定した着地値ではない＝通年予測の見通し）。
- A-9 は実行直前に `syncSalesFromSalesInput_` で SALES窓を `[fy-4/04, fy/03]` に再整列し、予測窓 `[fy/04, fy+1/03]` と
  同一fyで隣接非重複に確定する。よって全月 `forecast_open` となり実績は混ざらない。
- クライアント名は A-9 入口で正規化される（a9-client-normalize）ため、書き出される FORECAST_SNAPSHOT /
  履歴 / LANDING_FORECAST の client は ACTUAL_EVAL_MONTHLY と同表記になる。

**(b) 検証・学習経路（B-1 → B-2 → B-3、C-1）— 実績を使うが予測は書き換えない**
- B-1：実績を ACTUAL_EVAL_MONTHLY に取り込む（取込時に client は正規化済み）。
- B-2：提出済みの FORECAST_SNAPSHOT（`final_pred`=r[9]）と実績を `[client, target_month]` の複合キーで突合し
  EVAL_LOG / EVAL_COMPARE_MONTHLY に記録。snapshot 側 client も A-9 で正規化済みのため、生文字列突合でも一致する。
- B-3：EVAL_INSIGHTS に外れ要因・次アクションを整理。
- C-1：EVAL_LOG と impact履歴から reliability を学習。

---

## 3. CONFIG パラメータ（読取方式）

### 3.1 読取規約
- readConfigLabelMap_ が CONFIG A:B を読み、configKeyOf_ で「（」または「(」より前をキー化して完全一致マップを作る。
- read*FromConfig_ 系はすべてこのマップ経由。セル番地直読みは廃止。tuneRows に行を挿入しても壊れない。
- ラベル末尾の注記（全角括弧内）は自由に変更可。キー部分が一致すれば読める。
- **const 既定値は CONFIG セル欠落/不読時のフォールバック**。clasp push だけでは既存bookの CONFIG セルは
  更新されないため、const 既定の引き上げ（overlay-cap-raise の 0.40 / 0.05 等）を発効させるには
  A-1 再生成か該当セルの手動修正が必要。

### 3.2 主要キー（抜粋・現行同期）
- AI_WEIGHT / **AI_MAX_ABS_EFFECT（既定0.05＝±5% / 読取上限clampも0.05）**
- AI_MISSING_CONFIDENCE_DEFAULT（既定0.5 / 0で不採用）
- AI_TOTAL_NEUTRAL_THRESHOLD（既定0＝dead-zoneなし）
- AI_QUALITY_NEUTRAL_THRESHOLD / AI_QUALITY_PARTIAL_THRESHOLD
- AI_SCORE_BASIS（level/momentum）/ AI_MOMENTUM_LOOKBACK_QUARTERS / AI_MOMENTUM_MIN_HISTORY
- AI_WEIGHT_PROPOSAL_MIN / AI_WEIGHT_PROPOSAL_MAX
- **QUAL_SUBJECTIVE_MONTHLY_CAP（既定0.40 / 読取上限clampは2.0）** / QUAL_CALIBRATION_ENABLED
- SPOT_BG_* / KNOWN_SPOT_* / SEASONAL_*
- DLM_ENGINE_MODE / DLM_PRIMARY_SPOT_CAP_BASIS / DLM_BACKTEST_* / DLM_LOG_EPSILON_RATE
- RELIABILITY_APPLY_ENABLED（既定1）/ RELIABILITY_R_MIN / R_MAX / SHRINKAGE_K / MIN_SAMPLES / MIN_CHANGE
- POOL_MIN_CLIENTS（既定2）
- LMDI_DECOMPOSITION_ENABLED
- FORECAST_CLOSED_MONTH_MODE（actual/forecast、既定actual。§9参照）
- Vertex AI調査キー：VERTEX_PROJECT_ID / VERTEX_LOCATION / VERTEX_GEMINI_MODEL / VERTEX_DATASTORE_ID /
  VERTEX_SEARCH_LOCATION / VERTEX_SERVING_CONFIG（既定 default_search）/ AI_RESEARCH_ENABLED（既定1）

### 3.3 撤去済みCONFIG行（参考）
config-simplify / config-deadknob-removal / qual-share-const-removal により、担当者の B10 互換参照 /
QUAL_SUBJECTIVE_MAX_SCALE / QUAL_SHARE_* / QUARTERLY_REVIEW_PERIOD_MONTHS / BIAS_CORRECTION_* /
AI_DIRECTION_HIT_* / AI_EFFECT_MIN_MEANINGFUL / DLM_FORECAST_HORIZON は CONFIG から撤去済み。

---

## 4. 学習ループ（四半期レビュー C-1〜C-3）

### 4.1 データ収集（collectQuarterlyReviewData_）
- EVAL_LOG から client一致（`isSameClient_`）・scenario=neutral・constraint_relevant_flag=1 の行を抽出。
- 直近3ヶ月（last3）が揃わなければ ready:false（提案ゼロで正常終了）。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY / AI_SCORE_HISTORY を last3 で結合（いずれも client一致は `isSameClient_`）。

### 4.2 サンプル水増し対策（dedup / open限定）
- **forecast_source 列**：AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY 末尾に `forecast_source`
  （'forecast_open' / 'actual_closed'）を持つ。記録値は `result.sourceByMonth[i]`。
- **open限定**：C-1集計は `forecast_source==='forecast_open'` の行のみ採用。closed後の actual上書き行を自動除外。
- **dedup**：quant側は month単位で最新 run_at の1件、subjective側は (target_month, source_type, source_key) 単位で
  最新 run_at の1件。n が「A-9実行回数」に依存しなくなる。

### 4.3 信頼度提案（generateReliabilityProposals_）
- rHat = clamp(2h, R_MIN, R_MAX)。POOL_PRIOR と shrinkage_k で収縮 → rShrunk。
  rShrunk = clamp((n*rHat + k*rPool) / (n+k), R_MIN, R_MAX)。
- |rShrunk − 現在値| >= MIN_CHANGE のとき提案化。各bookのC-1で source_type別 hit/n を RELIABILITY_EVIDENCE に永続化。

### 4.4 適用（applyQuarterlyProposals）
- QUARTERLY_REVIEW の承認列に従う。承認かつ auto_update_enabled=1 のときのみ SOURCE_RELIABILITY を upsert ＋ 履歴記録。
- 同一 review_id の二重適用は早期終了。

### 4.5 AI_weight 提案（generateQuarterlyProposals_）
- AI方向一致率と mean|kAI-1| から ai_weight_override 提案を生成。
- 現在値の読取は **空欄=未設定→CONFIG既定、明示数値→その値**（calibration-blank-override-fix と同方針）。

### 4.6 calibration override の空欄/明示0 区別（calibration-blank-override-fix）
- `calOverrideNum_(v)`：`''`/`null`/`undefined` は null（未設定→無視）、数値は isFinite なら採用。
- 明示的に 0 を入れた場合は 0 が採用され、意図的にAIを止める運用は従来どおり有効。

---

## 5. 検証ポリシー（KPI）
- 計画用単一値＝P50。P10/P90は説明帯（hard gateではない）。
- 制約：annual_abs_error_rate <= 10% / half_wape <= 12%（将来目標10%）/ over-forecast rate <= 5%。
- 診断：月次APE、Q差分、定量寄与率、主観オーバーレイ率（AI除く）、AI寄与率、Known Spot寄与率、レンジ逸脱数。
- レンジ逸脱月（actualがP10-P90外）はB-3で追加調査を必須化。
- 1 client = 1 book。
- ※ overlay-cap-raise（cap 0.40 / AI ±5%）適用後は、主観/AIオーバーレイのあるbookで P50 が変化しうるため、
  本KPIの再確認（特に over-forecast rate の悪化有無）を1サイクル行うこと。

---

## 6. ログ／永続シート
- ユーザ表示シート：GUIDE / CONFIG / SALES_INPUT / SALES_MONTHLY / AI_RESEARCH（サマリービュー）/
  PRODUCT / CLIENT / OPINIONS / DEV_SPOT / OUTPUT / DASHBOARD（hideNonUserSheets_ で制御）。
- 内部管理（原則非表示）：RUN_LOG / FORECAST_SNAPSHOT / PROCESS_STATUS / AI_SCORE_HISTORY /
  AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY / CALIBRATION_STATE / CALIBRATION_HISTORY /
  QUARTERLY_REVIEW / QUARTERLY_REVIEW_LOG / DLM_STATE / BACKTEST_REPORT / SOURCE_RELIABILITY /
  RELIABILITY_EVIDENCE / POOL_PRIOR / POOL_REGISTRY / POOL_AGGREGATION_LOG / LANDING_FORECAST /
  AI_RESEARCH_STRUCTURED / AI_RESEARCH_TASK_LOG / AI_RESEARCH_RAW。
- **FORECAST_SNAPSHOT は15列**（snapshot_id / run_date / client / target_month / scenario / base_pred /
  subjective_adj / ai_adj / deterministic_adj / **final_pred(index 9)** / confidence_interval_lower /
  confidence_interval_upper / key_factors_json / subjective_input_date / calibration_applied_json）。
  client 列は A-9 入口で正規化された値が書かれる（a9-client-normalize）。これにより B-2 の生文字列複合キー
  突合（client+target_month）が ACTUAL_EVAL_MONTHLY と一致する。
- **AI_RESEARCH_RAW（旧 WEB + EXTERNAL の統合）**：headers は
  `client / as_of_date / axis / topic / direction / magnitude / uncertainty / relative_position /
  evidence / frozen_flag / frozen_at / note`。`axis`=web/rag で生応答の出所を区別する。
  `frozen_flag`/`frozen_at` は将来用の予約列（現状未使用 / 既定0・空）。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY 末尾に forecast_source（§4.2）。これらの client も
  A-9 入口で正規化された値で書かれる。

---

## 7. POOL_PRIOR 横断集約（3c-3c）
### 7.1 構成
- ハブbook初期化：adminSetupPoolHub（手動）。POOL_REGISTRY / POOL_AGGREGATION_LOG / POOL_PRIOR を作成。
- 集約実行：adminAggregatePoolPriorAcrossBooks（手動）。REGISTRYの enabled=1 book を openById で開き、
  各bookの RELIABILITY_EVIDENCE から source_type別の生 hit/n を読み、集約してPOOL_PRIORへ書き込み＋fan-out。

### 7.2 集約ロジック
- 集約入力：生 hit/n をプール（rShrunk再プール禁止）。
- 粒度：reliability:{source_type}（factor_product/factor_client/opinion/ai_topic）。person系キーは横断不可。
- pooled_value = clamp(2 * Σhit/Σn, R_MIN, R_MAX)。precision = SHRINKAGE_K（固定）。
- フェイルセーフ：n_clients < POOL_MIN_CLIENTS → written=false / reason=min_clients。Σn < MIN_SAMPLES → reason=min_samples。
  単一book読取失敗は status=excluded でログ記録し全体は止めない。

### 7.3 予測への波及
- POOL_PRIOR は C-1 の提案収縮（rShrunk）に効くのみ。予測値そのものは変えない（集約直後のA-9はno-op）。

---

## 8. バージョン整合と適用順序
- VERSION / BUILD_STAGE / 設計書版 / 手動チェックリストは各リリースで同期する。
- 現行コードは VERSION='2.3.37-dev' / BUILD_STAGE='overlay-cap-raise'。
  DLM_BUILD_STAGE は 'v8-step3c3c-1'（DLMロジック無変更のため据え置き）。
- 本改訂（設計書 v2.3.6）は overlay-cap-raise（2.3.37 / 主観月次cap 0.20→0.40・AI上限 ±3%→±5%）を
  ドキュメントへ反映する doc sync。**従来の doc sync と異なり予測値が変わる数値変更**であり、既存bookは
  A-1再生成または CONFIG セルの手動修正で発効する。
- 設計書ドリフトは既知の再発リスク。リリースごとに doc sync を独立タスクとして扱う。コードが
  2.3.33→2.3.37 と4増分先行していた点に留意し、今後は stage 確定のたびに doc を同期する。

---

## 9. 通年予測モード（FORECAST_CLOSED_MONTH_MODE）【確定：A=通年予測 / コード変更なし】
年間総計のセマンティクスは **「通年（12ヶ月）すべて予測の見通し」（選択肢A）** で確定。

### 9.1 確定した仕様
- **年度合計(P10/P50/P90)は、FORECAST_CLOSED_MONTH_MODE の値にかかわらず、常に12ヶ月すべての予測simから算出する。**
  closed月上書きは sim配列に触れず月次表示用配列にしか作用しない。
- よって本トグルは **「月次表示の見せ方」だけを切り替える**もので、計画の主数値（年度合計）の意味は変えない。

### 9.2 モード別の月次表示
- `actual`（既定）：closed月は実績で上書き表示（objOnly/mixed/regTotal）。
- `forecast`：closed月も予測のまま表示（通年予測モード）。経過月の実績は ActualClosed 列に参考併記。

### 9.3 closed月上書きの発火条件と、実運用での非発火
- actualモードの上書きは `sourceByMonth[i]==='actual_closed'` の月のみに作用する。
- A-9 経路では、timing（実績前に回す）と structure（syncSalesFromSalesInput_ による窓非重複）の二重で
  発火条件に到達しない（全月 forecast_open）。

### 9.4 回帰確認（1回のみ）
- OUTPUT「（参考）内訳とメモ（P50比較）」表の `ForecastSource` 列が、A-9 実行後に全行 `forecast_open` であることを1回確認。

### 9.5 表示位置と OUTPUT 表示整形（output-note-relocate / output-display-tidy）
- 年度合計行と月次表ヘッダの間は短い1行注記のみ。「年度≠月次合算」の長文説明は混合セクションの Scenario Split の下に配置。
- output-display-tidy（2.3.30）：OUTPUT 6行目は実警告があるときのみ赤字。row10 は F列。21〜23行は折り返し・幅縮小。
  初期設定ダイアログの黄色 notice は撤去。いずれも表示位置/書式のみで数値は不変。

---

## 10. AI調査パイプライン（Vertex AI / A-4 = runVertexAIResearch）

### 10.1 全体像
A-4はメーカー名（CONFIG!B2）について、4 topic（Market/Competitor/Channel/DX）ごとに以下を実行する：
1. **grounded web検索**（callVertexGeminiGrounded_）：Gemini + googleSearch（不可時 googleSearchRetrieval へフォールバック）。
2. **RAG検索**（callVertexSearchRAG_）：Vertex AI Search データストアを検索。DATASTORE_ID/SEARCH_LOCATION 未設定時はskip（web-only）。
3. **構造化**（callVertexGeminiStructured_）：web結果をevent行、RAG結果をbenchmark行として JSON 構造化。**temperature=0 固定**。

### 10.2 readiness 判定
- readVertexConfig_ が geminiReady（projectId+location+geminiModel）と ragReady（datastoreId+searchLocation）を分離。
- A-4の入口判定は geminiReady。RAG未設置でも web-only で実行できる。
- AI_RESEARCH_ENABLED=0 のときは「無効」メッセージで終了。geminiReady=false のときは「必須設定未入力」エラーで停止。

### 10.3 スコア化（event / benchmark）
- **event_score**：`direction符号 × (impact_score/100) × 50 × confidence` を ±50 にclamp。impact_score は 0〜100 の影響の大きさ（50は中立ではない）。
- **benchmark_score**：`(relative_percentile-50) × relative_confidence × quality倍率(high1/medium0.75/low0.5)` を ±50 にclamp。
  benchmark 行は相対位置の根拠があるときだけ書く（プレースホルダ禁止）。
- readAIResearchScores_ が topic別に blend：Market=0.65:0.35 / **Competitor=0.70:0.30** / Channel=0.65:0.35 / DX=0.50:0.50。
- 各topic final score は ±50（AI_TOPIC_SCORE_ABS_CAP）でclamp。

### 10.4 品質中立化と「根拠なき中立化／捏造ゼロ」の是正（ai-score-degroundless）
- 総合 dead-zone 中立化は撤去（AI_TOTAL_NEUTRAL_THRESHOLD 既定0）。上限は ±AI_MAX_ABS_EFFECT（±5%）のみ。
- event_only への構造的 ×0.5 ペナルティを撤去。信号の弱さは coverage 由来 qualityMultiplier のみで反映。
- honest-zero：web が真に neutral かつ benchmark 根拠なしのトピックは 0 になりうる（捏造しない設計どおり）。

### 10.5 失敗時の扱い（フェイルセーフ）
- web/rag/structure の各段で失敗してもrun全体は止めず、AI_RESEARCH_TASK_LOG に per-topic per-aspect 行を残す。
- 全topicで構造化行が0件なら AI_RESEARCH_STRUCTURED は上書きせず既存行を保持。1件以上で全置換し再描画。

### 10.6 ランタイムシート（ai-research-raw-merge / raw-migrate-header-match）
- AI_RESEARCH_STRUCTURED：A-1で先行作成。A-9が読む唯一の入口。
- Vertex実行時に遅延作成されるのは AI_RESEARCH_TASK_LOG / AI_RESEARCH_RAW の2枚。
- 旧2枚→RAW 移送（migrateLegacyAIResearchRawSheets_）は**ヘッダ名マッチ＋シート名由来 axis**（raw-migrate-header-match）。
  旧2枚の列順が RAW と異なっても同名ヘッダ列へ入り、axis はシート名（WEB→web / EXTERNAL→rag）から確定。
  frozen_flag は旧側に無ければ 0、frozen_at は空で補完。移送は一度だけ（旧タブ削除＝冪等）。

### 10.7 CONFIGキー（既定値）
- VERTEX_PROJECT_ID = forecast-agent-498907 / VERTEX_LOCATION = global / VERTEX_GEMINI_MODEL = gemini-3.1-pro-preview
- VERTEX_DATASTORE_ID = fujikeizai-portfolio-2025 / VERTEX_SEARCH_LOCATION = global / VERTEX_SERVING_CONFIG = default_search
- AI_RESEARCH_ENABLED = 1 / 必要OAuthスコープ：cloud-platform / script.external_request

### 10.8 RAGクエリのフレーム整合（rag-query-frame-align）
- `buildRagQuery_` は外部観測可能な「支援需要に影響しうる外部環境」に整合。topic 別の日本語展開語で外部環境を狙う。
  英語topic素語の連結や出版社名（弁別力なし）は撤去。本変更は A-4 側のクエリのみ。

---

## 11. AI調査サマリービュー（AI_RESEARCH シート）
A-4実行時に writeAIResearchSummaryView_ が AI_RESEARCH シートを再描画する（ユーザ表示シート）。

### 11.1 構成（3段）
- ① topic別サマリー（要約文）：report_text を topic単位で表示。全topicで空ならフェイルセーフ文。
- ② AIスコア サマリー（4軸）：topic別 Final Score / event_score / benchmark_score / 最新as_of / 備考。
- ③ スコア根拠（event / benchmark 明細）。

### 11.2 再描画の不変条件
- A-4を複数回実行しても追記でなく再描画（重複しない）。A-4失敗（outRows空）時は上書きせず前回内容を保持。
- 本ビューは表示専用で、予測コア（A-9）には影響しない。

---

## 12. 既知の latent issue（非ブロッキング / 別スコープで対応）

現状、ブロッキングな latent issue は無い。

解消済みメモ：
- **client名マッチの不整合**は client-match-unify（2.3.32）で解消。`isSameClient_` を `normalizeClientName_`
  経由に統一し、読取側の突合を全て `isSameClient_` に集約。さらに a9-client-normalize（2.3.33）で A-9 入口の
  永続クライアントを正規化し、FORECAST_SNAPSHOT 等と ACTUAL_EVAL_MONTHLY の表記を一致させた。
- **AI_RESEARCH_RAW 旧データ移送の列整合**は raw-migrate-header-match（2.3.31）で解消（ヘッダ名マッチ＋シート名由来 axis）。
- **FORECAST_SNAPSHOT の三角測量系 vestigial カラム**は snapshot-vestigial-removal で物理削除済み。
- LANDING_FORECAST / EVAL_LOG の再実行追記増殖は複合キー upsert 化で解消済み。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY は追記のままで、重複排除は C-1 読取側で行う。
- shadow mode 表示の `opsQuantOnly` エイリアス問題は objOnly 独立コピー化で解消済み。
- 旧Gem手動貼付経路・孤立関数（adminShowCloneGuide / importPastSalesToSalesTab）は物理削除済み。
- calibration override の空欄/明示0 取り違えは calibration-blank-override-fix で解消済み。

将来検討（実害なし / 別スコープ）：
- `normalizeClientName_` の正規化辞書が現状 ｳﾞｨｱﾄﾘｽ系のハードコードのみ。新メーカーで表記ゆれが出た場合の
  汎用正規化（全角/半角・法人格表記の一般則化）は未対応。必要が生じた時点で別スコープで対応する。
- `writeOutputFY_` の `warn_coerced` 判定（`coerceMatch`）は、Gem手動貼付経路撤去後の Vertex warning summary には
  `warn_coerced=` が含まれないため常に空振り（coerceCount=0 で if ブロック非発火）の死にコード。`match` が null を
  返すだけで無害だが、整理する場合は別スコープのデッドコード削除として扱う。

---

この v2.3.6 は、annual-forecast-mode（FORECAST_CLOSED_MONTH_MODE）の方針を「A=通年予測」とする正典と、
client-match-unify（2.3.32）/ a9-client-normalize（2.3.33）の正規化統一を維持しつつ、overlay-cap-raise
（2.3.37 / 主観月次cap 0.20→0.40・AI上限 ±3%→±5%）を現行実装へ同期したドキュメント改訂である。本改訂は
表示・構造のみの従来 doc sync とは異なり、主観/AIオーバーレイのあるbookで P10/P50/P90 が変わる数値変更であり、
既存bookは A-1再生成または CONFIG セルの手動修正で発効する。年度合計は常に12ヶ月すべての予測simから算出する
通年予測であり、実績は検証・学習経路（B/C）でのみ使用する。ブロッキングな残存 latent issue は無い。
