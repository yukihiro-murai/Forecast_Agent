# 売上予測スクリプト 設計書 v2.3.2（gem-path物理削除 / objOnly独立コピー化 反映版）

## 0. 文書の目的
この設計書は、実装（Forecast_Agent.js）の現行挙動を正確に記述する。
旧v7の「三角観測 w1/w2/w3/w4 + 逆sMAPE重み更新」は実装に存在しないため撤去済み。
3目的は不変：(1)予測精度向上 (2)透明化（根拠明示・再現性） (3)学習性（継続改善）。

対象実装：`Forecast_Agent.js`（VERSION='2.3.11-dev' / BUILD_STAGE='v8-a9-toast-fix'）
設計書版：v2.3.2（v2.3.1 からのドキュメント改訂。**コード変更は伴わない**。
gem-path物理削除 / objOnly独立コピー化までの設計記述を現行コード版へ同期）

### 0.0 v2.3.1 → v2.3.2 の更新点（この改訂で反映したもの）
- **objonly-dealias**：`runForecastFYCore_` の `objOnly` を `quantOnly` の独立コピー（`.slice()`）化。
  `closedMonthMode='actual'` の実績上書きが `quantOnly`（KPI診断）/ `opsQuantOnly`（DLM比較の旧Ops参照）を
  巻き込み変異させない配線になった。現行データフローでは actual_closed 月が立たないため挙動中立であり、
  予測コア（OUTPUTのP10/P50/P90）は不変。
- **gem-path-removal**：旧Gem手動貼付経路の関数群（`generateAIResearchTemplate` / `parseAIResearchPaste_` /
  `showPromptPreviewDialog_` / `buildAiParseWarningText_` / `pushInvalidSample_`）をコードから物理削除。
  A-4 は `runVertexAIResearch`、A-9 は `countAIResearchStructuredRows_` で件数把握する経路に確定。
  これは到達不能な旧経路の撤去であり、予測コア（OUTPUTのP10/P50/P90）は不変。

### 0.0' v2.3 → v2.3.1 の更新点（参考・再掲）
- **§9 annual-forecast-mode の方針を確定**。年度合計(P10/P50/P90)は「実績込みの着地予測」ではなく
  **「通年（12ヶ月）すべて予測の見通し」**（=選択肢A）で確定。これは現行実装の挙動そのものであり、
  **コード変更は発生しない**。§9 から「方針確認待ち」表記を撤去し、確定仕様として記述する。
- **「予測・提出経路」と「検証・学習経路」の分離を明文化（§2.7）**。予測（A-9）は実績が入る前に回す
  運用であり、かつ構造的にも予測窓と実績窓が非重複なため実績非依存。実績を使うのは検証・学習
  （B-1〜B-3 / C-1）だけで、そこでも提出済み予測を実績で書き換えることはしない（突合・学習のみ）。
- **closed月上書き（actualモード）の実運用での非発火を確定（§9.3）**。timing（実績前に回す）と
  structure（窓非重複）の二重で発火条件に到達しないことを正典化し、§9.2 の旧論点を解消する。

### 0.0'' v2.3 で反映済みの更新点（参考・再掲）
- **AI調査をVertex AI自動実行へ移行**。旧A-4「Gemプロンプト生成→TSV手動貼付→parse」はメニューから廃止し、
  A-4は `runVertexAIResearch`（Vertex grounded web検索 + Vertex AI Search RAG + 構造化出力）に置き換え済み（§10）。
  現行では当該関数群もコードから物理削除済み（メニュー非掲載に留まらず定義・参照とも0件）。
- **AI調査サマリービュー（AI_RESEARCH）を正典化**。`writeAIResearchSummaryView_` が3段ビューを再描画する（§11）。
- **新規シート3枚**：`AI_RESEARCH_TASK_LOG` / `AI_RESEARCH_WEB` / `AI_RESEARCH_EXTERNAL`（§6・§10）。
- **ランタイムシート初期化を正典化**（§1.1・§10.5）。
- **FORECAST_REPORT は撤去済み**（§6）。

---

## 0.5 v7からの主要な乖離（同期のために明記）
- **三角観測は廃止済み**。実装は「単一Opsモデル（線形トレンド×季節指数）＋残差Monte Carlo」が定量土台。
- **逆sMAPE重み更新（w_i）は存在しない**。重み更新ロジックは実装されていない。
  （FORECAST_SNAPSHOTに w1〜w4 列は残るが固定値の記録列であり、更新ロジックは無い＝vestigial。）
- **DLM（対数空間の状態空間モデル）が追加**され、CONFIGの DLM_ENGINE_MODE（off/shadow/primary）で制御。
- **主観は乗算係数（kProd/kClient/kOpinion/kAI）として反映**し、月次cap（QUAL_SUBJECTIVE_MONTHLY_CAP）でクリップ。
  v6の「主観キャリブレータ（オーバーレイ率ターゲット探索）」は撤去済み（cap pass-through）。
- **AIは予測係数ではなく、benchmark/event blend のスコアとして kAI に限定反映**。品質不足時は中立化（kAI=1.0）。
- **AI調査の取得経路が Vertex AI へ移行**。旧Gem手動貼付（TSV）経路はメニュー廃止後、
  現行では定義・参照とも0件（§10）。
- **信頼度（SOURCE_RELIABILITY）** が追加され、各ソース（factor_product/factor_client/opinion/ai_topic）の
  寄与に reliability_r を乗じる。CONFIG RELIABILITY_APPLY_ENABLED で制御（既定ON）。
- **LMDI分解** が追加され、主観乗算 Πk-1 を因子別に厳密加法分解（CONFIG LMDI_DECOMPOSITION_ENABLED、既定OFF）。
- **POOL_PRIOR のクライアント横断集約** が追加（中央集約book→各bookへfan-out / 手動実行）。

---

## 1. 実装スコープ

### 1.1 現行実装済み
- データ取込：外部実績SS → SALES_INPUT（A-2）→ SALES_MONTHLY 48ヶ月横持ち BASE/SPOT/TOTAL（A-3）
- 主観入力：PRODUCT / CLIENT / OPINIONS / DEV_SPOT（A-5〜A-8）
- AI調査（Vertex AI自動）：A-4 `runVertexAIResearch`。grounded web検索 + Vertex AI Search RAG +
  構造化出力 → AI_RESEARCH_STRUCTURED へ記録、AI_RESEARCH（サマリービュー）へ再描画（§10・§11）
- 予測実行：runPhase1Forecast（A-9）。OUTPUT / FORECAST_SNAPSHOT 更新
- ダッシュボード：updatePhase1Dashboard（A-10）
- 検証：実績取込（B-1）→ EVAL_LOG / EVAL_COMPARE_MONTHLY（B-2）→ EVAL_INSIGHTS（B-3）
- 四半期レビュー：runQuarterlyReview（C-1）→ applyQuarterlyProposals（C-2）→ ログ閲覧（C-3）
- 信頼度：SOURCE_RELIABILITY 適用（reliabilityApply）＋ C-1のreliability提案＋RELIABILITY_EVIDENCEへの raw hit/n 永続化
- 横断プール：POOL_PRIOR のクライアント横断集約（adminSetupPoolHub / adminAggregatePoolPriorAcrossBooks）
- **シート初期化**：AI_RESEARCH_STRUCTURED は A-1（`setupForecastBook` の order[] と `buildPhase1Sheets_`）で
  先行作成される。Vertex実行時に真に遅延作成されるのは AI_RESEARCH_TASK_LOG / AI_RESEARCH_WEB /
  AI_RESEARCH_EXTERNAL の3枚のみ。`ensureAIResearchRuntimeSheets_` は AI_RESEARCH_STRUCTURED については
  冪等に存在確認・ヘッダ整合する。
  POOL_PRIOR / POOL_REGISTRY / POOL_AGGREGATION_LOG / DLM_STATE / BACKTEST_REPORT / LANDING_FORECAST も
  管理関数や予測実行時に必要に応じて `getOrCreateSheet_` で作成する。

### 1.2 既定OFFのシャドウ機能（検証待ち）
- DLM_ENGINE_MODE = off（shadow/primaryはCONFIGで切替可）
- LMDI_DECOMPOSITION_ENABLED = 0
- FORECAST_CLOSED_MONTH_MODE = actual（**月次表示のみ**を切り替えるトグル。年度合計の算出には影響しない。詳細は§9）
- AI_SCORE_BASIS = level（momentumに切替でAIスコアを相対位置の変化として扱う）
- ※ RELIABILITY_APPLY_ENABLED は既定1（ON）。ただしSOURCE_RELIABILITY空ならno-op（予測不変）。
- ※ AI_RESEARCH_ENABLED は既定1。0でA-4のVertex調査をスキップし、AI_RESEARCH_STRUCTUREDの既存行のみ参照。

### 1.3 未実装（今後）
- 構造変化検知（CUSUM/Bai-Perron）、分位点回帰の本格適用、学習窓の動的最適化。
- POOL_PRIOR集約のprecision導出の経験ベイズ化（現状はΣhit/Σn加重＋固定precision=SHRINKAGE_K）。
- person系ソースの横断プール（現状は普遍キーのsource_type単位まで）。

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
- kAI = 1 + clamp(Σ(topicスコア × reliability) × AI_WEIGHT, ±AI_MAX_ABS_EFFECT)。AI合計が中立閾値未満なら 1.0。
- Monte Carlo（N_SIM=1000）で各simの total を生成し、月次P10/P50/P90を取る。
- 主観差分は月次cap（QUAL_SUBJECTIVE_MONTHLY_CAP、既定0.20）でクリップ（capHitを診断記録）。
- **kProd全月1.0のフェイルセーフ（旧hard throw撤去）**：PRODUCTに有効行があるのに
  kProdが全月1.0（直近12ヶ月closed BASE実績の無い製品＝weight=0）の場合、旧実装はthrowでA-9を停止していた。
  現行は **throwせず警告**（`productWeightWarning`）に置換し、A-9を完走させる。
  警告は weight=0 となった製品名を列挙し、OUTPUT上部の警告ブロックとRUN_LOG note（`prodw=...`）に残す。
  原則：中立（kProd=1.0）で予測は成立するため、エラーでメニュー実行をブロックしない。

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
  final blended score（benchmark/event blend）を返す。品質不足topicは中立化（multiplier=0/0.5）。
- AI_SCORE_BASIS=momentum のときは AI_SCORE_HISTORY の過去runとの差分（momentum）に切替（既定はlevel）。
- これがA-4（Vertex）で書かれた行を読む唯一の入口。A-9はVertex APIを呼ばず、シート上の構造化行のみ参照する。

### 2.7 年度合計のセマンティクスと2経路の分離（確定）
本ツールには性質の異なる2つの経路があり、実績の扱いが明確に違う。混同しないこと。

**(a) 予測・提出経路（A-9 / runForecastFYCore_）— 実績非依存**
- 入力は SALES_MONTHLY（過去48ヶ月のBASE/SPOT履歴）と主観入力・AI調査・DEV_SPOT。
- 予測対象は予測FYの12ヶ月（fy/04〜fy+1/03）。出力は OUTPUT と FORECAST_SNAPSHOT。
- **年度合計(P10/P50/P90)は、12ヶ月すべての予測simから算出する**（mixedは `aggregateAnnualSim_(totalCalibratedSimByMonth)`、
  客観のみは quant側simの年次集約）。経過月の実績で固定した着地値ではない＝**通年予測の見通し**（§9で確定した選択肢A）。
- 運用上、この経路は会社公式予測を役員会に提出する前に回すため、実績が出る前に実行される。
  加えて構造的にも、A-9 は実行直前に `syncSalesFromSalesInput_` で SALES窓を `[fy-4/04, fy/03]` に再整列し、
  予測窓 `[fy/04, fy+1/03]` とは同一fyで隣接非重複に確定する。よって `forecastMonthIndexesInSales` は全て -1、
  `sourceByMonth` は全月 `forecast_open` となり、予測窓に実績が混ざる余地はない。

**(b) 検証・学習経路（B-1 → B-2 → B-3、C-1）— 実績を使うが予測は書き換えない**
- B-1：実績を ACTUAL_EVAL_MONTHLY に取り込む。
- B-2：提出済みの FORECAST_SNAPSHOT と実績を突合し、EVAL_LOG / EVAL_COMPARE_MONTHLY に誤差を記録する。
- B-3：EVAL_INSIGHTS に外れ要因・次アクションを整理する。
- C-1：EVAL_LOG と impact履歴（AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY）から reliability を学習する。
- ここで実績は「答え合わせと学習の材料」であり「予測の入力」ではない。**提出済みの予測値を実績で
  retroactive に書き換えることはしない**。学習結果は次回以降のA-9の係数（reliability等）に反映される。

この分離により、「役員会に出す通年予測（実績非依存）」と「随時の検証・学習（実績使用）」は
別ワークフローとして両立する。年度合計が常に通年予測であること（a）と、実績を使った継続改善が回ること（b）は
矛盾しない。

---

## 3. CONFIG パラメータ（読取方式）

### 3.1 読取規約
- readConfigLabelMap_ が CONFIG A:B を読み、configKeyOf_ で「（」または「(」より前をキー化して完全一致マップを作る。
- readModelTuningFromConfig_ / readDlmEngineMode_ / readReliabilityApplyEnabled_ /
  readLmdiDecompositionEnabled_ / readDlmPrimarySpotCapBasis_ / readForecastClosedMonthMode_ /
  readAiScoreBasis_ / readVertexConfig_ はすべてこのマップ経由。
- セル番地直読みは廃止。tuneRows に行を挿入しても壊れない。
- ラベル末尾の注記（全角括弧内）は自由に変更可。キー部分が一致すれば読める。

### 3.2 主要キー（抜粋）
- AI_WEIGHT / AI_MAX_ABS_EFFECT / AI_TOTAL_NEUTRAL_THRESHOLD / AI_QUALITY_*_THRESHOLD
- AI_SCORE_BASIS（level/momentum）/ AI_MOMENTUM_LOOKBACK_QUARTERS / AI_MOMENTUM_MIN_HISTORY
- QUAL_SUBJECTIVE_MONTHLY_CAP / QUAL_CALIBRATION_ENABLED
- SPOT_BG_* / KNOWN_SPOT_* / SEASONAL_*
- DLM_ENGINE_MODE / DLM_PRIMARY_SPOT_CAP_BASIS / DLM_BACKTEST_*
- RELIABILITY_APPLY_ENABLED（既定1）/ RELIABILITY_R_MIN / R_MAX / SHRINKAGE_K / MIN_SAMPLES / MIN_CHANGE
- POOL_MIN_CLIENTS（横断集約の最低クライアント数、既定2）
- LMDI_DECOMPOSITION_ENABLED
- FORECAST_CLOSED_MONTH_MODE（actual/forecast、既定actual。§9参照）
- **Vertex AI調査キー（§10）**：VERTEX_PROJECT_ID / VERTEX_LOCATION / VERTEX_GEMINI_MODEL /
  VERTEX_DATASTORE_ID / VERTEX_SEARCH_LOCATION / AI_RESEARCH_ENABLED（既定1）

---

## 4. 学習ループ（四半期レビュー C-1〜C-3）

### 4.1 データ収集（collectQuarterlyReviewData_）
- EVAL_LOG から client一致・scenario=neutral・constraint_relevant_flag=1 の行を抽出（ヘッダidx参照）。
- 直近3ヶ月（last3）が揃わなければ ready:false（提案ゼロで正常終了）。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY / AI_SCORE_HISTORY を last3 で結合。

### 4.2 サンプル水増し対策（dedup / open限定）
背景：GUIDEは「確認→修正→再実行を前提とする」と明記しており、同じ月に対しA-9を複数回押す運用が正規。
旧実装はA-9のたびに履歴を追記（dedupなし）し、C-1が1行ずつ `n += 1` していたため、
入力を変えずA-9を数回押すだけで n が膨張し、提案が「高」信頼度に化ける欠陥があった。
さらに、closed月の actual上書きにより `pred_p50_quant_only` に actual が記録され、
C-1の month単位 last-write-wins が最新（actual上書き後）の行を拾って surprise=0 で当該月が脱落していた。

対策（現行 computeReliabilityHitStats_）：
- **forecast_source 列**：AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY の末尾に
  `forecast_source`（'forecast_open' / 'actual_closed'）を持つ。記録値は runForecastFYCore_ の
  result.sourceByMonth[i] をそのまま書く（既存列順は不変）。
- **open限定**：C-1集計は `forecast_source==='forecast_open'` の行のみ採用。
  これにより closed後の actual上書き行（surprise=0）が自動除外され、月脱落が解消する。
  forecast_source が空/欠落の旧（汚染）行も同フィルタで自動除外される。
- **dedup**：
  - quant側（AI_IMPACT_HISTORY）：month単位で最新 run_at の1件のみ採用（latestQuantByMonth）。
  - subjective側（SUBJECTIVE_IMPACT_HISTORY）：(target_month, source_type, source_key) 単位で
    最新 run_at の1件のみ採用（latestSubjectiveByUnit）。
  これにより n が「A-9実行回数」に依存しなくなり、実行回数非依存の評価になる。
- 触らない箇所：evalActualByMonth（EVAL_LOG由来のactual）は month単位 last-write-wins のままでよい
  （actualは確定値なので同月重複でも値は同じ）。予測コアの計算は一切変更していない。

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

---

## 5. 検証ポリシー（KPI）
- 計画用単一値＝P50。P10/P90は説明帯（hard gateではない）。
- 制約：annual_abs_error_rate <= 10% / half_wape <= 12%（将来目標10%）/ over-forecast rate <= 5%。
- 診断：月次APE、Q差分、定量寄与率、主観オーバーレイ率（表示のみ・制御目標ではない）、Known Spot寄与率、レンジ逸脱数。
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
  AI_RESEARCH_STRUCTURED / **AI_RESEARCH_TASK_LOG / AI_RESEARCH_WEB / AI_RESEARCH_EXTERNAL**。
- **FORECAST_REPORT は撤去済み**（sheet-consolidationで物理削除）。
- FORECAST_SNAPSHOT 末尾に calibration_applied_json を保存。
  ※ FORECAST_SNAPSHOT の三角測量系カラム（linear/robust/regime/simulation_pred / w1〜w4）は
  vestigial（更新ロジック無し・固定記録列）。撤去は別スコープで検討（§12）。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY 末尾に forecast_source（§4.2）。

---

## 7. POOL_PRIOR 横断集約（3c-3c）
v1.8では未実装と記載していたが、現行実装で完了している。

### 7.1 構成
- ハブbook初期化：adminSetupPoolHub（メニュー非掲載 / スクリプトエディタから手動）。
  POOL_REGISTRY / POOL_AGGREGATION_LOG / POOL_PRIOR を作成。
- 集約実行：adminAggregatePoolPriorAcrossBooks（同・手動）。
  REGISTRYの enabled=1 book を openById で開き、各bookの RELIABILITY_EVIDENCE から
  source_type別の生 hit/n を読み、集約してPOOL_PRIORへ書き込み＋各bookへfan-out。

### 7.2 集約ロジック
- 集約入力：生 hit/n をプール（rShrunk再プール禁止）。
- 粒度：reliability:{source_type}（factor_product/factor_client/opinion/ai_topic）。
  person系キーは横断不可（書かない）。ai_topicは普遍キーのため将来 topic単位拡張余地あり。
- pooled_value = clamp(2 * Σhit/Σn, R_MIN, R_MAX)。precision = SHRINKAGE_K（固定）。
- フェイルセーフ：
  - n_clients < POOL_MIN_CLIENTS → written=false / reason=min_clients（POOL_PRIORに書かない＝空のまま）。
  - Σn < MIN_SAMPLES → written=false / reason=min_samples。
  - 単一book読取失敗（権限/不存在）は status=excluded でログ記録し、全体は止めない（exclusion-with-logging）。
- 監査：POOL_AGGREGATION_LOG に per-book 行（ok/excluded/empty/no_columns）と
  per-scope 行（written/reason/n_clients/Σhit/Σn/pooled_value）を1 run_id で記録。

### 7.3 予測への波及
- POOL_PRIOR は C-1 の提案収縮（rShrunk）に効くのみ。予測値そのものは変えない（集約直後のA-9はno-op）。
- POOL_MIN_CLIENTS未満で空のまま据え置く設計は、auditability（1.0で埋めない）を保つため。

---

## 8. バージョン整合と適用順序
- VERSION / BUILD_STAGE / 設計書版 / 手動チェックリストは各リリースで同期する。
- 現行コードは VERSION='2.3.11-dev' / BUILD_STAGE='v8-a9-toast-fix'。
  DLM_BUILD_STAGE は 'v8-step3c3c-1'（DLMロジック無変更のため据え置き）。
- 本改訂（設計書 v2.3.2）は **コード変更を伴わない**。gem-path物理削除 / objOnly独立コピー化までの
  実装状態をドキュメントに反映する doc sync である。
- この設計書は §1〜§7・§9〜§11 の確定機能を記述対象とする。
- 設計書ドリフトは既知の再発リスク。リリースごとに doc sync を独立タスクとして扱う。

---

## 9. 通年予測モード（FORECAST_CLOSED_MONTH_MODE）【確定：A=通年予測 / コード変更なし】
本機能は実装済みであり、年間総計のセマンティクスを **「通年（12ヶ月）すべて予測の見通し」（選択肢A）** で確定した。
これは現行実装の挙動そのものであるため、確定に伴うコード変更は発生しない。

### 9.1 確定した仕様
- **年度合計(P10/P50/P90)は、FORECAST_CLOSED_MONTH_MODE の値にかかわらず、常に12ヶ月すべての予測simから算出する。**
  実装上、年度合計は sim配列（mixedは `totalCalibratedSimByMonth`、客観は quant側sim）の年次集約（`aggregateAnnualSim_`）に
  percentile を取ったものであり、`FORECAST_CLOSED_MONTH_MODE` の closed月上書きは sim配列に触れず、
  月次表示用の配列（objOnly/mixed/regTotal の月次値）にしか作用しない。
- よって本トグルは **「月次表示の見せ方」だけを切り替えるもの**であり、計画の主数値（年度合計）の意味は変えない。

### 9.2 モード別の月次表示
- `FORECAST_CLOSED_MONTH_MODE = actual`（既定）：closed月は実績で上書き表示（objOnly/mixed/regTotal）。
- `FORECAST_CLOSED_MONTH_MODE = forecast`：closed月も予測のまま表示（通年予測モード）。
  経過月の実績は ActualClosed 列に参考併記し、予測値には混ぜない。
- いずれのモードでも、年度合計は §9.1 のとおり12ヶ月すべての予測simから算出される（経過実績で固定した着地値ではない）。

### 9.3 closed月上書きの発火条件と、実運用での非発火
- actualモードの上書きは `sourceByMonth[i]==='actual_closed'` の月のみに作用する。
- `actual_closed` は `getForecastContext_` 上、「forecast月が SALES_MONTHLY ヘッダに存在し
  （forecastMonthIndexesInSales>=0）、かつ lastClosedMonthStart 以前」のときに立つ。
- **A-9 経路では、この条件は二重に成立しない（＝上書きブロックは発火しない）：**
  1. **timing**：会社公式予測を役員会へ提出する前に回す運用のため、予測FYの実績が出る前にA-9を実行する。
  2. **structure**：A-9 は実行直前に `syncSalesFromSalesInput_` で SALES窓を `[fy-4/04, fy/03]` に再整列し、
     予測窓 `[fy/04, fy+1/03]` と同一fyで隣接非重複に確定する。よって forecastMonthIndexesInSales は全て -1、
     sourceByMonth は全月 `forecast_open` となり、上書き対象の closed月が存在しない。
- この結果、旧 §9.2 で挙げていた「actualモードで mixed/regTotal まで実績上書きするのは回帰か」という論点は、
  **上書き対象の月がそもそも生まれない以上、実害として成立しない**ため解消とする。

### 9.4 回帰確認（1回のみ）
- OUTPUT「（参考）内訳とメモ（P50比較）」表の `ForecastSource` 列が、A-9 実行後に全行 `forecast_open` で
  あることを1回確認すれば、§9.3 の非発火（上書き対象月なし）はその場で確認できる。
- 上記が確認できれば、本§9に関する追加対応は不要（コード変更なしで決着）。

---

## 10. AI調査パイプライン（Vertex AI / A-4 = runVertexAIResearch）
v2.2まで設計書に一切記述が無かった領域。現行A-4の実体を正典化する。

### 10.1 全体像
A-4はメーカー名（CONFIG!B2）について、4 topic（Market/Competitor/Channel/DX）ごとに以下を実行する：
1. **grounded web検索**（callVertexGeminiGrounded_）：Gemini + googleSearch（不可時 googleSearchRetrieval へ
   フォールバック）で最新Web情報を取得。groundingメタからcitationを抽出。
2. **RAG検索**（callVertexSearchRAG_）：Vertex AI Search データストア（富士経済PDF / Mixonline Excel等）を
   検索し summary + citations + documents を取得。VERTEX_DATASTORE_ID/SEARCH_LOCATION 未設定時はskip（web-only）。
3. **構造化**（callVertexGeminiStructured_）：web結果をevent行、RAG結果をbenchmark行として
   responseMimeType=application/json で構造化し、AI_RESEARCH_STRUCTURED の行へ変換（buildVertexStructuredRows_）。

### 10.2 readiness 判定
- readVertexConfig_ が geminiReady（projectId+location+geminiModel）と ragReady（datastoreId+searchLocation）を分離。
- A-4の入口判定は geminiReady。RAG未設置でも web-only で実行できる。
- AI_RESEARCH_ENABLED=0 のときは「無効」メッセージで何もせず終了。
- geminiReady=false のときは「必須設定が未入力」エラーで停止（手動TSV経路には落ちない）。

### 10.3 スコア化（event / benchmark）
- event_score：方向(up/down/neutral) × |impact-50| × confidence を ±50 にclamp。
- benchmark_score：(relative_percentile-50) × relative_confidence × quality倍率(high1/medium0.75/low0.5) を ±50 にclamp。
- これらを readAIResearchScores_ が topic別に blend（topicごと固定比率）し、品質不足は中立化。
- AI_SCORE_HISTORY に topic別 blended_score を記録（momentum算出の履歴源）。

### 10.4 失敗時の扱い（フェイルセーフ）
- web/rag/structure の各段で失敗してもrun全体は止めず、AI_RESEARCH_TASK_LOG に per-topic per-aspect 行を残す。
- 全topicで構造化行が0件なら、AI_RESEARCH_STRUCTURED は**上書きせず既存行を保持**し、エラー通知して終了。
- 1件以上得られたら AI_RESEARCH_STRUCTURED を全置換し、サマリービュー（§11）を再描画。

### 10.5 ランタイムシート
- AI_RESEARCH_STRUCTURED：A-1で先行作成される構造化結果シート（event/benchmark行）。A-9が読む唯一の入口。
  `ensureAIResearchRuntimeSheets_` は冪等に存在確認・ヘッダ整合する。
- A-1では作らず、Vertex実行時に遅延作成されるのは以下の3枚：
  - AI_RESEARCH_TASK_LOG：run_id/topic/aspect(web/rag/structure)/status/duration/usage/citations/error を記録。
  - AI_RESEARCH_WEB：web検索の生応答・citation・promptをnote(JSON)で保存。
  - AI_RESEARCH_EXTERNAL：RAGの生応答・summary・documentsをnote(JSON)で保存。

### 10.6 CONFIGキー（既定値）
- VERTEX_PROJECT_ID = forecast-agent-498907
- VERTEX_LOCATION = global（grounding推奨）
- VERTEX_GEMINI_MODEL = gemini-3.1-pro-preview（またはgemini-3.5-flash）
- VERTEX_DATASTORE_ID = （空ならRAGスキップ）
- VERTEX_SEARCH_LOCATION = （空ならRAGスキップ）
- AI_RESEARCH_ENABLED = 1
- 必要OAuthスコープ：cloud-platform / script.external_request（appsscript.jsonに設定済み）

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
- **FORECAST_SNAPSHOT の三角測量系 vestigial カラム**：w1〜w4 / linear/robust/regime/simulation_pred は
  更新ロジック無し（固定記録列）。撤去は確認後の別スコープ。

解消済みメモ：
LANDING_FORECAST（`client|fy|target_month`）と EVAL_LOG（`client|target_month|scenario`）の
再実行追記増殖は複合キー upsert 化で解消済み。AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY は
設計どおり追記のままであり、重複排除は C-1 読取側（forecast_open の最新run_at 1件のみ採用）で行う。
また、shadow mode 表示の `opsQuantOnly` エイリアス問題は objOnly 独立コピー化で解消済み。
旧Gem手動貼付経路の到達不能関数群は物理削除済み。

---

この v2.3.2 は、annual-forecast-mode（FORECAST_CLOSED_MONTH_MODE）の方針を「A=通年予測」とする正典を維持しつつ、
gem-path物理削除と objOnly独立コピー化を現行実装へ同期したドキュメント改訂である。コード変更は伴わない。
年度合計は常に12ヶ月すべての予測simから算出する通年予測であり、closed月上書き（actualモード）は
実運用（A-9）では timing と structure の二重で発火しない。実績は検証・学習経路（B/C）でのみ使用し、
提出済みの予測を書き換えることはない。残存 latent issue は §12 の2件に絞り、各々別スコープで対応する。
