# 売上予測スクリプト 設計書 v1.8（現行エンジン同期版）

## 0. 文書の目的
この設計書は、実装（Forecast_Agent.js）の現行挙動を正確に記述する。
旧v7の「三角観測 w1/w2/w3/w4 + 逆sMAPE重み更新」は実装に存在しないため撤去した。
3目的は不変：(1)予測精度向上 (2)透明化（根拠明示・再現性） (3)学習性（継続改善）。

---

## 0.5 v7からの主要な乖離（同期のために明記）
- **三角観測は廃止済み**。実装は「単一Opsモデル（線形トレンド×季節指数）＋残差Monte Carlo」が定量土台。
- **逆sMAPE重み更新（w_i）は存在しない**。重み更新ロジックは実装されていない。
- **DLM（対数空間の状態空間モデル）が追加**され、CONFIGの DLM_ENGINE_MODE（off/shadow/primary）で制御。
- **主観は乗算係数（kProd/kClient/kOpinion/kAI）として反映**し、月次cap（QUAL_SUBJECTIVE_MONTHLY_CAP）でクリップ。
  v6の「主観キャリブレータ（オーバーレイ率ターゲット探索）」は撤去済み（cap pass-through）。
- **AIは予測係数ではなく、benchmark/event blend のスコアとして kAI に限定反映**。品質不足時は中立化（kAI=1.0）。
- **信頼度（SOURCE_RELIABILITY）** が追加され、各ソース（factor_product/factor_client/opinion/ai_topic）の
  寄与に reliability_r を乗じる。CONFIG RELIABILITY_APPLY_ENABLED で制御（v1.8で既定ON）。
- **LMDI分解** が追加され、主観乗算 Πk-1 を因子別に厳密加法分解（CONFIG LMDI_DECOMPOSITION_ENABLED、既定OFF）。

---

## 1. 実装スコープ

### 1.1 現行実装済み
- データ取込：外部実績SS → SALES_INPUT_MONTHLY（A-2）→ SALES 48ヶ月横持ち BASE/SPOT/TOTAL（A-3）
- 主観入力：FACTORS_PRODUCT / FACTORS_CLIENT / OPINIONS / DEV_SPOT（A-4〜A-7）
- AI調査：Gemプロンプト生成→TSV貼付→parse（A-8 / parseAIResearchPaste_）
- 予測実行：runPhase1Forecast（A-9）。OUTPUT / FORECAST_SNAPSHOT / FORECAST_REPORT 更新
- ダッシュボード：updatePhase1Dashboard（A-10）
- 検証：実績取込（B-1）→ EVAL_LOG / EVAL_COMPARE_MONTHLY（B-2）→ EVAL_INSIGHTS（B-3）
- 四半期レビュー：runQuarterlyReview（C-1）→ applyQuarterlyProposals（C-2）→ ログ閲覧（C-3）
- 信頼度：SOURCE_RELIABILITY 適用（reliabilityApply）＋ C-1のreliability提案

### 1.2 既定OFFのシャドウ機能（検証待ち）
- DLM_ENGINE_MODE = off（shadow/primaryはCONFIGで切替可）
- LMDI_DECOMPOSITION_ENABLED = 0
- ※ v1.8で RELIABILITY_APPLY_ENABLED は 0→1（既定ON）に変更

### 1.3 未実装（今後）
- POOL_PRIOR のクライアント横断自動更新（3c-3c）。現在 readPoolPrior_ は読むのみ、書込み関数なし。
- 構造変化検知（CUSUM/Bai-Perron）、分位点回帰の本格適用、学習窓の動的最適化。

---

## 2. 予測エンジンの実体（runForecastFYCore_）

### 2.1 定量土台
1. SALES から baseSeries48（BASE 48ヶ月）を読む。
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

### 2.4 信頼度（reliability）
- readReliabilityApplyEnabled_ が真のとき readSourceReliability_(client) を適用。
- getSourceReliability_(map, type, key)：未登録は 1.0（中立＝フェイルセーフ）。
- 適用先：factor_product:person / factor_client:person / opinion:person / ai_topic:topic。
- v1.8で既定ON。ただし SOURCE_RELIABILITY が空なら全ソース1.0でno-op（予測不変）。

### 2.5 LMDI分解（診断のみ）
- lmdiDecompose_：Πk-1 を kProd/kClient/kOpinion/kAI に厳密加法分解（全正値はLMDI-I、非正値は線形フォールバックで和保存）。
- 既定OFF。ONのときOUTPUTに寄与シェア／絶対レンジ／相対レンジを出力。

---

## 3. CONFIG パラメータ（読取方式）

### 3.1 読取規約（v1.8で堅牢化）
- readConfigLabelMap_ が CONFIG A:B を読み、configKeyOf_ で「（」または「(」より前をキー化して完全一致マップを作る。
- readModelTuningFromConfig_ / readDlmEngineMode_ / readReliabilityApplyEnabled_ /
  readLmdiDecompositionEnabled_ / readDlmPrimarySpotCapBasis_ はすべてこのマップ経由。
- セル番地直読み（B32〜B56）は廃止。tuneRows に行を挿入しても壊れない。
- ラベル末尾の注記（全角括弧内）は自由に変更可。キー部分が一致すれば読める。

### 3.2 主要キー（抜粋）
- AI_WEIGHT / AI_MAX_ABS_EFFECT / AI_TOTAL_NEUTRAL_THRESHOLD / AI_QUALITY_*_THRESHOLD
- QUAL_SUBJECTIVE_MONTHLY_CAP / QUAL_CALIBRATION_ENABLED
- SPOT_BG_* / KNOWN_SPOT_* / SEASONAL_*
- DLM_ENGINE_MODE / DLM_PRIMARY_SPOT_CAP_BASIS / DLM_BACKTEST_*
- RELIABILITY_APPLY_ENABLED（v1.8既定1）/ RELIABILITY_R_MIN / R_MAX / SHRINKAGE_K / MIN_SAMPLES / MIN_CHANGE
- LMDI_DECOMPOSITION_ENABLED

---

## 4. 学習ループ（四半期レビュー C-1〜C-3）

### 4.1 データ収集（collectQuarterlyReviewData_）
- EVAL_LOG から client一致・scenario=neutral・constraint_relevant_flag=1 の行を抽出（ヘッダidx参照）。
- 直近3ヶ月（last3）が揃わなければ ready:false（提案ゼロで正常終了）。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY / AI_SCORE_HISTORY を last3 で結合。

### 4.2 信頼度提案（generateReliabilityProposals_）
- 各 (source_type, source_key) について、push方向と「実績−定量(quant_only)」のサプライズ方向の的中率hを集計。
- rHat = clamp(2h, R_MIN, R_MAX)。POOL_PRIOR（横断事前）と shrinkage_k で収縮 → rShrunk。
- |rShrunk − 現在値| >= MIN_CHANGE のとき提案化。confidenceは n と変化量で判定。
- ※ プール入力は「生のhit/n」を母数とすべき（rShrunk再プールは自己強化になるため不可）。3c-3cの設計論点。

### 4.3 適用（applyQuarterlyProposals）
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
- 内部管理：RUN_LOG / FORECAST_SNAPSHOT / PROCESS_STATUS / AI_SCORE_HISTORY / AI_IMPACT_HISTORY /
  SUBJECTIVE_IMPACT_HISTORY / CALIBRATION_STATE / CALIBRATION_HISTORY / DLM_STATE / BACKTEST_REPORT /
  SOURCE_RELIABILITY / POOL_PRIOR / LANDING_FORECAST など（原則非表示）。
- FORECAST_SNAPSHOT 末尾に calibration_applied_json を保存。

---

## 7. 今後（3c-3c：POOL_PRIOR 横断集約）の設計論点（未確定）
1. transport：中央プールbook→各bookローカルPOOL_PRIORへfan-out書込み（推奨）／実行時openById読み／手動。
2. 集約入力：生hit/n をプール（rShrunk再プール禁止）。各bookのC-1にsource_type別hit/nの永続化が必要。
3. 粒度：まず reliability:{source_type}。ai_topicは普遍キーなので将来 reliability:ai_topic:{topic} に拡張可。person系は横断不可。
4. precision導出：初版はn加重平均＋保守的固定precision、将来は経験ベイズ。
5. フェイルセーフ：min_clients/min_samples未満は中立1.0据え置き、[R_MIN,R_MAX]クリップ、n_clients/updated_at記録。

---

この v1.8 は実装と設計書の乖離を解消し、信頼度ON化の前提を明文化した同期版。
3c-3c（横断プール）は本体が複数bookで効くことを確認後に着手する。
