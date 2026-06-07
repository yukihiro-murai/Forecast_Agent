# 売上予測スクリプト 設計書 v2.2（現行エンジン同期版）

## 0. 文書の目的
この設計書は、実装（Forecast_Agent.js）の現行挙動を正確に記述する。
旧v7の「三角観測 w1/w2/w3/w4 + 逆sMAPE重み更新」は実装に存在しないため撤去済み。
3目的は不変：(1)予測精度向上 (2)透明化（根拠明示・再現性） (3)学習性（継続改善）。

対象実装：`Forecast_Agent.js`（VERSION='2.2.2-dev' / BUILD_STAGE='v8-annual-forecast-mode'）
設計書版：v2.2（v1.8からの同期改訂。3c-3c横断プールと学習ループ修正を正典化）

### 0.0 v1.8からの主な更新点（この改訂で反映したもの）
- **3c-3c（POOL_PRIOR横断集約）を「実装済み」へ更新**。v1.8では未実装と記載していたが、
  `adminAggregatePoolPriorAcrossBooks` ほかで実装済み（§7）。
- **学習ループのサンプル水増し修正を正典化**。`forecast_source` 列の追加と、C-1集計の
  dedup/open限定ロジックを§4・§8に記述。
- **kProd全月1.0のフェイルセーフ化を正典化**（throw撤去→警告）。§2.3・§8に記述。
- **annual-forecast-mode（FORECAST_CLOSED_MONTH_MODE）は「実装済み・方針確認待ち」として§9に隔離**。
  既定OFF（actual）で従来挙動のため、本文の予測記述は従来どおりとする。

---

## 0.5 v7からの主要な乖離（同期のために明記）
- **三角観測は廃止済み**。実装は「単一Opsモデル（線形トレンド×季節指数）＋残差Monte Carlo」が定量土台。
- **逆sMAPE重み更新（w_i）は存在しない**。重み更新ロジックは実装されていない。
- **DLM（対数空間の状態空間モデル）が追加**され、CONFIGの DLM_ENGINE_MODE（off/shadow/primary）で制御。
- **主観は乗算係数（kProd/kClient/kOpinion/kAI）として反映**し、月次cap（QUAL_SUBJECTIVE_MONTHLY_CAP）でクリップ。
  v6の「主観キャリブレータ（オーバーレイ率ターゲット探索）」は撤去済み（cap pass-through）。
- **AIは予測係数ではなく、benchmark/event blend のスコアとして kAI に限定反映**。品質不足時は中立化（kAI=1.0）。
- **信頼度（SOURCE_RELIABILITY）** が追加され、各ソース（factor_product/factor_client/opinion/ai_topic）の
  寄与に reliability_r を乗じる。CONFIG RELIABILITY_APPLY_ENABLED で制御（既定ON）。
- **LMDI分解** が追加され、主観乗算 Πk-1 を因子別に厳密加法分解（CONFIG LMDI_DECOMPOSITION_ENABLED、既定OFF）。
- **POOL_PRIOR のクライアント横断集約** が追加（中央集約book→各bookへfan-out / 手動実行）。

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
- 信頼度：SOURCE_RELIABILITY 適用（reliabilityApply）＋ C-1のreliability提案＋RELIABILITY_EVIDENCEへの raw hit/n 永続化
- 横断プール：POOL_PRIOR のクライアント横断集約（adminSetupPoolHub / adminAggregatePoolPriorAcrossBooks）

### 1.2 既定OFFのシャドウ機能（検証待ち）
- DLM_ENGINE_MODE = off（shadow/primaryはCONFIGで切替可）
- LMDI_DECOMPOSITION_ENABLED = 0
- FORECAST_CLOSED_MONTH_MODE = actual（forecastに切替で通年予測モード。詳細は§9）
- ※ RELIABILITY_APPLY_ENABLED は既定1（ON）。ただしSOURCE_RELIABILITY空ならno-op（予測不変）。

### 1.3 未実装（今後）
- 構造変化検知（CUSUM/Bai-Perron）、分位点回帰の本格適用、学習窓の動的最適化。
- POOL_PRIOR集約のprecision導出の経験ベイズ化（現状はn加重平均＋固定precision=SHRINKAGE_K）。
- person系ソースの横断プール（現状は普遍キーのsource_type単位まで）。

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
- **kProd全月1.0のフェイルセーフ（旧hard throw撤去）**：FACTORS_PRODUCTに有効行があるのに
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

---

## 3. CONFIG パラメータ（読取方式）

### 3.1 読取規約
- readConfigLabelMap_ が CONFIG A:B を読み、configKeyOf_ で「（」または「(」より前をキー化して完全一致マップを作る。
- readModelTuningFromConfig_ / readDlmEngineMode_ / readReliabilityApplyEnabled_ /
  readLmdiDecompositionEnabled_ / readDlmPrimarySpotCapBasis_ / readForecastClosedMonthMode_ はすべてこのマップ経由。
- セル番地直読みは廃止。tuneRows に行を挿入しても壊れない。
- ラベル末尾の注記（全角括弧内）は自由に変更可。キー部分が一致すれば読める。

### 3.2 主要キー（抜粋）
- AI_WEIGHT / AI_MAX_ABS_EFFECT / AI_TOTAL_NEUTRAL_THRESHOLD / AI_QUALITY_*_THRESHOLD
- QUAL_SUBJECTIVE_MONTHLY_CAP / QUAL_CALIBRATION_ENABLED
- SPOT_BG_* / KNOWN_SPOT_* / SEASONAL_*
- DLM_ENGINE_MODE / DLM_PRIMARY_SPOT_CAP_BASIS / DLM_BACKTEST_*
- RELIABILITY_APPLY_ENABLED（既定1）/ RELIABILITY_R_MIN / R_MAX / SHRINKAGE_K / MIN_SAMPLES / MIN_CHANGE
- POOL_MIN_CLIENTS（横断集約の最低クライアント数、既定2）
- LMDI_DECOMPOSITION_ENABLED
- FORECAST_CLOSED_MONTH_MODE（actual/forecast、既定actual。§9参照）

---

## 4. 学習ループ（四半期レビュー C-1〜C-3）

### 4.1 データ収集（collectQuarterlyReviewData_）
- EVAL_LOG から client一致・scenario=neutral・constraint_relevant_flag=1 の行を抽出（ヘッダidx参照）。
- 直近3ヶ月（last3）が揃わなければ ready:false（提案ゼロで正常終了）。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY / AI_SCORE_HISTORY を last3 で結合。

### 4.2 サンプル水増し対策（dedup / open限定）★この改訂で正典化
背景：GUIDEは「確認→修正→再実行を前提とする」と明記しており、同じ月に対しA-9を複数回押す運用が正規。
旧実装はA-9のたびに履歴を追記（dedupなし）し、C-1が1行ずつ `n += 1` していたため、
入力を変えずA-9を数回押すだけで n が膨張し、提案が「高」信頼度に化ける欠陥があった。
さらに、closed月の actual上書きにより `pred_p50_quant_only` に actual が記録され、
C-1の month単位 last-write-wins が最新（actual上書き後）の行を拾って surprise=0 で当該月が脱落していた。

対策（現行 computeReliabilityHitStats_）：
- **forecast_source 列の追加**：AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY の末尾に
  `forecast_source`（'forecast_open' / 'actual_closed'）を追加。記録値は runForecastFYCore_ の
  result.sourceByMonth[i] をそのまま書く（既存列順は不変）。
- **open限定**：C-1集計は `forecast_source==='forecast_open'` の行のみ採用。
  これにより closed後の actual上書き行（surprise=0）が自動除外され、月脱落が解消する。
  forecast_source が空/欠落の旧（汚染）行も同フィルタで自動除外される。
- **dedup**：
  - quant側（AI_IMPACT_HISTORY）：month単位で最新 run_at の1件のみ採用。
  - subjective側（SUBJECTIVE_IMPACT_HISTORY）：(target_month, source_type, source_key) 単位で
    最新 run_at の1件のみ採用。
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
- 内部管理：RUN_LOG / FORECAST_SNAPSHOT / PROCESS_STATUS / AI_SCORE_HISTORY / AI_IMPACT_HISTORY /
  SUBJECTIVE_IMPACT_HISTORY / CALIBRATION_STATE / CALIBRATION_HISTORY / DLM_STATE / BACKTEST_REPORT /
  SOURCE_RELIABILITY / RELIABILITY_EVIDENCE / POOL_PRIOR / POOL_REGISTRY / POOL_AGGREGATION_LOG /
  LANDING_FORECAST など（原則非表示）。
- FORECAST_SNAPSHOT 末尾に calibration_applied_json を保存。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY 末尾に forecast_source（§4.2）。

---

## 7. POOL_PRIOR 横断集約（3c-3c）★この改訂で「実装済み」に更新
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
- 学習ループ修正（forecast_source追加・dedup・kProdフェイルセーフ）は本設計書v2.2で正典化済み。
- 現行コードは VERSION='2.2.2-dev' / BUILD_STAGE='v8-annual-forecast-mode'。
  この設計書はそのうち §1〜§7 の確定機能を記述対象とし、annual modeは§9で別扱いとする。

---

## 9. 【方針確認待ち】annual-forecast-mode（FORECAST_CLOSED_MONTH_MODE）
本機能は **実装済みだが、年間総計のセマンティクスに関する方針が未確定** のため、確定機能とは分けて記述する。
既定は actual（従来挙動）であり、本文（§2）の予測記述は既定モード前提で不変。

### 9.1 挙動
- `FORECAST_CLOSED_MONTH_MODE = actual`（既定）：closed月は実績で上書き表示（objOnly/mixed/regTotal）。
- `FORECAST_CLOSED_MONTH_MODE = forecast`：closed月も予測のまま表示（通年予測モード）。
  経過月の実績は ActualClosed 列に参考併記し、予測値には混ぜない。
- どちらのモードでも、年度合計(P10/P50/P90)は12ヶ月すべての予測simから算出される
  （経過実績で固定した着地値ではない）。

### 9.2 未確定の論点（実装に進む前に要確認）
- 年間総計を「実績込みの着地予測」として扱うのか、「通年（12ヶ月）すべて予測の見通し」として扱うのか。
  現行実装は後者（通年予測）を既定とするが、計画用途として前者が望ましい場面があるか要確認。
- OUTPUTの年度合計セマンティクスの表記（着地 vs 通年）をどちらに正規化するか。
- 既定モード（actual）でも closed月を mixed/regTotal まで実績上書きしている点が、
  従来挙動（objOnlyのみ上書き）からの変化かどうかを回帰確認で確定させる。

### 9.3 暫定方針
- 方針確定までは既定 actual を維持。forecastモードは検証用途に留める。
- 確定後、この§9を§2以降の本文へ統合し、設計書から「方針確認待ち」表記を外す。

---

この v2.2 は、3c-3c横断プールと学習ループ修正（forecast_source / dedup / kProdフェイルセーフ）を
コードと同期して正典化した改訂版。annual-forecast-mode は実装済みだが方針確認待ちとして§9に隔離している。
