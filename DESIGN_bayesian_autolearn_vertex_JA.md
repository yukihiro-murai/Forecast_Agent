# 設計仕様：ベイズ自動学習 + Vertexハイブリッド予測（B-5）

**このファイルの用途**：「実績を取り込むたびにパラメータがベイズ更新され、予測精度が自動改善する」
ゴールに向けた学習レイヤーの設計根拠を記述した参照資料。

対象：`Forecast_Agent.js`（VERSION='2.4.0-dev' / BUILD_STAGE='bayesian-autolearn-vertex-hybrid'）

---

## 0. スコープの宣言

本変更は **既存予測コア（Opsモデル・DLM・モンテカルロ混合）を書き換えず、その出力に乗る
「学習済み補正レイヤー」を追加する** ものとする。次は対象外:

- Ops trend/seasonal、DLM、残差分位点プールの計算自体の変更
- vNext（employee flow / Engine 0.5.0）への移植
- C-1 四半期レビュー（提案→人手承認）の廃止・置き換え

学習レイヤーは C-1 と**併存**する。C-1 はソース別信頼度の「提案→人が適用」、
B-5 は係数の「自動更新」で、更新対象が別（下表）。

| 学習対象 | 更新経路 | 保存先 |
|---|---|---|
| 全体補正係数 bias_correction_factor | B-5 月次（自動）/ C-2 四半期（人手） | CALIBRATION_STATE |
| 暦月別バイアス residual_month_bias_json | B-5 月次（自動） | CALIBRATION_STATE |
| ソース別信頼度 SOURCE_RELIABILITY | C-1 提案 → C-2 適用（人手） | SOURCE_RELIABILITY |
| Vertexアシスト信頼度（vertex_forecast） | C-1 と同じ hit-rate 機構で自動学習 | SOURCE_RELIABILITY |

---

## 1. B-5 月次ベイズ自動学習

### 1.1 入出力

- 入力：EVAL_LOG のうち `scenario=neutral` かつ `constraint_relevant_flag='1'` の
  「forecast_open 時点の中立予測 vs 実績」の月次ペア（target_month 単位で最新 evaluated_at のみ採用）。
- 出力：CALIBRATION_STATE の `bias_correction_factor` と `residual_month_bias_json` の更新。

### 1.2 計算（autoLearnComputeState_）

誤差符号 `e = (pred - actual) / |actual|`（正=過剰予測）。

- **EWMA**: 最新月を w=1 として半減期4か月 `w_i = 0.5^(i/4)`。
- **全体バイアス**: 事前分布 N(0, 精度 kG=3) のベイズ縮小
  `postBias = Σ w·e / (Σw + kG)`。サンプルが少ないほど0へ縮む（過学習防止）。
- **係数**: `targetFactor = clamp(1 - postBias, 0.75, 1.25)`。
  1回の更新幅は `±0.05` にクランプ（`newFactor = cur + clamp(target - cur)`）し、
  毎月の実行で漸近的に target へ収束する。暴走を防ぐ安全弁。
- **暦月別バイアス**: 暦月 m ごとに `b_m = clamp(-Σ_m w·e / (Σ_m w + kM=2), ±0.20)`。
  |b_m| < 0.005 は記録しない（ノイズ書き込み防止）。
- **発火条件**: 中立評価月 ≥3（AUTOLEARN_MIN_EVAL_MONTHS）。
- **停止弁**: CALIBRATION_STATE `auto_update_enabled=0` で B-5 は何も書かない
  （C-2 の手動経路は残る）。

### 1.3 予測への適用（runForecastFYCore_）

`sourceByMonth[i]==='forecast_open'` の月だけに
`k = biasFactor × (1 + clamp(monthBias[m], ±0.25)) × kVertex[i]` を
`p10/p50/p90/raw` に乗算。**closed 月の実績上書き値は再スケールしない**
（実績は確定値であり補正対象ではない — 旧実装は全月に掛けていたためここを修正）。

### 1.4 自動トリガ

- `runMonthlyAutoLearn`（メニュー B-5 / Web検証タブ）で手動実行。
- `updatePhase1EvaluationReport`（B-2）の末尾で `autoLearnAfterEvalReport_` が自動実行し、
  `learn_status`（PROCESS_STATUS）に結果を残す。**実績が入るたび学習が進む**構造。

---

## 2. Vertexアシスト（ルールベース × AI ハイブリッド）

### 2.1 位置づけ

A-9 の統計予測に対し、Vertex Gemini が「AIの月次調整案」を出し、
**制限付きで**統計予測へ掛け合わせる。A-4（runVertexAIResearch）の末尾で非同期に生成し
VERTEX_FORECAST_LOG へ記録、A-9 は最新の一致行だけを読む（A-9 内では Vertex を呼ばない。
レイテンシ・失敗を予測本線から切り離す）。

### 2.2 重み付け（信頼度連動）

`kVertex[i] = clamp(1 + adj[i] × vertexW, 0.7, 1.3)` where
`vertexW = CONFIG VERTEX_ASSIST_WEIGHT(0.5) × r_vertex × (0.5 + 0.5×confidence)`

- `adj[i]` は Gemini 返却の ±0.30 内調整案。
- `r_vertex` は SOURCE_RELIABILITY の `vertex_forecast` 信頼度 — 初期1.0で、
  A-9 が `vertexPushByMonth` を SUBJECTIVE_IMPACT_HISTORY（source_type='vertex_forecast'）
  に記録するため、**C-1 と同じ hit-rate 機構で実績ベースに自動学習される**。
  AI が外し続ければ r が下がり寄与が自動的に弱まる。
- `confidence` はモデル自身の確度（低いほど反映が弱い）。

### 2.3 生成プロンプト

client名・対象12月・直近12か月実績・傾き・季節強弱月・既知Spot・主観プッシュ・
AIトピックスコア・直近の予測ミス を渡し、JSON
`{monthly_adj:[12], confidence, rationale_ja}` を要求。
`VERTEX_FORECAST_ENABLED=1` で有効（既定0）。

---

## 3. カウンターファクト検算（学習効果バックテスト）

`runLearningBacktest_`：EVAL_LOG の中立月を古い順に walk-forward し、
当月までの学習係数 `k = factor × (1 + monthBias)` で stored final_pred を再スケール、
補正なし/ありの WAPE を比較する。**簡易検算**（学習係数のみ適用しモデル再推定はしない）
である旨を結果に明記する。必要評価月 ≥6。

---

## 4. UI 露出

- 検証タブ：「自動学習」カード — 現在の係数・暦月バイアス・Vertex状態・信頼度学習件数を表で表示、
  B-5 実行 / Vertexアシスト実行 / 学習効果検算 ボタン。検算結果は WAPE KPI＋月別表。
- 予測タブ：学習パラメータが実際に効いている時だけ「学習パラメータ（この予測へ自動適用）」
  カードを出す（未学習のデフォルト状態では出さない — 意味のない情報を置かない）。
- `learn_status` を PROCESS_STATUS／ナビのエラードット対象に追加。

---

## 5. 受け入れ基準

1. 中立評価が3か月未満では B-5 は書き込まない（insufficient_eval_months）。
2. 過剰予測が続くと係数が1回あたり最大0.05ずつ下方修正され、反復で target へ収束する。
3. closed 月の実績上書き値は補正で変わらない。
4. `auto_update_enabled=0` では B-5 は no-op、C-2 の手動経路は残る。
5. `VERTEX_FORECAST_ENABLED=0` では kVertex=1（無影響）でA-4も失敗しない（skipped）。
6. 検算カードは評価月不足時に理由を表示し、十分時は補正なし/あり WAPE を並べる。

---

## 6. 実機検証で発見・修正した根本原因（2026-10-01 追記）

学習ループの実経路検証（B-1→B-2→B-5→検算→A-9）で、**EVAL_LOG が永遠に空になる
事前バグ**を発見し修正した。これは自動学習だけでなく C-1 の信頼度集計や
EVAL_COMPARE_MONTHLY の月結合をも全て止めていた。

- 原因: FORECAST_SNAPSHOT.target_month へ 'yyyy/MM' 文字列を setValues すると
  Sheets の自動書式推定で **Date セル化** する列があった。一方 ACTUAL_EVAL_MONTHLY は
  文字列 'yyyy/MM'。`[client|month]` 結合で `String(Date)` ('Wed Apr 01 2026…')
  と '2026/04' が一致せず、EVAL 追記・比較表・C-1 last3 フィルタ・
  LANDING_FORECAST upsert の dedupe が全てミスしていた。
- 対処: `ymKey_(v)`（Date / 'YYYY/MM' / 'YYYY-MM' / 'YYYY/M' → 'yyyy/MM'）を
  全結合点（B-2 評価マージ・比較表・B-4・C-1・ランディング dedupe・
  collectNeutralEvalPairs_）に適用。書き込み側は SNAPSHOT/EVAL_LOG の
  target_month 列に `@` 書式を setValues 前に設定し新規 Date 化を抑止。
  `toMonthStart_` は realm 非依存の Date 判定に強化。
- 効果（実データ @20）: B-2 で EVAL_LOG 蓄積→内部自動学習が発火
  （n=10, factor 1.00→0.95→0.90 収束中、暦月バイアス 9 ヶ月学習済み）、
  検算 WAPE 0.612→0.584、A-9 で学習係数が予測へ自動適用され
  予測画面の「学習パラメータ」カードに表示されることを目視確認。
