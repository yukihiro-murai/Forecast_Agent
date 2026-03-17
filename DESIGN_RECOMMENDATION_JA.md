# 売上予測スクリプト 設計アドバイス（改訂版 v7 / 初動実装と強化計画の分離版）

## 0. 文書の目的
本設計は以下3目的を達成する。
1. **予測精度向上**
2. **透明化（根拠明示・再現性）**
3. **学習性（継続改善）**

---

## 0.5 背景と課題認識
- 単一回帰だけでは業界変化に追従しづらい。
- 未確定月（open）が予測を過少方向に歪める。
- 単発案件がベース需要を歪める。
- 主観・AIの過剰反映は再現性と学習性を下げる。

---

## 1. 実装スコープの明確化（最重要）

### 1.1 **初動実装（Phase 1: これを作る）**
- データ分離: `SALES_INPUT_MONTHLY` / `ACTUAL_EVAL_MONTHLY`
- UIボタン運用（①〜⑥）と依存チェック
- `PROCESS_STATUS` 管理
- ④予測実行フロー（open補正、スパイク分離、三角観測、simulation、補正、シナリオ生成）
- 重み更新（逆sMAPE + 変動制約）
- ログ・レポート（`RUN_LOG`, `FORECAST_SNAPSHOT`, `EVAL_LOG`, `FORECAST_REPORT`, `OVERRIDE_LOG`）
- Geminiコピペ運用（BIGM2Y関連性フィルタ + TSV検証）
- ダッシュボード（シナリオ帯・差分・根拠表示・KPI信号灯）

### 1.2 **強化計画（Phase 2以降: 今後やる）**
- 構造変化検知（CUSUM/Bai-Perron）
- 分位点回帰の本格適用・高度モデル（状態空間/階層ベイズ/因果推論）
- 全件長時間バッチの高度化（運用整合の確認後）
- 学習窓の動的最適化

### 1.3 **初動でやらないこと（明示）**
- 外部AI APIの完全自動連携
- スクリプト自動連鎖実行（ユーザ確認なし）

---

## 2. 設計原則（理由つき）
1. 単一モデル禁止（弱点相互補完）
2. input/eval分離（リーケージ防止）
3. UI実行主義（追跡可能性）
4. AI補助限定（客観優先）
5. 上限制御（モデル無効化防止）
6. 反復運用（確認→修正→再実行）
7. 自動連鎖禁止（誤連鎖防止）

---

## 3. 用語集
- `input`, `eval`, `open`, `closed`, `normalized_actual`
- `spike_amount`, `base_amount`, `deterministic_factors`
- `regime`, `regime_detection`, `regime_transition`
- `scenario`（nega/neutral/posi）
- `confidence_level`（low/mid/high）
- `relevance_score`（0-100）
- `PROCESS_STATUS`, `CLIENT_PARAMS`, `run_log`, `snapshot`, `override`
- `w1/w2/w3/w4`（線形/季節/レジーム/simulation）

---

## 4. データソース・シート定義

### 4.0 データソース定義
- 取得元: 外部スプレッドシート
- 取得方法: `SpreadsheetApp.openById`
- 取得粒度: 月次（client/product/month）
- 取得範囲: 初期24か月全量、以降差分取得
- 差分キー: `client + product + target_month` の未登録行のみ追加（同キー既存は更新）
- 権限: 実行ユーザに閲覧権限必須

### 4.1 予測入力
- `SALES_INPUT_MONTHLY`
- `client, product, target_month, input_amount, status`

### 4.2 検証実績
- `ACTUAL_EVAL_MONTHLY`
- `client, product, target_month, eval_actual_amount, actual_closed_flag`

### 4.3 必須シート
`AI_RESEARCH_PROMPT`, `AI_RESEARCH_PASTE`, `AI_RESEARCH_STRUCTURED`, `RUN_LOG`, `FORECAST_SNAPSHOT`, `EVAL_LOG`, `OVERRIDE_LOG`, `WEIGHT_UPDATE_LOG`, `SPIKE_LOG`, `PROCESS_STATUS`, `CLIENT_PARAMS`, `DETERMINISTIC_FACTORS`, `FORECAST_REPORT`, `DASHBOARD`, `CHANGELOG`

---

## 5. UIメニュー・依存関係

### 5.1 標準メニュー
1. ①予測入力売上取り込み
2. ②検証実績取り込み
3. ③AI調査テンプレ生成
4. ④予測実行
5. ⑤予測検証レポート更新
6. ⑥ダッシュボード更新

### 5.2 依存関係
| step | 前提 |
|---|---|
| ① | なし |
| ② | なし |
| ③ | ①成功（SALES_INPUT参照のため） |
| ④ | ①成功 |
| ⑤ | ②成功 + 過去④実行 |
| ⑥ | ④成功 |

### 5.3 ④対象選択仕様（確定）
- **Phase 1は単一クライアント選択のみ**（ドロップダウン）
- 全件一括はPhase 2以降の検討
- 初期選択は前回実行クライアント

---

## 6. スクリプト独立アーキテクチャ

### 6.1 原則
- スクリプト間の直接呼び出し禁止
- シートI/O連携
- `PROCESS_STATUS` による前提確認

### 6.2 I/O表
| script | button | input | output |
|---|---|---|---|
| `importSalesInput.gs` | ① | 外部ソース | `SALES_INPUT_MONTHLY` |
| `importActualEval.gs` | ② | 外部ソース | `ACTUAL_EVAL_MONTHLY` |
| `generateAIPrompt.gs` | ③ | `SALES_INPUT_MONTHLY` | `AI_RESEARCH_PROMPT` |
| `parseAIResearch.gs` | ③補助 | `AI_RESEARCH_PASTE` | `AI_RESEARCH_STRUCTURED` |
| `runForecast.gs` | ④ | `SALES_INPUT_MONTHLY`,`AI_RESEARCH_STRUCTURED`,`DETERMINISTIC_FACTORS` | `FORECAST_OUTPUT`,`FORECAST_SNAPSHOT`,`FORECAST_REPORT` |
| `runEvaluation.gs` | ⑤ | `ACTUAL_EVAL_MONTHLY`,`FORECAST_SNAPSHOT`,`OVERRIDE_LOG` | `EVAL_LOG`,`WEIGHT_UPDATE_LOG` |
| `updateDashboard.gs` | ⑥ | `FORECAST_REPORT`,`EVAL_LOG` | `DASHBOARD` |

### 6.3 PROCESS_STATUS仕様
- 列: `step_key,last_run_date,last_run_by,status,target_client,record_count,error_summary`
- 状態: `not_run/running/success/error`

### 6.4 GAS 6分制約
- ④は1クライアント単位
- `execution_duration_sec` を常時計測
- 5分超過頻発時は対象月/製品分割

---

## 7. open補正
- `open_month_set` を実行時判定
- `normalized_actual = (途中実績/経過営業日)*全営業日`
- `confidence_level`: <50 low / 50-80 mid / >80 high
- lowは学習不使用

---

## 8. スパイク分離・決定論要因
- 検出: 直近12か月中央値×2.0超
- 分離: `base_amount`, `spike_amount`
- 学習は`base_amount`のみ
- `SPIKE_LOG`記録 + ④で承認

### 8.5 DETERMINISTIC_FACTORS入力タイミング（確定）
- **Phase 1は事前手動入力方式（B案）**
- シート: `DETERMINISTIC_FACTORS`
- 列: `client,target_month,amount,reason,confirmed_by,input_date`
- ④開始時に適用一覧を確認ポップアップ表示

---

## 9. 予測エンジン

### 9.0 ④内部フロー
前提チェック → 対象選択 → open/closed判定 → スパイク検出/承認 → base算出 → 観測1 → 観測2 → 観測3（regime）→ simulation → 合成 → AI補正 → 主観補正入力 → 上限制御/減衰 → シナリオ生成 → 出力/ログ

### 9.1 三角観測
- `w1`: 線形回帰
- `w2`: ロバスト季節
- `w3`: レジーム補正
- `w4`: simulation

### 9.2 個別モデル仕様（Phase 1固定）
- 線形回帰: OLS, 時間変数, 学習窓は**全closed固定**
- ロバスト季節: メディアン季節指数 + Winsorize **P5/P95**
- レジーム: 直近6か月加重平均で補正

### 9.3 レジーム検知
- method: `moving_avg_gap`
- lookback: 6か月
- threshold: ±15%
- transition: 2か月漸進

### 9.4 合成
`final_base = w1*linear + w2*robust + w3*regime + w4*simulation`

### 9.5 重み更新
- `w_i = (1/sMAPE_i) / Σ(1/sMAPE_j)`
- 変動幅 ±0.10, 最低 0.05
- 初回更新条件: closed実績3か月以上
- それまでは `CLIENT_PARAMS` 初期値、未設定は `0.15/0.40/0.25/0.20`

### 9.6 simulation
- 実績<36か月: ブートストラップ
- 実績>=36か月: 分位点回帰追加比較
- `bootstrap_n` 初期値: **500**（`CLIENT_PARAMS`で可変）

### 9.7 予測ホライズン
- 基本: 当月+3か月
- 代替: 年度末
- 不確実性幅は先行月ごと+5%

### 9.8 シナリオ生成
- 中立: `final_base + deterministic_factors + subjective_adj + ai_adj`
- ネガ: `中立 - downside_width(P10)`
- ポジ: `中立 + upside_width(P90)`
- `P10/P90` は **simulation出力分布** から算出

---

## 10. AI調査 + 主観補正

### 10.1 AI運用
- Gemで出力しTSVを貼付

### 10.2 BIGM2Y関連性
- `https://bigm2y.com/service/` 前提
- 関連薄情報は除外

### 10.3 TSV
`client	as_of_date	topic	direction	estimated_impact_pct	confidence	evidence	time_horizon	business_relevance_reason	relevance_score`

### 10.4 検証
- 列数/必須/値域/重複
- `estimated_impact_pct` ±30超警告
- `relevance_score` 0-100外拒否

### 10.5 主観補正
- ④の最終出力前に入力
- `subjective_reason` 必須
- 保存先: `FORECAST_SNAPSHOT` と `OVERRIDE_LOG`

### 10.6 減衰
- `effective_impact = impact_pct * (0.5^(months_since/3))`
- AI: `as_of_date`基準
- 主観: `subjective_input_date`基準（補正なしはnull）
- AIデータは `as_of_date` 6か月超で補正対象外

### 10.7 補正制御
- 主観±8%、AI±5%
- `relevance_score < 60` は補正不使用

---

## 11. 透明化レポート・ダッシュボード

### 11.1 レポート必須
3シナリオ、各観測、重み、寄与率、補正量、差分要因、採用理由、前提条件

### 11.2 寄与率
- v6式を採用
- 分母が前期実績1%未満なら寄与率算出対象外（絶対額表示）

### 11.3 出力先
- `FORECAST_REPORT`（client×month×scenario）

### 11.4 ダッシュボード
- シナリオ帯
- 観測別トレンド
- 実績vs予測差分
- 根拠詳細
- KPI信号灯

---

## 12. ログ仕様
### 12.1 RUN_LOG
`run_id,run_at,run_by,function_name,client,status,count,model_version,parameters_snapshot_json,input_data_hash,execution_duration_sec,error_summary`

### 12.2 FORECAST_SNAPSHOT
`...,subjective_input_date`（補正なしはnull）

### 12.3 OVERRIDE_LOG
`override_type` は `subjective_input` / `manager_override` を区別

### 12.4 保持期間
FORECAST 24m / EVAL 36m / RUN 12m（超過はアーカイブ）

---

## 13. 学習ループ・KPI
- ②後にclosed増で候補生成
- ⑤で重み更新提案→承認時のみ反映
- KPI閾値: sMAPE, 捕捉率, 根拠欠落率
- `was_overridden` 別集計
- 赤信号時アクション5項目

---

## 14. CLIENT_PARAMS
- 初期は管理者投入
- 未設定は全体デフォルト
- ⑤提案で更新
- 推奨追加パラメータ: `winsor_p_low=5`, `winsor_p_high=95`, `bootstrap_n=500`, `ai_max_age_months=6`

---

## 15. ガバナンス
- 上書き理由必須
- 根拠セットで意思決定
- AIは補助情報
- 仕様変更時は `CHANGELOG` 同時更新（`change_date,changed_by,section,change_summary,reason`）

---

## 16. Phase移行条件
- P1→P2: 3か月安定 + sMAPE<=30
- P2→P3: 捕捉率>=70
- P3→P4: sMAPE 5%以上改善

---

## 17. 強化候補（今後の見通し）
- CUSUM/Bai-Perron
- 状態空間/階層ベイズ/因果推論
- 全件長時間実行の高度化
- 学習窓の動的最適化

---

## 18. 不要/後回し判断
- **後回し（不要ではない）**: 高度統計モデル、長時間自動連鎖、完全自動AI連携
- 理由: Phase 1の安定運用と現場定着を優先

---

このv7は「初動実装」と「今後強化」を明確に分離し、開発者が判断に迷わない実装仕様として整理した最終版。
