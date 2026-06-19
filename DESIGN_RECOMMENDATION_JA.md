# 売上予測スクリプト 設計書 v2.3.10（2.3.48 同期 / output-display-polish）

## 0. 文書の目的
この設計書は、実装（Forecast_Agent.js）の現行挙動を正確に記述する。
旧v7の「三角観測 w1/w2/w3/w4 + 逆sMAPE重み更新」は実装に存在しないため撤去済み。
3目的は不変：(1)予測精度向上 (2)透明化（根拠明示・再現性） (3)学習性（継続改善）。

対象実装：`Forecast_Agent.js`（VERSION='2.3.48-dev' / BUILD_STAGE='output-display-polish'）
設計書版：v2.3.10（v2.3.9 からのドキュメント改訂）。
DLM_BUILD_STAGE は 'v8-step3c3c-1'（DLMロジック無変更のため据え置き）。

本改訂は、設計書 v2.3.9 が記述する `2.3.46-dev / budget-nav-guard` から現行コード
`2.3.48-dev / output-display-polish` までの差分を同期する。コードはこの間に2増分（2.3.47〜2.3.48）進んでいる：
- **2.3.47 config-blank-guard**：CONFIG 数値読取の空セルガード（空/欠落→既定値へフォールバック）＋ coerceMatch 死にコード削除。
- **2.3.48 output-display-polish**：OUTPUT 6行目の常時黒字化（赤字判定 step3aHasError 死にコード削除）／A-10 ナビの範囲選択化／K列幅是正。

いずれもスコア算出式・予測コア・履歴スキーマは無変更で、A-9 の年度合計・月次 P10/P50/P90 は不変
（config-blank-guard のみ、空セルが残る異常 book では既定値フォールバックにより予測が変わりうる＝是正であり回帰ではない。
正常にシード済みの book では予測不変）。output-display-polish は表示・ナビゲーションのみで予測不変。

**前版での未反映の解消**：設計書 v2.3.9 が「未適用・未検証」と記していた config-blank-guard は本改訂時点で
**適用・確認済み**である。これに伴い、v2.3.9 の §3.3「getCfg 空セルの罠（現存）」と §12「coerceMatch 死にコード（現存）」は、
本 v2.3.10 では **是正済み／撤去済み** として記載を更新する（§3.3・§12）。ファイル先頭コメントの VERSION/BUILD_STAGE も
const と同期済み（2.3.48-dev / output-display-polish）。

### 0.0 v2.3.9 → v2.3.10 の更新点（この改訂で反映したもの）

この2増分（2.3.47〜2.3.48）は、CONFIG 数値読取の空セル是正（正常シード book では予測不変）と OUTPUT 表示・
ナビゲーションの微修正（予測不変）。スコア算出式・予測コア・履歴スキーマは無変更、A-9 の年度合計・月次 P10/P50/P90
は不変。本改訂で §3.3・§12 の「現存の罠／死にコード」記述を解消する。

#### 0.0.1 output-display-polish（2.3.48）— OUTPUT 表示・ナビゲーション微修正（予測不変）
- **OUTPUT 6行目（注記ブロック）を常時黒字（#000000）化**。旧 output-display-tidy（2.3.30）の
  「実警告（エンジン⚠ / productWeightWarning / cal⚠ / web/rag/structure_error>=1）のときのみ赤字（#b71c1c）」を撤回し、
  6行目は警告/エラーの有無にかかわらず目立たせない方針に変更。これに伴い赤字判定の死にコード `step3aHasError`
  （`/(?:web_error|rag_error|structure_error)=[1-9]/` の test 変数）を物理削除。AI 個別の警告強調は 21〜23 行に残る。
  `step3aWarn`（AI取込警告サマリー表示）と `productWeightWarning`/`productWeightWarningText` は引き続き使用・残存。
- **A-10「予算を策定」(`gotoBudgetEntry`) のスクロールを範囲選択化**。単一セル `H24` ではなく予算ブロック全体
  `H24:J40` を `activate()` で範囲選択し、ブロック全体が画面に収まるよう可視化（アクティブセルは左上 H24）。
  OUTPUT 不在 / A-9 未実行（タイトル空）のガードは不変（budget-nav-guard）。
- **OUTPUT の K 列（11列目）幅を 40px → 130px** に拡げ潰れを解消。H/I/J=160px、グラフ L列（12）開始は不変。
- 変更箇所は `writeOutputFY_` と `gotoBudgetEntry` の2関数のみ。予測ロジック・他シートは無変更。

#### 0.0.2 config-blank-guard（2.3.47）— CONFIG数値読取の空セルガード + coerceMatch死にコード削除
- `readModelTuningFromConfig_` 内の `getCfg(key, def)` と `readAiMissingConfidenceDefault_` で、CONFIG セルが
  **空文字列('')/null/undefined・キー欠落**のとき既定値へフォールバックするよう是正。
  旧来は `Number('')===0 → isFinite(0)===true` により「0 という上書き値」を採用し、既定値ではなくクランプ下限へ
  化けていた（calibration-blank-override-fix と同型）。**明示的に 0 を入れた場合は従来どおり 0 を採用**（空と明示0を区別）。
- 是正後の挙動例：`QUAL_SUBJECTIVE_MONTHLY_CAP` 空 → 0.40（旧: 0.01）、`AI_WEIGHT` 空 → 0.0008（旧: 0）、
  `AI_MISSING_CONFIDENCE_DEFAULT` 空 → 0.5（旧: 0=欠落 event 行は不採用）が正しく効く。
- **予測影響**：正常にセットアップ済み（全セルシード）の book では空セルが無いため発火せず、是正の前後で予測不変。
  空セルが残る異常 book でのみ既定値フォールバックにより予測が変わりうる（=是正であり回帰ではない）。
- `writeOutputFY_` 内の coerceMatch / coerceCount 死にコード（Vertex 警告サマリーに該当文字列が無く恒常的に
  非発火だった `warn_coerced=` 判定。`coerceCount=0` で `if(>=3)` が永久に走らない）を物理削除。

### 0.0' v2.3.8 → v2.3.9 の更新点（参考・再掲）

この5増分（2.3.42〜2.3.46）は、いずれも **予測不変**（スコア算出式・予測コア・履歴スキーマ無変更 /
A-9 の年度合計・月次 P10/P50/P90 不変）の表示・メニュー・ナビゲーション変更である。

#### 0.0'.1 budget-nav-guard（2.3.46）— A-10 ナビゲーションのガード（予測不変）
- `gotoBudgetEntry`（A-10 予算を策定）は OUTPUT の予算策定欄へカーソルを移動するナビゲーションのみ。
- OUTPUT シートが無い場合は注意ダイアログで終了。OUTPUT はあるが **A-9 未実行（タイトル空）** の場合も、
  空の予算欄へ飛ばさず「先に A-9 を実行」ダイアログで終了する（=本ガード追加）。
- A-9 実行後は OUTPUT を前面化し予算欄をアクティブにしてトースト表示。
  ※ 2.3.48（output-display-polish）で、アクティブ化対象が単一セル H24 から範囲 H24:J40 に変更された（§0.0.1）。

#### 0.0'.2 guide-a10-sync（2.3.45）— GUIDE 手順表の A-10 反映（表示のみ）
- GUIDE 手順表の A 行が A-1〜A-10 の10行になり、A-10 行（青背景）に「予算を策定」の説明が入る。
- B 行は A-10 の直下から B-1〜B-4、C 行は C-1〜C-3。メニュー A 群と GUIDE A 行が項目・順序とも一致する。

#### 0.0'.3 budget-entry-polish（2.3.44）— A-10 メニュー追加 + 予算欄の仕上げ（予測不変）
- メニュー A 群に **「A-10 予算を策定」（`gotoBudgetEntry`）** を A-9 の直下・区切り線の上に追加。
- OUTPUT 混合（本命）セクションの **H24:J24 に「予算を策定」見出し**（結合表示）を追加。客観のみセクションには出さない。
- 予算ブロック **H24:J40 を外周のみの控えめな枠線**（緑 #6aa84f / SOLID_MEDIUM・内側罫線なし）で囲む。
- `buildOUTPUT_` がチャートを除去するようにし、A-1 再セットアップ時に旧グラフが残らないようにした
  （A-9 後の再生成位置 L25 と食い違わない）。混合セクションのグラフは **L25**（startRow+1 / L列）に生成。

#### 0.0'.4 output-budget-columns（2.3.43）— 予算策定欄の追加（予測不変）
- OUTPUT 混合（本命）セクションに、編集可能な予算列を追加（**客観のみセクションには出さない**）。
  - **H列 Adopted Forecast（採用予測）**：月次の混合 Baseline(P50) を初期値（数値・手入力で上書き可 / 黄背景）。
  - **I列 Sales Uplift（営業上積み）**：手入力（黄背景）。空欄は 0 扱い。
  - **J列 Final Budget（最終予算）**：`=H+I`（月次）。確定額（緑背景・太字）。
- 年度合計バンドには H/I/J のヘッダと月次 SUM（`=SUM(H29:H40)` 等）を置く。
- チャートは **L列（col 12）開始** へ右移動し、K列（col 11）を余白として H/I/J（8〜10列）と重ならないようにした。
  ※ K列幅は 2.3.48（output-display-polish）で 40px→130px に是正された（§0.0.1）。

#### 0.0'.5 dashboard-menu-to-b（2.3.42）— ダッシュボード更新を B 群へ移設（予測不変）
- 予測ダッシュボード更新（`updatePhase1Dashboard`）を **旧 A-10 から B-3** へ移設。
- 検証インサイト更新（`updatePhase1LearningInsights`）は **B-4** へ繰り下げ。
- B 群は **B-1 実績取込 / B-2 検証レポート / B-3 ダッシュボード更新 / B-4 検証インサイト** の順。
- GUIDE 分類表で DASHBOARD を「事後検証用」に並べ、緑（#d9ead3）表示にする。OUTPUT は「出力用」のまま。
- DASHBOARD は引き続き既定非表示で、前面表示は **B-3 実行直後**（`updatePhase1Dashboard` 末尾の
  `showSheet → setActiveSheet`）になる。以降 A-4 等で `hideNonUserSheets_` が走ると再び隠れる。

### 0.0'' v2.3.7 → v2.3.8 の更新点（参考・再掲）— report-text-compress（2.3.41）
**変更スコープは A-4 の構造化プロンプトのみ**（`buildVertexStructureSystemInstruction_` と
`buildVertexStructureUserContent_` の2関数）。スコア算出式・予測コア・履歴スキーマは無変更で、A-9 の
P10/P50/P90 は不変。AI_RESEARCH_STRUCTURED の `report_text` 列の生成方針だけを変えた。

- **report_text を topic ごと 2〜3文・最大120字程度に圧縮**。競合・他社の製品別シェアや個別製品の販促状況
  などの**明細を列挙しない**。ソース文の**転記・引用をしない**（出典マーカー `[1]`/`[2]` 等の羅列を持ち込まない）。
  対象メーカー（client）の外部環境がどの向きに・どの程度動いているかという**結論を1文目に**書く。
- **非遡及性**：A-4 を再実行しない限り既存 `report_text` は前回の長文のまま残る。A-4 再実行で短文化する。
- **不変点**：スコア列（`event_score`/`benchmark_score`/`blended_score`）と coverage/quality/neutralized は
  構造的に変わらない（live検索の揺れ・要約変化に伴う微差は許容、算出式は不変）。

### 0.0''' v2.3.6 → v2.3.7 の更新点（参考・再掲）
- **spot-bg-sensitivity-down（2.3.39）**：背景SPOTの過去感度を引き下げ。`SPOT_BG_SHRINK` 0.50→0.35、
  `SPOT_SPIKE_MAD_K` 3.0→2.5。背景SPOTが効いている book では P10/P50/P90 が変化しうる（意図した数値変更）。
  `SPOT_BG_FLOOR_RATE`=0.15 / `SPOT_BG_CAP_RATE`=0.20 / known spot は不変。**const 既定は CONFIG 欠落時の
  フォールバックに過ぎず、既存bookは A-1 再生成か該当セルの手修正をしない限り旧値（0.50/3.0）のまま**。
- **dashboard-hide-insights-upsert（2.3.38）**：DASHBOARD 既定非表示化、EVAL_INSIGHTS の人手入力保持 upsert
  （複合キー `[client, target_month]`）。※前面表示タイミングは本書 0.0'.5（dashboard-menu-to-b）で A-10→B-3 に変更済み。

### 0.0'''' v2.3.5 → v2.3.6 の更新点（参考・再掲）— overlay-cap-raise（2.3.37）
- **予測値が変わる数値変更**。`QUAL_SUBJECTIVE_MONTHLY_CAP` 0.20→0.40、`AI_MAX_ABS_EFFECT` 0.03（±3%）→0.05（±5%）。
  主観入力・AI調査の効いている book では P10/P50/P90 が変化しうる。`AI_TOPIC_SCORE_ABS_CAP`（±50/4軸±200）と
  `AI_WEIGHT_DEFAULT`（0.0008）は不変。

### 0.0''''' v2.3.4 → v2.3.5（参考・再掲）
- **client-match-unify（2.3.32）**：読取側のクライアント突合を `isSameClient_`（両側 `normalizeClientName_` で正規化）に集約。
- **a9-client-normalize（2.3.33）**：A-9 入口で CONFIG!B2 を正規化し、永続クライアント値を正規化済みに統一。

### 0.0'''''' v2.3.3 → v2.3.4（参考・再掲）
- **output-display-tidy（2.3.30）**：表示専用整形（初期ダイアログ黄notice撤去、OUTPUT 6行目の赤字判定を実警告時のみ、
  row10 を F列、21〜23行を折り返し・幅縮小）。※ 6行目の赤字判定は 2.3.48（output-display-polish）で撤回（常時黒字化 / §9.5）。
- **raw-migrate-header-match（2.3.31）**：旧2枚→RAW の移送をヘッダ名マッチ＋シート名由来 axis に変更。

### 0.0''''''' v2.3.2 → v2.3.3（参考・再掲 / 抜粋）
- snapshot-vestigial-removal（2.3.12）：FORECAST_SNAPSHOT を15列化、`final_pred` は J列（index 9）。
- config-simplify（2.3.13）/ config-deadknob-removal（2.3.15）：未使用CONFIG行・互換参照を撤去。
- rag-config-defaults（2.3.16）：CONFIG に RAG 既定値を A-1 で自動投入。地域エンドポイント対応。
- ai-dx-confidence-diagnostics（2.3.21）：confDrop 可視化。
- output-note-relocate（2.3.22）：年度≠月次合算の長文説明を Scenario Split の下へ移設。
- ai-score-robustness（2.3.23）：event_score を `direction符号 × (impact/100) × 50 × conf` へ是正。`AI_MISSING_CONFIDENCE_DEFAULT`（既定0.5）導入。
- ai-score-degroundless（2.3.24）：根拠なき中立化／捏造ゼロを是正。
- calibration-blank-override-fix（2.3.25）：override 空欄判定を是正（`Number('')→0` の罠を回避）。
- rag-query-frame-align（2.3.26）：`buildRagQuery_` を AI調査の新フレームへ整合。
- qual-share-const-removal（2.3.27）/ orphan-fn-removal（2.3.28）：デッドコード・孤立関数を物理削除。
- ai-research-raw-merge（2.3.29）：生ログを `AI_RESEARCH_RAW` に集約。axis 列で区別。

---

## 0.5 v7からの主要な乖離（同期のために明記）
- **三角観測は廃止済み**。実装は「単一Opsモデル（線形トレンド×季節指数）＋残差Monte Carlo」が定量土台。
- **逆sMAPE重み更新（w_i）は存在しない**。FORECAST_SNAPSHOT は15列・`final_pred`はindex 9。
- **DLM（対数空間の状態空間モデル）が追加**され、CONFIGの DLM_ENGINE_MODE（off/shadow/primary）で制御。
- **主観は乗算係数（kProd/kClient/kOpinion/kAI）として反映**し、月次cap（QUAL_SUBJECTIVE_MONTHLY_CAP、既定0.40）でクリップ。
- **AIは benchmark/event blend のスコアとして kAI に限定反映**。総合中立化（dead-zone）は撤去（既定閾値0）。上限は ±AI_MAX_ABS_EFFECT（既定±5%）のみ。
- **背景SPOTの過去感度を引き下げ**（SPOT_BG_SHRINK=0.35 / SPOT_SPIKE_MAD_K=2.5）。floor=0.15・cap=0.20・known spot は不変。
- **AI調査の取得経路が Vertex AI へ移行**。生ログは `AI_RESEARCH_RAW`（axis=web/rag）。report_text は短文圧縮。
- **信頼度（SOURCE_RELIABILITY）/ LMDI分解 / POOL_PRIOR 横断集約** が追加。
- **クライアント名マッチを統一・正規化**（client-match-unify / a9-client-normalize）。
- **CONFIG数値読取の空セルガード**（config-blank-guard）。空文字列/欠落を既定値へフォールバックし、`Number('')===0` の罠を回避。明示0は採用（空と区別）。
- **DASHBOARD は既定非表示**。前面表示は **B-3（updatePhase1Dashboard）直後**（dashboard-menu-to-b で A-10→B-3 に移設）。
- **OUTPUT 混合セクションに予算策定欄（H/I/J）を追加**。**A-10「予算を策定」はその欄（H24:J40 を範囲選択）へのナビゲーションのみ**で予測不変。
- **OUTPUT 6行目を常時黒字化**（output-display-polish）。実警告時の赤字判定は撤回し、6行目は目立たせない。

---

## 1. 実装スコープ

### 1.1 現行実装済み
- データ取込：外部実績SS → SALES_INPUT（A-2）→ SALES_MONTHLY 48ヶ月横持ち BASE/SPOT/TOTAL（A-3）
- 主観入力：PRODUCT / CLIENT / OPINIONS / DEV_SPOT（A-5〜A-8）
- AI調査（Vertex AI自動）：A-4 `runVertexAIResearch`。grounded web検索 + Vertex AI Search RAG + 構造化出力
  → AI_RESEARCH_STRUCTURED へ記録、AI_RESEARCH（サマリービュー）へ再描画。report_text は短文圧縮（§10・§11）。
- 予測実行：runPhase1Forecast（A-9）。OUTPUT / FORECAST_SNAPSHOT 更新。入口で CONFIG!B2 を正規化。
- **予算策定欄（A-9 が OUTPUT に出力）**：混合セクションに編集可能な H/I/J 列を生成（§9.6）。
- **A-10 予算を策定（`gotoBudgetEntry`）**：OUTPUT の予算欄（H24:J40 を範囲選択）へ移動するナビゲーションのみ（予測非影響 / §9.6）。
- CONFIG数値読取の空セルガード：`readModelTuningFromConfig_`/`readAiMissingConfidenceDefault_` が空/欠落を既定値へフォールバック（§3.3）。
- ダッシュボード：updatePhase1Dashboard（**B-3**）。DASHBOARD は既定非表示で、B-3 実行直後のみ前面表示。
- 検証：実績取込（B-1）→ EVAL_LOG / EVAL_COMPARE_MONTHLY（B-2）→ DASHBOARD（B-3）→ EVAL_INSIGHTS（B-4）。
  B-4 は人手入力列を保持する upsert。
- 四半期レビュー：runQuarterlyReview（C-1）→ applyQuarterlyProposals（C-2）→ ログ閲覧（C-3）
- 信頼度：SOURCE_RELIABILITY 適用＋C-1のreliability提案＋RELIABILITY_EVIDENCE への raw hit/n 永続化
- 横断プール：POOL_PRIOR のクライアント横断集約（adminSetupPoolHub / adminAggregatePoolPriorAcrossBooks）
- シート初期化：AI_RESEARCH_STRUCTURED は A-1 で先行作成。Vertex実行時に遅延作成されるのは AI_RESEARCH_TASK_LOG /
  AI_RESEARCH_RAW の2枚。旧2枚→RAW の移送はヘッダ名マッチ＋シート名由来 axis。

### 1.2 既定OFF/中立のシャドウ機能（検証待ち）
- DLM_ENGINE_MODE = off（shadow/primaryはCONFIGで切替可）
- LMDI_DECOMPOSITION_ENABLED = 0
- FORECAST_CLOSED_MONTH_MODE = actual（**月次表示のみ**を切り替えるトグル。年度合計の算出には影響しない。§9）
- AI_SCORE_BASIS = level（momentumに切替でAIスコアを相対位置の変化として扱う）
- AI_TOTAL_NEUTRAL_THRESHOLD = 0（dead-zone なし＝既定）
- ※ RELIABILITY_APPLY_ENABLED は既定1（ON）。SOURCE_RELIABILITY空ならno-op（予測不変）。
- ※ AI_RESEARCH_ENABLED は既定1。0でA-4のVertex調査をスキップ。
- ※ AI_MISSING_CONFIDENCE_DEFAULT は既定0.5（0で従来どおり不採用）。空セルは config-blank-guard で既定0.5へフォールバック（§3.3）。

### 1.3 未実装（今後）
- 構造変化検知（CUSUM/Bai-Perron）、分位点回帰の本格適用、学習窓の動的最適化。
- POOL_PRIOR集約のprecision導出の経験ベイズ化（現状はΣhit/Σn加重＋固定precision=SHRINKAGE_K）。
- person系ソースの横断プール（現状は普遍キーのsource_type単位まで）。
- AI_RESEARCH_RAW の `frozen_flag` / `frozen_at` は将来のスナップショット凍結用の予約列（現状未使用 / 既定0・空）。
- `normalizeClientName_` の正規化辞書拡張（現状は ｳﾞｨｱﾄﾘｽ系のみハードコード）。
- キャップ値・背景SPOT感度の経験ベイズ的な自動調整（現状は CONFIG 固定値 + 手動調整）。
- **予算策定欄（H/I 手入力）の A-9 再実行をまたいだ保持**（現状は `resetOutputSheet_` で都度ワイプ＝一回入力前提で許容済み。任意拡張 / §9.6・§12）。
- enable系フラグ（RELIABILITY_APPLY_ENABLED / AI_RESEARCH_ENABLED 等）の `Number(x || 0)` 空セルガード（チューニング数値側は config-blank-guard で是正済み。enable側は別系統で現状維持 / §12）。

### 1.4 メニュー構成（実装同期 / 2.3.48）
- A-1 初期セットアップ（setupForecastBook）
- A-2 売上データを取り込む（importSalesInputMonthly）
- A-3 予測用に売上データを加工（aggregateSalesData）
- A-4 AI調査を取り込む（runVertexAIResearch / Vertex AI自動）
- A-5 製品ごとの動向を入力（openProductTrendEntryDialog）
- A-6 クライアント動向を入力（openClientTrendEntryDialog）
- A-7 担当者意見を入力（openOpinionsEntryDialog）
- A-8 開発/スポット要因を入力（openDevEntryDialog）
- A-9 予測を実行（runPhase1Forecast）
- **A-10 予算を策定（gotoBudgetEntry）** ← OUTPUT 予算欄 H24:J40 へのナビゲーションのみ（予測非影響）
- B-1 検証用に実績データを取り込み（importActualEvalMonthly）
- B-2 検証レポートを更新（updatePhase1EvaluationReport）
- **B-3 予測ダッシュボードを更新（updatePhase1Dashboard）** ← 旧 A-10 から移設
- **B-4 検証インサイトを更新（updatePhase1LearningInsights）** ← 旧 B-3 から繰り下げ
- C-1 四半期レビューを実行（runQuarterlyReview）
- C-2 承認済み提案を適用（applyQuarterlyProposals）
- C-3 過去の提案履歴を開く（openQuarterlyReviewLog）
- 管理者用（メニュー非掲載 / スクリプトエディタから手動）：
  adminSetupGuideOnly / adminInitDLMAndBacktest / adminSetupPoolHub / adminAggregatePoolPriorAcrossBooks

GUIDE 手順表は A 行 A-1〜A-10（10行）、B 行 B-1〜B-4、C 行 C-1〜C-3 で、メニュー A/B/C 群と項目・順序が一致する（guide-a10-sync）。
※ 2.3.47（config-blank-guard）・2.3.48（output-display-polish）でメニュー項目・順序は変わっていない。

---

## 2. 予測エンジンの実体（runForecastFYCore_）

### 2.1 定量土台
1. SALES_MONTHLY から baseSeries48（BASE 48ヶ月）を読む。
2. adjustForUnclosedMonths_：未確定月を同月トレンド係数で補完。補完後は途中実績を下回らない。
3. fitOpsModelTrendSeason_：OLS線形トレンド＋移動平均ベース季節指数（0.80〜1.20でクリップ）。
4. buildResidualPool_：確定月の残差%プール（MADクリップ＋中央値方向へ収縮）。
5. forecastByResidualQuantiles_ で P10/P50/P90 の定量土台を算出。
6. DLM_ENGINE_MODE=primary かつ実績充足時は、BASEを対数空間DLMの予測へ差し替え（shadowは比較のみ）。

### 2.2 SPOT
- 背景SPOT（未知の再発）：fitSpotRecurringModel_ が月別の期待値・発生確率・severity標本を作る（BASE P50比でcap）。
  - 月別の背景SPOT期待値 = `max(monthAvg × SPOT_BG_SHRINK, monthAvg × SPOT_BG_FLOOR_RATE)` を `BASE_P50 × SPOT_BG_CAP_RATE`
    で上限クリップ。スパイク判定は `med + SPOT_SPIKE_MAD_K × mad` 超を除外。**現行値：SHRINK=0.35 / SPIKE_MAD_K=2.5 / FLOOR=0.15 / CAP=0.20**。
- 既知SPOT（DEV_SPOT）：金額×確度を月別に固定加算。背景との二重計上はoffset率で調整。

### 2.3 主観乗算（forecastMonteCarloMixed_）
- kProd = 1 + Σ(製品構成比 × 製品step × reliability)
- kClient = 1 + Σ(client step × reliability)
- kOpinion = 担当者別の最新意見を ±5% jitter込みで合成（× reliability）
- kAI = 1 + clamp(Σ(topicスコア × reliability) × AI_WEIGHT, ±AI_MAX_ABS_EFFECT)。総合中立化（dead-zone）は既定OFF。
- Monte Carlo（N_SIM=1000）で各simの total を生成し、月次P10/P50/P90を取る。
- 主観差分は月次cap（QUAL_SUBJECTIVE_MONTHLY_CAP、**既定0.40**）でクリップ。cap は `±(quantOpsBase × cap)` を限度とする
  唯一の制御点（`applySubjectiveCap_`）。`QUAL_CALIBRATION_ENABLED=0` のときは cap 無効。
- **kProd全月1.0のフェイルセーフ**：PRODUCTに有効行があるのに kProdが全月1.0（直近12ヶ月closed BASE実績の無い製品＝weight=0）
  の場合、throwせず警告（`productWeightWarning`）に置換し A-9を完走させる。

### 2.4 信頼度（reliability）
- readReliabilityApplyEnabled_ が真のとき readSourceReliability_(client) を適用。未登録は 1.0（中立＝フェイルセーフ）。
- 適用先：factor_product:person / factor_client:person / opinion:person / ai_topic:topic。
- 既定ON。SOURCE_RELIABILITY が空なら全ソース1.0でno-op（予測不変）。クライアント絞り込みは `isSameClient_` 経由。

### 2.5 LMDI分解（診断のみ）
- lmdiDecompose_：Πk-1 を kProd/kClient/kOpinion/kAI に厳密加法分解。既定OFF。

### 2.6 AIスコアの予測反映
- readAIResearchScores_ が AI_RESEARCH_STRUCTURED から topic別 final blended score を返す。品質不足topicは中立化。
- 各軸 final score は ±AI_TOPIC_SCORE_ABS_CAP（=50）でクリップ（4軸合計±200）。
- AI_SCORE_BASIS=momentum のときは AI_SCORE_HISTORY の過去runとの差分（momentum）に切替（既定はlevel）。
  momentum の lookback は `as_of_date` 単位で dedup され、A-9実行回数では動かない。
- ※ report_text の圧縮（report-text-compress）は表示変更で、final score（event/benchmark）には影響しない。

### 2.7 年度合計のセマンティクスと2経路の分離（確定）
**(a) 予測・提出経路（A-9 / runForecastFYCore_）— 実績非依存**
- 年度合計(P10/P50/P90)は12ヶ月すべての予測simから算出する（経過実績で固定した着地値ではない）。
- A-9 は実行直前に `syncSalesFromSalesInput_` で SALES窓を `[fy-4/04, fy/03]` に再整列し、予測窓 `[fy/04, fy+1/03]` と
  隣接非重複に確定する。よって全月 `forecast_open` となり実績は混ざらない。
- クライアント名は A-9 入口で正規化される。

**(b) 検証・学習経路（B-1 → B-2 → B-3 → B-4、C-1）— 実績を使うが予測は書き換えない**
- B-1：実績を ACTUAL_EVAL_MONTHLY に取り込む（取込時に client は正規化済み）。
- B-2：FORECAST_SNAPSHOT（`final_pred`=r[9]）と実績を `[client, target_month]` 複合キーで突合し EVAL_LOG / EVAL_COMPARE_MONTHLY に記録。
- B-3：DASHBOARD を更新（B-2 の比較結果から KPI 集計）。DASHBOARD を前面表示。
- B-4：EVAL_INSIGHTS に外れ要因・次アクションを整理。人手入力列は保持する upsert（§4.7・§6）。
- C-1：EVAL_LOG と impact履歴から reliability を学習。

---

## 3. CONFIG パラメータ（読取方式）

### 3.1 読取規約
- readConfigLabelMap_ が CONFIG A:B を読み、configKeyOf_ で「（」または「(」より前をキー化して完全一致マップを作る。
- read*FromConfig_ 系はすべてこのマップ経由。セル番地直読みは廃止。tuneRows に行を挿入しても壊れない。
- **const 既定値は CONFIG セル欠落/不読/空セル時のフォールバック**。clasp push だけでは既存bookの CONFIG セルは更新されないため、
  const 既定の変更（overlay-cap-raise / spot-bg-sensitivity-down 等）を発効させるには A-1 再生成か該当セルの手修正が必要。
- **空文字列('')も既定値へフォールバック**（config-blank-guard / §3.3）。`Number('')===0` が isFinite を通り「0という上書き」
  として採用されクランプ下限へ化ける問題を回避。明示的に 0 を入れた場合は従来どおり 0 を採用（空と明示0を区別）。

### 3.2 主要キー（抜粋・現行同期）
- AI_WEIGHT / **AI_MAX_ABS_EFFECT（既定0.05＝±5%）** / AI_MISSING_CONFIDENCE_DEFAULT（既定0.5 / 0で不採用 / 空は既定0.5）
- AI_TOTAL_NEUTRAL_THRESHOLD（既定0）/ AI_QUALITY_NEUTRAL_THRESHOLD / AI_QUALITY_PARTIAL_THRESHOLD
- AI_SCORE_BASIS（level/momentum）/ AI_MOMENTUM_LOOKBACK_QUARTERS / AI_MOMENTUM_MIN_HISTORY
- **QUAL_SUBJECTIVE_MONTHLY_CAP（既定0.40 / 空は既定0.40）** / QUAL_CALIBRATION_ENABLED
- **SPOT_BG_SHRINK（既定0.35）** / SPOT_BG_FLOOR_RATE（0.15）/ SPOT_BG_CAP_RATE（0.20）/ **SPOT_SPIKE_MAD_K（既定2.5）** / KNOWN_SPOT_*
- DLM_ENGINE_MODE / DLM_PRIMARY_SPOT_CAP_BASIS / DLM_BACKTEST_* / DLM_LOG_EPSILON_RATE
- RELIABILITY_APPLY_ENABLED（既定1）/ RELIABILITY_R_MIN / R_MAX / SHRINKAGE_K / MIN_SAMPLES / MIN_CHANGE / POOL_MIN_CLIENTS
- LMDI_DECOMPOSITION_ENABLED / FORECAST_CLOSED_MONTH_MODE（actual/forecast、既定actual）
- Vertex AI調査キー：VERTEX_PROJECT_ID / VERTEX_LOCATION / VERTEX_GEMINI_MODEL / VERTEX_DATASTORE_ID /
  VERTEX_SEARCH_LOCATION / VERTEX_SERVING_CONFIG（既定 default_search）/ AI_RESEARCH_ENABLED（既定1）

### 3.3 CONFIG数値読取の空セルガード（config-blank-guard / 2.3.47・是正済み）
`readModelTuningFromConfig_` 内の `getCfg(key, def)` と `readAiMissingConfidenceDefault_` は、CONFIG のセルが
**空文字列('')/null/undefined・キー欠落**のとき既定値へフォールバックする。旧来は `Number('')===0 → isFinite(0)===true`
により「0 という上書き値」を採用し、既定値ではなくクランプ下限へ化けていた（calibration-blank-override-fix と同型）。
config-blank-guard でこれを是正済み：例として `QUAL_SUBJECTIVE_MONTHLY_CAP` 空 → 0.40（旧: 0.01 へ化けた）、
`AI_WEIGHT` 空 → 0.0008（旧: 0）、`AI_MISSING_CONFIDENCE_DEFAULT` 空 → 0.5（旧: 0=欠落 event 行は不採用）が正しく効く。
**明示的に 0 を入れた場合は従来どおり 0 を採用**（空と明示0を区別）。
正常にセットアップ済み（全セルシード）の book では空セルが無いため、是正の前後で予測は不変。空セルが残る異常 book でのみ
既定値フォールバックにより予測が変わりうる（=是正であり回帰ではない）。
※ enable系フラグ（RELIABILITY_APPLY_ENABLED / AI_RESEARCH_ENABLED 等）の `Number(x || 0)` 側は本ガードの対象外で、
空セルで OFF に倒れる挙動は現状維持（実害なし / §12）。

### 3.4 撤去済みCONFIG行（参考）
config-simplify / config-deadknob-removal / qual-share-const-removal により、担当者の B10 互換参照 /
QUAL_SUBJECTIVE_MAX_SCALE / QUAL_SHARE_* / QUARTERLY_REVIEW_PERIOD_MONTHS / BIAS_CORRECTION_* /
AI_DIRECTION_HIT_* / AI_EFFECT_MIN_MEANINGFUL / DLM_FORECAST_HORIZON は CONFIG から撤去済み。

---

## 4. 学習ループ（四半期レビュー C-1〜C-3）

### 4.1 データ収集（collectQuarterlyReviewData_）
- EVAL_LOG から client一致（`isSameClient_`）・scenario=neutral・constraint_relevant_flag=1 の行を抽出。
- 直近3ヶ月（last3）が揃わなければ ready:false（提案ゼロで正常終了）。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY / AI_SCORE_HISTORY を last3 で結合。

### 4.2 サンプル水増し対策（dedup / open限定）
- **forecast_source 列**：AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY 末尾に `forecast_source`（'forecast_open' / 'actual_closed'）。
- **open限定**：C-1集計は `forecast_source==='forecast_open'` の行のみ採用。closed後の actual上書き行を自動除外。
- **dedup**：quant側は month単位で最新 run_at の1件、subjective側は (target_month, source_type, source_key) 単位で最新 run_at の1件。

### 4.3 信頼度提案（generateReliabilityProposals_）
- rHat = clamp(2h, R_MIN, R_MAX)。POOL_PRIOR と shrinkage_k で収縮 → rShrunk = clamp((n*rHat + k*rPool)/(n+k), R_MIN, R_MAX)。
- |rShrunk − 現在値| >= MIN_CHANGE のとき提案化。source_type別 hit/n を RELIABILITY_EVIDENCE に永続化。

### 4.4 適用（applyQuarterlyProposals）
- QUARTERLY_REVIEW の承認列に従う。承認かつ auto_update_enabled=1 のときのみ SOURCE_RELIABILITY を upsert ＋ 履歴記録。
- 同一 review_id の二重適用は早期終了。

### 4.5 AI_weight 提案（generateQuarterlyProposals_）
- AI方向一致率と mean|kAI-1| から ai_weight_override 提案を生成。
- 現在値の読取は **空欄=未設定→CONFIG既定、明示数値→その値**（calibration-blank-override-fix と同方針）。

### 4.6 calibration override の空欄/明示0 区別（calibration-blank-override-fix）
- `calOverrideNum_(v)`：`''`/`null`/`undefined` は null（未設定→無視）、数値は isFinite なら採用。明示 0 は 0 を採用。

### 4.7 EVAL_INSIGHTS の人手入力保持 upsert（dashboard-hide-insights-upsert）
- **B-4**（`updatePhase1LearningInsights`）の書込は `upsertEvalInsightsRows_` 経由の複合キー `[client, target_month]` upsert。
- **機械算出列は毎回上書き**（actual_total / pred_p50 / diff / error_rate / 各 breach flag / insight / next_action 等）。
- **人手入力列は既存値があれば保持**（HUMAN_COLS = 0-index 14,15,16,18,19,20,21,22：cause_hypothesis / cause_bucket /
  impacted_assumption / action_type / next_cycle_reflection / owner / due_date / status）。
- これにより B-4 再実行でメンバーの原因分析・担当・期日・ステータスが消えない。
- ※ メニュー再編（dashboard-menu-to-b）で本機能の呼び出しは旧 B-3 → **B-4** に移った（挙動は不変）。

---

## 5. 検証ポリシー（KPI）
- 計画用単一値＝P50。P10/P90は説明帯（hard gateではない）。
- 制約：annual_abs_error_rate <= 10% / half_wape <= 12%（将来目標10%）/ over-forecast rate <= 5%。
- 診断：月次APE、Q差分、定量寄与率、主観オーバーレイ率（AI除く）、AI寄与率、Known Spot寄与率、レンジ逸脱数。
- レンジ逸脱月（actualがP10-P90外）はB-4で追加調査を必須化。
- 1 client = 1 book。
- ※ overlay-cap-raise / spot-bg-sensitivity-down 適用後は P50 が変化しうるため本KPIの再確認を1サイクル行うこと。
- ※ report-text-compress（2.3.41）・dashboard-menu-to-b〜budget-nav-guard（2.3.42〜2.3.46）・output-display-polish（2.3.48）は
  予測値に影響しないため本KPIへの影響は無い。config-blank-guard（2.3.47）は正常シード book で予測不変
  （空セルが残る異常 book でのみ既定値フォールバックで変わりうる＝是正。その場合は本KPIを1サイクル再確認）。
- ※ **予算策定欄（H/I/J）は計画値そのものではなく、ディレクター/営業が会社へ打ち上げる予算を OUTPUT 上で手入力・確定する欄**。
  計画の主数値は引き続き混合の年度合計 P50。予算欄は KPI 評価（B-2/B-3）の対象外（§9.6）。

---

## 6. ログ／永続シート
- ユーザ表示シート：GUIDE / CONFIG / SALES_INPUT / SALES_MONTHLY / AI_RESEARCH（サマリービュー）/ PRODUCT / CLIENT /
  OPINIONS / DEV_SPOT / OUTPUT（hideNonUserSheets_ の userVisible はこの10枚）。
  **DASHBOARD は既定非表示**（userVisible に含めない）。前面表示は **B-3（updatePhase1Dashboard）直後**の
  `showSheet → setActiveSheet`。以降 A-4 等で `hideNonUserSheets_` が走ると再び隠れる（dashboard-menu-to-b で
  前面表示タイミングを旧 A-10 → B-3 に変更済み）。
- 内部管理（原則非表示）：RUN_LOG / FORECAST_SNAPSHOT / PROCESS_STATUS / AI_SCORE_HISTORY / AI_IMPACT_HISTORY /
  SUBJECTIVE_IMPACT_HISTORY / CALIBRATION_STATE / CALIBRATION_HISTORY / QUARTERLY_REVIEW / QUARTERLY_REVIEW_LOG /
  DLM_STATE / BACKTEST_REPORT / SOURCE_RELIABILITY / RELIABILITY_EVIDENCE / POOL_PRIOR / POOL_REGISTRY /
  POOL_AGGREGATION_LOG / LANDING_FORECAST / AI_RESEARCH_STRUCTURED / AI_RESEARCH_TASK_LOG / AI_RESEARCH_RAW / DASHBOARD。
- **FORECAST_SNAPSHOT は15列**（snapshot_id / run_date / client / target_month / scenario / base_pred / subjective_adj /
  ai_adj / deterministic_adj / **final_pred(index 9)** / confidence_interval_lower / confidence_interval_upper /
  key_factors_json / subjective_input_date / calibration_applied_json）。client 列は A-9 入口で正規化された値。
- **OUTPUT 混合セクションの予算列 H/I/J**：A-9 が生成する編集可能欄。H=Adopted Forecast（初期値=月次P50・手入力可）、
  I=Sales Uplift（手入力・空欄0）、J=Final Budget（=H+I）。年度バンドは月次 SUM。表示・手入力欄であり予測非影響（§9.6）。
- **OUTPUT 6行目（注記ブロック）**：output-display-polish（2.3.48）で常時黒字（#000000）。実警告時の赤字判定は撤回（§9.5）。
- **AI_RESEARCH_STRUCTURED の report_text 列**：topicごとの短文要約（2〜3文・最大120字程度）。A-4 再実行で更新（非遡及）。
- **EVAL_INSIGHTS は24列**。機械算出列と人手入力列を持ち、B-4 は複合キー upsert で機械列を上書き・人手列を保持（§4.7）。
- **AI_RESEARCH_RAW（旧 WEB + EXTERNAL の統合）**：`client / as_of_date / axis / topic / direction / magnitude /
  uncertainty / relative_position / evidence / frozen_flag / frozen_at / note`。axis=web/rag で出所を区別。
- AI_IMPACT_HISTORY / SUBJECTIVE_IMPACT_HISTORY 末尾に forecast_source（§4.2）。

---

## 7. POOL_PRIOR 横断集約（3c-3c）
### 7.1 構成
- ハブbook初期化：adminSetupPoolHub（手動）。POOL_REGISTRY / POOL_AGGREGATION_LOG / POOL_PRIOR を作成。
- 集約実行：adminAggregatePoolPriorAcrossBooks（手動）。REGISTRYの enabled=1 book を openById で開き、
  各bookの RELIABILITY_EVIDENCE から source_type別の生 hit/n を読み、集約してPOOL_PRIORへ書き込み＋fan-out。

### 7.2 集約ロジック
- 集約入力：生 hit/n をプール（rShrunk再プール禁止）。粒度：reliability:{source_type}。person系キーは横断不可。
- pooled_value = clamp(2 * Σhit/Σn, R_MIN, R_MAX)。precision = SHRINKAGE_K（固定）。
- フェイルセーフ：n_clients < POOL_MIN_CLIENTS → written=false / reason=min_clients。Σn < MIN_SAMPLES → reason=min_samples。
  単一book読取失敗は status=excluded でログ記録し全体は止めない。

### 7.3 予測への波及
- POOL_PRIOR は C-1 の提案収縮（rShrunk）に効くのみ。予測値そのものは変えない（集約直後のA-9はno-op）。

---

## 8. バージョン整合と適用順序
- VERSION / BUILD_STAGE / 設計書版 / 手動チェックリストは各リリースで同期する。
- 現行コードは **VERSION='2.3.48-dev' / BUILD_STAGE='output-display-polish'**。DLM_BUILD_STAGE は 'v8-step3c3c-1'（据え置き）。
  ファイル先頭コメントの VERSION/BUILD_STAGE も const と同期済み。
- 本改訂（設計書 v2.3.10）は 2.3.47〜2.3.48 の2増分（config-blank-guard / output-display-polish）を反映する。
  config-blank-guard は正常シード book で予測不変・空セル異常 book で是正、output-display-polish は表示・ナビのみで予測不変。
- 設計書ドリフトは既知の再発リスク。リリースごとに doc sync を独立タスクとして扱う。今後も stage 確定のたびに doc を同期する。
- **次期予定**：現時点で未適用・別生成済みのプロンプトは無い。次の変更スコープ確定時に同期する。

---

## 9. 通年予測モード（FORECAST_CLOSED_MONTH_MODE）と OUTPUT 表示
年間総計のセマンティクスは **「通年（12ヶ月）すべて予測の見通し」（選択肢A）** で確定。

### 9.1 確定した仕様
- **年度合計(P10/P50/P90)は、FORECAST_CLOSED_MONTH_MODE の値にかかわらず、常に12ヶ月すべての予測simから算出する。**
  closed月上書きは sim配列に触れず月次表示用配列にしか作用しない。
- 本トグルは **「月次表示の見せ方」だけを切り替える**もので、計画の主数値（年度合計）の意味は変えない。

### 9.2 モード別の月次表示
- `actual`（既定）：closed月は実績で上書き表示（objOnly/mixed/regTotal）。
- `forecast`：closed月も予測のまま表示（通年予測モード）。経過月の実績は ActualClosed 列に参考併記。

### 9.3 closed月上書きの発火条件と、実運用での非発火
- actualモードの上書きは `sourceByMonth[i]==='actual_closed'` の月のみに作用する。
- A-9 経路では、timing（実績前に回す）と structure（syncSalesFromSalesInput_ による窓非重複）の二重で発火条件に到達しない（全月 forecast_open）。

### 9.4 回帰確認（1回のみ）
- OUTPUT「（参考）内訳とメモ（P50比較）」表の `ForecastSource` 列が、A-9 実行後に全行 `forecast_open` であることを1回確認。

### 9.5 表示位置と OUTPUT 表示整形（output-note-relocate / output-display-tidy / output-display-polish）
- 年度合計行と月次表ヘッダの間は短い1行注記のみ。「年度≠月次合算」の長文説明は混合セクションの Scenario Split の下に配置。
- output-display-tidy（2.3.30）：row10 は F列。21〜23行は折り返し・幅縮小。
- **output-display-polish（2.3.48）**：OUTPUT 6行目（注記ブロック）は、警告/エラーの有無にかかわらず**常に黒字（#000000）**。
  旧 output-display-tidy の「実警告時のみ赤字（#b71c1c）」は撤回し、6行目は目立たせない方針に変更。赤字判定の死にコード
  `step3aHasError` を削除。AI個別の警告強調（no_data の赤バナー等）は 21〜23 行に残る。`step3aWarn`（AI取込警告サマリー）と
  `productWeightWarning`（製品重み警告のテキスト）は6行目に表示されるが、行全体は赤くならない。

### 9.6 予算策定欄（output-budget-columns / budget-entry-polish / budget-nav-guard / output-display-polish）【予測不変】
OUTPUT 混合（本命）セクションに、会社へ打ち上げる予算を OUTPUT 上で確定するための編集可能欄を追加する。
**客観のみセクションには出さない。** 予測コア（A〜G列の年度合計・月次 P10/P50/P90）は本欄の追加で不変。

- **列構成（混合セクションのみ）**
  - **H列 Adopted Forecast（採用予測）**：初期値＝月次の混合 Baseline(P50)（数値・手入力で上書き可）。黄背景。
  - **I列 Sales Uplift（営業上積み）**：手入力（空欄は 0 扱い）。黄背景。
  - **J列 Final Budget（最終予算）**：月次 `=H+I`。緑背景・太字（確定額）。
- **レイアウト（A-9 出力時）**
  - 混合セクション startRow=24。H24:J24 に **「予算を策定」見出し**（結合表示）。
  - 年度バンド：ヘッダ H25:J25、月次SUM H26:J26（`=SUM(H29:H40)` 等）。
  - 月次バンド：ヘッダ H28:J28、月次データ H29:J40。H に月次P50初期値、I 空欄、J=H+I。
  - 予算ブロック **H24:J40 を外周のみの緑枠**（#6aa84f / SOLID_MEDIUM・内側罫線なし）で囲む。
  - 各セクションのチャートは **L列（col 12）開始**（混合は L25）。K列（col 11）は予算列とグラフの間の余白で、
    **幅 130px**（output-display-polish で 40px→130px に是正・潰れ解消）。H/I/J（8〜10列）と重ならない。
  - `buildOUTPUT_` は A-1 再セットアップ時に旧チャートを除去（A-9 後の再生成位置 L25 と食い違わない）。
- **A-10 予算を策定（gotoBudgetEntry）**：OUTPUT を前面化し **H24:J40 を範囲選択してアクティブ化**するナビゲーションのみ
  （予算ブロック全体を画面に収める。アクティブセルは左上 H24 / output-display-polish。旧来は単一セル H24 のみアクティブ化）。
  OUTPUT 不在は注意ダイアログ、OUTPUT はあるが A-9 未実行（タイトル空）も「先に A-9」ダイアログで終了する（budget-nav-guard）。
  本ナビは予測を一切実行しない。
- **非保持（現行仕様 / 許容済み）**：A-9 を再実行すると `resetOutputSheet_` が OUTPUT を全消去するため、
  H/I の手入力は消え、H は再び月次P50初期値・I は空欄に戻る（=一回入力前提）。EVAL_INSIGHTS の入力保持 upsert とは
  方針が異なるが、これは **許容方針として確認済み**。再実行をまたいだ保持（upsert/別シート化）は将来の任意拡張（§12）。
- **年度合計の非加法性（注意）**：年度 Adopted Forecast 合計（=Σ月次H＝Σ月次P50）は、混合セクション上部の年度
  Baseline(P50)（=annualMixedCalSim の P50）と一致しない。これは分位点非加法性に由来する想定どおりの差で、
  予算欄は加法的な月次積み上げ、Baseline(P50) は年度sim分位点という別概念のため（§12・将来検討）。

---

## 10. AI調査パイプライン（Vertex AI / A-4 = runVertexAIResearch）

### 10.1 全体像
A-4はメーカー名（CONFIG!B2）について、4 topic（Market/Competitor/Channel/DX）ごとに以下を実行する：
1. **grounded web検索**（callVertexGeminiGrounded_）：Gemini + googleSearch（不可時 googleSearchRetrieval へフォールバック）。
2. **RAG検索**（callVertexSearchRAG_）：Vertex AI Search データストアを検索。DATASTORE_ID/SEARCH_LOCATION 未設定時はskip（web-only）。
3. **構造化**（callVertexGeminiStructured_）：web結果をevent行、RAG結果をbenchmark行として JSON 構造化。**temperature=0 固定**。
   構造化プロンプトで report_text を短文圧縮で生成させる（§10.9）。

### 10.2 readiness 判定
- readVertexConfig_ が geminiReady（projectId+location+geminiModel）と ragReady（datastoreId+searchLocation）を分離。
- A-4の入口判定は geminiReady。RAG未設置でも web-only で実行できる。
- AI_RESEARCH_ENABLED=0 のときは「無効」メッセージで終了。geminiReady=false のときは「必須設定未入力」エラーで停止。

### 10.3 スコア化（event / benchmark）
- **event_score**：`direction符号 × (impact_score/100) × 50 × confidence` を ±50 にclamp。impact_score は 0〜100（50は中立ではない）。
- **benchmark_score**：`(relative_percentile-50) × relative_confidence × quality倍率(high1/medium0.75/low0.5)` を ±50 にclamp。根拠があるときだけ出力。
- readAIResearchScores_ が topic別に blend：Market=0.65:0.35 / **Competitor=0.70:0.30** / Channel=0.65:0.35 / DX=0.50:0.50。各topic ±50 にclamp。

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
- 旧2枚→RAW 移送はヘッダ名マッチ＋シート名由来 axis（WEB→web / EXTERNAL→rag）。frozen_flag は無ければ 0、frozen_at は空。移送は一度だけ。

### 10.7 CONFIGキー（既定値）
- VERTEX_PROJECT_ID = forecast-agent-498907 / VERTEX_LOCATION = global / VERTEX_GEMINI_MODEL = gemini-3.1-pro-preview
- VERTEX_DATASTORE_ID = fujikeizai-portfolio-2025 / VERTEX_SEARCH_LOCATION = global / VERTEX_SERVING_CONFIG = default_search
- AI_RESEARCH_ENABLED = 1 / 必要OAuthスコープ：cloud-platform / script.external_request

### 10.8 RAGクエリのフレーム整合（rag-query-frame-align）
- `buildRagQuery_` は外部観測可能な「支援需要に影響しうる外部環境」に整合。topic 別の日本語展開語で外部環境を狙う。

### 10.9 report_text の短文圧縮（report-text-compress / 2.3.41）
- 変更スコープは `buildVertexStructureSystemInstruction_` と `buildVertexStructureUserContent_` の2関数のプロンプトのみ。
- 圧縮ルール：topicごと2〜3文・最大120字程度、明細列挙禁止、ソース転記・出典マーカー禁止、結論を1文目に。
- 不変点：event_score / benchmark_score / blended_score の算出式、coverage / quality / neutralized 判定は変えない。
- 非遡及：A-4 を再実行しない限り既存 report_text は前回の長文のまま。
- 波及：AI_RESEARCH サマリービュー①（topic別要約文）が短くなり横伸びが解消。A-9 の OUTPUT は不変。

---

## 11. AI調査サマリービュー（AI_RESEARCH シート）
A-4実行時に writeAIResearchSummaryView_ が AI_RESEARCH シートを再描画する（ユーザ表示シート）。

### 11.1 構成（3段）
- ① topic別サマリー（要約文）：report_text を topic単位で表示。report-text-compress 後は短文。
- ② AIスコア サマリー（4軸）：topic別 Final Score / event_score / benchmark_score / 最新as_of / 備考。
- ③ スコア根拠（event / benchmark 明細）。個別の根拠は③および AI_RESEARCH_STRUCTURED の evidence / relative_reason 側で確認。

### 11.2 再描画の不変条件
- A-4を複数回実行しても追記でなく再描画（重複しない）。A-4失敗（outRows空）時は上書きせず前回内容を保持。表示専用で予測コアに影響しない。

---

## 12. 既知の latent issue（非ブロッキング / 別スコープで対応）

現存（2.3.48）の latent issue：
- **予算策定欄 H/I の A-9 再実行非保持**（§9.6）：`resetOutputSheet_` が都度ワイプするため H/I 手入力が消える。
  これは「一回入力前提」の現行仕様として **許容方針で確認済み**。再実行をまたいだ保持（upsert/別シート化）は将来の任意拡張で
  非ブロッキング（EVAL_INSIGHTS の保持 upsert とは方針が異なる）。
- **年度 Adopted Forecast 合計 ≠ 年度 Baseline(P50)**（§9.6）：分位点非加法性に由来する想定どおりの差。実害なし（H列ヘッダ NOTE に
  趣旨を補足することは可能だが必須ではない）。
- **enable系フラグの `Number(x || 0)` パターン**（RELIABILITY_APPLY_ENABLED / AI_RESEARCH_ENABLED 等）：空セルで OFF に倒れる。
  既定 ON の RELIABILITY_APPLY_ENABLED は空 SOURCE_RELIABILITY では no-op のため実害なし。
  ※ チューニング数値読取側の空セル罠は config-blank-guard（2.3.47）で是正済み。enable フラグ側は別系統で現状維持。

解消済みメモ：
- **CONFIG数値読取の空セルの罠**（`getCfg` / `readAiMissingConfidenceDefault_`）は **config-blank-guard（2.3.47）で是正**。
  空/欠落を既定値へフォールバックし、`Number('')===0 → isFinite` でクランプ下限へ化ける問題を解消。明示0は従来どおり採用（§3.3）。
- **OUTPUT の coerceMatch / coerceCount 死にコード**（`writeOutputFY_` 内の恒常非発火だった `warn_coerced=` 判定）は
  **config-blank-guard（2.3.47）で物理削除**。
- **OUTPUT 6行目の赤字判定 step3aHasError 死にコード**は **output-display-polish（2.3.48）で物理削除**
  （6行目の常時黒字化に伴う / §9.5）。
- client名マッチの不整合は client-match-unify（2.3.32）/ a9-client-normalize（2.3.33）で解消。
- AI_RESEARCH_RAW 旧データ移送の列整合は raw-migrate-header-match（2.3.31）で解消。
- FORECAST_SNAPSHOT の三角測量系 vestigial カラムは snapshot-vestigial-removal で物理削除済み。
- EVAL_INSIGHTS の再実行で人手入力が消える問題は dashboard-hide-insights-upsert（2.3.38）で解消（B-4 で複合キー upsert）。
- AI調査 report_text の長文化（サマリービュー①横伸び）は report-text-compress（2.3.41）で解消。
- calibration override の空欄/明示0 取り違えは calibration-blank-override-fix で解消済み（getCfg 側も config-blank-guard で同型是正済み）。

将来検討（実害なし / 別スコープ）：
- `normalizeClientName_` の正規化辞書が現状 ｳﾞｨｱﾄﾘｽ系のハードコードのみ。汎用正規化は未対応。
- VERSION の 2.3.40 飛ばし（2.3.39→2.3.41）は版番号の運用上の隙間で、動作影響は無い。

---

この v2.3.10 は、annual-forecast-mode（FORECAST_CLOSED_MONTH_MODE）の方針「A=通年予測」、client-match-unify /
a9-client-normalize の正規化統一、overlay-cap-raise（主観月次cap 0.40・AI上限 ±5%）、spot-bg-sensitivity-down
（背景SPOT shrink 0.35・spike MAD_K 2.5）、report-text-compress（AI調査要約の短文圧縮）、2.3.42〜2.3.46 の表示・メニュー・
ナビゲーション増分（DASHBOARD の B-3 移設、OUTPUT 混合セクションの予算策定欄 H/I/J 追加、GUIDE A-10 同期、A-10 ナビガード）を
維持しつつ、2.3.47〜2.3.48 の2増分（config-blank-guard による CONFIG 数値読取の空セルガード＋ coerceMatch 死にコード削除、
output-display-polish による OUTPUT 6行目の常時黒字化＋赤字判定 step3aHasError 死にコード削除・A-10 ナビの範囲選択化・K列幅是正）を
現行実装へ同期したドキュメント改訂である。config-blank-guard は正常シード済み book で予測不変（空セルが残る異常 book でのみ
既定値フォールバックで予測が変わりうる＝是正であり回帰ではない）、output-display-polish は表示・ナビゲーションのみで予測不変。
スコア算出式・予測コア・履歴スキーマは無変更、A-9 の年度合計・月次 P10/P50/P90 は不変。年度合計は常に12ヶ月すべての予測simから
算出する通年予測であり、実績は検証・学習経路（B/C）でのみ使用する。予算策定欄は OUTPUT 上での手入力確定欄であり計画の主数値
（混合 P50）そのものではない。ブロッキングな残存 latent issue は無い。
