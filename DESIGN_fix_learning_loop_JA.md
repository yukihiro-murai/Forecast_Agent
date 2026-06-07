# 設計仕様：学習ループのサンプル水増し修正 + kProd全月1.0のフェイルセーフ化

**このファイルの用途**：今回の修正の設計根拠（なぜ・何を・どこまで）を記述した参照資料。
**Codex CLI には渡さない**（実装プロンプトは別ファイル `PROMPT_fix_learning_loop.txt`）。
コードレビュー・将来の振り返り・設計書本体（v1.x）への反映時の出典として使う。

対象：`Forecast_Agent.js`（現状 VERSION='2.2.0-dev' / BUILD_STAGE='v8-step3c3c'）
修正後：VERSION='2.2.1-dev' / BUILD_STAGE='v8-step3c3c-fix-learnloop'

---

## 0. 修正スコープの宣言

本修正は **確定バグ2件のみ** を対象とする。挙動の最適化・リファクタリング・新機能は含めない。

- 修正A（高）：A-9再実行で信頼度学習のサンプル数 n が水増しされ、closed月が surprise=0 で脱落する。
- 修正B（中）：FACTORS_PRODUCT に有効行があるのに kProd が全月1.0だと hard throw でA-9全体が止まる。

横断プール（3c-3c）・年間総計セマンティクス・設計書の版ずれ（DESIGN v1.8 とコードの乖離）は
本修正の対象外。設計書同期は別途ドキュメント改訂で行う。

---

## 1. 修正A：学習ループのサンプル水増し

### 1.1 現象（バグの再現経路）

1. GUIDE は「確認→修正→再実行を前提とする（一発確定しない）」と明記。同じ月に対しA-9を複数回押す運用が正規。
2. `runPhase1Forecast` は毎回 `writeAIHistoriesForRun_` / `writeSubjectiveImpactHistory_` を
   `getLastRow()+1` に**追記**する（dedupなし）。
3. C-1の `computeReliabilityHitStats_` は `data.subjectiveImpacts` を run_id でまとめず
   1行ずつ `g.n += 1` する。結果 `n ≈ A-9実行回数 × 月数` に膨張。
4. `generateReliabilityProposals_` は `n >= RELIABILITY_MIN_SAMPLES(2)` で提案を通し、
   confidence判定も `n >= 6 / n >= 3` を使う。→ **入力を変えずA-9を数回押すだけで提案が通り「高」信頼度に化ける**。

### 1.2 併発する closed月脱落

`runForecastFYCore_` は closed月について `objOnly`(=`quantOnly` と同一参照)を actual で上書きする。
そのため closed後にA-9を再実行すると、AI_IMPACT_HISTORY の `pred_p50_quant_only` に **actual が記録**される。
C-1の `quantByMonth` は month単位 last-write-wins なので、最新（=actual上書き後）の行を拾い、
`surprise = sign(actual - quant) = 0` となってその月が hit集計から丸ごと落ちる。

### 1.3 採用する解（確定方針）

- **dedup単位**：`(client, target_month, source_type, source_key)`。同一キーは **最新 run_at の1件のみ採用**。
  → 同じ月を何度A-9しても、その月・そのソースのサンプルは1件。n が実行回数に依存しなくなる。
- **評価対象**：**その月が forecast_open だった時点の forecast値** を使う。
  closed後の actual上書き行は評価から除外する。

この2つを両立させるには「その行を記録したとき、その月が open だったか closed だったか」を
履歴行が保持している必要がある。現行スキーマには無いため、列を追加する（1.4）。

### 1.4 スキーマ変更（必要最小限）

以下2シートに `forecast_source` 列を**末尾に追加**する（既存列の順序は変えない）。

- `AI_IMPACT_HISTORY`：末尾に `forecast_source` を追加
  （現行ヘッダ末尾は `disabled_topics_count`）
- `SUBJECTIVE_IMPACT_HISTORY`：末尾に `forecast_source` を追加
  （現行ヘッダ末尾は `source_updated_at`）

記録値は `runForecastFYCore_` の `result.sourceByMonth[i]`（`'forecast_open'` / `'actual_closed'`）をそのまま書く。

### 1.5 後方互換（既存の汚染データの扱い）

既存bookには `forecast_source` 列が無い過去行が大量にある（=水増し済みの汚染データ）。
これらは **C-1集計の対象外**とする（forecast_source が空/欠落の行は読まない）。

理由：過去行は run_at 単位の重複を含み、open/closed の区別も付かない。安全側に倒し、
列追加後の新しい記録から正しくやり直す。`ensureSheetHeaders_` で列を足すと既存行の当該セルは空になるため、
「forecast_source が 'forecast_open' に正規一致する行のみ採用」とすれば自動的に旧行は除外される。

これは「挙動変更」に見えるが、バグの是正に不可分な配線であり、本修正スコープ内とする。
（誤った学習が止まり、正しい蓄積に切り替わるだけで、新しい価値判断は加えない。）

### 1.6 触らない箇所（誤修正防止）

- `evalActualByMonth`（EVAL_LOG由来の actual）は month単位 last-write-wins のままでよい。
  actual は確定値なので同月重複でも値は同じ。ここは変更しない。
- `generateQuarterlyProposals_` のAI方向一致率（hitRate）は比率ベースで n の絶対数に依存しないため、
  今回の dedup の主目的ではない。ただし同じ dedup済みデータを参照するよう整合は取る。
- 予測コア（`forecastMonteCarloMixed_` 等）の計算は一切変更しない。OUTPUTのP10/P50/P90は不変。

---

## 2. 修正B：kProd全月1.0のフェイルセーフ化

### 2.1 現象

`runForecastFYCore_` 末尾：

```
if (factorsProduct.length > 0 && mixed.diagnostics && mixed.diagnostics.kProdByMonth
    && mixed.diagnostics.kProdByMonth.every(k => Math.abs(Number(k || 1) - 1) < 1e-9)) {
  throw new Error('FACTORS_PRODUCT に有効行がありますが、kProd が全月1.0です。...');
}
```

`kProd` は `productWeights.has(p) ? get(p) : 0`。直近12ヶ月の closed BASE実績が無い製品
（新製品・SPOTのみ・履歴窓外）に要因を入れると weight=0 → kProd=1.0 → throw。
「factorが効かなかった」という正当な状況をA-9停止に変えている。

### 2.2 採用する解

throw を廃止し、**OUTPUTへの警告表示＋ログ記録**に置き換える（A-9は最後まで完走する）。
警告には「kProdが全月1.0だった」事実と、weight=0 となった製品名（productWeights に乗らなかった
FACTORS_PRODUCT の製品名）を列挙し、ユーザーが原因（製品名キーの不一致 or 履歴なし）を
自分で確認できるようにする。

フェイルセーフ原則：エラーでメニュー実行をブロックしない。中立（kProd=1.0）で予測は成立する。

### 2.3 警告の出し先

- OUTPUT 上部の既存警告ブロック（エンジン/cap/reliability行の近辺）に1行追加するのが自然。
  実装が重ければ、`runPhase1Forecast` の完了 toast / RUN_LOG の note への記録でも可。
  **最低限 RUN_LOG に残す**（監査証跡）。throwしないことだけは必須。

---

## 3. バージョン

- `VERSION = '2.2.1-dev'`
- `BUILD_STAGE = 'v8-step3c3c-fix-learnloop'`
- 既存の `DLM_BUILD_STAGE` は変更しない（DLMロジックは無変更のため）。

---

## 4. 受け入れ基準（要約）

1. A-9を入力不変で複数回実行 → C-1の reliability提案の n が増えない（実行回数非依存）。
2. closed月について、open時点の forecast値で surprise が評価され、hit集計から脱落しない。
3. 新規book（履歴なし製品にFACTORS_PRODUCT入力）でA-9が throw せず完走し、警告が出る。
4. 既存の forecast_source 無し履歴行はC-1集計に混ざらない。
5. OUTPUTのP10/P50/P90は本修正で変化しない（予測コア不変）。
