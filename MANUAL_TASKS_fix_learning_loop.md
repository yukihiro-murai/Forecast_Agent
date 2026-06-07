# 手動検証チェックリスト：学習ループ修正 + kProdフェイルセーフ化

**このファイルの用途**：Codexで `Forecast_Agent.js` を更新した後、Techiが手で行う検証手順。
上から順に実施する。検証は本番bookではなくコピー/検証用bookで行うのが安全。

適用後の期待状態：VERSION='2.2.1-dev' / BUILD_STAGE='v8-step3c3c-fix-learnloop'

---

## フェーズ0：反映確認

- [ ] PROMPT_fix_learning_loop.txt をCodex CLIに渡し、Forecast_Agent.js を更新
- [ ] VERSION が '2.2.1-dev'、BUILD_STAGE が 'v8-step3c3c-fix-learnloop' に上がっている
- [ ] DLM_BUILD_STAGE は変更されていない（'v8-step3c3c-1' のまま）
- [ ] clasp push でGASへ反映
- [ ] スプレッドシートを開き直し、Forecast Agent メニューが正常表示される

---

## フェーズ1：スキーマ変更の確認（修正A-1）

- [ ] 検証用bookで A-1 初期セットアップを実行
- [ ] AI_IMPACT_HISTORY のヘッダ末尾に `forecast_source` 列がある
- [ ] SUBJECTIVE_IMPACT_HISTORY のヘッダ末尾に `forecast_source` 列がある
- [ ] 両シートとも既存列の順序・名称が変わっていない（末尾追加のみ）

---

## フェーズ2：forecast_source が正しく記録される（修正A-2）

- [ ] 通常どおり A-2 → A-3 → A-4〜A-7（最低限の入力）→ A-8 → A-9 を実行
- [ ] AI_IMPACT_HISTORY の各行の forecast_source 列に 'forecast_open' または 'actual_closed' が入っている
- [ ] SUBJECTIVE_IMPACT_HISTORY の各行の forecast_source 列にも同様の値が入っている
- [ ] 予測実行月のうち、未来月は 'forecast_open'、確定済み月は 'actual_closed' になっている
      （実行時点の年月と予測対象FYの関係で確認）

---

## フェーズ3：n が水増しされないこと（修正A-3の核心）

目的：同じ入力でA-9を複数回押しても、C-1のサンプル数が増えないことを実証する。

前提：reliability提案を出すには直近3ヶ月のneutral実績が必要（A-9→B-1→B-2を3ヶ月分）。
この前提が未整備なら、まず下の「簡易確認」で代替する。

### 3-A：本検証（3ヶ月履歴がある場合）

- [ ] 入力を一切変えずに A-9 を **3回** 連続実行する
- [ ] C-1 四半期レビューを実行
- [ ] QUARTERLY_REVIEW_LOG の reliability:* 提案行の rationale にある `n=...` を確認
- [ ] その n が「対象月数」程度であり、「対象月数 × 3」になっていないこと
      （修正前なら3倍に膨らむ。修正後は実行回数に依存しない）

### 3-B：簡易確認（3ヶ月履歴がまだ無い場合）

- [ ] SUBJECTIVE_IMPACT_HISTORY を直接開く
- [ ] 同一 (client, target_month, source_type, source_key) の行が、A-9実行回数ぶん重複して
      存在することを確認（=ログ自体は追記され重複する。これは想定どおり）
- [ ] ※ C-1集計側で「最新run_atの1件のみ」に畳むのが修正の本体。ログ重複そのものは残る設計。
      3ヶ月履歴が揃ったら 3-A で最終確認する。

---

## フェーズ4：closed月が surprise=0 で脱落しないこと（修正A-3）

目的：実績確定後にA-9を再実行しても、その月が hit集計から落ちないことを確認する。

- [ ] ある月の実績を B-1 で取り込み、その月を closed にする
- [ ] その後 A-9 を再実行する（→ その月の forecast_source が 'actual_closed' で追記される）
- [ ] C-1 を実行
- [ ] 当該 closed月の評価が、open時点の forecast値（'forecast_open' 行）で行われている
      （AI_IMPACT_HISTORY で、同月に forecast_open 行と actual_closed 行が両方あり、
       C-1は forecast_open 行の pred_p50_quant_only を使っているはず）
- [ ] forecast_source が空の旧行（A-1再実行前の履歴がある場合）はC-1集計に混ざっていないこと

---

## フェーズ5：kProd全月1.0で止まらないこと（修正B）

目的：履歴のない製品にFACTORS_PRODUCTを入れてもA-9が完走することを確認する。

- [ ] FACTORS_PRODUCT に、SALES_INPUT_MONTHLYの直近12ヶ月closed BASEに存在しない製品名で
      有効行（担当者・日付・Step・理由すべて入力）を1件作る
      （例：新製品名、またはSPOTにしか出ない製品名、または明らかな打ち間違い名）
- [ ] 他に kProd を動かす有効な製品要因が無い状態にする（その1行だけが効く想定）
- [ ] A-9 を実行
- [ ] **エラーで止まらず、OUTPUTまで完走する**（修正前はここで throw して止まっていた）
- [ ] RUN_LOG の最新行 note に kProd全月1.0の警告（prodw=... 等）と該当製品名が残っている
- [ ] （OUTPUT表示も実装した場合）OUTPUT上部の警告ブロックに同趣旨の1行が出ている
- [ ] 製品名キーを正しい（履歴のある）名前に直して A-9 を再実行 → 警告が消え、kProd が動く

---

## フェーズ6：回帰確認（予測コア不変）

目的：本修正で予測値が変わっていないことを確認する。

- [ ] 修正前の任意の正常ケースで OUTPUT の年度合計 P10/P50/P90 を控えておく
      （フェーズ前に控え忘れた場合は、Gitで修正前コミットに一度戻して記録）
- [ ] 同一入力で修正後 A-9 を実行し、OUTPUT の P10/P50/P90 が一致する
      （Monte Carloの乱数で微小な揺れは出るが、構造的な差は出ないこと）
- [ ] DLMセクション（shadow/primary）の値が修正前と一致する

---

## 完了条件

- [ ] フェーズ3：A-9複数回でも n が増えない（または3-Bで配線確認）
- [ ] フェーズ4：closed月が open時点forecastで評価され脱落しない
- [ ] フェーズ5：履歴なし製品でA-9が完走し警告が出る
- [ ] フェーズ6：予測値が修正前と一致（コア不変）

ここまで確認できたら、設計書（DESIGN_RECOMMENDATION_JA）の版ずれ解消（v1.9化）に進む。
