# step 3c-3c-1 手動作業チェックリスト（raw hit/n 永続化の検証）

PROMPT_3c3c_step1.txt をCodexに渡してコードを更新した後、上から順に実施する。
本ステップは予測値が変わらないことが前提（変わったら何かが間違っている）。

---

## フェーズ0：版の整合（先に解消する）

- [ ] **VERSION の食い違いを解消**：v1.8 のコードは `2.1.1-dev`、`MANUAL_TASKS_v1.8.md` フェーズ0は
      `2.1.0-dev` を確認、と書いている。本ステップ適用後は `2.1.2-dev` になる。
      → `MANUAL_TASKS_v1.8.md` を `2.1.1-dev` に直す（コードを正とする）か、逆かを決め、
        今後のチェックリストの版番号を実コードに合わせる。放置すると検証が必ずFAILする。
- [ ] PROMPT_3c3c_step1.txt をCodex CLIに渡し Forecast_Agent.js を更新
- [ ] VERSION が `2.1.2-dev`、BUILD_STAGE / DLM_BUILD_STAGE が `v8-step3c3c-1` であること
- [ ] clasp push（または手動コピー）→ スプレッドシートを開き直し、メニューが正常表示されること

---

## フェーズ1：シート作成の確認

- [ ] A-1 初期セットアップを実行（検証用bookで）
- [ ] `RELIABILITY_EVIDENCE` シートが作成されている
- [ ] ヘッダが
      `client / source_type / source_key / quarter_label / quarter_end_month / n / hit / hit_rate / computed_at / run_id / note`
- [ ] タブ色が内部管理色、かつ非表示（ユーザーには見えない位置）になっている

---

## フェーズ2：全ソースの hit/n が残ることの確認（本ステップの核）

前提：実績3か月分について A-9 → B-1 → B-2 が済んでおり、EVAL_LOG に neutral 行が3か月分、
SUBJECTIVE_IMPACT_HISTORY / AI_IMPACT_HISTORY に同じ月のデータがあること。

- [ ] C-1 四半期レビューを実行
- [ ] `RELIABILITY_EVIDENCE` に source_type:source_key ごとの行が書かれている
- [ ] **安定ソース（reliability変化が MIN_CHANGE 未満で提案化されないソース）も、n>=1 なら記録されている**
      （ここが従来との差。QUARTERLY_REVIEW_LOG には提案だけ、こちらには全件）
- [ ] `n` `hit` が整数、`hit_rate = hit/n` になっている
- [ ] `quarter_label`（例 FY2026-Q1）と `quarter_end_month`（yyyy/MM）が正しい

---

## フェーズ3：upsert と予測不変の確認

- [ ] 同じ四半期で C-1 を再実行 → `RELIABILITY_EVIDENCE` の行が**重複せず上書き**される
- [ ] C-1 実行の前後で A-9 を実行し、OUTPUT の計画値（P10/P50/P90）が**1円も変わらない**
- [ ] （回帰確認）QUARTERLY_REVIEW に出る reliability 提案が、リファクタ前と同じ内容であること

---

## フェーズ4：横断（複数book）での蓄積確認

- [ ] 2冊目以降のbookでも同手順で C-1 を実行 → 各bookの `RELIABILITY_EVIDENCE` に
      そのbookぶんの hit/n が貯まる（book間で混ざらない）
- [ ] 集約に乗せたいbookが何冊あるか、それぞれの spreadsheet_id を控える
      （3c-3c-2 の POOL_REGISTRY 登録に使う）

---

## トラブルシュート：C-1 後に RELIABILITY_EVIDENCE が空

確認順：
1. data.ready か（実績3か月の neutral 行が EVAL_LOG にあるか。B-2 を3か月分実行しているか）
2. SUBJECTIVE_IMPACT_HISTORY に同じ月の push 行があるか（A-9 実行時にしか書かれない）
3. surprise方向が確定した月があるか（actual と pred_p50_quant_only が同値=flat の月は n に入らない）
→ いずれも「3か月の運用履歴」が前提。空でもエラーにはならない（フェイルセーフ）。

---

## 完了条件 → 3c-3c-2 へ

- [ ] 安定ソースを含む全ソースの hit/n が、提案有無に関係なく永続化される
- [ ] 予測値が不変であることを確認
- [ ] 集約対象bookの spreadsheet_id を把握

ここまで確認でき、かつ DESIGN_3c3c_JA.md 第8章の transport前提
（全bookが1サービスアカウント／実行者から openById で読み書き可能か）が確定したら、
3c-3c-2（中央集約 → fan-out）の Codex プロンプトを作成する。
