# v1.9 手動作業チェックリスト（横断プール集約の反映と検証）

このファイルは、Codexでコードを更新した後にTechiが手で行う作業と検証手順をまとめたもの。
上から順に実施すること。中央集約book（ハブ）の初期セットアップ詳細は **POOL_SETUP_v1.9.md** を参照。

横断プール（3c-3c）は **POOL_PRIOR が閾値を満たして埋まるまで予測も提案も一切変わらない**純加算的な機能。
検証は「集約が走る → POOL_PRIOR に書かれる → fan-out される → C-1提案が収縮する」の配線確認が中心。

---

## フェーズ0：コード反映前の確認

- [ ] PROMPT_v1.9.txt をCodex CLIに渡し、Forecast_Agent.js を更新
- [ ] VERSION が `2.2.0-dev`、BUILD_STAGE が `v8-step3c3c` に上がっていることを確認
- [ ] 予測ロジック・FORECAST_SNAPSHOT/REPORTの列・消費側（generateReliabilityProposals_）に差分が**無い**ことを確認
      （今回の差分は admin関数2本＋シート定数2つ＋CONFIG 1行＋HOW TO TESTのみのはず）
- [ ] clasp push で**全bookへ**反映（ハブ・全クライアントbookに同一コードが入ること）
- [ ] 各bookを開き直し、Forecast Agent メニューが従来通り（新項目が増えていないこと＝admin関数は非掲載）

> 注意：横断集約は各bookに同じコードが入っていることが前提。1冊でも旧コードのbookがあると、
> そのbookは RELIABILITY_EVIDENCE の形が違って集約から漏れる可能性がある。push漏れに注意。

---

## フェーズ1：ハブbookの用意（中央集約book）

詳細は POOL_SETUP_v1.9.md。ここでは要点だけ。

- [ ] ハブにするbookを1つ決める（既存のクライアントbookを兼用してよい／専用bookでもよい）
- [ ] ハブbookのスクリプトエディタで **adminSetupPoolHub** を1回だけ実行
- [ ] POOL_REGISTRY / POOL_AGGREGATION_LOG / POOL_PRIOR の3シートが作成されたことを確認
- [ ] POOL_REGISTRY に集約対象の client book を登録
      - book_id：各bookのURLの `/d/` と `/edit` の間の文字列
      - client_name：参考（任意）
      - enabled：集約に含めるなら 1
- [ ] 最低2冊以上を enabled=1 で登録（POOL_MIN_CLIENTS=2 のため、1冊では何も書かれない）

---

## フェーズ2：各bookにエビデンスが溜まっているか

横断集約の入力は各bookの RELIABILITY_EVIDENCE（raw hit/n）。これは C-1 実行時にしか書かれない。

- [ ] 登録した各 client book で C-1（四半期レビュー）が実行済みで、RELIABILITY_EVIDENCE に行があること
      - 各bookで RELIABILITY_EVIDENCE シートを開き、source_type / source_key / n / hit が入っていることを確認
      - 行が無いbook：そのbookで C-1 を実行（実績3ヶ月＝A-9→B-1→B-2 の履歴が前提。無いと「期間不足」で書かれない）
- [ ] 少なくとも2冊で、同じ source_type（例 opinion）に行があること
      （別々の source_type にしか行が無いと、その source_type の n_clients が1になり書かれない）

> エビデンスが揃わない場合は、横断集約を急がず、各bookで通常運用（A-9→B-1→B-2→C-1）を回して
> RELIABILITY_EVIDENCE を育てるのが先。集約は後からいつでも回せる。

---

## フェーズ3：集約の実行と監査ログ確認

- [ ] ハブbookのスクリプトエディタで **adminAggregatePoolPriorAcrossBooks** を実行
- [ ] 完了 alert に「対象book数 / ok数 / 除外数」と、書いた scope・pooled_value、書かなかった scope・理由が出ること
- [ ] POOL_AGGREGATION_LOG を開き、同じ run_id で以下が記録されていること
      - per-book 行：各登録bookの status（ok / excluded / empty / no_columns）、rows_read、rows_skipped
      - per-scope 行：reliability:factor_product / factor_client / opinion / ai_topic ごとの
        written（true/false）、reason、n_clients、sum_hit、sum_n、hit_rate、pooled_value、precision
- [ ] pooled_value が `clamp(2 × sum_hit/sum_n, R_MIN, R_MAX)` と手計算で一致すること（1つ抜き取り検算）
- [ ] precision が SHRINKAGE_K（既定4）であること

---

## フェーズ4：POOL_PRIOR への書き込みと fan-out 確認

- [ ] ハブbookの POOL_PRIOR に `reliability:{source_type}` 行が upsert されていること
      （pooled_value / precision / n_clients / updated_at / note）
- [ ] 登録した**各 client book の POOL_PRIOR にも同じ値が fan-out**されていること
      （client bookを開いて POOL_PRIOR を確認。ハブと同値・同 note）
- [ ] written=false だった scope（min_clients 等）は POOL_PRIOR に**書かれていない**（空のまま）こと（判断2）

> 運用注意（判断3）：クライアントbookの POOL_PRIOR は集約専用の従属データ。
> 手で編集しても次の集約で上書きされて消える。POOL_PRIOR を直接いじらないこと。

---

## フェーズ5：フェイルセーフの確認

- [ ] **min_clients 未達**：ある source_type の寄与が1冊だけになる状態を作り（例：1冊だけ enabled）、
      その source_type が written=false / reason=min_clients でログに出て、POOL_PRIOR に書かれないこと
- [ ] **1冊除外でも止まらない（判断4）**：POOL_REGISTRY に無効な book_id（権限なし or 存在しない）を1行 enabled=1 で入れ、
      実行 → その行が status=excluded でログに残り、他bookの集約と fan-out は正常完了すること
- [ ] 確認後、無効行は削除 or enabled=0 に戻す

---

## フェーズ6：予測・提案への波及確認（最重要）

- [ ] **no-op 確認**：集約直後に、いずれかの client book で A-9（予測）を実行 → OUTPUT の P10/P50/P90 が
      集約前と変わらないこと（POOL_PRIOR は C-1提案の収縮にだけ効き、予測値そのものは動かさない）
- [ ] **C-1波及確認**：同じ client book で C-1 を実行 → generateReliabilityProposals_ の提案 rShrunk が
      POOL_PRIOR の pooled_value 方向へ収縮していること
      - 比較方法：POOL_PRIOR を空にした状態の提案値（≒ rHat 寄り）と、埋めた状態の提案値（pooled_value 寄り）を見比べる
      - 収縮の強さは precision（=k=4）と各bookの n で決まる。n が小さいほど pooled へ強く寄る

---

## トラブルシュート

**Q. adminAggregatePoolPriorAcrossBooks が「POOL_REGISTRY を作成してください」で止まる**
→ ハブbookで adminSetupPoolHub を未実行、または enabled=1 の行が無い。フェーズ1をやり直す。

**Q. per-scope 行が全部 written=false（min_clients）になる**
→ 各source_typeに寄与している distinct book が2冊未満。原因は次のいずれか：
  1. 登録bookが1冊しか enabled=1 になっていない
  2. 2冊登録しているが、片方の RELIABILITY_EVIDENCE が空（C-1未実行）
  3. 2冊にエビデンスはあるが、別々の source_type にしか行が無い（例：A bookは opinion のみ、B bookは ai_topic のみ）
→ フェーズ2に戻り、同じ source_type に2冊以上の行がある状態を作る。

**Q. 特定bookが status=excluded になる**
→ POOL_AGGREGATION_LOG の reason を確認。多くは book_id の打ち間違い or 共有権限なし。
  ハブを実行するアカウントが、その client book に編集権限を持っているか確認（方式A前提）。

**Q. fan-out されない（ハブのPOOL_PRIORには書かれるが client book に反映されない）**
→ そのbookが status='ok' でなかった可能性（fan-out は ok bookのみ対象）。per-book 行の status と
  fanout_status を確認。

---

## 完了条件

- [ ] フェーズ3：集約が走り、POOL_AGGREGATION_LOG に per-book / per-scope が記録される
- [ ] フェーズ4：ハブと各 client book の POOL_PRIOR に同値が入る（written=true の scope のみ）
- [ ] フェーズ5：min_clients 未達は書かない／1冊除外で止まらない、を確認
- [ ] フェーズ6：no-op（予測不変）と、C-1提案が pooled_value へ収縮することを確認

ここまで確認できたら横断プールの本体は完成。次の検討は fast-follow（ai_topic 単位プール／経験ベイズ precision）。
当面は四半期に1回 adminAggregatePoolPriorAcrossBooks を回す運用でよい（高頻度実行は不要）。
