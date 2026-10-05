# POOL_SETUP_v1.9 — 中央集約book（ハブ）セットアップ手順

横断プール（3c-3c）の「方式A：中央集約 + fan-out」で使うハブbookを立ち上げる手順。
**最初に1回だけ**やる作業。以後は四半期に1回 ハブで集約関数を回すだけ。

---

## 0. 用語と全体像

- **ハブbook**：各クライアントbookの RELIABILITY_EVIDENCE を集めて、各bookの POOL_PRIOR に書き戻す中心。
- **方式A**：ハブが各bookを openById で読み書きする。クライアントbookは何も実行しない（C-1で勝手にエビデンスが溜まるだけ）。
- **手動・低頻度**：集約はメニューに出さない admin関数を、スクリプトエディタから手で実行。四半期に1回想定。

```
[client book A] ─┐ RELIABILITY_EVIDENCE(raw hit/n)
[client book B] ─┼─→ (ハブが openById で読む) ─→ source_type単位で Σhit/Σn ─→ pooled_value
[client book C] ─┘                                                              │
                                                                                ├─→ ハブの POOL_PRIOR に upsert
                                                                                └─→ 各 client book の POOL_PRIOR に fan-out
```

---

## 1. ハブbookをどれにするか

2択。どちらでも動く。

- **(推奨) 既存のクライアントbookを兼用**：すでに同じ Forecast_Agent.js が入っているので追加のclasp設定が不要。
  「主要な1社のbook」をハブ兼用にするのが最も手間が少ない。
- **専用bookを新設**：集約専用に空のbookを作り、同じスクリプトを bind/push する。
  クライアントデータと混ざらず概念的にきれいだが、clasp のpush先（scriptId）が増える。

> 単独技術者運用なら兼用で十分。専用にしたい場合のみ、新bookにスクリプトを紐付けて push しておくこと。

ハブにしたbookには、後述の admin関数で **POOL_REGISTRY / POOL_AGGREGATION_LOG / POOL_PRIOR** が作られる。
（クライアント運用シートには影響しない。兼用でも既存シートは壊れない。）

---

## 2. ハブの初期化（adminSetupPoolHub）

1. ハブbookを開く
2. 拡張機能 → Apps Script でスクリプトエディタを開く
3. 関数一覧から **adminSetupPoolHub** を選んで実行
   - 初回は権限承認ダイアログが出る → 承認
4. 確認ダイアログで OK
5. 完了後、ハブbookに次の3シートができていることを確認：
   - **POOL_REGISTRY**（登録表）
   - **POOL_AGGREGATION_LOG**（集約の監査ログ）
   - **POOL_PRIOR**（横断事前の中央マスタ）

> adminSetupPoolHub と adminAggregatePoolPriorAcrossBooks は **メニューに出ない**。
> adminInitDLMAndBacktest と同じく、スクリプトエディタから関数を選んで実行する運用。

---

## 3. book_id の調べ方

各クライアントbookのURLを見る：

```
https://docs.google.com/spreadsheets/d/【ここがbook_id】/edit#gid=0
```

`/d/` と `/edit` の間の長い文字列が book_id。これを POOL_REGISTRY に登録する。

---

## 4. POOL_REGISTRY への登録

POOL_REGISTRY シートに、集約したいクライアントbookを1行ずつ記入。

| 列 | 内容 | 例 |
|---|---|---|
| book_id | 各bookのID（手順3） | 1AbC...xyz |
| client_name | 参考用の名前（任意） | ○○製薬 |
| enabled | 集約対象なら 1、外すなら 0/空 | 1 |
| note | メモ（任意） | 2026Q1から参加 |

ルール：
- **enabled=1 の行だけ集約対象**。一時的に外したいbookは 0 にする（行は消さなくてよい）。
- **最低2冊を enabled=1** にする。POOL_MIN_CLIENTS=2 のため、1冊だけだと何も書かれない（フェイルセーフ）。
- book_id は文字列。先頭の `=` などは付けない。
- ハブbook自身を登録してもよい（自分も1クライアントなら）。二重には集計されない設計。

---

## 5. 前提：各bookにエビデンスがあること

集約の入力は各bookの RELIABILITY_EVIDENCE。これは **各bookで C-1 を実行したときだけ**書かれる。

- 登録した各bookで、過去に A-9（予測）→ B-1（実績取込）→ B-2（検証）→ C-1（四半期レビュー）を回していること。
- C-1が「期間不足」で終わるbook（実績3ヶ月未満）はエビデンスが無い → 集約に寄与しない。
- 急ぐ必要はない。各bookで通常運用を続ければ RELIABILITY_EVIDENCE は自然に溜まる。

---

## 6. 集約の実行（四半期に1回）

1. ハブbookのスクリプトエディタで **adminAggregatePoolPriorAcrossBooks** を実行
2. 完了 alert を確認（対象book数 / ok数 / 除外数、書いた scope と pooled_value、書かなかった scope と理由）
3. POOL_AGGREGATION_LOG で詳細を確認（→ MANUAL_TASKS_v1.9 フェーズ3〜5）
4. ハブと各 client book の POOL_PRIOR に値が入っていることを確認

これで、各bookの次回 C-1 から、横断 pooled_value が reliability 提案の収縮に使われる。

---

## 7. 運用上の注意

- **POOL_PRIOR（client book側）は触らない**：集約専用の従属データ。手編集しても次の集約で上書きされて消える。
  reliability を手で固定したい場合は POOL_PRIOR ではなく SOURCE_RELIABILITY を使う（こちらは集約で消えない）。
- **コードは全bookで同一に保つ**：clasp push 漏れがあると、旧コードのbookはエビデンス形式が違って集約から漏れる。
- **権限**：ハブを実行するアカウントが、登録した全 client book に編集権限を持っていること（方式Aの前提）。
  権限の無いbookは status=excluded でログに残り、他bookの集約は止まらない（が、そのbookには反映されない）。
- **頻度**：横断priorはゆっくりしか動かないので、四半期に1回で十分。トリガー自動化はしない（失敗が見えにくいため手動）。

---

## 8. 次のステップ（fast-follow / 今回はやらない）

複数bookで集約が安定したら検討：
- `reliability:ai_topic:{topic}` の topic 単位プール（ai_topic は person と違い横断可能）。
- 経験ベイズによる precision の動的化（クライアント間分散から収縮の強さを決める）。

いずれも消費側の小改修が要るので、本体（source_type 単位プール）が実運用で問題ないことを確認してから。
