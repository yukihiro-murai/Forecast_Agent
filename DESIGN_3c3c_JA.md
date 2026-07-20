# step 3c-3c 設計書（POOL_PRIOR クライアント横断集約）

DESIGN_RECOMMENDATION_JA.md v1.8 の第7章「未確定」を確定させ、実装に落とす。
本書は v1.8 本体（信頼度ON・cap pass-through・LMDI）には手を入れない。追加のみ。

---

## 0. 目的と非目的

- 目的: 各book（=1クライアント）が単独では持てない reliability の事前分布（prior）を、
  全bookの「生の的中実績（hit/n）」を足し合わせて作り、各bookのC-1の収縮先に供給する。
- 非目的: 予測エンジン（BASE/SPOT/主観/AI/cap）の挙動変更。本改修は C-1 の prior 供給のみ。
- フェイルセーフ最優先: プールが無い・薄い・壊れている場合は、すべて中立 1.0（＝従来 no-op）に倒す。

---

## 1. 現状の前提と差分（v1.8 実装に対して）

### 1.1 既にある配線（変更不要）
- `POOL_PRIOR` シート: `pool_scope, param_key, pooled_value, precision, n_clients, updated_at, updated_by, note`（作成済み）
- `readPoolPrior_(scope)`: ローカルbookの POOL_PRIOR を読む（実装済み・読むだけ）
- `generateReliabilityProposals_`: prior を収縮に使用済み
  `rShrunk = clamp((n·rHat + k·rPool) / (n + k), R_MIN, R_MAX)`
  - 現状 POOL_PRIOR が空 → `rPool=1.0` / `k=RELIABILITY_SHRINKAGE_K(4)` で中立収縮（=フェイルセーフ既定）

### 1.2 欠落（3c-3cの前提ブロッカー）
- **raw hit/n がどこにも永続化されていない。**
  hit/n は `generateReliabilityProposals_` 内の `grouped` に一度だけ計算され、
  `if (g.n < MIN_SAMPLES) return;` と `if (delta < MIN_CHANGE) return;` を通過した
  提案だけが `diagnostic_metrics_json` 経由で `QUARTERLY_REVIEW_LOG` に残る。
  → 安定ソース（変化が小さいソース）の hit/n は捨てられている。
- `POOL_PRIOR` への**書き手（writer）が存在しない**。`readPoolPrior_` の対になる writer が無い。

### 1.3 結論
3c-3c は2段に分割する。3c-3c-1（永続化）はどのtransportを選んでも必須・単一bookで完結・予測不変。
先に3c-3c-1を入れ、複数bookで証拠が貯まることを確認してから3c-3c-2（集約・fan-out）に進む。

---

## 2. 分割

| ステップ | 内容 | 前提 | リスク | 予測挙動 |
|---|---|---|---|---|
| 3c-3c-1 | 全ソースの raw hit/n を `RELIABILITY_EVIDENCE` に永続化 | 単一book・実績3か月 | 低（追加のみ） | 不変 |
| 3c-3c-2 | 各bookの evidence を中央集約 → `POOL_PRIOR` 書込み → 各bookへ fan-out | 複数book・transport確定 | 中（横断書込み） | priorが変わるとC-1提案が変わる（A-9の計画値は承認後のみ変化） |

---

## 3. 5つの設計論点の確定

### 3.1 transport（論点1）
**採用: 中央プールbook → 各book POOL_PRIOR への fan-out 書込み（設計書の推奨案）。**
- 中央集約用に1冊「pool book」を用意。`aggregatePoolPrior_()`（管理者メニュー外・手動）が
  各bookの `RELIABILITY_EVIDENCE` を `openById` で読み、source_type別に集計して
  pool book の `POOL_PRIOR` に書く。
- `distributePoolPrior_()` が pool book の `POOL_PRIOR` を各bookの `POOL_PRIOR` へ upsert（fan-out）。
- **前提（要確認）**: 全bookが1つのサービスアカウント／実行者から `openById` で読み書き可能なこと。
  この前提が崩れる場合のみ、代替（実行時 openById 読み／手動コピー）に切替。
- book一覧は pool book の `POOL_REGISTRY`（client, spreadsheet_id, enabled）で管理（ハードコードしない）。

### 3.2 集約入力（論点2）
**raw hit/n のみ。rShrunk の再プールは禁止（自己強化ループ防止）。**
- 各bookは `RELIABILITY_EVIDENCE` に `n, hit` を生のまま残す（3c-3c-1）。
- 集約は `Σhit / Σn` で行い、reliability_r 値（rShrunk）は集計対象にしない。これは選択ではなく正しさの制約。

### 3.3 粒度（論点3）
**まず `reliability:{source_type}` のtype単位。**
- person系（factor_product / factor_client / opinion）: 横断キーが無いため、book内でも横断でも
  source_key（人名）を巻き上げて type 単位で集計。
- ai_topic: 普遍キーなので type単位（`reliability:ai_topic`）でv1運用。将来 `reliability:ai_topic:{topic}` に拡張可。
- 現行 `readPoolPrior_(\`reliability:${g.type}\`)` と一致 → 読み側は無改修。

### 3.4 precision導出（論点4）
**値=n加重平均、precision=保守的固定。**
- `rHat_pool = clamp(2 · Σhit / Σn, R_MIN, R_MAX)`
  （= 各bookの per-client rHat を n加重平均したものと数学的に同値）
- `precision k_pool = min(n_clients, RELIABILITY_SHRINKAGE_K)`
  - client数が少ないうちは prior の重みを抑える。K（既定4）を超えない。
- 経験ベイズ・leave-one-out はv1では実装しない（将来）。
- **v1既知の簡略化**: book c の証拠が c 自身の prior にも含まれる（部分的自己参照）。
  precisionをKで頭打ちにすることで影響を抑える。leave-one-outは将来課題として明記。

### 3.5 フェイルセーフ（論点5）
- `n_clients < MIN_CLIENTS`（新設 `RELIABILITY_MIN_CLIENTS`、既定2）→ そのscopeは pool を書かない（中立1.0据え置き）。
- `Σn < RELIABILITY_MIN_SAMPLES` → 同上。
- 書く場合も `[R_MIN, R_MAX]` クリップ。
- `n_clients`, `updated_at`, `updated_by` を必ず記録。
- 集約・fan-out のどこかで例外が出ても、各bookの予測は POOL_PRIOR 既定（無ければ1.0）で従来通り動く。

---

## 4. データ構造

### 4.1 新シート `RELIABILITY_EVIDENCE`（各book・内部管理・非表示）
各bookのC-1実行時に、その四半期の raw hit/n を全ソースぶん残す。

| 列 | 内容 |
|---|---|
| client | CONFIG!B2 |
| source_type | factor_product / factor_client / opinion / ai_topic |
| source_key | 人名 or topic名 |
| quarter_label | 例 FY2026-Q1 |
| quarter_end_month | yyyy/MM |
| n | 有効サンプル数（surprise方向が確定した月数の積み上げ） |
| hit | push方向 == surprise方向 だった回数 |
| hit_rate | hit / n（人間可読用） |
| computed_at | 実行日時 |
| run_id | C-1 の review_id |
| note | 任意 |

- upsertキー: `(client, source_type, source_key, quarter_label)`。同一四半期の再実行は上書き。
- **MIN_SAMPLES / MIN_CHANGE のフィルタは適用しない**（提案可否と証拠保存は別事）。n ≥ 1 を全件残す。

### 4.2 既存 `POOL_PRIOR`（pool book = 中央、各book = fan-out先）
スキーマ変更なし。`pool_scope=reliability:{type}` / `param_key=reliability_r` /
`pooled_value=rHat_pool` / `precision=k_pool` / `n_clients` を書く。

### 4.3 新シート `POOL_REGISTRY`（pool bookのみ）
| 列 | 内容 |
|---|---|
| client | メーカー名 |
| spreadsheet_id | そのbookのID |
| enabled | 1/0 |
| note | 任意 |

---

## 5. アルゴリズム

### 5.1 3c-3c-1（永続化・各book）
1. C-1 (`runQuarterlyReview`) が `data.ready`（実績3か月）のとき実行。
2. `computeReliabilityHitStats_(data)` を呼ぶ（現 `generateReliabilityProposals_` の `grouped` を関数化）。
   - 戻り値: `[{source_type, source_key, n, hit, hit_rate}, ...]`（フィルタ前・全件）。
3. `writeReliabilityEvidence_(client, quarterLabel, quarterEndMonth, stats, runId)` で upsert。
4. 提案生成（`generateReliabilityProposals_`）は同じ stats を使うよう差し替え（再計算しない）。
5. **永続化は提案有無に関係なく走る。**

### 5.2 3c-3c-2（集約・fan-out・pool book／管理者手動）
1. `aggregatePoolPrior_()`:
   - `POOL_REGISTRY` の enabled book を列挙。
   - 各bookの `RELIABILITY_EVIDENCE` を `openById` で読む。
   - source_type別に `Σhit, Σn`、`n_clients`（Σn>0 のbook数）を集計。
   - `rHat_pool = clamp(2·Σhit/Σn, R_MIN, R_MAX)`、`k_pool = min(n_clients, SHRINKAGE_K)`。
   - `n_clients < MIN_CLIENTS` or `Σn < MIN_SAMPLES` の scope はスキップ。
   - pool book の `POOL_PRIOR` を upsert。
2. `distributePoolPrior_()`:
   - pool book の `POOL_PRIOR` を読み、`POOL_REGISTRY` の各bookの `POOL_PRIOR` に upsert（fan-out）。
3. 各bookの次回 C-1 が `readPoolPrior_` で新しい rPool/precision を拾い、収縮に反映。

---

## 6. CONFIG 追加キー
| キー | 既定 | 用途 |
|---|---|---|
| RELIABILITY_MIN_CLIENTS | 2 | pool書込みに必要な最小client数 |

※ `RELIABILITY_MIN_SAMPLES`(2) / `RELIABILITY_SHRINKAGE_K`(4) / `R_MIN`(0) / `R_MAX`(1.5) は既存を流用。

---

## 7. 検証順序（MANUAL_TASKS と対応）
1. 3c-3c-1 を1冊に適用 → `RELIABILITY_EVIDENCE` が出来る。
2. 実績3か月で C-1 → **安定ソースを含む全ソース**の hit/n が残ることを確認（提案ゼロでも残る）。
3. A-9 の計画値が**変わらない**ことを確認（永続化は予測不変）。
4. 2冊目以降でも evidence が book別に貯まることを確認。
5. 複数bookが揃ってから 3c-3c-2（集約・fan-out）に着手。

---

## 8. 未確定（3c-3c-2 着手前にTechiが確定する1点）
- transport前提: 全bookが1つのサービスアカウント／実行者から `openById` で読み書き可能か。
  - 可能 → 本書の fan-out 案で 3c-3c-2 を実装。
  - 不可 → 実行時 openById 読み or 手動コピーに切替（その場合 3c-3c-2 の設計を差し替え）。
