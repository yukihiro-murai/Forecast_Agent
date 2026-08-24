# Forecast vNext — レビュー／監査用ブリーフィング（他AIエージェント向け）

**最終更新:** 2026-08-24 JST  
**対象ブランチ:** `cursor/7dfecef5`（`origin/cursor/7dfecef5` と同期想定）  
**作成目的:** チャット履歴なし・端末固有メモリなしで、**このリポジトリだけ**を読んで設計レビュー・監査・次作業ができる状態にする。

> **読み順（必須）**  
> 1. 本ファイル（全体コンテキスト）  
> 2. [`AGENTS.md`](./AGENTS.md)（Git／反映の運用契約）  
> 3. [`DESIGN_adaptive_learning_JA.md`](./DESIGN_adaptive_learning_JA.md)（本丸の設計思想）  
> 4. [`DESIGN_deploy_truth_JA.md`](./DESIGN_deploy_truth_JA.md)（反映の真実契約）  
> 5. [`README_VNEXT_JA.md`](./README_VNEXT_JA.md)（製品概要）  
> 6. 必要に応じて [`AI_HANDOFF.md`](./AI_HANDOFF.md)（古いライブメモ。**版番号は古くなりうる**）

---

## 1. 一言で何か

Forecast vNext は、製薬営業まわりの **クライアント×会計年度（FY）売上**について、一点予測ではなく **条件付き分布（P10–P90 等）** と根拠レイヤを出し、**正式予算は学習に戻さない**運用を持つ Google Apps Script / Spreadsheet システムである。

**本丸（2026-08-20 合意）は「数字当て」ではなく、観測→仮説→検証→次予測への反映が回り続ける適応学習プロセスである。**

---

## 2. ステークホルダーと使い方の前提

| 役割 | 主な入口 |
|---|---|
| 現場社員 | Web入口（Portal `/exec`）の 01〜03 Bot／吹き出し導線 |
| 管理ハブ担当 | Hub Spreadsheet 右案内（サイドバー）＋メニュー「年度予算策定」 |
| 開発（Cursor 等） | リポジトリ編集 → `npm run reflect` / `clasp push` → Hub で反映操作 |

**ユーザー（村上）の好み:** Cursor と対話しながら小さく直す。毎回の反映で迷う・止まる・偽の完了表示が続くと強くストレスになる。エージェントは **画面に無い UI 名を案内しない**こと。

**デバイス:** Mac mini / MacBook Pro の2台並行。GitHub `origin` が正。詳細は `AGENTS.md`。

---

## 3. 設計思想（監査で崩してはいけない契約）

### 3.1 予測・学習

- 未来予測の真の目的は直接最適化できない → **代理目的**の順:  
  **①区間校正 → ②層別バイアス → ③制約付き点誤差 → ④情報ギャップ**
- **System track**（システム推奨 vs 実績）= 学習に使う  
- **Budget track**（正式予算 vs 実績）= 監査用。**学習に戻さない**。年度途中の会社予算補正は原則しない
- 公式 vintage の後書き禁止。immutable FORECAST_RUN / release 契約を守る
- 人ゲート: 閾値超過・Model Release・重要ポリシーのみ
- ベイジアンは段階計画 Phase A（証拠配線）→ B（状態事後）→ C（候補重み）→ D（領域）。重い専用学習基盤は必須としない

詳細: [`DESIGN_adaptive_learning_JA.md`](./DESIGN_adaptive_learning_JA.md)  
Obsidian ADR: `45_MESO_decisions/Forecast vNext の本丸は数字当てではなく適応学習プロセスである.md`

### 3.2 UI（Web入口）

- **大きく変えない:** 01〜03 の Bot／キャラ付き吹き出し、「対話している」構図  
- **変えてよい:** 配色・文言・ボタン・機能増減  
- リード文から「キャラ付き」表現は除去済み（Portal 1.7.13）

### 3.3 物理／権限分離

- 共有ドライブ「年度予算策定」  
- 01 管理ハブ / 02 申請入口 / 03 クライアント年度ブック / 04 テンプレート  
- Client ブックに Admin・ZAC・Vertex 秘密を置かない  
- 中央 clasp ソース → Hub runtime コピー → Portal / Client bundle 配備

詳細: [`README_VNEXT_JA.md`](./README_VNEXT_JA.md)

### 3.4 反映（Deploy Truth）

`clasp push` だけでは **Web入口は変わらない**。  
経路: **中央 clasp → Hub runtime 同期 →（先に）Portal →（必要なときだけ）Client Template/Model → 確認**

偽陽性の歴史: Hub 内期待値 ≈ live Portal（両方古い）で ✓ が出た。  
対策: 中央マーカー（Portal 版 / `VN_ADMIN_RUNTIME_BUILD_STAMP`）との照合。

**致命的な実装落とし穴:** Hub 同期後も **開いたままの案内サイドバーは古い JS のまま**次ステップを続ける。旧 UI は「2/4 Client release…」で Template 再作成に入り停止する。  
対策（コード済み）: Hub 内容が変わったら `requiresSidebarReload` で停止して開き直し要求。Client 重い再作成は `allowHeavyRelease=true` のときだけ。

詳細: [`DESIGN_deploy_truth_JA.md`](./DESIGN_deploy_truth_JA.md)  
Cursor ルール: [`.cursor/rules/forecast-reflect.mdc`](./.cursor/rules/forecast-reflect.mdc)

---

## 4. 現在のリポジトリ状態（コード側・2026-08-24）

| 項目 | 値 |
|---|---|
| Remote | `git@github.com:yukihiro-murai/Forecast_Agent.git` |
| 作業ブランチ | `cursor/7dfecef5` |
| 直近の意図コミット群 | 適応学習 Phase A、反映 UX 連続修正、`npm run reflect`、2/4 停止対策 |
| Client runtime（コード） | `vnext-client-1.8.1` |
| Portal runtime（コード） | `vnext-portal-1.7.13` |
| Build stamp | `20260824-stop-after-hub-sync` |
| 中央 clasp Script ID | `1CkHthmMuU5r66ZpWJLw4bXrhNhDzcHCjBb2o1sFdIR1I0p1wNAao_erV` |
| Admin Hub Spreadsheet | `1baEZe6xYQ9KWyMMBk7kzH50v4dTtBPk9kWHK3qT7ID8` |
| Hub URL | https://docs.google.com/spreadsheets/d/1baEZe6xYQ9KWyMMBk7kzH50v4dTtBPk9kWHK3qT7ID8/edit |
| Admin Hub Script ID | `1ANVtOBzPo90DveLcmTYokpEycr4iOphMaKZrhyjvzUizEQUjGqHRScxw` |
| 社員 Portal Spreadsheet（参考・旧手渡し） | `16uiBgEF6Pz6N3UUpPU1DPJ6axelYRDJ3WuXFJUWy1dU` |

**注意:** Sheets / GAS のライブ ACTIVE pair・Portal ピン版は Git と一致しないことがある。監査時は Hub 案内の版表示と `/exec` の実表示を照合すること。`AI_HANDOFF.md` の版番号は古い。

---

## 5. いまユーザーが直面している状況（2026-08-24）

1. **対話開発＋反映のストレス**が主訴。場所のないボタン案内・画面に無いメニュー案内・反映ハングが連続した。  
2. 何度も **「2/4 Client release（Template/Model）を反映しています…」** で停止。これは **現行コードのメッセージではなく、開いたままの古い案内 JS** の兆候。  
3. 正しいユーザー操作（正本）:  
   - Hub: **「年度予算策定」→「案内を開く」→ 右パネル上部の緑ボタン**（「反映する」または「開発反映を実行」）  
   - Hub 同期直後に「開き直して」と出たら従う  
   - Web入口 **Cmd+Shift+R**  
   - メニュー「Web入口を最新版にする」は **Hub 最新同期＋シート再オープン後にしか出ない**。見えないときは無視。  
4. ライブで Portal がまだ `1.7.12` のまま・Hub 期待が `1.7.13`、というズレが観測されたことがある。Web 文言（「キャラ付き」削除）がハードリロード後も古い場合は、Portal 反映未完了。  
5. 適応学習 Phase A は **コード配線済み**（Hub シート LEARNING_*、サイドバーカード、区間ミス時 logSigma widen、テスト）。本番データでの「次 run が変わる」実測は未完了の可能性大。クライアント年度ブック数が 0 の Pilot 状態もありうる。

---

## 6. 主要コンポーネント地図

| パス | 役割 |
|---|---|
| `VNext_Admin.js` / `VNext_AdminSidebar.html` | Hub 業務・反映・学習ダッシュボード |
| `VNext_Engine.js` | 分布シミュレーション・学習証拠の適用 |
| `VNext_PortalRuntimeBundle.js` / `portal_runtime/` | 社員 Web入口・申請 |
| `VNext_ClientRuntimeBundle.js` / `client_runtime/` | クライアント年度ブック UI |
| `scripts/reflect.mjs` | Cursor 用最短 clasp＋可能なら Hub 同期 |
| `tests/vnext-*.mjs` | 契約／統合テスト |
| `DESIGN_*.md` | 設計正本 |

---

## 7. 適応学習 Phase A（実装済み範囲）

- Hub: `LEARNING_POLICY_JSON`、`LEARNING_OBSERVATION`、`LEARNING_EVIDENCE`  
- 振り返り評価から証拠 append → 次 forecast に `learningEvidence`  
- Engine: 前回区間ミス時に `logSigma` を widen（Budget は経路に入れない）  
- テスト: `tests/vnext-learning-phase-a.test.mjs`  
- **未着手（意図的）:** Phase B の層別自動状態更新、横断プール本実装、領域タグ実装

---

## 8. レビュー／監査で見てほしい論点（提案）

### 優先高（運用安定・ユーザー苦痛）

1. **反映パイプラインの信頼性:** Hub 同期とサイドバー寿命、Portal-first、Client 重処理のオプトインが十分か。さらに自動化（API executable）の要否。  
2. **エージェント案内契約:** 画面に無い UI を案内しないルールがコード／ドキュメント／テストで守れているか。  
3. **偽の「反映済み」:** 中央マーカー照合の抜け穴が残っていないか。  

### 優先中（設計整合）

4. 適応学習 Phase A のスキーマ・二重トラック分離・人間ゲートが設計書と一致しているか。  
5. Web入口 UI 制約（キャラ導線維持）と文言変更の境界。  
6. Client / Portal / Admin の権限・scope 分離が破られていないか。  

### 優先低〜次回

7. Phase B 以降のベイジアン状態更新の設計穴。  
8. `AI_HANDOFF.md` の陳腐化（ライブ版の単一正本化）。  

### 監査でやってほしくないこと

- ユーザー未承認の承認／差戻し／公式化／force push  
- Web入口のキャラ対話構図の全面刷新  
- Budget を学習に戻す変更  
- 「一点誤差だけ」を本丸にした最適化提案  

---

## 9. 検証コマンド（ローカル）

```bash
git fetch origin && git status -sb
node tests/vnext-integration.test.mjs
node tests/vnext-learning-phase-a.test.mjs
# Portal/Client を触った場合
cd portal_runtime && npm test && cd ..
cd client_runtime && npm test && cd ..
npm run reflect   # または clasp push
```

ライブ確認は Hub 案内の版表示と Web入口ハードリロードが必要（Git だけでは不可）。

---

## 10. 関連ドキュメント索引

| 文書 | 内容 |
|---|---|
| 本ファイル | 監査用の単一入口 |
| `AGENTS.md` | マルチマシン Git・反映案内契約 |
| `DESIGN_adaptive_learning_JA.md` | 適応学習の本丸 |
| `DESIGN_deploy_truth_JA.md` | 反映の真実 |
| `DESIGN_3c3c_JA.md` / `DESIGN_fix_learning_loop_JA.md` | 周辺設計 |
| `README_VNEXT_JA.md` | 製品説明 |
| `AI_HANDOFF.md` | 旧ライブ引継ぎ（要照合） |
| `.cursor/rules/forecast-reflect.mdc` | Cursor 常時ルール |

---

## 11. レビュー依頼文テンプレ（他エージェントへ貼る用）

```
リポジトリ Forecast_Agent のブランチ cursor/7dfecef5 をレビューしてください。
最初に REVIEW_AUDIT_BRIEF_JA.md を読み、続けて AGENTS.md / DESIGN_adaptive_learning_JA.md /
DESIGN_deploy_truth_JA.md を読んでください。
特に (1) 反映パイプラインの安定性 (2) 適応学習 Phase A の設計整合 (3) ユーザー案内の誤誘導リスク
を優先し、問題を優先度付きで指摘してください。破壊的操作や Budget を学習に戻す提案はしないでください。
```
