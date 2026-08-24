# AGENTS.md — マルチマシン運用ルール (全AIエージェント共通)

このリポジトリで作業する全ての AI コーディングエージェント (Claude Code / Codex / Grok Build / Hermes Agent / Cursor 等) は、コードに触れる前にまずこのファイルを読むこと。

## マルチマシン Git 同期プロトコル (必須・スキップ禁止)

このリポジトリは Mac mini と MacBook Pro の2台で並行運用されている。
**GitHub (origin) が唯一の正であり、ローカルは常に古い可能性がある。** ローカルの状態だけを見て作業を始めてはならない。

### 作業開始時 — どんな小さな作業でも必ず実行

1. `git branch --show-current` — 今いるブランチが依頼内容と合っているか確認する。
2. `git status` — 未コミット変更・untracked ファイルの有無を確認する。
3. `git fetch origin` → `git status` で ahead/behind を確認し、behind なら `git pull --ff-only`。
4. 別マシンで作られたブランチを引き継ぐ場合は fetch 後に `git switch <branch>`。
5. fast-forward できない (履歴が分岐している)、または未コミット変更とリモート更新が衝突しそうな場合は、**勝手に merge / rebase / stash / checkout せず**、状況を報告してユーザーの指示を待つ。

### 作業終了時・中断時 — 必ず実行

1. 変更を論理単位で commit する。中断時は WIP でも commit する — **push していない作業は、もう一台のマシンからは存在しないのと同じ**。
2. `git pull --ff-only` (push 直前の最終同期)。
3. `git push origin <branch>` (新規ブランチは `git push -u origin <branch>`)。
4. push が拒否されても force push はしない。fetch して状況を報告する。
5. `.clasp.json` があるプロジェクトでは、git push 後に `npm run reflect`（失敗時は `clasp push`）で中央 GAS に載せ、ユーザーには **「Cursor反映 → 反映する」1操作**だけ案内する。ネストした「システムの手入れ」案内はしない。

### 禁止事項

- `git push --force` / `--force-with-lease` (ユーザーの明示指示がある場合のみ例外)
- fetch / pull を省略して作業を始めること
- push を完了していないのに「作業完了」と報告すること
- コンフリクトの独断解決 (解決方針を先に報告し、合意してから解決する)

## Cursor 対話開発（反映のストレス低減）

ユーザーは対話で小さく直しながら進める。毎回の反映で迷わせない。

- エージェント: `npm run reflect` までやる（clasp + 可能なら Hub 同期）
- ユーザー: 管理ハブの **「反映する」** → Web入口 **Cmd+Shift+R** だけ
- 詳細契約: `DESIGN_deploy_truth_JA.md` / `.cursor/rules/forecast-reflect.mdc`

## 共通開発規約

ワークスペースルート (`~/Documents/GAS/`) の `AGENTS.md`・`CLAUDE.md`・`GIT_SYNC_RULES.md` に、GAS 共通のコーディング規約 (V8 / プレーンJS / Logger.log / Script Properties 等) と同期ルールの詳細版がある。読める環境では必ず参照すること。このリポジトリ固有の README・規約ファイルがあればそちらも必ず読むこと。
