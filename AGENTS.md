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
5. `.clasp.json` があるプロジェクトでは、git push 後に `clasp status` で対象ファイルを確認してから `clasp push` で GAS に反映する。`.clasp.json` が無いプロジェクトでは、GAS エディタへの手動反映が別途必要なことをユーザーに伝える。
6. 社員ポータルの Web 入口（`/exec`）まで反映する場合は、Hub ボタンの代わりに `node scripts/publish_employee_portal.mjs` を使う。未ログインなら先に `node scripts/gas_agent_login.mjs`（計画の承認・差戻し・公式化は含まない）。

### 禁止事項

- `git push --force` / `--force-with-lease` (ユーザーの明示指示がある場合のみ例外)
- fetch / pull を省略して作業を始めること
- push を完了していないのに「作業完了」と報告すること
- コンフリクトの独断解決 (解決方針を先に報告し、合意してから解決する)

## 共通開発規約

ワークスペースルート (`~/Documents/GAS/`) の `AGENTS.md`・`CLAUDE.md`・`GIT_SYNC_RULES.md` に、GAS 共通のコーディング規約 (V8 / プレーンJS / Logger.log / Script Properties 等) と同期ルールの詳細版がある。読める環境では必ず参照すること。このリポジトリ固有の README・規約ファイルがあればそちらも必ず読むこと。
