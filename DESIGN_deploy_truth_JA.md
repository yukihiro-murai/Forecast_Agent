# 開発反映の「真実」契約（Deploy Truth）

最終更新: 2026-08-20

## 問題

`clasp push` だけでは **Web入口は変わらない**。  
Hub の簡易確認が「Hub 内の期待値 ≒ いまの Portal」だけを見ると、**中央が新しいのに両方とも古い**場合に ✓ が出る（偽陽性）。

## 契約（これだけ覚える）

1. **Cursor:** 編集 → test → `clasp push`
2. **Hub:** 「開発反映を実行」（1/4 Hub → 2/4 Client → 3/4 Portal → 4/4 確認）
3. **確認:** Portal 版が中央期待と一致 + Web入口をハード再読み込み

「いま反映済みか確認」は **監査**。反映そのものではない。

## 自動検査

- Hub は中央 clasp の `VN_ADMIN_PORTAL_RUNTIME_VERSION` / `VN_ADMIN_RUNTIME_BUILD_STAMP` を読む
- Hub コードが中央より古い、または live Portal が中央期待と違う → **未反映（✗）**
- ✓ は「中央・Hub・Portal が同じ版」のときだけ

## 運用メモ

- clasp のあとに確認ボタンだけ押しても Web は変わらない
- 反映が止まったら進捗（1/4〜）を見て、止まった段から再実行
