# Forecast vNext と旧画面（2026-10-05 にアーカイブ）

売上予測と予算策定は新しいアプリ Trends2Targets（`app/`）だけで行う（村井さん、2026-10-05）。
Forecast vNext は 2026-10-01 から凍結していたので、コード・テスト・手順書をここへ移した（消していない。元の場所からそのまま移した）。

- `0_VNext_Naming.js`・`VNext_*.js`・`VNext_*Sidebar.html` … 管理ハブ・年度ブック・サイドバー
- `client_runtime/`・`portal_runtime/`・`scripts/` … クライアント用・入口用のランタイムと道具
- `Forecast_WebAppUI.html` … 旧 Web アプリの画面（空模様の判定の元。新アプリは `app/src/UI.html`）
- `tests/` … 上のもののテスト（vnext-*）、旧画面の空模様のテスト（forecast-weather）、旧 Web アプリの操作の記録のテスト（forecast-webaudit）
- `README_root.md`（元のルートの README）・`README_VNEXT_JA.md`・`AI_HANDOFF.md`・`DESIGN_admin_sidebar_simplification_JA.md`・`DESIGN_portal_entry_ux_audit_JA.md`
- `POOL_SETUP_v1.9.md`・`MANUAL_TASKS_*.md` … 管理ハブと旧ブックで行う手作業の手順

ここにあるテストは、元の場所（ルート）を前提にしたパスで書かれているので、このままでは動かない。
すでに公開されている vNext の Apps Script プロジェクト（ある場合）は、ここへ移しても変わらない（Drive・Apps Script 側には何もしていない）。
