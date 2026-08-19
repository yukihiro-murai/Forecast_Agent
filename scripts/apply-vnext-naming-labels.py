#!/usr/bin/env python3
"""Apply user-facing Japanese label updates (display only)."""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# Longer phrases first. Sheet name 2_予測と計画 is intentionally unchanged.
REPLACEMENTS: list[tuple[str, str]] = [
    ("Forecast Owner", "予算策定担当"),
    ("正式計画", "正式予算"),
    ("計画案", "予算案"),
    ("管理者の確認", "管理ハブの確認"),
    ("管理者attestation", "管理ハブ担当者 attestation"),
    ("管理者として", "管理ハブ担当者として"),
    ("管理者監査ログ", "管理ハブ監査ログ"),
    ("管理者向け詳細", "管理ハブ担当者向け詳細"),
    ("管理者向け", "管理ハブ担当者向け"),
    ("管理者設定", "管理ハブ設定"),
    ("管理者メール", "管理ハブ担当者メール"),
    ("管理者から", "管理ハブから"),
    ("管理者が", "管理ハブが"),
    ("管理者用", "管理ハブ用"),
    ("管理者へ", "管理ハブ担当者へ"),
    ("管理者の", "管理ハブの"),
    ("社員からの依頼", "申請入口からの依頼"),
    ("社員からの作成依頼", "申請入口からの作成依頼"),
    ("社員テスト", "受入試験"),
    ("社員book", "クライアント年度ブック"),
    ("社員用runtime", "クライアント年度ブック用 runtime"),
    ("従業員テスト", "受入試験"),
    ("従業員画面", "クライアント年度ブック画面"),
    ("Admin runtime", "管理ハブ runtime"),
    ("Admin専用", "管理ハブ専用"),
    ("生成Hubの", "生成した管理ハブの"),
    ("生成Hub", "生成した管理ハブ"),
    ("Hubの", "管理ハブの"),
    ("Hub ID", "管理ハブ Spreadsheet ID"),
    ("Admin Hubの", "管理ハブの"),
    ("Admin Hub", "管理ハブ"),
    ("Book ID", "クライアント年度ブック ID"),
    ("この book", "このクライアント年度ブック"),
    ("このbook", "このクライアント年度ブック"),
    ("このブック", "このクライアント年度ブック"),
    ("book・履歴", "Spreadsheet・履歴"),
    ("計画だけ修正", "予算案だけ修正"),
    ("既存の計画", "既存のクライアント年度ブック"),
    ("同じ年度の計画または作成依頼", "同じ年度のクライアント年度ブックまたは作成依頼"),
    ("予算計画の承認", "年度予算の承認"),
    ("受付・計画状況", "受付・申請状況"),
    ("提出した計画を", "提出した予算案を"),
    ("計画の保存", "予算の保存"),
    ("予測と計画を開け", "予測シートを開け"),
    ("Forecast vNext Admin Hub", "年度予算策定 管理ハブ"),
    ("管理者", "管理ハブ担当者"),
]

TARGETS = [
    ROOT / "VNext_UX.js",
    ROOT / "client_runtime/src/VNext_UX.js",
    ROOT / "VNext_HelpSidebar.html",
    ROOT / "client_runtime/src/VNext_HelpSidebar.html",
    ROOT / "VNext_PlanSidebar.html",
    ROOT / "client_runtime/src/VNext_PlanSidebar.html",
    ROOT / "VNext_InputSidebar.html",
    ROOT / "client_runtime/src/VNext_InputSidebar.html",
    ROOT / "VNext_AdminSidebar.html",
    ROOT / "portal_runtime/src/Portal_Core.js",
    ROOT / "portal_runtime/src/Portal_UX.js",
    ROOT / "portal_runtime/src/Portal_CreateSidebar.html",
    ROOT / "portal_runtime/src/Portal_Entry.html",
    ROOT / "VNext_Admin.js",
]


def apply(text: str) -> str:
    for old, new in REPLACEMENTS:
        text = text.replace(old, new)
    return text


def main() -> int:
    changed = 0
    for path in TARGETS:
        if not path.exists():
            print(f"skip missing {path}", file=sys.stderr)
            continue
        original = path.read_text(encoding="utf-8")
        updated = apply(original)
        if updated != original:
            path.write_text(updated, encoding="utf-8")
            changed += 1
            print(f"updated {path.relative_to(ROOT)}")
    print(f"done ({changed} files)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
