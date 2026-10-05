# Forecast Agent

売上予測と予算策定は、新しいアプリ **Trends2Targets**（[`app/`](./app/README.md)）だけで行います。
構成・公開・テストの手順は [`app/README.md`](./app/README.md)、これからの計画は [`app/DESIGN_evolution_JA.md`](./app/DESIGN_evolution_JA.md) にあります。

## このリポジトリにあるもの

| 場所 | 中身 |
|---|---|
| `app/` | Trends2Targets（Apps Script の Web アプリ。`app/.clasp.json`） |
| `Forecast_Agent.js`・`Forecast_WebApp.js` | 計算の元。`app/src/LegacyEngine.js` にそのまま包んで使う（変えたら `node app/tools/build-engine.mjs`） |
| `tests/` | 計算の元のテスト（`node tests/forecast-autolearn.test.mjs` など） |
| `assets/`・`characters.config.json` | 案内役のキャラの素材（`app/src/Assets.js` の元） |
| `DESIGN_RECOMMENDATION_JA.md` ほか `DESIGN_*.md` | 計算の設計（三角測量・学習・横断プール）と、データ基盤の設計（`DESIGN_data_platform_JA.md`） |
| `archive/` | 使わなくなったものの写し（消さずに残す） |

## 使わないもの

- 旧アプリ（「クライアント別売上予測」ブックのメニュー・旧 Web アプリ）と旧スプレッドシートは使わず、リリースもしていません（2026-10-04）。
  新アプリから届かない旧来の関数は `archive/legacy-2026-10-04/` へ移しました。
- Forecast vNext（管理ハブ・年度ブック・申請入口・クライアント用と入口用のランタイム）は 2026-10-01 から凍結していて、使いません。
  コード・テスト・手順書は `archive/vnext-2026-10-05/` へ移しました。
- ルートの `.clasp.json`・`.claspignore`・`appsscript.json` は旧アプリの Apps Script プロジェクトのものです。ルートからは `clasp push` しないでください（新アプリは `app/` で push します）。
