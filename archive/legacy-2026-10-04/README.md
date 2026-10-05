# 旧来のコードの元（2026-10-04）

旧アプリ（「クライアント別売上予測」ブックのメニュー・旧 Web アプリ）と旧スプレッドシートは使わず、リリースもしていない（村井さん、2026-10-04）。
新アプリ Trends2Targets（`app/`）は `Forecast_Agent.js`・`Forecast_WebApp.js` を計算の元として包む（`app/src/LegacyEngine.js`）。

そのため、新アプリから届かない 45 の関数（ブックのメニュー `onOpen`・入力のダイアログ・メニュー版の取り込み・管理ハブの POOL 集約・旧 Web アプリの `doGet` など、約 53KB）を除いた。
ここにあるのは、除く前のそのままの写し:

- `Forecast_Agent.js` / `Forecast_WebApp.js` … 除く前の旧来のコード
- `forecast-favicon.test.mjs` … 旧 Web アプリの `doGet` のテスト（新アプリでは動かない）

どの関数を除いたかは `node app/tools/engine-usage.mjs` の考え方（新アプリが呼ぶ関数から名前でたどり、届かないもの）で決めた。
新アプリのテスト（`app/tests/app-engine.test.mjs` の 7）が、届かない関数が残っていないことを確かめる。
`Forecast_WebAppUI.html` は 2026-10-05 に vNext と一緒に `archive/vnext-2026-10-05/` へ移した。
