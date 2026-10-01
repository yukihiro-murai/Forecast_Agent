# キャラクター素材（売上予測 Webアプリ）

案内役「よみ」（気象観測ロボ）と、予測の当たり具合を表す天気キャラ 8 種の素材。v8（2026-10-01 確定）。

- **正本は `src/cast.py`**（図形 = body と、顔 = キャラクター図鑑の部品の指定 fs）。SVG・PNG は手で直さず、`src/build.py` で作り直す。
- 図鑑（`~/Documents/GAS/shared/character-library/gen.py`）にも同じ body / fs で登録済み（シリーズ「天気予報 (Forecast_Agent)」）。図鑑が作る SVG とこのフォルダの `svg/` は一致する。直すときは両方を揃える。
- このフォルダは GAS に上げない（`.claspignore` の `assets/**`）。

## 書き出し

```bash
cd assets/characters/src
python3 build.py
```

Playwright 同梱の chrome-headless-shell（`~/Library/Caches/ms-playwright/`）で PNG にする。確認用 HTML は一時ディレクトリに作り、ここには残さない。
`build.py` は `Forecast_WebApp.js` の「タブのアイコン」区間（`FORECAST_FAVICON_URL`）と `Forecast_WebAppUI.html` の「キャラの絵」区間（`CHAR_SVG` / `YOMI_POSE`）も書き換える。

## キャラ一覧

| id | 名前 | 意味 | SVG 正本 | 透過 PNG |
|---|---|---|---|---|
| yomi | よみ | 案内役（天気を伝える側） | `svg/yomi.svg` | `png/yomi_160.png` / `_512.png` |
| kaisei | 快晴（かいせい） | 予測がよく当たっている | `svg/kaisei.svg` | `png/kaisei_*.png` |
| harenochi | 晴れのち曇り（はれのちくもり） | 当たっているが崩れ始めの兆し | `svg/harenochi.svg` | `png/harenochi_*.png` |
| kumori | 曇り（くもり） | ずれが出ている | `svg/kumori.svg` | `png/kumori_*.png` |
| ame | 雨（あめ） | 大きく外れている | `svg/ame.svg` | `png/ame_*.png` |
| sekka | 雪（せっか） | 数字が冷え込んでいる | `svg/sekka.svg` | `png/sekka_*.png` |
| taifuu | 台風（たいふう） | ずれが大きく荒れている | `svg/taifuu.svg` | `png/taifuu_*.png` |
| tenpen | 天変地異（てんぺんちい） | ごくまれな桁違いの外れ | `svg/tenpen.svg` | `png/tenpen_*.png` |
| mikakunin | 未確認・霧（みかくにん） | データ不足でまだ判定できない（悪い予測とは別系統） | `svg/mikakunin.svg` | `png/mikakunin_*.png` |

意味の詳しい条件（どの指標で、どの順で決めるか）は `CHARACTER_DESIGN_STATUS.md` の「表示条件」。2026-10-01 に画面へ実装した。
雪の id が `sekka` なのは、図鑑に `yuki`（ゆき）が既にいるため。

## よみ 5 ポーズ

| ファイル（SVG 正本） | 透過 PNG | ポーズ | 用途 |
|---|---|---|---|
| `yomi/svg/yomi_guide.svg` | `yomi/png/yomi_guide_160.png` / `_512.png` | 通常案内（片手を上げる） | ホームの「次の一手」など、ふだんの案内 |
| `yomi/svg/yomi_observe.svg` | `yomi/png/yomi_observe_*.png` | 観測中（虫めがね＋受信中の電波） | 取り込み・計算・読み込みの待ち時間 |
| `yomi/svg/yomi_discover.svg` | `yomi/png/yomi_discover_*.png` | 変化発見（両手を上げて「!」、風速計が傾く） | 前回からの変化・注意点を知らせるとき |
| `yomi/svg/yomi_explain.svg` | `yomi/png/yomi_explain_*.png` | 説明中（指し棒でグラフ札を指す） | 数値・グラフ・検証結果の読み方を説明するとき |
| `yomi/svg/yomi_done.svg` | `yomi/png/yomi_done_*.png` | 確認完了（にっこり＋チェック札） | 保存・実行・確認が終わったとき |

`yomi_guide` は `svg/yomi.svg`（図鑑の既定の絵）と同じ。表情だけを変えたいときは図鑑の表情違い 6 種（`characterSvg('yomi', 'happy')` など）を使う。

## サイズと用途

| 種類 | サイズ | 容量の目安 | 用途 |
|---|---|---|---|
| SVG 正本（`svg/`・`yomi/svg/`） | 80×80 viewBox | 0.6〜1.8 KB | 画面への埋め込み（どの大きさでもくっきり）。図鑑と同じ絵 |
| 透過 PNG `*_160.png` | 160×160 px | 3〜8 KB | 画面の 72〜80px 表示の 2 倍解像度。SVG を使えない場所（メール・外部ツール） |
| 透過 PNG `*_512.png` | 512×512 px | 10〜27 KB | 資料・スライド・印刷 |
| `favicon/yomi_favicon.svg` | 32×32 viewBox | 0.8 KB | タブのアイコンの原図（頭と顔を枠いっぱいに描いた専用の絵。16px でつぶれない） |
| `favicon/yomi_favicon_64.png` | 64×64 px | 2.7 KB | タブのアイコン。`Forecast_WebApp.js` の `FORECAST_FAVICON_URL` に data URI で埋め込み（末尾 `#favicon.png`） |
| `Characters.html` | — | 約 100 KB | 図鑑からの抜粋（`characters.config.json` → `sync.mjs`）。9 体 + 表情違い + 補助関数 |
| `review/*.png` | 2300px 幅 | 200〜330 KB | 確認用の一覧（`chars_v8.png` 確定版・`yomi_poses.png` ポーズとタブ・`chars_v8_r1.png` 第1回の確認・`placement_mock.png` 画面配置と表示条件の案・`placement_impl.png` 実装後の実際の画面。配置案と画面写真は build.py では作らない） |

## 画面での使われ方（2026-10-01）

タブのアイコン＝よみ、ホーム「次の一手」＝よみ（通常案内／空模様が変わった月は変化発見）、検証画面の上部＝最新月の天気キャラ、
データが無いときの空の表示＝未確認・霧、処理中＝よみ（観測中）、完了の通知＝よみ（確認完了）。
表示条件と実装場所は `CHARACTER_DESIGN_STATUS.md` の「表示条件」「画面への配置」。絵は `Forecast_WebAppUI.html` の「キャラの絵」区間（`build.py` が書き換える）。

## 画面で使うときの注意

- **GAS のファイルを増やさない。** 管理ハブの複製は、プロジェクトのファイルが `VNEXT_ADMIN_RUNTIME_FILE_TYPES_`（`VNext_ClientRuntimeProvisioning.js`）の 21 ファイルと完全に一致することを検査する。`Characters.html` などを別ファイルとして GAS に足すと複製が止まるため、使うときは既存の HTML（`Forecast_WebAppUI.html`）の中に埋め込むか、許可リストの変更とあわせて行う。タブのアイコンも同じ理由で `Forecast_WebApp.js` の中に置いている。
- タブのアイコンは `HtmlOutput.setFaviconUrl` で渡す（ページ内の `<link rel="icon">` は Apps Script に無視される）。URL の末尾は `#favicon.png`（拡張子が無い URL は黙って捨てられる）。
- アニメーションは控えめに（既存の `.charbox` の流儀: ふわっと数回・クリックで再生）。

## 旧版

`archive/v7/` に v7（2026-09-30 時点）の描画ページ・一覧 PNG・SVG を残している。
