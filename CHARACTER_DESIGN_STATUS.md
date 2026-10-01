# キャラクターデザイン作業 — 引き継ぎ記録（2026-10-01 更新）

Forecast_Agent 売上予測Webアプリの案内役「よみ」と天気キャラ 8 種の作業記録。2026-09-30 に VM 上の別エージェントから引き継ぎ、2026-10-01 に Claude Code（Mac mini）が v8 で確定させた。

## 依頼の全体像（ユーザー指定・要約）

- 案内役キャラ「よみ」（ひらがな表記）：天気全体を観測し・読み解き・利用者に伝えるキャスター役。天気そのものを表すキャラとは見た目と役割を分ける。「気象観測ロボット」案をユーザーが選択済み。
- 天気キャラ：快晴・晴れのち曇り・曇り・雨・雪・台風・天変地異 ＋ 未確認・霧（データ不足用）。
- 制約：雲をよみ本体・帽子の中心モチーフにしない／カエル不使用／親しみやすくシンプル／小さな表示でも識別可能／フラット・単純図形・テカリ禁止・青系統。
- 計算・保存・権限・予測判定ロジックは変更禁止。本番デプロイはユーザーの最終確認後のみ。キャラクター図鑑への反映が必須。

## 現在の状態（2026-10-01）

| 項目 | 状態 |
|---|---|
| デザイン（9 体） | **v8 で確定**（ユーザー OK 2026-10-01）。正本 `assets/characters/src/cast.py` |
| よみ 5 ポーズ | 作成済み（通常案内・観測中・変化発見・説明中・確認完了）。SVG 正本＋透過 PNG（160/512px） |
| 素材の整理 | `assets/characters/README.md` にファイル名・サイズ・用途。旧 v7 は `assets/characters/archive/v7/` |
| 図鑑登録 | 済み。`shared/character-library/gen.py` にシリーズ「天気予報 (Forecast_Agent)」9 体＋目の部品 2 種（`half` 半目・`swirl` ぐるぐる目）。図鑑 485 体／ERROR は登録前と同じ 29 件（既存キャラ分。追加 9 体の ERROR/WARN は 0） |
| Forecast_Agent への同期 | 済み。`characters.config.json` → `assets/characters/Characters.html`（GAS には上げない。下の「注意」） |
| タブのアイコン（ファビコン） | **本番反映済み（2026-10-01、公開 @29）**。`doGet` → `webSetFavicon_` → `HtmlOutput.setFaviconUrl(FORECAST_FAVICON_URL)`。HEAD と公開版 29 の 21 ファイルがローカルと一致。タブでの見え方は利用者のブラウザで確認（アプリ内ブラウザは社内 SSO で未ログイン） |
| 表示条件・画面への配置 | **未実装（ユーザー確認待ち）**。下の提案 |

## v8 デザインの要点

| id | キャラ | v8 での形 |
|---|---|---|
| yomi | よみ | 青いドーム頭＋紺の画面の顔（水色の目）＋頭上に風車型風向風速計（プロペラ＋尾翼、アメダスの形）＋短い手足 |
| kaisei | 快晴 | 太陽（長短交互の光線）＋半目＋ニヤッと口（傲慢顔） |
| harenochi | 晴れのち曇り | 後ろの太陽＝半目＋への字（ムッと睨む）、前の雲＝閉じ目＋ニヤッ（しれっと）。v7 の id `harenchi` から改名 |
| kumori | 曇り | 平たい雲＋大きな半目＋一文字の口 |
| ame | 雨 | 濃い青（#3F74D6）の雲＋「> <」の目＋開いた口（大泣き）＋涙＋しずく形の雨粒 |
| sekka | 雪 | 雲なしの六花（各枝に小枝 2 対）＋淡い六角の芯。図鑑に `yuki`（ゆき）が既にあるため id は `sekka` |
| taifuu | 台風 | 竜巻型。互い違いのレンズ帯でねじれ、上の帯に顔（点目＋〜の口） |
| tenpen | 天変地異 | 燃える尾の岩の彗星＋ぐるぐる目（表情違いでも目は回ったまま）＋火花 |
| mikakunin | 未確認・霧 | 波打つ霧の帯の切れ間から目だけがのぞく（口なし） |

顔はすべて図鑑の共通部品（目2点＋口1本）。図鑑が生成する SVG と `assets/characters/svg/` は一致する（直すときは cast.py と gen.py を両方揃える）。

## 書き出し・検証

```bash
cd assets/characters/src && python3 build.py   # svg/ png/ yomi/ favicon/ review/ と Forecast_WebApp.js のアイコン区間を作り直す
node tests/forecast-favicon.test.mjs            # アイコンの URL 形式・64x64 PNG・doGet の呼び出し・別ファイルを作らないこと
```

PNG 化は Playwright 同梱の chrome-headless-shell（`~/Library/Caches/ms-playwright/`）を直接呼ぶ。ユーザーは HTML/PDF を開けないので、確認には PNG を送る。

## 注意（2026-10-01 に判明）

- **GAS のファイルを増やさない。** 管理ハブの複製は、プロジェクトのファイルが `VNEXT_ADMIN_RUNTIME_FILE_TYPES_`（`VNext_ClientRuntimeProvisioning.js`）の 21 ファイルと完全に一致することを本番でも検査する。新しい `.js` / `.html` を足すと複製が「exactly the 21 clasp-target files」で止まる。タブのアイコンは `Forecast_WebApp.js` の中の自動生成区間に置き、`Characters.html` は `assets/characters/`（GAS 対象外）に出している。画面に組み込むときは既存の HTML に埋め込むか、許可リストの変更とあわせて行う。
- `.claspignore` に `assets/**` を追加した（以前は `!**/*.html` で `assets/characters/chars_v7.html` が GAS 対象に入っていた。本番には上がっていなかった）。
- タブのアイコンは `HtmlOutput.setFaviconUrl` で渡す。ページ内の `<link rel="icon">` は Apps Script に無視される。URL の末尾は `#favicon.png`（拡張子の無い URL・素の data URI は黙って捨てられる。Tanka の 2026-09-30 の記録と同じ）。
- 既存テストのうち 4 件（`vnext-admin-runtime-copy` / `vnext-empty-pilot-upgrade` / `vnext-integration` / `vnext-uat-feedback`）は 3554fdf の時点で古いまま失敗していたが、4008318（別セッション、テストのみ）で現行コードに合わせて直り、2026-10-01 時点で全 11 件 PASS。

## 本番反映（ファビコン）— 2026-10-01 実施済み（公開 @29）

ユーザーの「デプロイして」で実施。push 前の照合で、本番 HEAD とローカルの差分は `Forecast_WebApp.js`（アイコン）と `Forecast_WebAppUI.html`（ページ内アイコン 1 行）だけだった。
`clasp push`（21 ファイル）→ 再取得して 21/21 一致 → `clasp deploy -i AKfycbzKsq… -d "クライアント別売上予測 Webアプリ（タブのアイコンをよみに f0595e4）"` で @28 → @29 → 版 29 を取得して 21/21 一致。
**push の直前に `clasp status` が 66 ファイルになっていた。** 別セッションが `.claude/worktrees/<name>/`（リポジトリの中）に作業用ワークツリーを作り、その .js/.html が `!**/*.js` で拾われていた。
`.claspignore` に `.claude/**` を足して 21 ファイル（許可リストと完全一致）に戻してから push した（f0595e4）。push 前は必ず件数と中身（21 件・許可リスト一致）を確かめる。

以降、同じ手順で反映するとき:

1. `clasp status`（21 ファイル）→ `clasp push`（HEAD に反映）。
2. HEAD のデプロイ（`clasp deployments` の `@HEAD`: `AKfycby2pARjKmxoBcOC8-qFjG9QTYieqJ5MtE-yI7-h4BiW`）の `/dev` を開き、外枠の HTML に `link[rel~=icon]` と PNG の base64（`iVBORw0KGgo`）・`#favicon.png` があることを確かめる。
3. 公開デプロイ `AKfycbzKsqTkHbiOS96tG9WHO1rveH8TOOJIchm9EzSeJNPCu2Z5rLEKxWoCzl3JoSWSemogmg`（現行 @28）を `clasp deploy -i <id> -d "<説明>"` で同じ URL のまま版上げし、`/exec` でも同じ確認をする。タブのアイコンは Chrome が覚えているので、変わらなければタブを閉じて開き直す。

## 表示条件の提案（未確定・ユーザー確認待ち）

検証画面が既に持っている値（`webParseEval_` の `insights[]` の制約超過フラグ `annualBreach` / `halfBreach` / `overBreach` / `rangeBreach`、`compare[]` の `ape` / `signedErr`、`hasActuals`）を**読むだけ**で決める。計算・判定ロジックは変えない。最新月で判定し、上から順に当てはまったものを出す。

| 優先 | キャラ | 条件（案） |
|---|---|---|
| 1 | 未確認・霧 | 実績が未取込、または実績のある月が 3 か月未満 |
| 2 | 天変地異 | 最新月の APE が 150% 以上（ごくまれ） |
| 3 | 台風 | 直近 3 か月で誤差の向き（上振れ・下振れ）が入れ替わり、かつ APE 30% 以上が 2 か月以上 |
| 4 | 雪 | 過大（予測が実績を上回る）の超過がある（数字の冷え込み） |
| 5 | 雨 | 制約超過が 2 つ以上 |
| 6 | 曇り | 制約超過が 1 つ |
| 7 | 晴れのち曇り | 超過は無いが、APE が 2 か月続けて悪化している（崩れの兆し） |
| 8 | 快晴 | 超過なし |

未回答の確認事項：上の条件と数値（3 か月・150%・30%）、天変地異を通常運用で出すか、画面への配置。

## 画面への配置の提案（未確定）

- タブのアイコン：よみ（実装済み・本番未反映）
- ホーム「次の一手」：今の矢印キャラ（`CHAR_ARROW`）を よみ（通常案内）に。前回から天気が変わった月は よみ（変化発見）
- 検証画面の上部：最新月の天気キャラ＋ひとこと（判定の理由になったフラグを添える）
- データが無いときの空の表示（今の歯車キャラ `CHAR_GEAR`）：未確認・霧 または よみ（観測中）
- 実行中の待ち表示：よみ（観測中）／完了の通知：よみ（確認完了）

## 参考資料

- `assets/characters/README.md`（素材一覧・サイズ・用途）
- `~/Documents/GAS/shared/character-library/README.md`（図鑑の使い方）
- `DESIGN_bayesian_autolearn_vertex_JA.md`（予測・学習ロジックの設計）／`AI_HANDOFF.md`（Webアプリ全体の引き継ぎ）
