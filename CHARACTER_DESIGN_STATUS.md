# キャラクターデザイン作業 — 引き継ぎ記録（2026-09-30）

他のAIエージェントへ引き継ぐための現状記録。Forecast_Agent 売上予測Webアプリのキャラクター「よみ」＋天気キャラ7種のデザイン作業。

## 依頼の全体像（ユーザー指定・要約）

- 案内役キャラ「よみ」（ひらがな表記）：天気全体を観測し・読み解き・利用者に伝えるキャスター役。天気そのものを表すキャラとは見た目と役割を分ける。
- 「人工衛星」vs「気象観測ロボット」の2案を同タッチでラフ比較 → **ユーザーが「気象観測ロボット」を選択済み**
- 天気キャラクター7種：快晴・晴れのち曇り・曇り・雨・雪・台風・天変地異（＋未確認・データ不足用に霧を追加提案済み）
- 制約：雲をよみ本体・帽子の中心モチーフにしない／カエル不使用／親しみやすくシンプル／小さな表示でも識別可能
- 成果物の順序：①よみ2案比較ラフ＋推奨理由 ②天気7種一覧ラフ ③既存画面モックアップ ④意味・表示条件案 → 方向確認後に5ポーズ（通常案内・観測中・変化発見・説明中・確認完了）＋SVG正本＋透過PNG
- 計算・保存・権限・予測判定ロジックは変更禁止。本番反映はデザイン＋表示条件の確認後。
- **キャラクター図鑑（GASリポジトリ直下のライブラリ）への反映が必須**

## 現在の状態

- ①②は提示済み。③モックアップ・④表示条件案はドキュメントとして提示済みだが未レビュー。
- キャラSVGは **v7**（`assets/characters/` 以下）。ユーザーは品質向上に限界を感じ、他エージェントへ引き継ぎを指示。
- **未着手**：よみ5ポーズ展開、ポーズ別透過PNG、図鑑登録（下記手順）、表示条件の実装、既存画面への組み込み。

## v7 デザイン現況（`assets/characters/chars_v7.html` が描画確認用ソース）

| id | キャラ | 現状 | ユーザー指摘済みの方向 |
|---|---|---|---|
| yomi | よみ | ベレー帽状の傾いた受信皿＋ウィンク＋側面風速計＋紺スクリーン顔 | 「もっと個性的で可愛く」「一目で気象探知機とわかる」→ 方向は観測機器然、更なる個性を求められている |
| kaisei | 快晴 | 標準太陽＋半目＋片側にやり（傲慢顔） | 傲慢な表情で確定方向。微調整余地あり |
| harenchi | 晴れのち曇り | 後ろの太陽がムッと顔・前の雲がしれっとふさぐ「喧嘩」構図 | 「晴れと曇りが喧嘩している感じに」で方向確定。太陽顔を一度修正済み |
| kumori | 曇り | 平たい雲＋大きな半目・口（眠そう） | 方向性OK、表情大きめで確定 |
| ame | 雨 | 青系雲(#5b7fc4)＋泣き顔＋涙滴 | 青色強化で確定方向 |
| yuki | 雪 | 六角結晶（枝分かれ対称）＋中心顔 | 「雲は不要、きちんとした雪の結晶」で確定 |
| taifuu | 台風 | 上広がり→下収束の竜巻型積層＋顔 | 「この路線でいいが改善して」→ 竜巻路線で更に磨く |
| tenpen | 天変地異 | ぐるぐる目の彗星＋火花 | 特に指摘なし。現行維持 |
| mikakunin | 未確認・霧 | 多層波線の霧＋覗く目のみ（?廃止済み） | 「？は不要」「もっと霧らしく」で方向確定、改善余地あり |

全キャラ共通ルール：フラット・単純図形、ハイライト/テカリ禁止、80×80 viewBox、インク色は紺系統（#1f3a5f/#4A5F70/#33465a）で統一、天気キャラは「記号＋顔」でよみ（案内役）と役割を分ける。

## ファイル配置

```
assets/characters/
  chars_v7.html        # 全キャラ描画確認ページ（ブラウザで開くと一覧表示）
  chars_v7.png         # 透過背景の一覧レンダリング
  chars_v7_white.png   # 白背景の一覧レンダリング
  svg/                 # 個別SVG正本（9ファイル: yomi/kaisei/harenchi/kumori/ame/yuki/taifuu/tenpen/mikakunin）
```

描画→PNG化の手順（このVMで再現可能）:
```bash
# Playwrightでレンダリング
python3 - <<'EOF'
from playwright.sync_api import sync_playwright
import time
with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={'width':1150,'height':420}, device_scale_factor=2)
    pg.goto('file:///home/ubuntu/repos/Forecast_Agent/assets/characters/chars_v7.html')
    pg.evaluate("document.body.style.background='#fff'")
    time.sleep(0.8)
    pg.screenshot(path='/tmp/chars_white.png', full_page=True)
    pg.evaluate("document.body.style.background='transparent'")
    time.sleep(0.3)
    pg.screenshot(path='/tmp/chars_trans.png', full_page=True, omit_background=True)
    b.close()
EOF
```
注意：ユーザーはHTML/PDFをエディタでソース表示してしまうため、成果物はPNG画像で送ること（`message_user`のattachmentsにPNGパス）。

## キャラクター図鑑（ライブラリ）への登録手順

- **正本はユーザー側GASワークスペースの `~/Documents/GAS/shared/character-library/gen.py`**（このVMからは触れない）。生成物のヘッダに「手で編集しない。正本を編集して再同期」とある。
- リポジトリ内の生成物（このリポジトリの抜粋）:
  - `portal_runtime/characters.config.json`（現在3 id: yajirushi, kurippu, haguruma）
  - `portal_runtime/src/Characters.html`（生成されたビュー）
- 登録作業：①`portal_runtime/characters.config.json` に新キャラを追加（スキーマ：name/shape/color/ink/status/family/screen/place/face/persona/quirk/catch/tone{first,second,ending,style}/likes/weak/birthday/rarity/friends/lines/talk{hello,...}/svg/moods{happy,surprised,sleepy,worried,excited,wink 各svg}）②`src/Characters.html` 側にも同内容を反映 ③**ユーザーへ `gen.py` に転記する登録ブロック（JSON）を提示**し、正本側で再生成してもらう。参考実例は `~/repos/Chaos2Order/Characters.html`（24id収録）のエントリ構造。

## 表示条件の提案（未確定・ユーザー確認待ち）

- 既存指標に紐づく「表示判定のみ」の案（計算変更なし）：
  - 快晴 = 年間誤差率≤10% かつ半期WAPE≤12% かつ過大率≤5%
  - 晴れのち曇り = 直近で指標が閾値を割り込み始めた（時間的悪化の予兆。中間スコアではない）
  - 曇り = いずれかの制約超過（constraint_attention）
  - 雨 = 複数制約超過 or rangeOutside が顕著
  - 雪/台風/天変地異 = 要因別・極端ケース（天変地異はAPE>150%級のレア演出枠を提案）
  - 未確認・霧 = EVAL_LOG月数不足・データ未取込（悪い予測とは別系統）
- 未回答の確認事項：晴れのち曇りの定義（私は時間的悪化案を提示）、天変地異の使用範囲（通常状態かデモ限定か）、各キャラの配置位置。

## よみ 残タスク

- 5ポーズ：通常案内・観測中・変化発見・説明中・確認完了（ポーズ定義のみ確定、未制作）
- ポーズ別透過PNG＋SVG正本の整理（ファイル名・サイズ・用途を明記して提出）
- 図鑑登録（上記手順）
- 既存画面への組み込み（ホーム「次の一手」・検証画面等。`Forecast_WebAppUI.html` の `.charbox` パターン準拠：アニメーションは控えめ・クリックで再再生）

## 運用・デプロイ手順（既存ワークフロー）

```bash
cd /home/ubuntu/repos/Forecast_Agent
git add -A && git commit -m "..."
git pull --ff-only && git push origin master
clasp push
clasp deploy --deploymentId AKfycbzKsqTkHbiOS96tG9WHO1rveH8TOOJIchm9EzSeJNPCu2Z5rLEKxWoCzl3JoSWSemogmg --description "クライアント別売上予測 Webアプリ"
```
ライブURL = 上記deploymentId + `/exec`（現行 @28）。
Playwright認証: `~/.clasprc.json` の `tokens.default` の client_id/secret/refresh_token を `https://oauth2.googleapis.com/token` にPOST→access_token→`browser.new_context(extra_http_headers={'Authorization':'Bearer '+token})`。

## 参考資料

- `DESIGN_bayesian_autolearn_vertex_JA.md`（予測・学習ロジックの設計）
- `AI_HANDOFF.md`（Webアプリ全体の引き継ぎ）
- `~/repos/Chaos2Order/Characters.html`（図鑑の実例24件 — タッチ・スキーマの参照元）
