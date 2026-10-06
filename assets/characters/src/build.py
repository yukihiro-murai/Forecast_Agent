# -*- coding: utf-8 -*-
"""Trends2Targets（旧 Forecast_Agent）キャラ素材の書き出し — 正本は cast.py。手で SVG / PNG を直さず、ここから作り直す。

    python3 build.py              # すべて書き出す (下の一覧)
    python3 build.py --review r2  # 確認用の一覧 PNG を review/chars_v8_r2.png にも作る (確認ラウンドごと)

書き出すもの (assets/characters/ 以下。サイズと用途は README.md):
  svg/<id>.svg                    9 体の SVG 正本 (80×80)
  png/<id>_160.png / _512.png     9 体の透過 PNG
  yomi/svg/yomi_<pose>.svg        よみ 5 ポーズの SVG 正本
  yomi/png/yomi_<pose>_160.png / _512.png   5 ポーズの透過 PNG
  favicon/yomi_favicon.svg / yomi_favicon_64.png   タブのアイコン (32 升の専用の絵)
  review/chars_v8.png / review/yomi_poses.png      確認用の一覧 (白背景)
  ../../Forecast_WebApp.js        「タブのアイコン」区間の FORECAST_FAVICON_URL (64×64 PNG の data URI + #favicon.png) を書き換える
                                  (計算の元。変えたら node app/tools/build-engine.mjs で LegacyEngine.js を作り直す)
  ../../app/src/Assets.js         新アプリのタブのアイコン (APP_FAVICON_URL) と画面のキャラ (APP_UI_CHARS_JS)。ファイルごと生成

PNG 化はヘッドレス Chromium (Playwright 同梱の chrome-headless-shell) で行う。
確認用の HTML は一時ディレクトリに作り、リポジトリには残さない。
"""
import base64, glob, html, os, re, struct, subprocess, sys, tempfile, zlib
import cast

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.dirname(HERE)                      # assets/characters
REPO = os.path.dirname(os.path.dirname(ASSETS))     # リポジトリ直下
V7_HTML = os.path.join(ASSETS, 'archive', 'v7', 'chars_v7.html')
PNG_SIZES = (160, 512)
FAVICON_PX = 64
WEBAPP_JS = os.path.join(REPO, 'Forecast_WebApp.js')
APP_ASSETS_JS = os.path.join(REPO, 'app', 'src', 'Assets.js')  # 新アプリ（段階1〜）
DATA_URL_PREFIX = 'data:image/png;base64,'
# Apps Script の setFaviconUrl は末尾が画像の拡張子でない URL を黙って捨てる。# 以降は URL の断片で画像データではない
DATA_URL_SUFFIX = '#favicon.png'

DESC = {
    'yomi': '気象観測ロボ・案内役', 'mousho': '炎の光線＋汗＋舌出し', 'kaisei': '半目＋ニヤッ（傲慢）', 'harenochi': '太陽ムッ × 雲しれっと',
    'kumori': '平たい雲・大きな半目', 'ame': '濃い青・大泣き', 'sekka': '雲なしの六花',
    'taifuu': '竜巻型・ねじれ帯', 'tenpen': 'ぐるぐる目の彗星', 'mikakunin': '霧の帯から目だけ',
}
V7_ID = {'sekka': 'yuki', 'harenochi': 'harenchi'}  # v7 で別名だった id


def chrome():
    pats = [os.path.expanduser('~/Library/Caches/ms-playwright/chromium_headless_shell-*/chrome-headless-shell-mac-*/chrome-headless-shell'),
            os.path.expanduser('~/.cache/ms-playwright/chromium_headless_shell-*/chrome-headless-shell-linux*/chrome-headless-shell')]
    for p in pats:
        hits = sorted(glob.glob(p))
        if hits:
            return hits[-1]
    # 無ければ Mac の Google Chrome を使う（2026-10-06。同じ PNG になることを確かめた）
    mac = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
    if os.path.exists(mac):
        return mac
    raise SystemExit('chrome-headless-shell が見つからない (npx playwright install chromium-headless-shell)')


def shoot(html_text, png_path, w, h, scale=2, transparent=False):
    os.makedirs(os.path.dirname(png_path), exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        hp = os.path.join(td, 'page.html')
        open(hp, 'w', encoding='utf-8').write(html_text)
        args = [chrome(), '--headless', '--disable-gpu', '--hide-scrollbars', f'--force-device-scale-factor={scale}',
                f'--window-size={w},{h}', f'--screenshot={png_path}']
        if transparent:
            args.append('--default-background-color=00000000')
        r = subprocess.run(args + ['file://' + hp], capture_output=True, text=True, timeout=120)
        if not os.path.exists(png_path):
            raise SystemExit(r.stderr[-800:])
    return png_path


def png_info(path):
    """(幅, 高さ, 四隅がすべて透明か)。8bit RGBA・インターレースなしの PNG だけを読む"""
    b = open(path, 'rb').read()
    w, h = struct.unpack('>II', b[16:24])
    if b[24] != 8 or b[25] != 6 or b[28] != 0:
        return w, h, False
    pos, idat = 8, b''
    while pos < len(b):
        n = struct.unpack('>I', b[pos:pos + 4])[0]
        if b[pos + 4:pos + 8] == b'IDAT':
            idat += b[pos + 8:pos + 8 + n]
        pos += 12 + n
    raw = zlib.decompress(idat)
    stride = w * 4
    rows, prev = [], bytearray(stride)
    for y in range(h):
        ft, line = raw[y * (stride + 1)], bytearray(raw[y * (stride + 1) + 1:(y + 1) * (stride + 1)])
        for i in range(stride):
            a = line[i - 4] if i >= 4 else 0
            up = prev[i]
            c = prev[i - 4] if i >= 4 else 0
            if ft == 1: line[i] = (line[i] + a) & 255
            elif ft == 2: line[i] = (line[i] + up) & 255
            elif ft == 3: line[i] = (line[i] + ((a + up) >> 1)) & 255
            elif ft == 4:
                p = a + up - c; pa, pb, pc = abs(p - a), abs(p - up), abs(p - c)
                line[i] = (line[i] + (a if pa <= pb and pa <= pc else up if pb <= pc else c)) & 255
        rows.append(line); prev = line
    corners = [rows[0][3], rows[0][stride - 1], rows[h - 1][3], rows[h - 1][stride - 1]]
    return w, h, all(a == 0 for a in corners)


def transparent_png(svg_text, path, px):
    page = ('<!DOCTYPE html><html><head><meta charset="utf-8"><style>html,body{margin:0;padding:0;background:transparent;overflow:hidden}'
            f'svg{{width:{px}px;height:{px}px;display:block}}</style></head><body>{svg_text}</body></html>')
    shoot(page, path, px, px, scale=1, transparent=True)
    w, h, clear = png_info(path)
    if (w, h) != (px, px) or not clear:
        raise SystemExit(f'{path}: {w}x{h} 透明={clear} (期待 {px}x{px} の透過 PNG)')
    return path


def write(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    open(path, 'w', encoding='utf-8').write(text)


def v7_svgs():
    if not os.path.exists(V7_HTML):
        return {}
    src = open(V7_HTML, encoding='utf-8').read()
    return {m.group(1): re.sub(r"'\s*\n\+'", '', m.group(2)) for m in re.finditer(r"CHARS\.(\w+)='(.*?)';\n", src, re.S)}


CSS = '''body{background:#fff;margin:0;padding:28px 34px 30px;font-family:"Hiragino Kaku Gothic ProN","Hiragino Sans","Noto Sans JP",sans-serif;color:#1f3a5f}
h1{font-size:21px;margin:0 0 4px;letter-spacing:.02em} .note{font-size:13px;color:#5A6B7E;margin:0 0 22px}
.grid{display:grid;grid-template-columns:repeat(5,196px);gap:26px 14px}
.cell{text-align:center;padding:14px 6px 12px;border-radius:14px;background:#F7FAFD}
.big{width:132px;height:132px;margin:0 auto} .big svg,.sm svg,.pv svg{width:100%;height:100%;display:block}
.nm{font-size:16px;font-weight:800;margin-top:8px} .ds{font-size:12px;color:#5A6B7E;margin-top:3px;min-height:16px}
.sm{display:flex;gap:14px;justify-content:center;align-items:flex-end;margin-top:10px;height:48px}
.s48{width:48px;height:48px;display:inline-block} .s32{width:32px;height:32px;display:inline-block}
.prev{margin-top:26px;border-top:1px solid #DCE4EC;padding-top:14px}
.prev h2{font-size:13px;color:#8796A8;margin:0 0 10px;font-weight:700}
.prow{display:flex;gap:20px;flex-wrap:nowrap} .pc{text-align:center;width:88px;opacity:.9}
.pv{width:60px;height:60px;margin:0 auto} .pn{font-size:11px;color:#8796A8;margin-top:3px}
.legend{font-size:11.5px;color:#8796A8;margin-top:6px}'''


def cells_html(items):
    """items = [(svg, 名前, 説明)]"""
    return ''.join(
        f'<div class="cell"><div class="big">{s}</div><div class="nm">{html.escape(n)}</div><div class="ds">{html.escape(d)}</div>'
        f'<div class="sm"><span class="s48">{s}</span><span class="s32">{s}</span></div></div>' for s, n, d in items)


def review_page(title, note, with_v7=False):
    prev = ''
    v7 = v7_svgs() if with_v7 else {}
    if v7:
        prev = ('<div class="prev"><h2>参考：前回 v7</h2><div class="prow">' + ''.join(
            f'<div class="pc"><div class="pv">{v7.get(V7_ID.get(c[0], c[0]), "")}</div><div class="pn">{html.escape(c[1])}</div></div>'
            for c in cast.CAST) + '</div></div>')
    big = cells_html([(cast.svg_of(c), c[1], DESC[c[0]]) for c in cast.CAST])
    return (f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>{CSS}</style></head><body>'
            f'<h1>{html.escape(title)}</h1><p class="note">{html.escape(note)}</p><div class="grid">{big}</div>'
            '<div class="legend">各キャラ下の小さい図は 48px・32px 表示（小さくても見分けられるかの確認用）</div>'
            f'{prev}</body></html>')


def poses_page():
    big = cells_html([(cast.pose_svg(p), p[1], p[4]) for p in cast.YOMI_POSES])
    fav = cast.FAVICON_SVG
    tabs = ''.join(
        f'<span style="display:inline-flex;gap:10px;align-items:center;background:{bg};padding:8px 12px;border-radius:8px;margin-right:10px">'
        f'<span style="width:16px;height:16px;display:inline-block">{fav}</span><span style="width:32px;height:32px;display:inline-block">{fav}</span>'
        f'<span style="font-size:12px;color:{fg}">クライアント別売上予測</span></span>'
        for bg, fg in [('#FFFFFF;border:1px solid #DCE4EC', '#333'), ('#DEE1E6', '#333'), ('#35363A', '#eee')])
    return (f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>{CSS} .tabs svg{{width:100%;height:100%;display:block}}</style></head><body>'
            '<h1>よみ 5ポーズ ＋ タブのアイコン</h1>'
            '<p class="note">顔はキャラクター図鑑の共通部品（目2点＋口1本）。手と小物でポーズを分けています。</p>'
            f'<div class="grid">{big}</div>'
            '<div class="legend">各ポーズ下の小さい図は 48px・32px 表示</div>'
            '<div class="prev tabs"><h2>タブのアイコン（ファビコン）— 16px・32px（明るいタブ・非選択・暗いタブ）</h2>'
            f'<div style="display:flex;align-items:center">{tabs}</div></div></body></html>')


FAVICON_RE = re.compile(r"(// ===== タブのアイコン（自動生成: assets/characters/src/build\.py。この区間は手で編集しない） =====\n(?://.*\n)*)const FORECAST_FAVICON_URL = '[^']*';\n(// ===== /タブのアイコン =====)")


def write_favicon_url(data_url):
    """Forecast_WebApp.js の「タブのアイコン」区間だけを書き換える (区間が無ければ止める)"""
    src = open(WEBAPP_JS, encoding='utf-8').read()
    new, n = FAVICON_RE.subn(lambda m: m.group(1) + "const FORECAST_FAVICON_URL = '" + data_url + "';\n" + m.group(2), src)
    if n != 1:
        raise SystemExit('Forecast_WebApp.js に「タブのアイコン」区間が無い (または 2 つ以上ある)')
    if new != src:
        open(WEBAPP_JS, 'w', encoding='utf-8').write(new)




def write_app_assets(data_url):
    """新アプリ (app/src) のキャラ素材。ファイルごと生成する (手で編集しない)"""
    import json
    if not os.path.isdir(os.path.dirname(APP_ASSETS_JS)):
        return
    chars = {c[0]: cast.svg_of(c) for c in cast.CAST}
    poses = {p[0]: cast.pose_svg(p) for p in cast.YOMI_POSES}
    js = lambda d: json.dumps(d, ensure_ascii=False, separators=(',', ':'))
    ui = 'var CHAR_SVG = ' + js(chars) + ';\nvar YOMI_POSE = ' + js(poses) + ';'
    if '<?' in ui or '</script' in ui.lower():
        raise SystemExit('Assets: GAS テンプレートや script を壊す文字列を含む')
    text = '\n'.join([
        '/**',
        ' * Assets.js — 画面のキャラ素材（自動生成: Trends2Targets/assets/characters/src/build.py。手で編集しない）。',
        ' * APP_FAVICON_URL: タブのアイコン（よみの頭・64x64 PNG の data URI。末尾の #favicon.png は消さない）',
        ' * APP_UI_CHARS_JS: UI.html に差し込む CHAR_SVG（9 体）と YOMI_POSE（5 ポーズ）',
        ' */',
        'const APP_FAVICON_URL = ' + json.dumps(data_url) + ';',
        'const APP_UI_CHARS_JS = ' + json.dumps(ui, ensure_ascii=False) + ';',
        ''])
    old = open(APP_ASSETS_JS, encoding='utf-8').read() if os.path.exists(APP_ASSETS_JS) else None
    if old != text:
        open(APP_ASSETS_JS, 'w', encoding='utf-8').write(text)


def main():
    out = []
    for c in cast.CAST:
        s = cast.svg_of(c)
        write(os.path.join(ASSETS, 'svg', c[0] + '.svg'), s + '\n')
        for px in PNG_SIZES:
            out.append(transparent_png(s, os.path.join(ASSETS, 'png', f'{c[0]}_{px}.png'), px))
    for p in cast.YOMI_POSES:
        s = cast.pose_svg(p)
        write(os.path.join(ASSETS, 'yomi', 'svg', f'yomi_{p[0]}.svg'), s + '\n')
        for px in PNG_SIZES:
            out.append(transparent_png(s, os.path.join(ASSETS, 'yomi', 'png', f'yomi_{p[0]}_{px}.png'), px))
    write(os.path.join(ASSETS, 'favicon', 'yomi_favicon.svg'), cast.FAVICON_SVG + '\n')
    fav_png = transparent_png(cast.FAVICON_SVG, os.path.join(ASSETS, 'favicon', f'yomi_favicon_{FAVICON_PX}.png'), FAVICON_PX)
    data_url = DATA_URL_PREFIX + base64.b64encode(open(fav_png, 'rb').read()).decode('ascii') + DATA_URL_SUFFIX
    write_favicon_url(data_url)
    write_app_assets(data_url)
    os.makedirs(os.path.join(ASSETS, 'review'), exist_ok=True)
    shoot(review_page('売上予測キャラクター v8（よみ＋天気8種）',
                      'フラット・単純図形・テカリなし。顔はキャラクター図鑑の共通部品（目2点＋口1本）で統一しています。'),
          os.path.join(ASSETS, 'review', 'chars_v8.png'), 1150, 740)
    shoot(poses_page(), os.path.join(ASSETS, 'review', 'yomi_poses.png'), 1150, 560)
    if '--review' in sys.argv:
        name = sys.argv[sys.argv.index('--review') + 1]
        shoot(review_page('売上予測キャラクター v8 案（よみ＋天気8種）',
                          'フラット・単純図形・テカリなし。顔はキャラクター図鑑の共通部品（目2点＋口1本）で統一し、そのまま図鑑に登録できる形にしています。',
                          with_v7=True),
              os.path.join(ASSETS, 'review', f'chars_v8_{name}.png'), 1150, 860)
    print(f'svg {len(cast.CAST)} 体 + ポーズ {len(cast.YOMI_POSES)} / png {len(out) + 1} 枚 / favicon → {os.path.relpath(WEBAPP_JS, REPO)} ({len(data_url)} 文字)')


if __name__ == '__main__':
    main()
