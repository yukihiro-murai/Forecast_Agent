# -*- coding: utf-8 -*-
"""v8 キャラの書き出し — SVG 正本 (svg/<id>.svg) と確認用の一覧 PNG。

    python3 build.py                 # svg/ を書き出し、一覧 PNG (白背景) を ../review/chars_v8.png に作る
    python3 build.py --name r2       # 一覧 PNG の名前を review/chars_v8_r2.png にする (確認ラウンドごと)

PNG 化はヘッドレス Chromium (Playwright 同梱の chrome-headless-shell) で行う。
確認用の HTML は一時ディレクトリに作り、リポジトリには残さない (HTML は .claspignore 次第で GAS に上がるため)。
"""
import glob, html, os, re, subprocess, sys, tempfile
import cast

HERE = os.path.dirname(os.path.abspath(__file__))
V8 = os.path.dirname(HERE)
V7_HTML = os.path.join(os.path.dirname(V8), 'chars_v7.html')

DESC = {
    'yomi': '気象観測ロボ・案内役', 'kaisei': '半目＋ニヤッ（傲慢）', 'harenochi': '太陽ムッ × 雲しれっと',
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
    raise SystemExit('chrome-headless-shell が見つからない (npx playwright install chromium-headless-shell)')


def shoot(html_text, png_path, w, h, scale=2, transparent=False):
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


def v7_svgs():
    if not os.path.exists(V7_HTML):
        return {}
    src = open(V7_HTML, encoding='utf-8').read()
    out = {}
    for m in re.finditer(r"CHARS\.(\w+)='(.*?)';\n", src, re.S):
        out[m.group(1)] = re.sub(r"'\s*\n\+'", '', m.group(2))
    return out


def write_svgs():
    os.makedirs(os.path.join(V8, 'svg'), exist_ok=True)
    for c in cast.CAST:
        open(os.path.join(V8, 'svg', c[0] + '.svg'), 'w', encoding='utf-8').write(cast.svg_of(c) + '\n')
    return [c[0] for c in cast.CAST]


def review_page(title, note):
    v7 = v7_svgs()
    big = ''.join(
        f'<div class="cell"><div class="big">{cast.svg_of(c)}</div><div class="nm">{html.escape(c[1])}</div>'
        f'<div class="ds">{html.escape(DESC[c[0]])}</div>'
        f'<div class="sm"><span class="s48">{cast.svg_of(c)}</span><span class="s32">{cast.svg_of(c)}</span></div></div>'
        for c in cast.CAST)
    prev = ''
    if v7:
        prev = ('<div class="prev"><h2>参考：前回 v7</h2><div class="prow">' + ''.join(
            f'<div class="pc"><div class="pv">{v7.get(V7_ID.get(c[0], c[0]), "")}</div><div class="pn">{html.escape(c[1])}</div></div>'
            for c in cast.CAST) + '</div></div>')
    css = '''body{background:#fff;margin:0;padding:28px 34px 30px;font-family:"Hiragino Kaku Gothic ProN","Hiragino Sans","Noto Sans JP",sans-serif;color:#1f3a5f}
h1{font-size:21px;margin:0 0 4px;letter-spacing:.02em} .note{font-size:13px;color:#5A6B7E;margin:0 0 22px}
.grid{display:grid;grid-template-columns:repeat(5,196px);gap:26px 14px}
.cell{text-align:center;padding:14px 6px 12px;border-radius:14px;background:#F7FAFD}
.big{width:132px;height:132px;margin:0 auto} .big svg,.sm svg,.pv svg{width:100%;height:100%;display:block}
.nm{font-size:16px;font-weight:800;margin-top:8px} .ds{font-size:12px;color:#5A6B7E;margin-top:3px}
.sm{display:flex;gap:14px;justify-content:center;align-items:flex-end;margin-top:10px;height:48px}
.s48{width:48px;height:48px;display:inline-block} .s32{width:32px;height:32px;display:inline-block}
.prev{margin-top:26px;border-top:1px solid #DCE4EC;padding-top:14px}
.prev h2{font-size:13px;color:#8796A8;margin:0 0 10px;font-weight:700}
.prow{display:flex;gap:20px;flex-wrap:nowrap} .pc{text-align:center;width:88px;opacity:.9}
.pv{width:60px;height:60px;margin:0 auto} .pn{font-size:11px;color:#8796A8;margin-top:3px}
.legend{font-size:11.5px;color:#8796A8;margin-top:6px}'''
    return (f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>{css}</style></head><body>'
            f'<h1>{html.escape(title)}</h1><p class="note">{html.escape(note)}</p>'
            f'<div class="grid">{big}</div>'
            '<div class="legend">各キャラ下の小さい図は 48px・32px 表示（小さくても見分けられるかの確認用）</div>'
            f'{prev}</body></html>')


def main():
    name = ''
    if '--name' in sys.argv:
        name = '_' + sys.argv[sys.argv.index('--name') + 1]
    ids = write_svgs()
    os.makedirs(os.path.join(V8, 'review'), exist_ok=True)
    png = shoot(review_page('売上予測キャラクター v8 案（よみ＋天気8種）',
                            'フラット・単純図形・テカリなし。顔はキャラクター図鑑の共通部品（目2点＋口1本）で統一し、そのまま図鑑に登録できる形にしています。'),
                os.path.join(V8, 'review', f'chars_v8{name}.png'), 1150, 860)
    print('svg:', len(ids), '件 →', os.path.join(V8, 'svg'))
    print('png:', png)


if __name__ == '__main__':
    main()
