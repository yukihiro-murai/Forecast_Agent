# -*- coding: utf-8 -*-
"""Forecast_Agent 売上予測Webアプリのキャラ v8 — 図形の正本。

各キャラは body (顔なしの図形, 80×80) + fs (図鑑の顔部品の指定) + ink (顔の色) で定義する。
図鑑 (shared/character-library/gen.py) の add(body=..., fs=...) と同じ形なので、そのまま登録できる。
書き出し: python3 build.py (svg/ と一覧 PNG を作る)
"""
import math
import facekit as kit
NAVY = '#0F3557'; INK = '#1f3a5f'; SLATE = '#33465a'; WHITE = '#ffffff'
def f(v): return kit._f(round(v, 2))

# ---------- よみ: 気象観測ロボ (ドーム頭 + 風車型風向風速計) ----------
Y_BLUE = '#29A0E2'; Y_CYAN = '#8FE3FF'; Y_LIGHT = '#7FC8F0'
def aerovane(cx=40, top=2, mast_bottom=24, body=NAVY, prop=Y_LIGHT):
    y = top + 11
    return (f'<path d="M{cx} {mast_bottom}V{y+3}" stroke="{body}" stroke-width="2.8" stroke-linecap="round"/>'
            f'<path d="M{cx-10} {y}a3.2 3.2 0 0 1 3.2-3.2h13l5 3.2-5 3.2h-13a3.2 3.2 0 0 1-3.2-3.2z" fill="{body}"/>'
            f'<path d="M{cx+7} {y-2.6}l5.5-8.4h4.2l-2.4 8.4zM{cx+7} {y+2.6}l5.5 8.4h4.2l-2.4-8.4z" fill="{body}"/>'
            f'<rect x="{cx-14.6}" y="{y-8.5}" width="3.4" height="17" rx="1.7" fill="{prop}"/>'
            f'<circle cx="{cx-11}" cy="{y}" r="2.6" fill="{body}"/>')
def arm(x1, y1, x2, y2, color=NAVY):
    return f'<path d="M{f(x1)} {f(y1)}L{f(x2)} {f(y2)}" stroke="{color}" stroke-width="3.4" stroke-linecap="round"/><circle cx="{f(x2)}" cy="{f(y2)}" r="3.6" fill="{color}"/>'
YOMI_SHELL = (f'<rect x="27" y="63" width="8" height="10" rx="3" fill="{NAVY}"/><rect x="45" y="63" width="8" height="10" rx="3" fill="{NAVY}"/>'
              f'<path d="M14 47a26 24 0 0 1 52 0v5a14 14 0 0 1-14 14H28a14 14 0 0 1-14-14z" fill="{Y_BLUE}"/>'
              f'<rect x="21" y="35" width="38" height="24" rx="10" fill="{NAVY}"/>')
yomi_body = aerovane() + arm(16, 52, 9.4, 61) + arm(64, 49, 70.8, 38.8) + YOMI_SHELL
YOMI_FS = {'eyes': 'dots', 'mouth': 'smile', 'x': 40, 'y': 45, 'gap': 8, 'r': 3.5, 'my': 7.5, 'w': 10}

# ---------- 太陽 (快晴・晴れのち曇り 共通) ----------
SUN = '#F7B52C'; RAY = '#F39A1E'
def sun_rays(cx, cy, r_in, long, short, n=8, sw=4.4, color=RAY, start=-90, only=None):
    out = []
    for i in range(n):
        if only and not only(i): continue
        a = math.radians(start + i * 360 / n); L = long if i % 2 == 0 else short
        out.append(f'M{f(cx + math.cos(a)*r_in)} {f(cy + math.sin(a)*r_in)}L{f(cx + math.cos(a)*(r_in+L))} {f(cy + math.sin(a)*(r_in+L))}')
    return f'<path d="{"".join(out)}" stroke="{color}" stroke-width="{sw}" stroke-linecap="round"/>'
kaisei_body = sun_rays(40, 41, 23, 11, 7) + f'<circle cx="40" cy="41" r="19" fill="{SUN}"/>'
KAISEI_FS = {'eyes': 'half', 'mouth': 'smirk', 'x': 40, 'y': 38.5, 'gap': 7.5, 'r': 3, 'my': 9, 'w': 10}

# ---------- 晴れのち曇り: 後ろの太陽がムッと睨み、前の雲はしれっと ----------
HN_CLOUD = '#A8BBCF'
harenochi_body = (sun_rays(26, 25, 16, 7, 5, n=8, sw=3.8, only=lambda i: i in (0, 1, 5, 6, 7))
                  + f'<circle cx="26" cy="25" r="14" fill="{SUN}"/>'
                  + kit.EYES['half'](28.5, 21.5, INK, 4.8, 2.3) + kit.MOUTHS['frown'](28.5, 26.5, INK, 6.5)
                  + f'<path d="M25 66c-7 0-11-5-11-10.5S18 45 25 45c.6-8 7.2-14 15.2-14 6.4 0 11.8 3.8 14 9.4'
                    f'C61.2 40.2 67 45.6 67 52.6c3.4 1.6 5.5 4.6 5.5 8 0 3-2.4 5.4-6.5 5.4z" fill="{HN_CLOUD}"/>')
HARENOCHI_FS = {'eyes': 'closed', 'mouth': 'smirk', 'x': 46, 'y': 51, 'gap': 7, 'r': 2.6, 'my': 7.5, 'w': 9}

# ---------- 曇り: 平たい雲 + 大きな半目 ----------
CLOUD_G = '#8FA4BA'
kumori_body = (f'<path d="M15 59c-6.5 0-10-4.4-10-9.4s3.8-9.2 9.6-9.6C15.4 33 21 28 28 28c2.8 0 5.4.8 7.6 2.3C38.4 24.6 44.4 21 51.2 21'
               f'c9.2 0 16.6 7 17.2 15.9 4.8 1.4 8.1 5.6 8.1 10.8 0 6.4-4.8 11.3-11 11.3z" fill="{CLOUD_G}"/>')
KUMORI_FS = {'eyes': 'half', 'mouth': 'line', 'x': 41, 'y': 43, 'gap': 9.5, 'r': 3.6, 'my': 9, 'w': 11}

# ---------- 雨: 青い雲が大泣き (涙 + 雨粒) ----------
RAIN = '#3F74D6'; TEAR = '#BFE0FF'
def drop(cx, cy, s, color):
    return (f'<path d="M{f(cx)} {f(cy-s*1.45)}C{f(cx+s*0.35)} {f(cy-s*0.8)} {f(cx+s)} {f(cy-s*0.2)} {f(cx+s)} {f(cy+s*0.35)}'
            f'a{f(s)} {f(s)} 0 0 1 -{f(2*s)} 0C{f(cx-s)} {f(cy-s*0.2)} {f(cx-s*0.35)} {f(cy-s*0.8)} {f(cx)} {f(cy-s*1.45)}z" fill="{color}"/>')
ame_body = (f'<path d="M19 50c-7 0-11.5-5-11.5-11S12 28 19.5 28c.6-9.5 8.4-17 18-17 7.4 0 13.6 4.4 16.4 10.7'
            f'C61.6 21.9 69 28.4 69 37.2c3.2 1.9 4.5 5.2 4.5 8.1 0 2.7-2 4.7-6 4.7z" fill="{RAIN}"/>'
            + drop(24, 64, 3.6, RAIN) + drop(40, 70, 3.6, RAIN) + drop(56, 64, 3.6, RAIN)
            + drop(30.2, 41.6, 2.1, TEAR) + drop(49.8, 41.6, 2.1, TEAR))
AME_FS = {'eyes': 'squint', 'mouth': 'open', 'x': 40, 'y': 33, 'gap': 9, 'r': 2.6, 'my': 9, 'w': 9}

# ---------- 雪: 雲なしの六花 (枝分かれ対称) + 六角の芯 ----------
SNOW = '#6EA9DF'; SNOW_IN = '#E3F0FB'; SNOW_INK = '#2F5D8A'
SNOW_ARM = ('<path d="M40 30.5V8.5 M40 23l-5.5-5.2M40 23l5.5-5.2 M40 15.5l-4-3.8M40 15.5l4-3.8" '
            f'stroke="{SNOW}" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round" fill="none"/>')
_hex = ' '.join(f'{f(40 + 13.5*math.cos(math.radians(60*i-90)))},{f(40 + 13.5*math.sin(math.radians(60*i-90)))}' for i in range(6))
sekka_body = (''.join(f'<g transform="rotate({60*i} 40 40)">{SNOW_ARM}</g>' for i in range(6))
              + f'<polygon points="{_hex}" fill="{SNOW_IN}" stroke="{SNOW}" stroke-width="3.2" stroke-linejoin="round"/>')
SEKKA_FS = {'eyes': 'arcs', 'mouth': 'smile', 'x': 40, 'y': 38.5, 'gap': 5.6, 'r': 2.2, 'my': 6.5, 'w': 7}

# ---------- 台風: 上広がり→下収束のレンズ帯 (互い違いでねじれる) ----------
TY1 = '#3B5BA5'; TY2 = '#5674BD'; TYW = '#9DB0DA'
def sweep(cx, cy, w, h, color, lean=0):
    x1, x2 = cx - w/2, cx + w/2
    return (f'<path d="M{f(x1)} {f(cy - h*0.25 - lean)}C{f(cx - w*0.25)} {f(cy - h*1.05)} {f(cx + w*0.25)} {f(cy - h*1.05)} {f(x2)} {f(cy - h*0.25 + lean)}'
            f'C{f(cx + w*0.22)} {f(cy + h*0.95)} {f(cx - w*0.22)} {f(cy + h*0.95)} {f(x1)} {f(cy - h*0.25 - lean)}z" fill="{color}"/>')
taifuu_body = (f'<path d="M6 33c-3-2-3.5-5.5-1-8M74 28c3 1.5 3.6 5 1.2 7.8M14 55c-3-1.2-4-4-2.4-6.6" stroke="{TYW}" stroke-width="2.6" fill="none" stroke-linecap="round"/>'
               + sweep(41.5, 71, 9, 3.6, TY2, 0.5) + sweep(38, 63.5, 17, 5.2, TY1, 1) + sweep(42.5, 54.5, 26, 6.8, TY2, 1.2)
               + sweep(37.5, 44, 36, 8.2, TY1, 1.5) + sweep(42, 32, 47, 9.6, TY2, 1.8) + sweep(40, 19, 60, 13, TY1))
TAIFUU_FS = {'eyes': 'dots', 'mouth': 'wave', 'x': 40, 'y': 17, 'gap': 8, 'r': 2.7, 'my': 6, 'w': 9}

# ---------- 天変地異: ぐるぐる目の彗星 + 火花 ----------
ROCK = '#6C7E96'; FIRE1 = '#F08A24'; FIRE2 = '#F7B52C'
tenpen_body = (f'<path d="M47 36C35 33 20 23 7 7c15 13 29 18 41 20z" fill="{FIRE1}"/>'
               f'<path d="M45 45C36 42 24 34 13 21c12 10 24 15 33 16z" fill="{FIRE2}"/>'
               f'<path d="M46.5 33.5l15.5 1.8 9 10.2-1.6 13.4-11.8 7.6-14.6-1.9-8.2-11 2.4-13.8z" fill="{ROCK}" stroke="{ROCK}" stroke-width="3" stroke-linejoin="round"/>'
               f'<path d="M73.5 33l3.5-5M76 41.5l4.2-1.6M70 27.5l1-5" stroke="{FIRE1}" stroke-width="2.4" stroke-linecap="round"/>')
TENPEN_FS = {'eyes': 'swirl', 'mouth': 'o', 'x': 55, 'y': 47, 'gap': 7, 'r': 3.2, 'my': 10, 'w': 9, 'keep_eyes': True}

# ---------- 未確認・霧: 霧の帯の切れ間から目だけのぞく ----------
FOG1 = '#A9BACB'; FOG2 = '#C6D2DE'; FOG_IN = '#E1E8EF'
def fog(d, color, sw=7.5, op=1):
    o = '' if op == 1 else f' opacity="{op}"'
    return f'<path d="{d}" stroke="{color}" stroke-width="{sw}" fill="none" stroke-linecap="round"{o}/>'
mikakunin_body = (fog('M24 14c6-3 12-3 18 0s12 3 18 0', FOG2, 6.5)
                  + fog('M8 26c7-3.6 14-3.6 21 0s14 3.6 21 0 10-3 14-1', FOG1)
                  + fog('M10 40h60', FOG_IN, 13)
                  + fog('M12 54c7-3.6 14-3.6 21 0s14 3.6 21 0s9-3 14-1.5', FOG1)
                  + fog('M24 66c5-2.6 10-2.6 15 0s10 2.6 15 0', FOG2, 6.5))
MIKAKUNIN_FS = {'eyes': 'dots', 'mouth': 'none', 'x': 40, 'y': 40, 'gap': 7.5, 'r': 3.2, 'my': 9, 'w': 9, 'no_mouth': True}

CAST = [
    ('yomi', 'よみ', yomi_body, YOMI_FS, Y_CYAN),
    ('kaisei', '快晴', kaisei_body, KAISEI_FS, INK),
    ('harenochi', '晴れのち曇り', harenochi_body, HARENOCHI_FS, INK),
    ('kumori', '曇り', kumori_body, KUMORI_FS, SLATE),
    ('ame', '雨', ame_body, AME_FS, WHITE),
    ('sekka', '雪', sekka_body, SEKKA_FS, SNOW_INK),
    ('taifuu', '台風', taifuu_body, TAIFUU_FS, WHITE),
    ('tenpen', '天変地異', tenpen_body, TENPEN_FS, WHITE),
    ('mikakunin', '未確認・霧', mikakunin_body, MIKAKUNIN_FS, SLATE),
]
def svg_of(c, mood=None):
    """図鑑の compose() と同じ合成 (body + 顔)。mood で表情違い"""
    return kit.compose(c[2], c[3], c[4], mood)

