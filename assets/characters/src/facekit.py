# -*- coding: utf-8 -*-
"""顔パーツ合成 — キャラクター図鑑 (~/Documents/GAS/shared/character-library/gen.py) の顔部品の写し。

gen.py の「目2点・口1本」部品 (eyes / mouths / face_svg / compose / face_label) をそのまま写し、
予報シリーズで使う 2 種類の目を足している:
  half  … 半目 (まぶたの一文字 + 下半分の黒目)。快晴の傲慢顔・曇りの眠そう顔・晴れのち曇りの太陽
  swirl … ぐるぐる目 (外へ広がる 2 回転強の渦)。天変地異
図鑑へ登録するときは、この 2 つを gen.py の EYES に足す (それ以外は gen.py と同じ実装なので見た目は一致する)。
"""

# 顔パーツ (共通ルール: 目2点・口1本まで)。中心 (cx, cy) を受け取って描く
def eyes(cx, cy, ink, gap=8, r=2.8):
    return f'<circle cx="{cx-gap}" cy="{cy}" r="{r}" fill="{ink}"/><circle cx="{cx+gap}" cy="{cy}" r="{r}" fill="{ink}"/>'
def mouth_line(cx, cy, ink, w=12):
    return f'<path d="M{cx-w/2} {cy}h{w}" stroke="{ink}" stroke-width="3" stroke-linecap="round"/>'
def mouth_smile(cx, cy, ink, w=12):
    return f'<path d="M{cx-w/2} {cy}q{w/2} 4 {w} 0" stroke="{ink}" stroke-width="2.5" fill="none" stroke-linecap="round"/>'
def mouth_frown(cx, cy, ink, w=14):
    return f'<path d="M{cx-w/2} {cy+3}q{w/2} -6 {w} 0" stroke="{ink}" stroke-width="3" fill="none" stroke-linecap="round"/>'
def mouth_o(cx, cy, ink, r=3.2):
    return f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{ink}" stroke-width="2.5"/>'
def eyes_closed(cx, cy, ink, gap=8):
    return f'<path d="M{cx-gap-4} {cy}h8M{cx+gap-4} {cy}h8" stroke="{ink}" stroke-width="3" stroke-linecap="round"/>'
def eyes_squint(cx, cy, ink):  # > <
    return f'<path d="M{cx-14} {cy-5}l7 5-7 5M{cx+14} {cy-5}l-7 5 7 5" stroke="{ink}" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"/>'
def wink(cx, cy, ink, gap=8):
    return f'<circle cx="{cx-gap}" cy="{cy}" r="2.8" fill="{ink}"/><path d="M{cx+gap-4} {cy}h8" stroke="{ink}" stroke-width="3" stroke-linecap="round"/>'

def svg(body):
    return f'<svg viewBox="0 0 80 80" xmlns="http://www.w3.org/2000/svg">{body}</svg>'

# ---------------- 顔の合成 (目2点・口1本のルールを関数で守る。body + fs から既定の顔と表情違いを作る) ----------------
EYE_KINDS = ['dots', 'closed', 'squint', 'wink', 'peek', 'arcs']
MOUTH_KINDS = ['line', 'smile', 'frown', 'o', 'open', 'cat', 'wave', 'grin', 'smirk', 'tongue', 'none']
MOODS = {  # 表情バリエーション (自動生成)。既定の顔はキャラ固有 (fs 指定)
    'happy': ('arcs', 'smile'), 'surprised': ('dots', 'o'), 'sleepy': ('closed', 'line'),
    'worried': ('dots', 'frown'), 'excited': ('squint', 'grin'), 'wink': ('wink', 'smile'),
}
MOOD_LABEL = {'happy': 'うれしい', 'surprised': 'びっくり', 'sleepy': 'ねむい', 'worried': 'しんぱい', 'excited': 'はりきり', 'wink': 'ウインク'}
TALK_KEYS = [('hello', '初対面のあいさつ'), ('morning', '朝いちばん'), ('monday', '月曜'), ('friday', '金曜'), ('empty', '何も無いとき (0 件)'),
             ('loading', '処理中・待たせるとき'), ('done', '完了・成功'), ('error', '失敗・エラー'), ('praise', '使う人をほめる'),
             ('nudge', 'そっと促す (未読・未入力)'), ('secret', 'レアに出る裏話'), ('night', '夜・残業中'), ('rest', '使いすぎ・休憩のすすめ')]
MOOD_OF_TALK = {'hello': None, 'morning': 'happy', 'monday': 'sleepy', 'friday': 'excited', 'empty': 'surprised', 'loading': 'worried', 'done': 'happy',
                'error': 'worried', 'praise': 'wink', 'nudge': None, 'secret': 'wink', 'night': 'sleepy', 'rest': 'sleepy'}
RARITY_LABEL = {'common': 'いつも', 'rare': 'たまに (週 1 くらい)', 'secret': 'ごくまれ (1%)'}

def _f(v):
    return f'{v:.1f}'.rstrip('0').rstrip('.') if isinstance(v, float) else str(v)
def f_eyes_dots(x, y, ink, gap=8, r=2.8):
    return f'<circle cx="{_f(x-gap)}" cy="{_f(y)}" r="{_f(r)}" fill="{ink}"/><circle cx="{_f(x+gap)}" cy="{_f(y)}" r="{_f(r)}" fill="{ink}"/>'
def f_eyes_closed(x, y, ink, gap=8, r=2.8):
    h = max(6, r * 2.8); sw = max(2, r)
    return f'<path d="M{_f(x-gap-h/2)} {_f(y)}h{_f(h)}M{_f(x+gap-h/2)} {_f(y)}h{_f(h)}" stroke="{ink}" stroke-width="{_f(sw)}" stroke-linecap="round"/>'
def f_eyes_squint(x, y, ink, gap=8, r=2.8):  # > <
    o, i, h = gap + 6, gap - 1, 5
    return (f'<path d="M{_f(x-o)} {_f(y-h)}l{_f(o-i)} {_f(h)}-{_f(o-i)} {_f(h)}M{_f(x+o)} {_f(y-h)}l-{_f(o-i)} {_f(h)} {_f(o-i)} {_f(h)}" '
            f'stroke="{ink}" stroke-width="{_f(max(2, r))}" fill="none" stroke-linecap="round" stroke-linejoin="round"/>')
def f_eyes_wink(x, y, ink, gap=8, r=2.8):
    h = max(6, r * 2.8)
    return f'<circle cx="{_f(x-gap)}" cy="{_f(y)}" r="{_f(r)}" fill="{ink}"/><path d="M{_f(x+gap-h/2)} {_f(y)}h{_f(h)}" stroke="{ink}" stroke-width="{_f(max(2, r))}" stroke-linecap="round"/>'
def f_eyes_peek(x, y, ink, gap=8, r=2.8):  # 片目を見開いて覗き込む
    R = r * 2.1
    return (f'<circle cx="{_f(x-gap)}" cy="{_f(y)}" r="{_f(R)}" fill="none" stroke="{ink}" stroke-width="{_f(max(2, r))}"/>'
            f'<circle cx="{_f(x-gap)}" cy="{_f(y)}" r="{_f(r*0.7)}" fill="{ink}"/><circle cx="{_f(x+gap)}" cy="{_f(y)}" r="{_f(r)}" fill="{ink}"/>')
def f_eyes_arcs(x, y, ink, gap=8, r=2.8):  # ^ ^ (にっこり閉じ目)
    h = max(6, r * 2.6); d = h * 0.6
    return (f'<path d="M{_f(x-gap-h/2)} {_f(y+1)}q{_f(h/2)} -{_f(d)} {_f(h)} 0M{_f(x+gap-h/2)} {_f(y+1)}q{_f(h/2)} -{_f(d)} {_f(h)} 0" '
            f'stroke="{ink}" stroke-width="{_f(max(2, r))}" fill="none" stroke-linecap="round"/>')
def f_eyes_yen(x, y, ink, gap=8, r=2.8):  # ¥ の目 (お金のマーク。深い V の両腕 + 縦棒 + 横棒 2 本。r で拡縮。横棒は線幅より広い間隔で潰れないようにし、縦棒は下の横棒よりさらに下へ伸ばす)
    H = r * 2.5; y0 = y - H / 2; aw = r; bw = r * 0.8; sw = max(1.6, r * 0.5)
    yj = y0 + H * 0.44; b1 = y0 + H * 0.54; b2 = y0 + H * 0.87; yb = y0 + H * 1.14
    def one(cx):
        return (f'<path d="M{_f(cx-aw)} {_f(y0)}L{_f(cx)} {_f(yj)}L{_f(cx+aw)} {_f(y0)}'
                f'M{_f(cx)} {_f(yj)}V{_f(yb)}'
                f'M{_f(cx-bw)} {_f(b1)}h{_f(bw*2)}M{_f(cx-bw)} {_f(b2)}h{_f(bw*2)}" '
                f'stroke="{ink}" stroke-width="{_f(sw)}" fill="none" stroke-linecap="round" stroke-linejoin="round"/>')
    return one(x - gap) + one(x + gap)
def f_eyes_none(x, y, ink, gap=8, r=2.8):  # 目なし (レンズが光って目が見えない眼鏡等)
    return ''
EYES = {'dots': f_eyes_dots, 'closed': f_eyes_closed, 'squint': f_eyes_squint, 'wink': f_eyes_wink, 'peek': f_eyes_peek, 'arcs': f_eyes_arcs, 'yen': f_eyes_yen, 'none': f_eyes_none}
def f_mouth_line(x, y, ink, w=12):
    return f'<path d="M{_f(x-w/2)} {_f(y)}h{_f(w)}" stroke="{ink}" stroke-width="3" stroke-linecap="round"/>'
def f_mouth_smile(x, y, ink, w=12):
    return f'<path d="M{_f(x-w/2)} {_f(y)}q{_f(w/2)} {_f(w/3)} {_f(w)} 0" stroke="{ink}" stroke-width="2.5" fill="none" stroke-linecap="round"/>'
def f_mouth_frown(x, y, ink, w=12):
    return f'<path d="M{_f(x-w/2)} {_f(y+3)}q{_f(w/2)} -{_f(w/2.4)} {_f(w)} 0" stroke="{ink}" stroke-width="3" fill="none" stroke-linecap="round"/>'
def f_mouth_o(x, y, ink, w=12):
    return f'<circle cx="{_f(x)}" cy="{_f(y)}" r="{_f(max(2.2, w/3.8))}" fill="none" stroke="{ink}" stroke-width="2.5"/>'
def f_mouth_open(x, y, ink, w=12):
    return f'<ellipse cx="{_f(x)}" cy="{_f(y)}" rx="{_f(max(2.4, w/4))}" ry="{_f(max(2.8, w/3.4))}" fill="{ink}"/>'
def f_mouth_cat(x, y, ink, w=12):  # ω
    return f'<path d="M{_f(x-w/2)} {_f(y)}q{_f(w/4)} {_f(w/2.4)} {_f(w/2)} 0t{_f(w/2)} 0" stroke="{ink}" stroke-width="2.5" fill="none" stroke-linecap="round"/>'
def f_mouth_wave(x, y, ink, w=12):  # 〜
    return f'<path d="M{_f(x-w/2)} {_f(y)}q{_f(w/4)} -{_f(w/3)} {_f(w/2)} 0t{_f(w/2)} 0" stroke="{ink}" stroke-width="2.5" fill="none" stroke-linecap="round"/>'
def f_mouth_grin(x, y, ink, w=12):
    W = w * 1.3
    return f'<path d="M{_f(x-W/2)} {_f(y-1)}q{_f(W/2)} {_f(W/1.8)} {_f(W)} 0" stroke="{ink}" stroke-width="2.5" fill="none" stroke-linecap="round"/>'
def f_mouth_smirk(x, y, ink, w=12):  # ニヤッ (左端は緩く下げ、右端だけ吊り上げる非対称の笑み)
    W = w * 1.3
    return (f'<path d="M{_f(x-W/2)} {_f(y+1.5)}q{_f(W*0.6)} {_f(W/2.6)} {_f(W)} -{_f(W/5)}" stroke="{ink}" stroke-width="2.5" fill="none" stroke-linecap="round"/>'
            f'<path d="M{_f(x+W/2)} {_f(y+1.5-W/5)}l{_f(W/7)} -{_f(W/8)}" stroke="{ink}" stroke-width="2.5" fill="none" stroke-linecap="round"/>')
TONGUE = '#f4739a'  # 舌はどのキャラでも同じ桃色 (ink だと「舌」に見えない。顔の色ルールの唯一の例外)
def f_mouth_tongue(x, y, ink, w=12):  # 舌出し (上の口線 + 顎の下まで垂れる丸い桃色の舌。線も輪郭も無し)
    tw, th = max(9.0, w / 1.5), max(13.0, w / 1.15); r = tw / 2
    return (f'<path d="M{_f(x-w/2)} {_f(y-1.5)}q{_f(w/2)} {_f(w/3.2)} {_f(w)} 0" stroke="{ink}" stroke-width="2.4" fill="none" stroke-linecap="round"/>'
            f'<path d="M{_f(x-r)} {_f(y+0.5)}v{_f(th-r)}a{_f(r)} {_f(r)} 0 0 0 {_f(tw)} 0v-{_f(th-r)}z" fill="{TONGUE}"/>')
def f_mouth_none(x, y, ink, w=12):
    return ''
MOUTHS = {'line': f_mouth_line, 'smile': f_mouth_smile, 'frown': f_mouth_frown, 'o': f_mouth_o, 'open': f_mouth_open,
          'cat': f_mouth_cat, 'wave': f_mouth_wave, 'grin': f_mouth_grin, 'smirk': f_mouth_smirk, 'tongue': f_mouth_tongue, 'none': f_mouth_none}

def face_svg(fs, ink, mood=None):
    """fs = {eyes, mouth, x, y, gap, r, my, w}。mood を渡すと目と口を差し替える。
    fs['no_mouth']=True のときは表情を変えても口を描かない (くちばし等を body 側に持つキャラ用)。
    fs['keep_eyes']=True のときは表情を変えても目を差し替えない (目そのものがキャラの個性のキャラ用)。
    既定は従来どおり mood が口も差し替えるので、既存キャラの見た目は変わらない。"""
    x = fs.get('x', 40); y = fs.get('y', 40); gap = fs.get('gap', 8); r = fs.get('r', 2.8)
    my = fs.get('my', 11); w = fs.get('w', 12)
    eyes = fs.get('eyes', 'dots'); mouth = fs.get('mouth', 'smile')
    if mood:
        mood_eyes, mouth = MOODS[mood]
        if not fs.get('keep_eyes'): eyes = mood_eyes
        if mood == 'surprised': r = r * 1.15
        if mood == 'sleepy': w = max(6, w * 0.6)
    if fs.get('no_mouth'): mouth = 'none'
    return EYES[eyes](x, y, ink, gap, r) + MOUTHS[mouth](x, y + my, ink, w)
def compose(body, fs, ink, mood=None):
    return svg(body + face_svg(fs, ink, mood))
def face_label(fs):
    e = {'dots': '目2点', 'closed': '閉じ目', 'squint': '「> <」の目', 'wink': 'ウインク', 'peek': '片目見開き', 'arcs': 'にっこり目', 'yen': '¥の目', 'none': '目なし'}[fs.get('eyes', 'dots')]
    m = {'line': '一文字の口', 'smile': '笑い口', 'frown': 'への字', 'o': '「お」の口', 'open': '開いた口', 'cat': 'ωの口', 'wave': '〜の口', 'grin': 'にかっと口', 'smirk': 'ニヤッと口', 'tongue': '舌出し', 'none': '口なし'}[fs.get('mouth', 'smile')]
    return e if m == '口なし' else f'{e} + {m}'


def f_eyes_half(x, y, ink, gap=8, r=2.8):
    """半目 (まぶたの一文字 + 下半分の黒目)。傲慢・眠そう用"""
    w = r * 2.9; sw = max(2, r * 0.8)
    def one(cx):
        return (f'<path d="M{_f(cx-w/2)} {_f(y)}h{_f(w)}" stroke="{ink}" stroke-width="{_f(sw)}" stroke-linecap="round"/>'
                f'<path d="M{_f(cx-r)} {_f(y)}a{_f(r)} {_f(r)} 0 0 0 {_f(2*r)} 0z" fill="{ink}"/>')
    return one(x - gap) + one(x + gap)
def f_eyes_swirl(x, y, ink, gap=8, r=2.8):
    """ぐるぐる目 (目を回している)。外へ広がる 2 回転強の渦"""
    import math
    R = r * 1.75; sw = max(1.3, r * 0.47); turns = 2.25; n = 48
    def one(cx):
        pts = []
        for i in range(n + 1):
            t = i / n; a = t * turns * 2 * math.pi; rr = R * (0.12 + 0.88 * t)
            pts.append(f'{_f(round(cx + rr * math.cos(a), 2))} {_f(round(y + rr * math.sin(a), 2))}')
        return f'<path d="M{" L".join(pts)}" stroke="{ink}" stroke-width="{_f(round(sw, 2))}" fill="none" stroke-linecap="round" stroke-linejoin="round"/>'
    return one(x - gap) + one(x + gap)
EYES['half'] = f_eyes_half
EYES['swirl'] = f_eyes_swirl
