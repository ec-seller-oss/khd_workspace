"""
京橋クリニック 公式LINE リッチメニュー（本物・LINE仕様 2500x1686px / 3列x2行）
そのままLINE公式アカウントManagerにアップロード可能。
クリーム白×レンガ赤（KHDデザイン）。アンケートの主因に直結した6タイル。
出力: kyobashi_richmenu.png
"""
from PIL import Image, ImageDraw, ImageFont

W, H = 2500, 1686
BG   = (0xF9, 0xF6, 0xEF)
RED  = (0xAA, 0x2E, 0x26)
REDD = (0x8C, 0x24, 0x1D)
INK  = (0x1A, 0x1A, 0x1A)
GRY  = (0x6E, 0x6E, 0x6E)
LN   = (0xDA, 0xD6, 0xCF)
CARD = (0xF1, 0xEC, 0xE1)
WHT  = (0xFF, 0xFF, 0xFF)

FB = "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc"   # bold
FM = "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc"   # regular
def f(path, sz): return ImageFont.truetype(path, sz, index=0)

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)

# 上部アクセント帯（クリニック名）
HEAD = 0  # メニューはグリッドのみ（ヘッダー帯はLINEトーク側に出るため最小に）

cols, rows = 3, 2
cw, ch = W // cols, H // rows

tiles = [
    ("01", "順番・待ち時間", "当日の順番をLINEで確認"),
    ("02", "健診・各種予約", "健診・化学物質ほか"),
    ("03", "Web問診", "来院前にスマホで記入"),
    ("04", "診療時間・休診", "月火木金 9:30-18:00／土日祝休"),
    ("05", "アクセス・電話", "地図・お問い合わせ"),
    ("06", "よくある質問", "本人確認・受付ルール 他"),
]

f_no    = f(FB, 70)
f_title = f(FB, 92)
f_sub   = f(FM, 50)

def center_x(draw, text, font, box_x, box_w):
    bb = draw.textbbox((0, 0), text, font=font)
    tw = bb[2] - bb[0]
    return box_x + (box_w - tw) // 2

for i, (no, title, sub) in enumerate(tiles):
    r, c = divmod(i, cols)
    x0, y0 = c * cw, r * ch
    # タイル枠（交互の地色でカード感）
    fill = CARD if (i % 2 == 0) else BG
    d.rectangle([x0, y0, x0 + cw, y0 + ch], fill=fill)
    # 赤い番号バッジ（円）
    bcx, bcy, br = x0 + 130, y0 + 130, 56
    d.ellipse([bcx - br, bcy - br, bcx + br, bcy + br], fill=RED)
    nb = d.textbbox((0, 0), no, font=f_no)
    d.text((bcx - (nb[2]-nb[0]) // 2, bcy - (nb[3]-nb[1]) // 2 - nb[1]), no, font=f_no, fill=WHT)
    # タイトル（中央寄せ）
    tx = center_x(d, title, f_title, x0, cw)
    d.text((tx, y0 + ch // 2 - 70), title, font=f_title, fill=INK)
    # サブ（中央寄せ）
    sx = center_x(d, sub, f_sub, x0, cw)
    d.text((sx, y0 + ch // 2 + 60), sub, font=f_sub, fill=GRY)

# 罫線（タイル区切り：レンガ赤の細線）
for c in range(1, cols):
    d.line([(c * cw, 40), (c * cw, H - 40)], fill=LN, width=4)
for r in range(1, rows):
    d.line([(40, r * ch), (W - 40, r * ch)], fill=LN, width=4)
# 外枠
d.rectangle([0, 0, W - 1, H - 1], outline=RED, width=10)

img.save("kyobashi_richmenu.png", "PNG")
print("saved kyobashi_richmenu.png", img.size)
