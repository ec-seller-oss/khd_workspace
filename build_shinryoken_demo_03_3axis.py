# -*- coding: utf-8 -*-
"""
診療圏デモ画像3：立地を数字で見る3軸（①診療圏人口 ②競合 ③動線）
企画#2「腕より立地」の 2/5 用。デザインシステム=クリーム白#F9F6EF×レンガ赤#AA2E26
出力: shinryoken_demo_03_3axis.png (1600x900)
"""
from PIL import Image, ImageDraw, ImageFont

CREAM = (249, 246, 239)
BRICK = (170, 46, 38)
DARK = (43, 43, 43)
MUTED = (138, 129, 122)
WHITE = (255, 255, 255)
LINE = (224, 216, 204)
CARDBG = (255, 255, 255)
NUMBG = (245, 224, 220)

W6 = "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc"
W3 = "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc"

def f(bold, size):
    return ImageFont.truetype(W6 if bold else W3, size)

def text(d, xy, s, font, fill, anchor="la"):
    d.text(xy, s, font=font, fill=fill, anchor=anchor)

def build(path):
    Wd, Ht = 1600, 900
    img = Image.new("RGB", (Wd, Ht), CREAM)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 14, Ht], fill=BRICK)

    # 見出し
    text(d, (70, 60), "立地を数字で見る、3つの軸", f(True, 54), DARK)
    text(d, (70, 138), "この3つを掛けると「1日に何人来うるか」が概算で出る。なんとなくの好立地は、数字で崩れる。", f(False, 26), MUTED)
    d.line([70, 196, Wd - 70, 196], fill=LINE, width=2)

    cards = [
        ("①", "診療圏人口", "年齢別人口 × 受療率", "その街に「患者になりうる人」が\n何人いるかを年齢層別に積み上げる"),
        ("②", "競合", "数 と 距離（半径2km）", "同じ科が近くにあると\n競合按分で取り分が一気に減る"),
        ("③", "動線", "通勤・生活の人の流れ", "毎日通る導線上にあるか。\n認知されやすい場所が強い"),
    ]

    cw, ch = 440, 470
    gap = 40
    total = cw * 3 + gap * 2
    x0 = (Wd - total) // 2
    top = 250

    for i, (num, title, sub, desc) in enumerate(cards):
        x = x0 + i * (cw + gap)
        # カード
        d.rounded_rectangle([x, top, x + cw, top + ch], radius=24, fill=CARDBG, outline=LINE, width=3)
        # 番号バッジ
        bx, by, br = x + cw / 2, top + 96, 56
        d.ellipse([bx - br, by - br, bx + br, by + br], fill=NUMBG)
        text(d, (bx, by), num, f(True, 64), BRICK, anchor="mm")
        # タイトル
        text(d, (x + cw / 2, top + 210), title, f(True, 40), DARK, anchor="ma")
        # サブ（式）
        text(d, (x + cw / 2, top + 272), sub, f(True, 26), BRICK, anchor="ma")
        d.line([x + 50, top + 322, x + cw - 50, top + 322], fill=LINE, width=2)
        # 説明（2行）
        for j, ln in enumerate(desc.split("\n")):
            text(d, (x + cw / 2, top + 350 + j * 40), ln, f(False, 25), DARK, anchor="ma")

    # かけ算記号（カード間）
    for i in range(2):
        mx = x0 + cw + gap / 2 + i * (cw + gap)
        text(d, (mx, top + ch / 2), "×", f(True, 60), MUTED, anchor="mm")

    # 下部メッセージ
    text(d, (Wd / 2, top + ch + 48), "＝ 1日に来うる延べ外来（概算）", f(True, 38), DARK, anchor="ma")

    text(d, (70, Ht - 54), "※ イメージ図。実数は候補地ごとの診療圏調査で算出。宅建士が物件・契約までワンストップで対応。", f(False, 22), MUTED)
    img.save(path, "PNG")
    print("saved", path)

if __name__ == "__main__":
    build("/Users/kikuchikenta/01_honbu_docs_automation/shinryoken_demo_03_3axis.png")
