# -*- coding: utf-8 -*-
"""
診療圏調査デモ画像 2点 生成（X/note/YouTube 横断再利用）
デザインシステム: クリーム白 #F9F6EF × レンガ赤 #AA2E26
画像1: 受療率積み上げ（年齢別女性人口×受療率→約21人/日）
画像2: 競合按分 Before/After（21人/日 → 4〜6人/日）
"""
from PIL import Image, ImageDraw, ImageFont

# --- 配色 ---
CREAM = (249, 246, 239)
BRICK = (170, 46, 38)
DARK = (43, 43, 43)
MUTED = (138, 129, 122)
WHITE = (255, 255, 255)
LINE = (224, 216, 204)
# 積み上げセグメント（暖色グラデ）
SEG = [(205, 184, 154), (224, 176, 128), (217, 140, 95), (199, 91, 74), (170, 46, 38)]

W6 = "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc"
W3 = "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc"

def f(bold, size):
    return ImageFont.truetype(W6 if bold else W3, size)

def text(d, xy, s, font, fill, anchor="la"):
    d.text(xy, s, font=font, fill=fill, anchor=anchor)

def tw(d, s, font):
    b = d.textbbox((0, 0), s, font=font)
    return b[2] - b[0]

# ===================== 画像1: 受療率積み上げ =====================
def build_image1(path):
    Wd, Ht = 1600, 900
    img = Image.new("RGB", (Wd, Ht), CREAM)
    d = ImageDraw.Draw(img)

    # 左サイドのアクセントバー
    d.rectangle([0, 0, 14, Ht], fill=BRICK)

    # 見出し
    text(d, (70, 64), "「人口 × 受療率」で1日の外来を積み上げる", f(True, 52), DARK)
    text(d, (70, 138), "年齢別の女性人口に受療率（全国患者調査）を掛け、昼夜補正0.84で1日の延べ外来を算出", f(False, 26), MUTED)
    d.line([70, 196, Wd - 70, 196], fill=LINE, width=2)

    # 積み上げデータ
    segs = [
        ("15〜19歳", 1),
        ("20〜29歳", 6),
        ("30〜39歳", 9),
        ("40〜49歳", 4),
        ("50歳以上", 1),
    ]
    total = sum(v for _, v in segs)  # 21

    # 積み上げ棒（縦）
    bx = 180
    bw = 280
    top = 260
    bottom = 770
    bh = bottom - top
    cur = bottom
    for i, (label, val) in enumerate(segs):
        seg_h = bh * val / total
        y0 = cur - seg_h
        d.rectangle([bx, y0, bx + bw, cur], fill=SEG[i])
        # セグメント内ラベル
        cy = (y0 + cur) / 2
        lbl = f"{label}  {val}人"
        col = WHITE if i >= 3 else DARK
        text(d, (bx + bw / 2, cy), lbl, f(True, 26), col, anchor="mm")
        cur = y0
    # 棒の枠
    d.rectangle([bx, top, bx + bw, bottom], outline=LINE, width=2)
    text(d, (bx + bw / 2, bottom + 28), "年齢層別の積み上げ", f(False, 24), MUTED, anchor="ma")

    # 右側：合計コールアウト
    cx = 880
    text(d, (cx, 300), "積み上げ合計（昼夜補正後）", f(False, 30), DARK)
    text(d, (cx, 350), "約", f(True, 60), DARK)
    text(d, (cx + 90, 330), "21", f(True, 150), BRICK)
    text(d, (cx + 340, 350), "人 / 日", f(True, 60), DARK)
    text(d, (cx, 540), "→ 人口だけ見れば「十分やれる立地」に見える。", f(True, 32), DARK)
    text(d, (cx, 600), "　　でも、これは競合を入れる前の数字。", f(False, 30), MUTED)

    # 注記
    text(d, (70, Ht - 56), "※ イメージ図。実数は候補地ごとの診療圏調査で算出（年齢別人口×受療率×昼夜補正）。", f(False, 22), MUTED)
    img.save(path, "PNG")
    print("saved", path)

# ===================== 画像2: 競合按分 Before/After =====================
def build_image2(path):
    Wd, Ht = 1600, 900
    img = Image.new("RGB", (Wd, Ht), CREAM)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 14, Ht], fill=BRICK)

    text(d, (70, 64), "人口は多いのに「勝てる立地」とは限らない", f(True, 52), DARK)
    text(d, (70, 138), "半径2km内に既存の産婦人科が1軒 → 競合按分で新規の取り分は一気に圧縮", f(False, 26), MUTED)
    d.line([70, 196, Wd - 70, 196], fill=LINE, width=2)

    # Before カード
    def card(x, y, w, h, fill, outline):
        d.rounded_rectangle([x, y, x + w, y + h], radius=24, fill=fill, outline=outline, width=3)

    cw, ch = 520, 420
    by = 300
    # Before
    bx = 110
    card(bx, by, cw, ch, WHITE, LINE)
    text(d, (bx + cw / 2, by + 50), "Before（人口ベース）", f(True, 34), MUTED, anchor="mm")
    text(d, (bx + cw / 2, by + 200), "21", f(True, 200), MUTED, anchor="mm")
    text(d, (bx + cw / 2, by + 350), "人 / 日", f(True, 44), MUTED, anchor="mm")

    # 矢印
    ax0 = bx + cw + 40
    ax1 = ax0 + 160
    ay = by + ch / 2
    d.line([ax0, ay, ax1 - 30, ay], fill=BRICK, width=14)
    d.polygon([(ax1, ay), (ax1 - 40, ay - 30), (ax1 - 40, ay + 30)], fill=BRICK)
    text(d, ((ax0 + ax1) / 2, ay - 70), "競合按分", f(True, 28), BRICK, anchor="mm")
    text(d, ((ax0 + ax1) / 2, ay + 60), "競合1軒", f(False, 24), MUTED, anchor="mm")

    # After（強調）
    axx = ax1 + 40
    card(axx, by, cw, ch, BRICK, BRICK)
    text(d, (axx + cw / 2, by + 50), "After（競合按分後）", f(True, 34), (245, 220, 215), anchor="mm")
    text(d, (axx + cw / 2, by + 200), "4〜6", f(True, 160), WHITE, anchor="mm")
    text(d, (axx + cw / 2, by + 350), "人 / 日（月100〜160人）", f(True, 36), WHITE, anchor="mm")

    # 下部メッセージ
    text(d, (Wd / 2, by + ch + 70), "立地は勘じゃなく数字で勝てる。競合まで数値化して初めて判断できる。", f(True, 34), DARK, anchor="ma")

    text(d, (70, Ht - 56), "※ イメージ図。実数は候補地ごとの診療圏調査で算出。宅建士が物件・契約までワンストップで対応。", f(False, 22), MUTED)
    img.save(path, "PNG")
    print("saved", path)

if __name__ == "__main__":
    build_image1("/Users/kikuchikenta/01_honbu_docs_automation/shinryoken_demo_01_stacking.png")
    build_image2("/Users/kikuchikenta/01_honbu_docs_automation/shinryoken_demo_02_competition.png")
