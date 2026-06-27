# -*- coding: utf-8 -*-
"""ゆーし専用クイックガイド（LINE送付用・縦長PNG）KHDクリーム白×レンガ赤"""
from PIL import Image, ImageDraw, ImageFont

W, Hh = 1080, 1660
BG=(0xF9,0xF6,0xEF); RED=(0xAA,0x2E,0x26); REDD=(0x8C,0x24,0x1D)
INK=(0x1A,0x1A,0x1A); GRY=(0x6E,0x6E,0x6E); CARD=(0xF1,0xEC,0xE1)
CARDLN=(0xE1,0xDA,0xCB); REDBG=(0xF4,0xE4,0xE2); WHT=(255,255,255)
FB="/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc"   # bold
FH="/System/Library/Fonts/ヒラギノ角ゴシック W8.ttc"   # heavy
FR="/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc"   # regular
def f(path,sz): return ImageFont.truetype(path,sz,index=0)

img=Image.new("RGB",(W,Hh),BG); d=ImageDraw.Draw(img)

# ── ヘッダー（赤帯）──
d.rectangle([0,0,W,250],fill=RED)
d.text((60,58),"パソコンで クロードを使う",font=f(FH,66),fill=WHT)
d.text((62,152),"ゆーし専用 かんたんガイド",font=f(FB,40),fill=(0xF2,0xD8,0xD6))
# サブ一言
d.text((60,288),"むずかしいことは 考えなくてOK 🙆",font=f(FB,40),fill=REDD)

# ── 4ステップ ──
steps=[
 ("クロームを ひらく","カラフルな丸いマークのアプリ"),
 ("「母艦Mac3」を おす","さいしょだけ数字パスワード（菊池さんに聞く）"),
 ("いつもの クロードに 話しかける","「○○して」と ふつうに書くだけ"),
 ("おわったら まどを 閉じる","ほぞんは クロードがやるので 何もしなくてOK"),
]
y=370; ch=210; gap=28; x0=50; x1=W-50
for i,(main,sub) in enumerate(steps):
    yy=y+(ch+gap)*i
    d.rounded_rectangle([x0,yy,x1,yy+ch],radius=22,fill=CARD,outline=CARDLN,width=2)
    d.rounded_rectangle([x0,yy,x0+12,yy+ch],radius=0,fill=RED)
    # 番号の丸
    cx,cy,r=x0+115,yy+ch//2,58
    d.ellipse([cx-r,cy-r,cx+r,cy+r],fill=RED)
    d.text((cx,cy-4),str(i+1),font=f(FH,64),fill=WHT,anchor="mm")
    # 本文
    tx=x0+215
    d.text((tx,yy+48),main,font=f(FB,46),fill=INK)
    d.text((tx,yy+122),sub,font=f(FR,32),fill=GRY)

# ── 安心ボックス ──
by=y+(ch+gap)*4+8
d.rounded_rectangle([x0,by,x1,by+250],radius=22,fill=REDBG)
d.rounded_rectangle([x0,by,x0+12,by+250],radius=0,fill=RED)
lines=[
 "こわれません。まちがえても 大丈夫",
 "むずかしい設定は ぜんぶ菊池さんが やりました",
 "こまったら 菊池さんに LINE",
]
for j,ln in enumerate(lines):
    ly=by+38+j*68
    d.text((x0+45,ly-4),"✓",font=f(FH,40),fill=RED)
    d.text((x0+110,ly),ln,font=f(FB,38),fill=REDD if j==2 else INK)

# ── フッター ──
d.line([50,Hh-78,W-50,Hh-78],fill=CARDLN,width=2)
d.text((60,Hh-58),"KHD ｜ クロード かんたんガイド ｜ 菊池研太",font=f(FR,28),fill=GRY)

out="/Users/kikuchikenta/01_honbu_docs_automation/yuushi_quickguide.png"
img.save(out,"PNG")
print("saved:",out)
