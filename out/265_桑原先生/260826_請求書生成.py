import sys, types
for m in ['cryptography','cryptography.hazmat','cryptography.hazmat.primitives',
          'cryptography.hazmat.primitives.hashes','cryptography.hazmat.bindings',
          'cryptography.hazmat.bindings._rust']:
    sys.modules.setdefault(m, types.ModuleType(m))
from fpdf import FPDF

F = "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"
L, R = 18, 192          # left / right margin x
def yen(n): return f"{n:,}"

p = FPDF(format="A4")
p.add_font("ipa", "", F)
p.set_auto_page_break(False)
p.add_page()
p.set_margins(L, 18, 18)

# ---- title
p.set_font("ipa", "", 22)
p.set_xy(L, 18); p.cell(R-L, 11, "請  求  書", align="C")
p.set_line_width(0.5); p.line(L, 31, R, 31)

# ---- header block
p.set_font("ipa", "", 13)
p.set_xy(L, 39); p.cell(95, 8, "テナントアシスト・ウイン株式会社　御中")
p.set_line_width(0.3); p.line(L, 47.5, L+95, 47.5)
p.set_font("ipa", "", 9)
p.set_xy(L, 51); p.cell(95, 5, "下記のとおりご請求申し上げます。")

x2 = 118
p.set_font("ipa", "", 9)
for i, t in enumerate(["請求日：2026年8月26日", "請求書番号：260826-01"]):
    p.set_xy(x2, 39 + i*5); p.cell(74, 5, t)
p.set_font("ipa", "", 10)
p.set_xy(x2, 51); p.cell(74, 5.5, "KIKUCHIホールディングス株式会社")
p.set_font("ipa", "", 9)
for i, t in enumerate(["東京都江東区東陽5-29-41", "ローズハイム東陽町302",
                       "代表取締役　菊池　研太", "TEL：080-6047-2797",
                       "登録番号：【★要記入】"]):
    p.set_xy(x2, 57 + i*4.8); p.cell(74, 4.8, t)

# ---- total box
y = 86
p.set_line_width(0.6); p.set_fill_color(242, 242, 242)
p.rect(L, y, R-L, 15, style="DF")
p.set_font("ipa", "", 13)
p.set_xy(L+4, y); p.cell(60, 15, "ご請求金額（税込）")
p.set_font("ipa", "", 17)
p.set_xy(L, y); p.cell(R-L-6, 15, f"￥{yen(12610)}  －", align="R")

# ---- subject
p.set_font("ipa", "", 9.5)
p.set_xy(L, 106)
p.multi_cell(R-L, 5.5, "件名：佐久市 開業用地 現地調査に係る交通費実費（管理番号265／桑原先生 開業支援）")

# ---- items table
cols = [(22, "日付"), (84, "摘要"), (16, "数量"), (26, "単価"), (26, "金額")]
y = 118
p.set_font("ipa", "", 9); p.set_line_width(0.25); p.set_fill_color(232, 232, 232)
x = L
for w, h in cols:
    p.rect(x, y, w, 8, style="DF"); p.set_xy(x, y); p.cell(w, 8, h, align="C"); x += w

rows = [("8月6日", "新幹線運賃　東京→佐久平（はくたか555号・自由席）", "E57977", 6040),
        ("8月6日", "新幹線運賃　佐久平→東京（あさま616号・指定席）", "E10541", 6570)]
y += 8
for d, desc, ref, amt in rows:
    h = 12
    x = L
    for w, _ in cols:
        p.rect(x, y, w, h); x += w
    p.set_font("ipa", "", 9)
    p.set_xy(L, y); p.cell(cols[0][0], h, d, align="C")
    p.set_xy(L+cols[0][0]+2, y+1.5); p.cell(80, 5, desc)
    p.set_font("ipa", "", 7.5)
    p.set_xy(L+cols[0][0]+2, y+6); p.cell(80, 5, f"えきねっと予約番号 {ref}")
    p.set_font("ipa", "", 9)
    p.set_xy(L+cols[0][0]+cols[1][0], y); p.cell(cols[2][0], h, "1", align="C")
    p.set_xy(L+cols[0][0]+cols[1][0]+cols[2][0]-3, y); p.cell(cols[3][0], h, yen(amt), align="R")
    p.set_xy(L+cols[0][0]+cols[1][0]+cols[2][0]+cols[3][0]-3, y); p.cell(cols[4][0], h, yen(amt), align="R")
    y += h

for _ in range(2):                       # blank rows
    x = L
    for w, _ in cols:
        p.rect(x, y, w, 8); x += w
    y += 8

for label, val, bold in [("小計（税込実費）", 12610, False), ("合計", 12610, True)]:
    p.rect(L, y, cols[0][0]+cols[1][0]+cols[2][0]+cols[3][0], 9, style="DF")
    p.rect(L+cols[0][0]+cols[1][0]+cols[2][0]+cols[3][0], y, cols[4][0], 9)
    p.set_font("ipa", "", 10 if bold else 9.5)
    p.set_xy(L, y); p.cell(cols[0][0]+cols[1][0]+cols[2][0]+cols[3][0]-4, 9, label, align="R")
    p.set_xy(L+cols[0][0]+cols[1][0]+cols[2][0]+cols[3][0]-3, y); p.cell(cols[4][0], 9, yen(val), align="R")
    y += 9

# ---- notes
y += 8
p.set_line_width(0.3)
p.rect(L, y, R-L, 46)
p.set_font("ipa", "", 9.5)
p.set_xy(L+3, y+2); p.cell(100, 5, "【お振込先】")
p.set_font("ipa", "", 9)
p.set_xy(L+3, y+7);  p.cell(160, 5, "金融機関名：【★要記入】　　支店名：【★要記入】")
p.set_xy(L+3, y+12); p.cell(160, 5, "種別・口座番号：【★要記入】　　口座名義：【★要記入】")
p.set_font("ipa", "", 9.5)
p.set_xy(L+3, y+18); p.cell(100, 5, "【備考】")
p.set_font("ipa", "", 8.5)
notes = ["・2026年8月6日、貴社ご依頼（2026年7月24日付）による佐久市・御代田町の開業用地 現地調査の交通費実費です。",
         "・上記は鉄道運賃の実費（えきねっと決済額）であり、消費税相当額は運賃に含まれます。",
         "・えきねっとの利用明細を別途ご提示可能です。",
         "・振込手数料は貴社にてご負担をお願いいたします。"]
for i, t in enumerate(notes):
    p.set_xy(L+3, y+23.5+i*5); p.cell(R-L-6, 5, t)

p.output("260826_請求書（TAW）交通費.pdf")
print("done")
