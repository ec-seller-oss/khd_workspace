"""patrol.py のオフライン検証：PDF解析(FIELD_RE)と新着差分(diff_new)。
ブラウザ部分はREINSに到達できないので触らない。"""
import sys, os, json, shutil, tempfile
from pathlib import Path

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.pagesizes import A4

HERE = Path(__file__).resolve().parent
TMP = Path(tempfile.mkdtemp())
# patrol.py を隔離環境で読み込む（ROOT配下にstate等を作るのでコピーして使う）
work = TMP / "reins"
work.mkdir()
shutil.copy(HERE.parent / "patrol.py", work / "patrol.py")
sys.path.insert(0, str(work))
import patrol

pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))

PAGES = [
    {"物件番号": "300123456789", "価格": "6,480", "所在地": "東京都墨田区東向島３丁目",
     "交通": "東武伊勢崎線 東向島 徒歩6分", "土地面積": "78.55", "建物面積": "102.34",
     "築年月": "2019年3月", "構造": "木造2階建", "取引態様": "専任媒介", "現況": "空家",
     "想定利回り": "8.4"},
    {"物件番号": "300987654321", "価格": "12,800", "所在地": "神奈川県横浜市旭区柏町１−２",
     "交通": "相鉄本線 三ツ境 徒歩8分", "土地面積": "210.10", "建物面積": "198.76",
     "築年月": "1998年11月", "構造": "鉄骨造3階建", "取引態様": "一般媒介", "現況": "空家",
     "想定利回り": "11.2"},
]

def make_pdf(path, pages):
    c = canvas.Canvas(str(path), pagesize=A4)
    for rec in pages:
        c.setFont("HeiseiKakuGo-W5", 11)
        y = 800
        c.drawString(40, y, "レインズ 図面")
        for k, v in rec.items():
            y -= 22
            unit = {"価格": " 万円", "土地面積": "㎡", "建物面積": "㎡", "想定利回り": "％"}.get(k, "")
            unit = "%" if k == "想定利回り" else unit
            c.drawString(40, y, f"{k}：{v}{unit}")
        c.showPage()
    c.save()

pdf = TMP / "zumen.pdf"
make_pdf(pdf, PAGES)

fail = 0

# --- 1) PDF解析 ---
items = patrol.parse_pdf(pdf)
print(f"[1] parse_pdf → {len(items)}ページ")
assert len(items) == 2, items
for got, want in zip(items, PAGES):
    for k, v in want.items():
        if got.get(k) != v:
            print(f"   ❌ {k}: 期待 {v!r} / 実際 {got.get(k)!r}")
            fail += 1
    if got["_needs_ai"]:
        print(f"   ❌ _needs_ai が立った（物件番号を拾えていない）")
        fail += 1
if fail == 0:
    print("   ✅ 全11項目 × 2ページ 一致")

# --- 2) 新着差分：1回目は全件新着 ---
fresh1 = patrol.diff_new(items, "KHD_TEST")
print(f"[2] 1回目 diff_new → 新着 {len(fresh1)}件（期待2）")
if len(fresh1) != 2:
    fail += 1

# --- 3) 2回目は同じPDFなら新着0 ---
fresh2 = patrol.diff_new(patrol.parse_pdf(pdf), "KHD_TEST")
print(f"[3] 2回目 diff_new → 新着 {len(fresh2)}件（期待0）")
if len(fresh2) != 0:
    fail += 1

# --- 4) 物件を1件足したら、その1件だけ新着 ---
NEW = dict(PAGES[0]); NEW["物件番号"] = "300555000111"; NEW["所在地"] = "東京都墨田区京島１丁目"
pdf2 = TMP / "zumen2.pdf"
make_pdf(pdf2, PAGES + [NEW])
fresh3 = patrol.diff_new(patrol.parse_pdf(pdf2), "KHD_TEST")
print(f"[4] 1件追加後 → 新着 {len(fresh3)}件（期待1）: "
      f"{[r.get('物件番号') for r in fresh3]}")
if [r.get("物件番号") for r in fresh3] != ["300555000111"]:
    fail += 1

# --- 5) 検索ラベルが違えば記憶は独立 ---
fresh4 = patrol.diff_new(items, "KHD_OTHER")
print(f"[5] 別ラベル → 新着 {len(fresh4)}件（期待2＝ラベルごとに独立）")
if len(fresh4) != 2:
    fail += 1

# --- 6) 出力ファイル生成 ---
txt = patrol.write_output(patrol.parse_pdf(pdf2), fresh3, "KHD_TEST")
body = Path(txt).read_text(encoding="utf-8")
print(f"[6] write_output →\n" + "\n".join("      " + l for l in body.splitlines()[:9]))
if "300555000111" not in body or "新着 1件" not in body:
    print("   ❌ 出力に新着物件が載っていない")
    fail += 1

# --- 7) メンテ時間判定 ---
import datetime as dt
class FakeDT(dt.datetime):
    _now = None
    @classmethod
    def now(cls, tz=None): return cls._now
real = patrol.datetime
for h, want in [(23, True), (2, True), (6, True), (7, False), (10, False), (22, False)]:
    FakeDT._now = FakeDT(2026, 8, 13, h, 30)
    patrol.datetime = FakeDT
    got = patrol.in_maintenance()
    if got != want:
        print(f"   ❌ {h}時: 期待 {want} / 実際 {got}")
        fail += 1
patrol.datetime = real
print("[7] in_maintenance 23-7時判定 ✅" if fail == 0 else "[7] 判定に失敗あり")

# --- 8) .env のガード（別プロセスで起動して終了コードを見る） ---
import subprocess
def run_with_env(env_body):
    (work / ".env").write_text(env_body, encoding="utf-8")
    return subprocess.run([sys.executable, str(work / "patrol.py"), "--dry-run"],
                          capture_output=True, text=True, timeout=120)

CASES = [
    ("IDもPWも無い", "SAVED_SEARCH_LABEL=X\n", "REINS_USER_ID / REINS_PASSWORD がありません"),
    ("保存済み検索名が無い", "REINS_USER_ID=a1\nREINS_PASSWORD=p\n", "SAVED_SEARCH_LABEL がありません"),
    ("雛形の説明文が残っている",
     "REINS_USER_ID=（自分がログイン権限を持つ業者ID）\nREINS_PASSWORD=p\nSAVED_SEARCH_LABEL=X\n",
     "雛形のままです"),
]
print("\n[8] .envガード")
for name, body, expect in CASES:
    r = run_with_env(body)
    ok = r.returncode == 1 and expect in r.stdout
    print(f"   {'✅' if ok else '❌'} {name} → rc={r.returncode}")
    if not ok:
        print(f"      期待: {expect!r} / 実際: {r.stdout.strip()[:200]}")
        fail += 1

print("\n" + ("🎉 全テスト通過" if fail == 0 else f"❌ 失敗 {fail}件"))
sys.exit(1 if fail else 0)
