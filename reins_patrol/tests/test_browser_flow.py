"""patrol.py のブラウザ操作をモック画面で通しで実行する。
REINSには到達できないので、同じ構造(ID/PW+ガイドライン同意→売買物件検索→
検索条件を表示→保存済み検索select→読込+confirm→検索→タブ→ページ内全選択→
図面一括取得→一括取得でPDFダウンロード)のローカルHTMLを相手に、
login/open_saved_search/download_tab_pdfs のコード自体が動くかを検証する。"""
import http.server, os, shutil, socketserver, sys, tempfile, threading
from pathlib import Path

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.pagesizes import A4

HERE = Path(__file__).resolve().parent
SITE = HERE / "mock_reins"
TMP = Path(tempfile.mkdtemp())
work = TMP / "reins"; work.mkdir()
# 隔離コピーを読む（patrol.py は自分の隣に downloads/ state/ を作るため）
shutil.copy(HERE.parent / "patrol.py", work / "patrol.py")
sys.path.insert(0, str(work))
import patrol

# --- モック用の図面PDFを2本作る（タブごと） ---
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
TABS = {
    0: [{"物件番号": "300111222333", "価格": "6,480", "所在地": "東京都墨田区東向島３丁目",
         "交通": "東武伊勢崎線 東向島 徒歩6分", "土地面積": "78.55", "建物面積": "102.34",
         "築年月": "2019年3月", "構造": "木造2階建", "取引態様": "専任媒介", "現況": "空家"},
        {"物件番号": "300444555666", "価格": "9,800", "所在地": "東京都墨田区京島１丁目",
         "交通": "京成押上線 京成曳舟 徒歩4分", "土地面積": "95.00", "建物面積": "120.00",
         "築年月": "2015年8月", "構造": "木造3階建", "取引態様": "一般媒介", "現況": "居住中"}],
    1: [{"物件番号": "300777888999", "価格": "12,800", "所在地": "神奈川県横浜市旭区柏町１−２",
         "交通": "相鉄本線 三ツ境 徒歩8分", "土地面積": "210.10", "建物面積": "198.76",
         "築年月": "1998年11月", "構造": "鉄骨造3階建", "取引態様": "専任媒介", "現況": "空家"}],
}
for i, recs in TABS.items():
    c = canvas.Canvas(str(SITE / f"zumen_tab{i}.pdf"), pagesize=A4)
    for rec in recs:
        c.setFont("HeiseiKakuGo-W5", 11); y = 800
        c.drawString(40, y, "レインズ 図面")
        for k, v in rec.items():
            y -= 22
            u = {"価格": " 万円", "土地面積": "㎡", "建物面積": "㎡"}.get(k, "")
            c.drawString(40, y, f"{k}：{v}{u}")
        c.showPage()
    c.save()

# --- ローカルHTTPサーバ（file://だとダウンロード挙動が変わるため） ---
os.chdir(SITE)
class Quiet(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *a): pass
srv = socketserver.TCPServer(("127.0.0.1", 0), Quiet)
port = srv.server_address[1]
threading.Thread(target=srv.serve_forever, daemon=True).start()
base = f"http://127.0.0.1:{port}"
print(f"mock REINS: {base}")

# --- patrol の向き先だけ差し替える（コード本体は一切いじらない） ---
patrol.LOGIN_URL = f"{base}/login.html"
patrol.USER_ID = "a0303xxxxxx21"
patrol.PASSWORD = "dummy-pass"
LABEL = "KHD03:墨田区 戸建 3LDK 〜6500万"

from playwright.sync_api import sync_playwright
fail = 0
with sync_playwright() as pw:
    # 母艦では playwright install 済みのchromiumが自動で使われる。
    # 別の場所のchromiumを使いたい時だけ PW_CHROMIUM に絶対パスを入れる。
    launch_kw = {"headless": True, "args": patrol.LAUNCH_ARGS}
    if os.environ.get("PW_CHROMIUM"):
        launch_kw["executable_path"] = os.environ["PW_CHROMIUM"]
    browser = pw.chromium.launch(**launch_kw)
    ctx = browser.new_context(user_agent=patrol.UA, locale="ja-JP",
                              viewport={"width": 1440, "height": 900},
                              accept_downloads=True)
    ctx.add_init_script(patrol.INIT_JS)
    page = ctx.new_page()

    print("\n=== login() ===")
    patrol.login(page)
    if "menu.html" not in page.url:
        print(f"   ❌ ログイン後に遷移していない: {page.url}"); fail += 1
    else:
        print("   ✅ ID/PW入力・ガイドライン同意・ログインボタン → メニュー到達")

    # ステルス設定が効いているか（実サイトの検知回避の要）
    wd = page.evaluate("() => navigator.webdriver")
    langs = page.evaluate("() => navigator.languages")
    print(f"   navigator.webdriver={wd} / languages={langs}")
    if wd is not None or langs[0] != "ja-JP":
        print("   ❌ init script が効いていない"); fail += 1

    print("\n=== open_saved_search() ===")
    patrol.open_saved_search(page, LABEL)
    if "results.html" not in page.url:
        print(f"   ❌ 検索結果に到達していない: {page.url}"); fail += 1
    else:
        print("   ✅ 売買物件検索→条件展開→保存済み検索選択→読込(confirm)→検索")

    print("\n=== download_tab_pdfs() ===")
    pdfs = []
    for i in range(2):
        pdfs += patrol.download_tab_pdfs(page, i)
    print(f"   取得PDF {len(pdfs)}本: {[p.name for p in pdfs]}")
    if len(pdfs) != 2:
        print("   ❌ 2タブ分のPDFが落ちていない"); fail += 1

    print("\n=== 存在しないタブ #5 ===")
    if patrol.download_tab_pdfs(page, 5) != []:
        print("   ❌ 存在しないタブで空配列を返していない"); fail += 1
    else:
        print("   ✅ スキップして継続")

    print("\n=== 保存済み検索名が間違っている場合 ===")
    page.goto(f"{base}/search.html")
    try:
        patrol.open_saved_search(page, "存在しない検索名")
        print("   ❌ 例外が出ていない"); fail += 1
    except RuntimeError as e:
        head = str(e).splitlines()[0]
        print(f"   ✅ RuntimeError: {head}")
        if "KHD01" not in str(e):
            print("   ❌ 画面上の候補一覧が示されていない"); fail += 1

    ctx.close(); browser.close()

# --- 落ちたPDFを解析して差分まで通す ---
print("\n=== parse_pdf → diff_new（通し） ===")
items = []
for p in pdfs:
    items += patrol.parse_pdf(p)
print(f"   解析 {len(items)}件: {[r.get('物件番号') for r in items]}")
if len(items) != 3:
    print("   ❌ 3件にならない"); fail += 1
fresh = patrol.diff_new(items, LABEL)
print(f"   1回目の新着 {len(fresh)}件（期待3）")
if len(fresh) != 3: fail += 1
fresh2 = patrol.diff_new(items, LABEL)
print(f"   2回目の新着 {len(fresh2)}件（期待0）")
if len(fresh2) != 0: fail += 1

srv.shutdown()
print("\n" + ("🎉 ブラウザ通し 全テスト通過" if fail == 0 else f"❌ 失敗 {fail}件"))
sys.exit(1 if fail else 0)
