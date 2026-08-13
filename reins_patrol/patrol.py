#!/usr/bin/env python3
"""
KHD REINS 巡回スクリプト（自社実装）
================================================================================
保存済み検索を実行 → 図面PDFを一括取得 → 物件情報を構造化 → 新着だけ通知。

伊東さんのツールに依存しない自前実装。仕様のみ参考にし、コードは共有していない。

使い方:
    python3 patrol.py --headed --pause-after-login   # 初回：各ステップで止めて目視
    python3 patrol.py --headed                        # ブラウザ表示で通し実行
    python3 patrol.py                                 # ヘッドレス（cron用）
    python3 patrol.py --dry-run                       # ログインまでで止める

終了コード:
    0 = 正常  /  5 = REINSメンテナンス時間帯（正常スキップ）  /  1 = エラー
"""
from __future__ import annotations

import argparse
import json
import os
import random
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime, time as dtime
from pathlib import Path

import pdfplumber
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout, Page

ROOT = Path(__file__).resolve().parent
DL_DIR = ROOT / "downloads"
OUT_DIR = ROOT / "output"
LOG_DIR = ROOT / "logs"
STATE = ROOT / "state" / "seen.json"     # 既出物件番号の記憶（新着差分用）
for d in (DL_DIR, OUT_DIR, LOG_DIR, STATE.parent):
    d.mkdir(parents=True, exist_ok=True)

load_dotenv(ROOT / ".env")
USER_ID = os.environ.get("REINS_USER_ID", "")
PASSWORD = os.environ.get("REINS_PASSWORD", "")
SAVED_SEARCH = os.environ.get("SAVED_SEARCH_LABEL", "")
NOTIFY_EMAIL = os.environ.get("NOTIFY_EMAIL", "").strip()

LOGIN_URL = "https://system.reins.jp/login/main/KG/GKG001200"

# ---------------------------------------------------------------------------
# セレクタは全部ここ。画面が変わったらこの dict だけ直す。
# ※ボタン名は東日本レインズの表示に合わせてある。初回は --headed で必ず目視すること。
# ---------------------------------------------------------------------------
SEL = {
    "user_id":      "input[name*='userId' i], input[name*='login' i], input[type='text']",
    "password":     "input[type='password']",
    "agree_label":  "ガイドラインを遵守",          # チェックボックスのラベル文言
    "login_btn":    "ログイン",
    "menu_baibai":  r"売買\s*物件検索",
    "expand_cond":  r"検索条件を表示",
    "load_btn":     r"読[込み]+",
    "search_btn":   "検索",
    "tab_pattern":  r"売外全|売外一|住宅以外建物全部|住宅以外建物一部",
    "select_all":   r"ページ内全選択",
    "zumen_batch":  r"図面一括取得",
    "download_btn": r"^一括取得",
    "maintenance":  ("サービス提供時間外", "定期メンテナンス"),
}

# 検知リスク低減（伊東さん運用を踏襲）。偽装は"足さない"—目立つため。
LAUNCH_ARGS = ["--disable-blink-features=AutomationControlled"]
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36")
INIT_JS = """
Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
Object.defineProperty(navigator, 'languages', { get: () => ['ja-JP','ja','en-US','en'] });
Object.defineProperty(navigator, 'plugins', { get: () => [
  {name:'PDF Viewer'},{name:'Chrome PDF Viewer'},{name:'Chromium PDF Viewer'}
]});
window.chrome = window.chrome || { runtime: {} };
"""

PAUSE = False   # --pause-after-login で True


def log(msg: str) -> None:
    print(f"[{datetime.now().isoformat(timespec='seconds')}] {msg}", flush=True)


def pause(label: str) -> None:
    if PAUSE:
        input(f"    ⏸  {label} … Enterで次へ")


def human(a: float = 0.6, b: float = 1.8) -> None:
    """機械的な等間隔アクセスを避けるための小さな揺らぎ。"""
    time.sleep(random.uniform(a, b))


def shot(page: Page, label: str) -> Path | None:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    p = LOG_DIR / f"{ts}_{label}.png"
    try:
        page.screenshot(path=str(p), full_page=True)
        log(f"    screenshot: {p.name}")
        return p
    except Exception as e:
        log(f"    screenshot失敗: {e}")
        return None


def in_maintenance() -> bool:
    """毎日23:00〜翌7:00はREINSメンテナンス。叩きに行かない。"""
    now = datetime.now().time()
    return now >= dtime(23, 0) or now < dtime(7, 0)


class Maintenance(Exception):
    pass


# ---------------------------------------------------------------------------
# ブラウザ操作
# ---------------------------------------------------------------------------
def login(page: Page) -> None:
    log("STEP 1: ログインページ")
    page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
    human()

    body = page.locator("body").inner_text(timeout=10000)
    if any(k in body for k in SEL["maintenance"]):
        shot(page, "maintenance")
        raise Maintenance("REINSメンテナンス時間帯")

    log("STEP 2: ID/PW入力")
    uid = page.locator(SEL["user_id"]).first
    uid.wait_for(state="visible", timeout=20000)
    uid.fill(USER_ID)
    human(0.3, 0.7)
    page.locator(SEL["password"]).first.fill(PASSWORD)
    human(0.3, 0.7)

    log("STEP 3: ガイドライン遵守チェック")
    try:
        page.locator("label", has_text=re.compile(SEL["agree_label"])).first.click(timeout=8000)
    except Exception:
        # ラベルが取れない画面向けフォールバック
        try:
            page.locator("input[type='checkbox']").first.check(timeout=5000)
            log("    (ラベル不可→checkbox直接)")
        except Exception as e:
            log(f"    ⚠️ チェックボックス操作をスキップ: {e}")
    human()

    log("STEP 4: ログイン")
    page.get_by_role("button", name=re.compile(SEL["login_btn"])).first.click()
    page.wait_for_load_state("networkidle", timeout=60000)
    shot(page, "after_login")
    pause("ログイン後")


def open_saved_search(page: Page, label: str) -> None:
    log("STEP 5: 売買物件検索へ")
    page.get_by_text(re.compile(SEL["menu_baibai"])).first.click()
    page.wait_for_load_state("networkidle", timeout=30000)
    human()

    log("STEP 6: 検索条件セクションを展開")
    try:
        page.get_by_text(re.compile(SEL["expand_cond"])).first.click(timeout=6000)
        human()
    except Exception:
        log("    (既に展開済みか、リンク無し)")
    shot(page, "search_form")
    pause("検索フォーム表示")

    log(f"STEP 7: 保存済み検索『{label}』を選択")
    sel = page.locator("select").first
    sel.wait_for(state="visible", timeout=15000)
    options = [o.strip() for o in sel.locator("option").all_text_contents()]

    def norm(s: str) -> str:
        return re.sub(r"\s+", "", s).replace("（", "(").replace("）", ")")

    target, tn = None, norm(label)
    for o in options:
        if not o:
            continue
        if tn == norm(o) or tn in norm(o) or norm(o).startswith(tn):
            target = o
            break
    if target is None:
        raise RuntimeError(
            f"保存済み検索が見つかりません。\n"
            f"  .envの指定: 『{label}』\n"
            f"  画面の候補: {options[:8]}\n"
            f"  → 表示名をそのままコピペして .env の SAVED_SEARCH_LABEL に入れてください"
        )
    log(f"    一致: 『{target}』")
    sel.select_option(label=target)
    human()

    log("STEP 8: 読込 → 確認ダイアログOK")
    page.once("dialog", lambda d: d.accept())
    page.get_by_role("button", name=re.compile(SEL["load_btn"])).first.click()
    try:
        page.get_by_role("button", name=re.compile(r"^\s*OK\s*$")).first.click(timeout=5000)
    except Exception:
        log("    (モーダルOKなし＝ネイティブダイアログで処理済み)")
    human()

    log("STEP 9: 検索実行")
    page.get_by_role("button", name=re.compile(SEL["search_btn"])).last.click()
    page.wait_for_load_state("networkidle", timeout=90000)
    human(1.5, 3.0)
    shot(page, "results")
    pause("検索結果")


def download_tab_pdfs(page: Page, tab_index: int) -> list[Path]:
    """指定タブで ページ内全選択 → 図面一括取得 → 一括取得 を実行。"""
    log(f"STEP 10-{tab_index}: タブ #{tab_index}")
    tabs = page.get_by_text(re.compile(SEL["tab_pattern"]))
    if tabs.count() <= tab_index:
        log(f"    タブ #{tab_index} は存在しない → スキップ")
        return []
    tabs.nth(tab_index).click()
    page.wait_for_load_state("networkidle", timeout=30000)
    human()

    try:
        page.get_by_role("button", name=re.compile(SEL["select_all"])).first.click(timeout=8000)
        human()
        page.get_by_role("button", name=re.compile(SEL["zumen_batch"])).first.click(timeout=8000)
        human(1.0, 2.0)
    except PWTimeout:
        log("    件数0（選択ボタンが出ない）→ スキップ")
        return []

    log("    ダウンロード待ち（最大3分）")
    with page.expect_download(timeout=180000) as info:
        page.get_by_role("button", name=re.compile(SEL["download_btn"])).first.click()
    dl = info.value
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = DL_DIR / f"tab{tab_index}_{ts}_{dl.suggested_filename}"
    dl.save_as(str(dst))
    log(f"    保存: {dst.name}")
    human(2.0, 4.0)
    return [dst]


# ---------------------------------------------------------------------------
# PDF → 構造化
# ---------------------------------------------------------------------------
FIELD_RE = {
    "物件番号":   r"物件番号[:：\s]*([0-9]{9,14})",
    "価格":       r"価格[:：\s]*([0-9,]+)\s*万円",
    "所在地":     r"所在地[:：\s]*([^\n]{4,40})",
    "交通":       r"(?:交通|沿線)[:：\s]*([^\n]{4,40})",
    "土地面積":   r"土地面積[:：\s]*([0-9,.]+)\s*(?:㎡|m2)",
    "建物面積":   r"建物面積[:：\s]*([0-9,.]+)\s*(?:㎡|m2)",
    "築年月":     r"築年月[:：\s]*([^\n]{4,20})",
    "構造":       r"構造[:：\s]*([^\n]{2,20})",
    "取引態様":   r"取引態様[:：\s]*([^\n]{2,12})",
    "現況":       r"現況[:：\s]*([^\n]{2,12})",
    "想定利回り": r"(?:想定利回り|利回り)[:：\s]*([0-9.]+)\s*%",
}


def parse_pdf(pdf: Path) -> list[dict]:
    """pdfplumberでページ単位にテキスト抽出→正規表現。取れない項目はClaudeに投げる。"""
    items: list[dict] = []
    with pdfplumber.open(pdf) as doc:
        for i, page in enumerate(doc.pages, 1):
            text = page.extract_text() or ""
            if len(text.strip()) < 30:
                items.append({"_page": i, "_needs_ai": True, "_src": pdf.name, "_text": ""})
                continue
            rec: dict = {"_page": i, "_src": pdf.name}
            for k, pat in FIELD_RE.items():
                m = re.search(pat, text)
                if m:
                    rec[k] = m.group(1).strip()
            rec["_needs_ai"] = "物件番号" not in rec
            rec["_text"] = text[:1800]
            items.append(rec)
    return items


def ai_fill(items: list[dict]) -> list[dict]:
    """正規表現で取れなかったページだけ claude -p に投げて補完する。"""
    if not shutil.which("claude"):
        log("    claude CLI 無し → AI補完スキップ")
        return items
    targets = [r for r in items if r.get("_needs_ai")]
    if not targets:
        return items
    log(f"    AI補完: {len(targets)}ページ")
    for rec in targets:
        if not rec.get("_text"):
            continue
        prompt = (
            "以下はREINSの物件図面から抽出したテキストです。"
            "物件番号・価格(万円)・所在地・交通・土地面積・建物面積・築年月・構造・取引態様・現況 を"
            "JSONオブジェクト1個だけで返してください。値が読み取れない項目はキーごと省略。"
            "前置き・説明・コードフェンスは一切不要。\n\n" + rec["_text"]
        )
        try:
            r = subprocess.run(["claude", "-p", prompt], capture_output=True,
                               text=True, timeout=180)
            if r.returncode == 0:
                m = re.search(r"\{.*\}", r.stdout, re.S)
                if m:
                    rec.update(json.loads(m.group(0)))
                    rec["_needs_ai"] = False
        except Exception as e:
            log(f"      AI補完失敗(p{rec['_page']}): {e}")
    return items


# ---------------------------------------------------------------------------
# 新着差分 ＆ 通知
# ---------------------------------------------------------------------------
def load_seen() -> dict:
    if STATE.exists():
        try:
            return json.loads(STATE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def diff_new(items: list[dict], label: str) -> list[dict]:
    """前回までに見た物件番号と突き合わせ、新着だけ返す。ここが毎日回す価値の中心。"""
    seen = load_seen()
    known = set(seen.get(label, []))
    fresh, ids = [], []
    for r in items:
        no = r.get("物件番号")
        if not no:
            continue
        ids.append(no)
        if no not in known:
            fresh.append(r)
    seen[label] = sorted(set(list(known) + ids))
    STATE.write_text(json.dumps(seen, ensure_ascii=False, indent=1), encoding="utf-8")
    return fresh


def write_output(items: list[dict], fresh: list[dict], label: str) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    (OUT_DIR / f"all_{ts}.json").write_text(
        json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")
    txt = OUT_DIR / f"new_{ts}.txt"
    lines = [f"■ REINS巡回 {datetime.now():%Y-%m-%d %H:%M}",
             f"  検索: {label}",
             f"  取得 {len(items)}件 / うち新着 {len(fresh)}件", ""]
    if not fresh:
        lines.append("  新着なし")
    for r in fresh:
        lines.append(f"── {r.get('物件番号','(番号不明)')}")
        for k in ("価格", "所在地", "交通", "土地面積", "建物面積",
                  "築年月", "構造", "取引態様", "現況", "想定利回り"):
            if r.get(k):
                lines.append(f"   {k}: {r[k]}")
        lines.append("")
    txt.write_text("\n".join(lines), encoding="utf-8")
    log(f"    出力: {txt.name}")
    return txt


def notify(txt: Path, fresh_count: int) -> None:
    if not NOTIFY_EMAIL:
        log("    NOTIFY_EMAIL未設定 → 通知スキップ（output/を直接見る運用）")
        return
    if fresh_count == 0:
        log("    新着0件 → 通知しない")
        return
    log(f"    通知先 {NOTIFY_EMAIL} / 新着{fresh_count}件 ※送信手段は未実装（要自前化）")


# ---------------------------------------------------------------------------
def main() -> int:
    global PAUSE
    ap = argparse.ArgumentParser(description="KHD REINS巡回")
    ap.add_argument("--headed", action="store_true", help="ブラウザを表示する")
    ap.add_argument("--pause-after-login", action="store_true", help="各ステップで停止")
    ap.add_argument("--dry-run", action="store_true", help="ログインまでで終了")
    args = ap.parse_args()
    PAUSE = args.pause_after_login

    if not USER_ID or not PASSWORD:
        log("❌ .env に REINS_USER_ID / REINS_PASSWORD がありません")
        return 1
    if not SAVED_SEARCH:
        log("❌ .env に SAVED_SEARCH_LABEL がありません")
        return 1
    if in_maintenance() and not args.headed:
        log("REINSメンテナンス時間帯(23:00-7:00) → スキップ")
        return 5

    log(f"=== REINS巡回 開始 / 検索: {SAVED_SEARCH} ===")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=not args.headed, args=LAUNCH_ARGS)
        ctx = browser.new_context(user_agent=UA, locale="ja-JP",
                                  viewport={"width": 1440, "height": 900},
                                  accept_downloads=True)
        ctx.add_init_script(INIT_JS)
        page = ctx.new_page()
        try:
            login(page)
            if args.dry_run:
                log("--dry-run: ログイン確認まで完了")
                return 0
            open_saved_search(page, SAVED_SEARCH)

            pdfs: list[Path] = []
            for i in range(2):          # 売外全 / 売外一
                pdfs += download_tab_pdfs(page, i)
            if not pdfs:
                log("PDFなし（該当0件）。正常終了")
                return 0

            log("STEP 11: PDF解析")
            items: list[dict] = []
            for p in pdfs:
                items += parse_pdf(p)
            items = ai_fill(items)

            fresh = diff_new(items, SAVED_SEARCH)
            log(f"    取得{len(items)}件 / 新着{len(fresh)}件")
            txt = write_output(items, fresh, SAVED_SEARCH)
            notify(txt, len(fresh))
            log("=== 完了 ===")
            return 0

        except Maintenance as e:
            log(f"メンテナンス: {e}")
            return 5
        except Exception as e:
            log(f"❌ エラー: {type(e).__name__}: {e}")
            shot(page, "error")
            log("   → logs/ のスクショをClaudeに見せてSEL辞書を直してください")
            return 1
        finally:
            if not args.pause_after_login:
                ctx.close()
                browser.close()


if __name__ == "__main__":
    sys.exit(main())
