# -*- coding: utf-8 -*-
"""
🤖 Coupang Wing 精算MSF 自動ダウンロード (01_経営管理 CFO)
─────────────────────────────────────────────────────────
【何をするか】
  Coupang Wing (wing.coupang.com) に自動ログインし、
  2アカウント × 直近2ヶ月分の「精算状況→決済確定」Excel を DL して
  Drive の _精算実額MSF/ フォルダへ自動配置する。

  → ゆーしが月に1回やっていた手作業を完全代替。

【初回設定（一度だけ）】
  cp ~/.config/khd/coupang_template.json ~/.config/khd/coupang.json
  # エディタで email/password/LINE token を埋める
  chmod 600 ~/.config/khd/coupang.json

【実行】
  python3 scripts/wing_auto_download.py
  python3 scripts/wing_auto_download.py --dry-run  # DLせずセッション確認のみ

【月次自動実行】
  run_monthly_ec_pipeline.sh から呼ばれる（毎月25日 9:00 cron）

【失敗時】
  - セッション切れ / 2FA要求 → LINE通知 → 菊池が1回手動ログイン後に再実行
  - DL後ファイルが見つからない → ログ確認
"""

import json, os, sys, re, shutil, glob, time, pickle, traceback, argparse
from pathlib import Path
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ────── 定数 ──────────────────────────────────────────
CFG_PATH   = Path.home() / ".config/khd/coupang.json"
SESS_DIR   = Path.home() / ".config/khd/wing_sessions"
MSF_DIR    = Path("/Users/kikuchikenta/Library/CloudStorage/"
                  "GoogleDrive-ec-seller@kikuchi-hd.net/"
                  "共有ドライブ/01_個人/2025_帳票、明細/韓国輸出売上/_精算実額MSF")
WING_BASE  = "https://wing.coupang.com"
LOGIN_URL  = f"{WING_BASE}/login"
SETTLE_URL = f"{WING_BASE}/vendor/partner/settlement/settlement-status"
DL_WAIT_S  = 60   # ダウンロード待機タイムアウト（秒）
NAV_WAIT   = 15000  # ナビゲーション待機（ms）

LOG_PREFIX = "[wing_auto_dl]"


def log(msg):
    print(f"{LOG_PREFIX} {msg}", flush=True)


# ────── 設定読み込み ──────────────────────────────────
def load_config():
    if not CFG_PATH.exists():
        sys.exit(
            f"❌ 設定ファイルが見つかりません: {CFG_PATH}\n"
            f"   cp ~/.config/khd/coupang_template.json ~/.config/khd/coupang.json\n"
            f"   してから email/password を記入してください。"
        )
    with open(CFG_PATH) as f:
        return json.load(f)


# ────── 日付範囲計算 ──────────────────────────────────
def calc_date_range():
    """先月1日 〜 今月末（2ヶ月窓）"""
    today = date.today()
    start = (today.replace(day=1) - relativedelta(months=1))  # 先月1日
    end   = (today.replace(day=1) + relativedelta(months=1) - timedelta(days=1))  # 今月末
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


# ────── セッション保存・ロード ────────────────────────
def session_path(account_name):
    SESS_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w]", "_", account_name)
    return SESS_DIR / f"{safe}_session.json"


def save_session(context, account_name):
    path = session_path(account_name)
    cookies = context.cookies()
    path.write_text(json.dumps(cookies, ensure_ascii=False, indent=2))
    log(f"セッション保存: {path}")


def load_session(context, account_name):
    path = session_path(account_name)
    if path.exists():
        try:
            cookies = json.loads(path.read_text())
            context.add_cookies(cookies)
            log(f"セッション復元: {path}")
            return True
        except Exception as e:
            log(f"セッション読み込みエラー（無視）: {e}")
    return False


# ────── ログイン ─────────────────────────────────────
def do_login(page, email, password, account_name):
    log(f"ログイン開始: {account_name} ({email})")
    page.goto(LOGIN_URL, timeout=NAV_WAIT)
    page.wait_for_load_state("domcontentloaded")

    # メールアドレス入力
    try:
        page.fill('input[type="email"], input[name="email"], input[id*="email"], input[id*="Email"]', email, timeout=8000)
    except PWTimeout:
        log("⚠️ email フィールド未検出。セレクタを確認してください。")
        page.screenshot(path=f"/tmp/wing_login_error_{account_name}.png")
        raise

    # パスワード入力
    page.fill('input[type="password"]', password, timeout=5000)

    # ログインボタン
    for sel in [
        'button[type="submit"]',
        'button:has-text("로그인")',  # 韓国語: ログイン
        'button:has-text("Login")',
        'input[type="submit"]',
    ]:
        try:
            page.click(sel, timeout=3000)
            break
        except PWTimeout:
            continue

    # ログイン完了待ち（URLが変わるか、エラーが出るか）
    try:
        page.wait_for_url(lambda u: "login" not in u, timeout=15000)
        log(f"ログイン成功: {account_name}")
    except PWTimeout:
        # 2FA や CAPTCHA の可能性
        page.screenshot(path=f"/tmp/wing_login_2fa_{account_name}.png")
        log(f"⚠️ ログイン後の遷移なし。2FA/CAPTCHA の可能性あり。")
        log(f"   スクリーンショット: /tmp/wing_login_2fa_{account_name}.png")
        raise RuntimeError(
            f"Wingログイン失敗 ({account_name}): 2FAまたはCAPTCHAが必要かもしれません。\n"
            f"一度ブラウザで手動ログインして、~/.config/khd/wing_sessions/ を再生成してください。"
        )


def ensure_logged_in(page, context, email, password, account_name):
    """セッション確認 → 必要なら再ログイン"""
    page.goto(SETTLE_URL, timeout=NAV_WAIT)
    page.wait_for_load_state("domcontentloaded")
    if "login" in page.url.lower():
        log(f"セッション切れ: {account_name} → 再ログイン")
        do_login(page, email, password, account_name)
        save_session(context, account_name)
    else:
        log(f"セッション有効: {account_name}")


# ────── 精算状況ページ操作 ──────────────────────────
def set_filters_and_download(page, start_date, end_date, account_name, download_dir, dry_run=False):
    """
    精算状況ページで検索条件を設定してExcelをDL
    返り値: DLされたファイルパス or None
    """
    log(f"精算状況ページ: {SETTLE_URL}")
    page.goto(SETTLE_URL, timeout=NAV_WAIT)
    page.wait_for_load_state("networkidle", timeout=20000)

    # ── 基準日: 「決済日」選択 ──────────────────────
    # セレクトボックスやラジオボタン。複数パターンで試す。
    for sel in [
        'select[name*="dateType"], select[id*="dateType"]',
        '[class*="dateType"] select',
        'select:near(:text("기준일"), 200)',
    ]:
        try:
            page.select_option(sel, label="결제 완료일", timeout=3000)  # 「決済完了日」韓国語
            log("基準日: 決済完了日 選択")
            break
        except Exception:
            pass

    # ── 開始日・終了日 ──────────────────────────────
    # date input を探して直接入力
    date_inputs = page.locator('input[type="date"], input[placeholder*="YYYY"], input[placeholder*="yyyy"]').all()
    if len(date_inputs) >= 2:
        date_inputs[0].fill(start_date)
        date_inputs[1].fill(end_date)
        log(f"期間: {start_date} 〜 {end_date}")
    else:
        # フォールバック: 韓国語ラベル付近のテキスト入力
        for idx, target in enumerate([start_date, end_date]):
            for sel in [
                f'input[id*="startDate"], input[id*="fromDate"]' if idx == 0 else 'input[id*="endDate"], input[id*="toDate"]',
            ]:
                try:
                    inp = page.locator(sel).first
                    inp.triple_click()
                    inp.fill(target)
                    log(f"日付{idx+1}入力: {target}")
                    break
                except Exception:
                    pass

    # ── 精算状態: 「決済確定」選択 ──────────────────
    for sel in [
        'select[name*="status"], select[id*="status"]',
        '[class*="status"] select',
    ]:
        try:
            # 韓国語で「결제 확정」= 決済確定
            page.select_option(sel, label="결제 확정", timeout=3000)
            log("精算状態: 決済確定 選択")
            break
        except Exception:
            try:
                page.select_option(sel, label="결제확정", timeout=1000)
                break
            except Exception:
                pass

    # ── 検索ボタン ──────────────────────────────────
    for sel in [
        'button:has-text("검색")',   # 韓国語: 検索
        'button:has-text("조회")',   # 韓国語: 照会
        'button[type="submit"]:visible',
        'input[type="submit"]:visible',
    ]:
        try:
            page.click(sel, timeout=3000)
            log("検索実行")
            break
        except Exception:
            pass

    page.wait_for_load_state("networkidle", timeout=20000)
    time.sleep(2)  # 検索結果の描画待ち

    if dry_run:
        log(f"[DRY RUN] ダウンロードステップをスキップ")
        page.screenshot(path=f"/tmp/wing_dryrun_{account_name}.png")
        log(f"スクリーンショット: /tmp/wing_dryrun_{account_name}.png")
        return None

    # ── ダウンロードボタン ───────────────────────────
    downloaded_file = None
    with page.expect_download(timeout=DL_WAIT_S * 1000) as dl_info:
        clicked = False
        for sel in [
            # ページ右上のバルクDLボタン（優先）
            'button:has-text("엑셀 다운로드 목록")',   # 決済管理エクセルダウンロードリスト
            'button:has-text("엑셀 다운로드")',
            'button:has-text("Excel")',
            'button:has-text("excel")',
            '[class*="excel-download"] button',
            '[class*="excelDownload"] button',
            'a[href*="excel"], a[href*="xlsx"]',
        ]:
            try:
                page.click(sel, timeout=3000)
                log(f"ダウンロードボタン クリック: {sel}")
                clicked = True
                break
            except Exception:
                pass
        if not clicked:
            log("⚠️ ダウンロードボタンが見つかりません。スクリーンショットを確認してください。")
            page.screenshot(path=f"/tmp/wing_no_dl_btn_{account_name}.png")
            raise RuntimeError(f"DLボタン未検出 ({account_name}): /tmp/wing_no_dl_btn_{account_name}.png を確認")

    dl = dl_info.value
    tmp_path = Path(dl.path())
    suggested_name = dl.suggested_filename or tmp_path.name
    dest_path = Path(download_dir) / suggested_name

    # 重複ファイル名の場合はサフィックスを付ける
    if dest_path.exists():
        stem = dest_path.stem; suffix = dest_path.suffix
        i = 1
        while dest_path.exists():
            dest_path = Path(download_dir) / f"{stem}_{i}{suffix}"
            i += 1

    shutil.move(str(tmp_path), str(dest_path))
    log(f"✅ DL完了: {dest_path.name} → {download_dir}")
    return str(dest_path)


# ────── アカウント1件処理 ──────────────────────────
def process_account(pw, account_cfg, start_date, end_date, download_dir, dry_run=False):
    name  = account_cfg["name"]
    email = account_cfg["login_email"]
    pwd   = account_cfg["login_password"]
    SESS_DIR.mkdir(parents=True, exist_ok=True)

    log(f"=== {name} 処理開始 ===")
    browser = pw.chromium.launch(
        headless=True,
        downloads_path=str(download_dir),
    )
    context = browser.new_context(accept_downloads=True)
    load_session(context, name)
    page = context.new_page()

    try:
        ensure_logged_in(page, context, email, pwd, name)
        result = set_filters_and_download(page, start_date, end_date, name, download_dir, dry_run)
        return result
    finally:
        context.close()
        browser.close()


# ────── メイン ───────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Coupang Wing MSF 自動ダウンロード")
    parser.add_argument("--dry-run", action="store_true", help="実際にDLせずセッション確認のみ")
    parser.add_argument("--account", type=str, default=None, help="特定アカウント名のみ実行（例: クーパン1）")
    args = parser.parse_args()

    cfg = load_config()
    MSF_DIR.mkdir(parents=True, exist_ok=True)
    start_date, end_date = calc_date_range()
    log(f"対象期間: {start_date} 〜 {end_date}")
    log(f"保存先: {MSF_DIR}")

    results = []
    errors  = []

    with sync_playwright() as pw:
        for acct in cfg["accounts"]:
            if args.account and acct["name"] != args.account:
                continue
            try:
                r = process_account(pw, acct, start_date, end_date, str(MSF_DIR), args.dry_run)
                results.append((acct["name"], r))
            except Exception as e:
                log(f"❌ {acct['name']} でエラー: {e}")
                errors.append((acct["name"], str(e)))

    # ── 結果サマリー + 通知 ──────────────────────────
    try:
        from ec_notify import notify
    except ImportError:
        def notify(m, **kw): print(f"[notify] {m}")

    if errors:
        err_msg = "\n".join(f"・{n}: {e[:80]}" for n, e in errors)
        msg = (
            f"⚠️ Coupang Wing DL 一部失敗\n"
            f"成功: {len(results)}件 / 失敗: {len(errors)}件\n"
            f"失敗内容:\n{err_msg}\n\n"
            f"▶ /tmp/ のスクリーンショットを確認するか\n"
            f"　Wingに手動ログイン後に再実行してください。"
        )
        notify(msg, urgent=True)
        log(msg)
        sys.exit(1)
    else:
        ok_files = "\n".join(f"・{n}: {Path(r).name if r else '(dry-run)'}" for n, r in results)
        msg = (
            f"✅ Wing MSF DL完了 ({len(results)}アカウント)\n"
            f"期間: {start_date}〜{end_date}\n"
            f"{ok_files}\n"
            f"→ ec_settlement_recon_v2.py を実行します"
        )
        notify(msg)
        log(msg)


if __name__ == "__main__":
    main()
