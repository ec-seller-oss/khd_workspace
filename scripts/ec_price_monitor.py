# -*- coding: utf-8 -*-
"""
Amazon JP 価格監視 → Coupang Wing 価格自動更新スクリプト — 03_事業運営 / EC韓国輸出
────────────────────────────────────────────────────────────────────────────────
【何をするか】
  1. DBの products テーブルから監視対象ASIN(amazon_in_stock=1, auto_price_update=1)を取得
  2. Amazon JP ページをスクレイピングして現在価格を取得
  3. 利益率を計算し、閾値に応じて Coupang Wing の掲載価格を自動更新する
  4. 価格変動を price_log テーブルに全件記録
  5. 結果をLINEで通知（異常時は緊急通知）

【利益率計算式】
  純収入(KRW) = coupang_price_krw × (1 - 手数料率0.11)
  原価(KRW)   = amazon_price_jpy / fx_rate
  margin_pct  = (純収入 - 原価) / 純収入 × 100

【自動更新ロジック】
  margin_pct < min_margin_pct (デフォルト5%) → 価格引き上げ
  margin_pct < 0%                             → 掲載一時停止 + 緊急通知
  margin_pct > 20%                            → 価格引き下げ提案通知（競争力アップ）

【他スクリプトとの関係】
  - ec_automation_db.py  … SQLite DB操作ヘルパー（商品・価格ログ）
  - ec_notify.py         … LINE通知ヘルパー
  - wing_auto_download.py … Coupang Wing Playwright操作の参考実装

【認証設定（初回のみ）】
  ~/.config/khd/coupang.json に email/password/LINE token を記入

【実行コマンド】
  # 全ASIN一括チェック（1回実行）
  python3 scripts/ec_price_monitor.py --check-once

  # 特定ASINのみ確認（テスト用）
  python3 scripts/ec_price_monitor.py --asin B0002S31E2,B08KR4KWXD --check-once

  # 更新せず結果だけ表示
  python3 scripts/ec_price_monitor.py --dry-run --check-once

【自動実行タイミング】
  2時間毎にcronから呼ぶ:
    0 */2 * * * /usr/local/bin/python3 /Users/kikuchikenta/01_honbu_docs_automation/scripts/ec_price_monitor.py --check-once >> /tmp/ec_price_monitor.log 2>&1

【Open API 対応について】
  Coupang Wing Open API が利用可能になった場合、CoupangPricingClient クラスの
  update_price() / delist() メソッドを差し替えるだけで対応可能な設計にしている。
  現状はPlaywrightによるWeb操作でフォールバック。
"""

import json
import os
import re
import sys
import time
import traceback
import argparse
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple

# requests（標準インストール済み）
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False
    print("[ec_price_monitor] ⚠️ requests 未インストール。Playwright のみで取得します。")

# BeautifulSoup4（任意、なければPlaywrightフォールバック）
try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False
    print("[ec_price_monitor] bs4 未インストール。Playwright でAmazon価格を取得します（低速）。")

# Playwright
try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    HAS_PLAYWRIGHT = True
except ImportError:
    HAS_PLAYWRIGHT = False
    print("[ec_price_monitor] ⚠️ Playwright 未インストール。requests のみで取得します。")

# 自作モジュール（同じ scripts/ ディレクトリから）
_SCRIPTS_DIR = Path(__file__).parent
sys.path.insert(0, str(_SCRIPTS_DIR))
from ec_automation_db import upsert_product, log_price, get_conn
from ec_notify import notify

# ────── 定数・設定 ────────────────────────────────────────────────────────
CFG_PATH    = Path.home() / ".config/khd/coupang.json"
SESS_DIR    = Path.home() / ".config/khd/wing_sessions"
WING_BASE   = "https://wing.coupang.com"
LOGIN_URL   = f"{WING_BASE}/login"
AMAZON_BASE = "https://www.amazon.co.jp/dp"

# デフォルト設定（DBのproductsテーブルの値が優先）
DEFAULT_FX_RATE         = 0.109    # 円→ウォン近似（円/ウォン）
DEFAULT_MARKUP_RATE     = 1.25     # 仕入れ価格に対する掛け率
DEFAULT_MIN_MARGIN_PCT  = 5.0      # 最低粗利率(%)。これを下回ったら価格引き上げ
HIGH_MARGIN_THRESHOLD   = 20.0     # これを超えたら価格引き下げ提案
COUPANG_FEE_RATE        = 0.11     # Coupang手数料率（11%）

# Amazon スクレイピング
AMAZON_PRICE_SELECTORS = [
    "#priceblock_ourprice",
    "#priceblock_dealprice",
    "span.a-price-whole",
    "#apex_desktop span.a-price .a-price-whole",
    "#apex_desktop .a-price .a-price-whole",
    ".a-price .a-price-whole",
]
AMAZON_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja-JP,ja;q=0.9,en;q=0.8",
}

NAV_WAIT_MS = 15000
LOG_PREFIX  = "[ec_price_monitor]"

POLL_INTERVAL_SEC = 2 * 60 * 60  # 2時間（常駐モード用）


def log(msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{LOG_PREFIX} {ts} {msg}", flush=True)


# ────── 設定読み込み ──────────────────────────────────────────────────────
def load_config() -> Dict[str, Any]:
    if not CFG_PATH.exists():
        sys.exit(
            f"❌ 設定ファイルが見つかりません: {CFG_PATH}\n"
            f"   cp ~/.config/khd/coupang_template.json ~/.config/khd/coupang.json"
        )
    with open(CFG_PATH) as f:
        return json.load(f)


# ────── DBから監視対象商品を取得 ──────────────────────────────────────────
def get_target_products(asin_filter: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """
    products テーブルから監視対象を取得する。
    条件: amazon_in_stock=1 AND auto_price_update=1
    asin_filter が指定された場合はその ASIN のみ対象。
    """
    with get_conn() as conn:
        if asin_filter:
            placeholders = ",".join(["?"] * len(asin_filter))
            rows = conn.execute(
                f"SELECT * FROM products WHERE asin IN ({placeholders})",
                asin_filter
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM products WHERE amazon_in_stock=1 AND auto_price_update=1"
            ).fetchall()
    return [dict(r) for r in rows]


# ────── Amazon価格取得 ────────────────────────────────────────────────────

def _parse_price_from_html(html: str) -> Optional[float]:
    """
    Amazon商品ページのHTMLから価格を抽出する。
    複数セレクタを順に試す。
    """
    if not HAS_BS4:
        return None

    soup = BeautifulSoup(html, "html.parser")

    for sel in AMAZON_PRICE_SELECTORS:
        # CSSセレクタ形式をBeautifulSoupのselect()で処理
        try:
            els = soup.select(sel)
            for el in els:
                txt = el.get_text(strip=True)
                # "1,234" や "¥1,234" 等から数字を抽出
                price_str = re.sub(r"[^\d]", "", txt)
                if price_str and len(price_str) >= 2:
                    price = float(price_str)
                    if price > 0:
                        return price
        except Exception:
            continue

    return None


def fetch_amazon_price_requests(asin: str) -> Optional[float]:
    """
    requests + BeautifulSoup4 でAmazon JP から価格取得を試みる。
    失敗(bot検知含む)の場合はNoneを返す。
    """
    if not HAS_REQUESTS:
        return None

    url = f"{AMAZON_BASE}/{asin}"
    try:
        resp = requests.get(url, headers=AMAZON_HEADERS, timeout=15)
        if resp.status_code != 200:
            log(f"  Amazon HTTP {resp.status_code}: {asin}")
            return None

        # bot検知チェック（"robot" や "captcha" が含まれる場合）
        if "robot" in resp.text.lower() or "captcha" in resp.text.lower():
            log(f"  Amazon bot検知の可能性: {asin} → Playwrightにフォールバック")
            return None

        price = _parse_price_from_html(resp.text)
        if price:
            log(f"  Amazon価格(requests): ¥{price:,.0f} / {asin}")
        return price

    except Exception as e:
        log(f"  requests取得エラー ({asin}): {e}")
        return None


def fetch_amazon_price_playwright(page: Any, asin: str) -> Optional[float]:
    """
    Playwright で Amazon JP から価格取得（bot検知回避フォールバック）。
    """
    if not HAS_PLAYWRIGHT:
        return None

    url = f"{AMAZON_BASE}/{asin}"
    try:
        page.goto(url, timeout=NAV_WAIT_MS)
        page.wait_for_load_state("domcontentloaded")

        # CAPTCHA確認
        if "captcha" in page.url.lower() or "captcha" in (page.title() or "").lower():
            log(f"  ⚠️ Amazon CAPTCHA検知: {asin}")
            return None

        # 各セレクタを試す
        for sel in AMAZON_PRICE_SELECTORS:
            try:
                el = page.locator(sel).first
                txt = el.inner_text(timeout=3000).strip()
                price_str = re.sub(r"[^\d]", "", txt)
                if price_str and len(price_str) >= 2:
                    price = float(price_str)
                    if price > 0:
                        log(f"  Amazon価格(Playwright): ¥{price:,.0f} / {asin}")
                        return price
            except Exception:
                continue

        log(f"  Amazon価格セレクタ全失敗: {asin}")
        return None

    except Exception as e:
        log(f"  Playwright取得エラー ({asin}): {e}")
        return None


def fetch_amazon_price(asin: str, pw_page: Optional[Any] = None) -> Optional[float]:
    """
    Amazon JP 商品の現在価格を取得する。
    1. requests + bs4 で試みる
    2. 失敗した場合は Playwright で再試行（pw_page が渡された場合）
    """
    # まず requests で試みる
    price = fetch_amazon_price_requests(asin)
    if price is not None:
        return price

    # Playwrightフォールバック
    if pw_page is not None:
        log(f"  Playwrightフォールバック: {asin}")
        price = fetch_amazon_price_playwright(pw_page, asin)
        if price is not None:
            return price

    log(f"  ❌ Amazon価格取得失敗: {asin}")
    return None


# ────── 利益率計算 ────────────────────────────────────────────────────────

def calc_margin(
    amazon_price_jpy: float,
    coupang_price_krw: float,
    fx_rate: float,
    fee_rate: float = COUPANG_FEE_RATE
) -> Tuple[float, float]:
    """
    利益率を計算する。
    戻り値: (margin_pct, net_revenue_krw)
      margin_pct      : 粗利率(%)
      net_revenue_krw : 手数料控除後のCoupang純収入(KRW)
    """
    net_revenue_krw = coupang_price_krw * (1.0 - fee_rate)
    cost_krw = amazon_price_jpy / fx_rate
    if net_revenue_krw <= 0:
        return -100.0, net_revenue_krw
    margin_pct = (net_revenue_krw - cost_krw) / net_revenue_krw * 100.0
    return margin_pct, net_revenue_krw


def calc_new_price(
    amazon_price_jpy: float,
    fx_rate: float,
    markup_rate: float,
    fee_rate: float = COUPANG_FEE_RATE
) -> float:
    """
    適正な新掲載価格(KRW)を計算する。
    new_price = (Amazon仕入れ価格KRW × markup_rate) / (1 - 手数料率)
    """
    cost_krw    = amazon_price_jpy / fx_rate
    new_price   = (cost_krw * markup_rate) / (1.0 - fee_rate)
    # 10ウォン単位に丸める
    return round(new_price / 10) * 10


# ────── Coupang Wing 価格更新クライアント ────────────────────────────────
class CoupangPricingClient:
    """
    Coupang Wing の価格操作インターフェース。

    現在は Playwright による Web操作で実装している。
    将来 Open API が利用可能になった場合は、このクラスの内部実装を
    差し替えるだけで対応できるよう抽象化している。

    使い方:
      client = CoupangPricingClient(page, account_name)
      success = client.update_price(item_id, new_price_krw)
      success = client.delist(item_id)
    """

    def __init__(self, page: Optional[Any], account_name: str):
        self.page = page
        self.account_name = account_name

    def update_price(
        self,
        asin: str,
        coupang_item_id: Optional[str],
        new_price_krw: float,
        dry_run: bool = False
    ) -> bool:
        """
        Coupang Wing の掲載価格を更新する。
        coupang_item_id が不明な場合は ASIN でページ検索を試みる。
        """
        if dry_run:
            log(f"  [DRY RUN] 価格更新スキップ: {asin} → ₩{new_price_krw:,.0f}")
            return True

        if self.page is None:
            log(f"  ❌ Playwright未初期化のため価格更新不可: {asin}")
            return False

        # Wing商品編集ページを開く
        # item_id が不明な場合は商品検索ページから ASIN で検索する
        if coupang_item_id:
            edit_url = f"{WING_BASE}/vendor/partner/items/{coupang_item_id}/edit"
        else:
            edit_url = f"{WING_BASE}/vendor/partner/items?keyword={asin}"

        try:
            self.page.goto(edit_url, timeout=NAV_WAIT_MS)
            self.page.wait_for_load_state("networkidle", timeout=20000)

            # ASINで検索した場合、商品一覧から対象をクリック
            if not coupang_item_id:
                try:
                    self.page.click(
                        f'[class*="item"]:has-text("{asin}"), a:has-text("{asin}")',
                        timeout=5000
                    )
                    self.page.wait_for_load_state("networkidle", timeout=15000)
                except PWTimeout:
                    log(f"  ⚠️ 商品が見つかりません（Wing上）: {asin}")
                    return False

            # 価格フィールドを探して更新
            price_str = str(int(new_price_krw))
            updated = False
            for sel in [
                'input[name*="price"], input[id*="price"]',
                'input[placeholder*="원"], input[placeholder*="가격"]',
                '[class*="price-input"] input',
            ]:
                try:
                    inp = self.page.locator(sel).first
                    inp.triple_click(timeout=3000)
                    inp.fill(price_str, timeout=3000)
                    log(f"  価格フィールド入力: {sel} → ₩{new_price_krw:,.0f}")
                    updated = True
                    break
                except Exception:
                    continue

            if not updated:
                log(f"  ❌ 価格フィールド未検出: {asin}")
                self.page.screenshot(path=f"/tmp/wing_price_edit_{asin}.png")
                return False

            # 保存ボタンをクリック
            for save_sel in [
                'button:has-text("저장")',    # 保存（韓国語）
                'button:has-text("적용")',    # 適用
                'button[type="submit"]:visible',
            ]:
                try:
                    self.page.click(save_sel, timeout=3000)
                    self.page.wait_for_load_state("networkidle", timeout=15000)
                    log(f"  ✅ 価格更新完了: {asin} → ₩{new_price_krw:,.0f}")
                    return True
                except Exception:
                    continue

            log(f"  ⚠️ 保存ボタン未検出: {asin}")
            return False

        except Exception as e:
            log(f"  ❌ 価格更新エラー ({asin}): {e}")
            traceback.print_exc()
            return False

    def delist(
        self,
        asin: str,
        coupang_item_id: Optional[str],
        dry_run: bool = False
    ) -> bool:
        """
        Coupang Wing で商品を一時停止（非表示）にする。
        赤字（margin < 0）の場合に呼ばれる。
        """
        if dry_run:
            log(f"  [DRY RUN] 掲載停止スキップ: {asin}")
            return True

        if self.page is None:
            log(f"  ❌ Playwright未初期化のため掲載停止不可: {asin}")
            return False

        if coupang_item_id:
            url = f"{WING_BASE}/vendor/partner/items/{coupang_item_id}/edit"
        else:
            url = f"{WING_BASE}/vendor/partner/items?keyword={asin}"

        try:
            self.page.goto(url, timeout=NAV_WAIT_MS)
            self.page.wait_for_load_state("networkidle", timeout=20000)

            # 「판매 중지」(販売停止) ボタン or トグルを探す
            for sel in [
                'button:has-text("판매 중지")',
                'button:has-text("비활성")',
                '[class*="delist"], [class*="deactivate"]',
                'input[type="checkbox"][name*="active"]',
            ]:
                try:
                    el = self.page.locator(sel).first
                    el.click(timeout=3000)
                    self.page.wait_for_load_state("networkidle", timeout=15000)
                    log(f"  ✅ 掲載停止完了: {asin}")
                    return True
                except Exception:
                    continue

            log(f"  ⚠️ 掲載停止ボタン未検出: {asin}")
            self.page.screenshot(path=f"/tmp/wing_delist_{asin}.png")
            return False

        except Exception as e:
            log(f"  ❌ 掲載停止エラー ({asin}): {e}")
            traceback.print_exc()
            return False


# ────── セッション管理（wing_auto_download.pyと同じ方式） ─────────────────

def session_path(account_name: str) -> Path:
    SESS_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w]", "_", account_name)
    return SESS_DIR / f"{safe}_session.json"


def load_session(context: Any, account_name: str) -> bool:
    path = session_path(account_name)
    if path.exists():
        try:
            cookies = json.loads(path.read_text())
            context.add_cookies(cookies)
            log(f"セッション復元: {path.name}")
            return True
        except Exception as e:
            log(f"セッション読み込みエラー（無視）: {e}")
    return False


def save_session(context: Any, account_name: str) -> None:
    path = session_path(account_name)
    cookies = context.cookies()
    path.write_text(json.dumps(cookies, ensure_ascii=False, indent=2))
    log(f"セッション保存: {path.name}")


def do_login(page: Any, email: str, password: str, account_name: str) -> None:
    log(f"Wingログイン: {account_name}")
    page.goto(LOGIN_URL, timeout=NAV_WAIT_MS)
    page.wait_for_load_state("domcontentloaded")

    try:
        page.fill(
            'input[type="email"], input[name="email"], input[id*="email"], input[id*="Email"]',
            email, timeout=8000
        )
    except PWTimeout:
        page.screenshot(path=f"/tmp/wing_price_login_err_{account_name}.png")
        raise

    page.fill('input[type="password"]', password, timeout=5000)

    for sel in [
        'button[type="submit"]',
        'button:has-text("로그인")',
        'button:has-text("Login")',
    ]:
        try:
            page.click(sel, timeout=3000)
            break
        except PWTimeout:
            continue

    try:
        page.wait_for_url(lambda u: "login" not in u, timeout=15000)
        log(f"Wingログイン成功: {account_name}")
    except PWTimeout:
        page.screenshot(path=f"/tmp/wing_price_2fa_{account_name}.png")
        raise RuntimeError(
            f"Wingログイン失敗 ({account_name}): 2FA/CAPTCHA の可能性。\n"
            f"スクリーンショット: /tmp/wing_price_2fa_{account_name}.png"
        )


def ensure_wing_logged_in(
    page: Any, context: Any,
    email: str, password: str, account_name: str
) -> None:
    page.goto(WING_BASE, timeout=NAV_WAIT_MS)
    page.wait_for_load_state("domcontentloaded")
    if "login" in page.url.lower():
        do_login(page, email, password, account_name)
        save_session(context, account_name)
    else:
        log(f"Wingセッション有効: {account_name}")


# ────── 1商品の価格チェック＆更新 ──────────────────────────────────────────

def check_and_update_product(
    product: Dict[str, Any],
    pricing_client: CoupangPricingClient,
    pw_page: Optional[Any],
    dry_run: bool = False
) -> Dict[str, Any]:
    """
    1商品の価格確認・利益率計算・必要に応じた価格更新を実行する。
    戻り値: 処理結果サマリー辞書
    """
    asin             = product["asin"]
    item_id          = product.get("coupang_item_id")
    coupang_price    = product.get("coupang_price_krw", 0) or 0
    fx_rate          = product.get("fx_rate", DEFAULT_FX_RATE) or DEFAULT_FX_RATE
    markup_rate      = product.get("markup_rate", DEFAULT_MARKUP_RATE) or DEFAULT_MARKUP_RATE
    min_margin_pct   = product.get("min_margin_pct", DEFAULT_MIN_MARGIN_PCT) or DEFAULT_MIN_MARGIN_PCT
    product_name_jp  = product.get("product_name_jp", asin)

    result: Dict[str, Any] = {
        "asin":            asin,
        "product_name_jp": product_name_jp,
        "action":          "no_change",
        "amazon_price":    None,
        "margin_pct":      None,
        "new_price":       None,
        "error":           None,
    }

    log(f"チェック開始: {asin} / {product_name_jp[:30]}")

    # Amazon価格取得
    amazon_price = fetch_amazon_price(asin, pw_page)
    if amazon_price is None:
        result["error"] = "Amazon価格取得失敗"
        result["action"] = "skip"
        log_price(asin, None, coupang_price, fx_rate, None, "price_fetch_failed")
        # 連続失敗カウントをDBに記録
        upsert_product({
            "asin": asin,
            "price_check_fail": (product.get("price_check_fail", 0) or 0) + 1,
            "last_price_check": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        return result

    result["amazon_price"] = amazon_price

    # 利益率計算
    margin_pct, net_revenue = calc_margin(amazon_price, coupang_price, fx_rate)
    result["margin_pct"] = margin_pct

    log(
        f"  Amazon: ¥{amazon_price:,.0f} / "
        f"Coupang: ₩{coupang_price:,.0f} / "
        f"fx: {fx_rate} / "
        f"粗利率: {margin_pct:.1f}%"
    )

    # 価格更新ロジック
    action = "no_change"

    if margin_pct < 0:
        # 赤字 → 掲載停止 + 緊急通知
        log(f"  ❌ 赤字 ({margin_pct:.1f}%) → 掲載停止")
        success = pricing_client.delist(asin, item_id, dry_run=dry_run)
        action = "delisted"
        result["action"] = action
        if not dry_run:
            upsert_product({
                "asin":           asin,
                "amazon_price_jpy": amazon_price,
                "coupang_active": 0,
                "last_price_check": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "price_check_fail": 0,
            })
            notify(
                f"❌ 赤字商品を掲載停止しました\n"
                f"ASIN: {asin}\n"
                f"商品: {product_name_jp[:40]}\n"
                f"Amazon: ¥{amazon_price:,.0f} / Coupang: ₩{coupang_price:,.0f}\n"
                f"粗利率: {margin_pct:.1f}%",
                urgent=True
            )

    elif margin_pct < min_margin_pct:
        # 利益率不足 → 価格引き上げ
        new_price = calc_new_price(amazon_price, fx_rate, markup_rate)
        log(f"  ⬆️ 利益率不足 ({margin_pct:.1f}% < {min_margin_pct}%) → ₩{new_price:,.0f} に引き上げ")
        success = pricing_client.update_price(asin, item_id, new_price, dry_run=dry_run)
        action = "updated"
        result["action"] = action
        result["new_price"] = new_price
        if not dry_run and success:
            upsert_product({
                "asin":             asin,
                "amazon_price_jpy": amazon_price,
                "coupang_price_krw": new_price,
                "last_price_check": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "price_check_fail": 0,
            })

    elif margin_pct > HIGH_MARGIN_THRESHOLD:
        # 利益率過剰 → 引き下げ提案通知（自動更新はしない）
        suggested_price = calc_new_price(
            amazon_price, fx_rate,
            markup_rate=markup_rate * 0.9  # 10%引き下げ提案
        )
        log(f"  💡 高利益率 ({margin_pct:.1f}%) → 価格引き下げ提案: ₩{suggested_price:,.0f}")
        action = "suggest_lower"
        result["action"] = action
        result["new_price"] = suggested_price
        if not dry_run:
            upsert_product({
                "asin":             asin,
                "amazon_price_jpy": amazon_price,
                "last_price_check": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "price_check_fail": 0,
            })
            notify(
                f"💡 価格引き下げ提案（競争力アップ）\n"
                f"ASIN: {asin}\n"
                f"商品: {product_name_jp[:40]}\n"
                f"現在粗利率: {margin_pct:.1f}%\n"
                f"提案価格: ₩{suggested_price:,.0f}（現在₩{coupang_price:,.0f}）\n"
                f"→ 手動で確認・更新してください"
            )

    else:
        # 利益率適正 → 変更なし
        log(f"  ✅ 利益率適正 ({margin_pct:.1f}%) → 変更なし")
        action = "no_change"
        result["action"] = action
        if not dry_run:
            upsert_product({
                "asin":             asin,
                "amazon_price_jpy": amazon_price,
                "last_price_check": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "price_check_fail": 0,
            })

    # 価格ログを記録
    if not dry_run:
        log_price(asin, amazon_price, coupang_price, fx_rate, margin_pct, action)
    else:
        log(f"  [DRY RUN] price_log記録スキップ: action={action}")

    return result


# ────── メイン処理 ────────────────────────────────────────────────────────

def run_once(
    asin_filter: Optional[List[str]] = None,
    dry_run: bool = False
) -> None:
    """1回分の価格チェック処理を実行する"""
    log("価格監視チェック開始")
    cfg = load_config()

    products = get_target_products(asin_filter)
    if not products:
        log("監視対象商品なし（DBに products が登録されていないか、auto_price_update=0）")
        return

    log(f"監視対象: {len(products)} 商品")

    results: List[Dict[str, Any]] = []
    errors: List[str] = []

    # アカウントを1つ選んでWingのセッションを確立
    # 価格更新はWingが必要。requestsのみの場合はPricingClientにNoneを渡す。
    accounts = cfg.get("accounts", [])
    wing_account = accounts[0] if accounts else None

    if HAS_PLAYWRIGHT and wing_account:
        # Playwright を起動してWingにログイン
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            context = browser.new_context()
            load_session(context, wing_account["name"])
            page = context.new_page()

            # Amazon価格取得用ページ（別コンテキストが望ましいが簡略化）
            amazon_page = context.new_page()

            try:
                ensure_wing_logged_in(
                    page, context,
                    wing_account["login_email"],
                    wing_account["login_password"],
                    wing_account["name"]
                )

                pricing_client = CoupangPricingClient(page, wing_account["name"])

                for product in products:
                    try:
                        result = check_and_update_product(
                            product, pricing_client,
                            pw_page=amazon_page,
                            dry_run=dry_run
                        )
                        results.append(result)
                        time.sleep(2)  # Amazon へのアクセスを少し間隔を開ける
                    except Exception as e:
                        err = f"{product['asin']}: {e}"
                        log(f"❌ {err}")
                        errors.append(err)

            finally:
                amazon_page.close()
                context.close()
                browser.close()

    else:
        # Playwright なしの場合: requestsのみで価格取得、Wing操作は不可
        log("⚠️ Playwright未使用: Amazon価格取得のみ実行（Wing価格更新は不可）")
        pricing_client = CoupangPricingClient(None, "")
        for product in products:
            try:
                result = check_and_update_product(
                    product, pricing_client,
                    pw_page=None,
                    dry_run=dry_run
                )
                results.append(result)
                time.sleep(1)
            except Exception as e:
                err = f"{product['asin']}: {e}"
                log(f"❌ {err}")
                errors.append(err)

    # ────── 結果サマリーを通知 ──────────────────────────────────────────
    updated   = [r for r in results if r["action"] == "updated"]
    delisted  = [r for r in results if r["action"] == "delisted"]
    no_change = [r for r in results if r["action"] == "no_change"]
    suggested = [r for r in results if r["action"] == "suggest_lower"]
    skipped   = [r for r in results if r["action"] == "skip"]

    log(
        f"結果: 更新{len(updated)}件 / 停止{len(delisted)}件 / "
        f"変更なし{len(no_change)}件 / 提案{len(suggested)}件 / "
        f"スキップ{len(skipped)}件 / エラー{len(errors)}件"
    )

    # 更新があった場合のみLINE通知
    if updated or delisted or errors:
        lines: List[str] = [f"価格監視チェック完了 ({len(products)} 商品)"]
        if updated:
            lines.append(f"\n⬆️ 価格引き上げ {len(updated)} 件")
            for r in updated[:5]:  # 最大5件まで列挙
                lines.append(
                    f"・{r['asin']} {r['product_name_jp'][:20]}\n"
                    f"  → ₩{r['new_price']:,.0f} (粗利{r['margin_pct']:.1f}%→更新)"
                )
        if delisted:
            lines.append(f"\n❌ 赤字停止 {len(delisted)} 件")
        if errors:
            lines.append(f"\n⚠️ エラー {len(errors)} 件")
            lines.extend(f"・{e[:80]}" for e in errors[:3])

        notify("\n".join(lines), urgent=bool(delisted or errors))

    elif suggested and not dry_run:
        # 提案のみの場合は既に個別通知済みなのでここでは通知不要
        pass
    else:
        log("通知なし（変更・エラーなし）")

    log("価格監視チェック完了")


def main() -> None:
    parser = argparse.ArgumentParser(description="Amazon JP 価格監視 → Coupang Wing 価格自動更新")
    parser.add_argument(
        "--asin", type=str, default=None,
        help="特定ASINのみチェック（カンマ区切り、例: B0002S31E2,B08KR4KWXD）"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="更新せず結果だけ表示（テスト確認用）"
    )
    parser.add_argument(
        "--check-once", action="store_true",
        help="1回だけ実行して終了（cronから呼ぶ時用）"
    )
    args = parser.parse_args()

    asin_filter: Optional[List[str]] = None
    if args.asin:
        asin_filter = [a.strip() for a in args.asin.split(",") if a.strip()]
        log(f"ASINフィルタ: {asin_filter}")

    if args.check_once:
        run_once(asin_filter=asin_filter, dry_run=args.dry_run)
        return

    # 常駐ポーリングモード（2時間間隔）
    log(f"常駐モード起動（{POLL_INTERVAL_SEC // 60} 分間隔）。Ctrl+C で停止。")
    while True:
        try:
            run_once(asin_filter=asin_filter, dry_run=args.dry_run)
        except KeyboardInterrupt:
            log("停止シグナル受信。終了します。")
            break
        except Exception as e:
            log(f"❌ 予期しないエラー: {e}")
            traceback.print_exc()
            notify(f"価格監視 予期しないエラー\n{str(e)[:200]}", urgent=True)

        log(f"次回チェックまで {POLL_INTERVAL_SEC // 60} 分待機...")
        try:
            time.sleep(POLL_INTERVAL_SEC)
        except KeyboardInterrupt:
            log("停止シグナル受信。終了します。")
            break


if __name__ == "__main__":
    main()
