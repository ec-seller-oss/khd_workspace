# -*- coding: utf-8 -*-
"""
🗄️ EC自動化 状態DB (SQLite) — 01_経営管理 CFO
─────────────────────────────────────────────
全自動化パイプラインの「神経中枢」。
受注・仕入・出荷・価格の状態をすべてここで管理する。

テーブル:
  orders    … Coupang注文の全ライフサイクル
  products  … 出品商品 (ASIN → Coupang掲載情報)
  price_log … 価格変動履歴
  purchase_log … Amazon発注ログ
"""
import sqlite3, os
from contextlib import contextmanager
from datetime import datetime

DB_PATH = os.path.expanduser(
    "~/01_honbu_docs_automation/ec_automation.db"
)

SCHEMA = """
-- ─────────────────────────────────────────────────────────
-- 受注テーブル（Coupang注文 → Amazon発注 → HANIRO → 追跡）
-- ─────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS orders (
    coupang_order_no  TEXT PRIMARY KEY,  -- Coupang注文番号 (B列)
    account           TEXT,              -- クーパン1 / クーパン2
    status            TEXT DEFAULT 'new',
    -- new → purchased → haniro_registered → shipped → delivered → settled

    -- 商品情報
    asin              TEXT,
    product_name_jp   TEXT,
    product_name_kr   TEXT,
    quantity          INTEGER DEFAULT 1,
    sale_price_krw    REAL,
    sale_date         TEXT,
    ship_deadline     TEXT,             -- 発送期限 (V列)

    -- 顧客情報
    customer_name     TEXT,
    postal_code       TEXT,
    address           TEXT,
    phone             TEXT,
    customs_no        TEXT,             -- 個人通関番号
    delivery_msg      TEXT,

    -- 仕入れ情報
    amazon_account    TEXT,             -- Amazon1 / Amazon2
    amazon_url        TEXT,
    amazon_order_no   TEXT,             -- O列に書くやつ
    cost_jpy          REAL,
    purchased_at      TEXT,

    -- HANIRO / 代行
    haniro_registered INTEGER DEFAULT 0,
    haniro_batch_id   TEXT,             -- CSVアップロードのバッチID

    -- 追跡
    tracking_no       TEXT,
    tracking_updated  INTEGER DEFAULT 0,
    coupang_tracking_pushed INTEGER DEFAULT 0,

    -- タイムスタンプ
    created_at        TEXT DEFAULT (datetime('now','localtime')),
    updated_at        TEXT DEFAULT (datetime('now','localtime'))
);

-- ─────────────────────────────────────────────────────────
-- 商品テーブル（ASIN → Coupang掲載情報 + 価格監視）
-- ─────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS products (
    asin              TEXT PRIMARY KEY,
    coupang_item_id   TEXT,             -- CoupangのitemId (Open API or Wing)
    product_name_jp   TEXT,
    product_name_kr   TEXT,

    -- 価格
    amazon_price_jpy  REAL,             -- 最新のAmazon価格
    coupang_price_krw REAL,             -- 現在のCoupang掲載価格
    fx_rate           REAL DEFAULT 0.11,-- 現在の為替レート(円→ウォン近似)
    markup_rate       REAL DEFAULT 1.25,-- 仕入れ価格に対する掛け率

    -- 在庫状況
    amazon_in_stock   INTEGER DEFAULT 1,
    coupang_active    INTEGER DEFAULT 1,

    -- 採算ライン
    min_margin_pct    REAL DEFAULT 5.0, -- 最低粗利率(%) 下回ったらアラート
    auto_price_update INTEGER DEFAULT 1,-- 1=価格自動更新

    -- 監視
    last_price_check  TEXT,
    last_listed_at    TEXT,
    price_check_fail  INTEGER DEFAULT 0,

    created_at        TEXT DEFAULT (datetime('now','localtime')),
    updated_at        TEXT DEFAULT (datetime('now','localtime'))
);

-- ─────────────────────────────────────────────────────────
-- 価格ログ（変動の全履歴）
-- ─────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS price_log (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    asin              TEXT,
    checked_at        TEXT,
    amazon_price_jpy  REAL,
    coupang_price_krw REAL,
    fx_rate           REAL,
    margin_pct        REAL,
    action_taken      TEXT   -- 'updated' / 'alert_sent' / 'no_change' / 'delisted'
);

-- ─────────────────────────────────────────────────────────
-- Amazon発注ログ（購入履歴）
-- ─────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS purchase_log (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    coupang_order_no  TEXT,
    asin              TEXT,
    amazon_account    TEXT,
    amazon_order_no   TEXT,
    cost_jpy          REAL,
    status            TEXT,  -- 'success' / 'failed' / 'skipped'
    note              TEXT,
    purchased_at      TEXT DEFAULT (datetime('now','localtime'))
);

-- ─────────────────────────────────────────────────────────
-- HANIROバッチログ
-- ─────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS haniro_batch_log (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id          TEXT,
    csv_path          TEXT,
    order_count       INTEGER,
    uploaded          INTEGER DEFAULT 0,
    created_at        TEXT DEFAULT (datetime('now','localtime'))
);
"""

INDEXES = """
CREATE INDEX IF NOT EXISTS idx_orders_status ON orders(status);
CREATE INDEX IF NOT EXISTS idx_orders_asin   ON orders(asin);
CREATE INDEX IF NOT EXISTS idx_price_log_asin ON price_log(asin, checked_at);
"""


@contextmanager
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    """初回DB作成（冪等）"""
    with get_conn() as conn:
        conn.executescript(SCHEMA)
        conn.executescript(INDEXES)
    print(f"✅ DB初期化完了: {DB_PATH}")


# ────── CRUD ヘルパー ──────────────────────────────────

def upsert_order(data: dict):
    """注文の挿入 or 更新"""
    data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cols = list(data.keys())
    placeholders = ", ".join(["?"] * len(cols))
    updates = ", ".join(f"{c}=excluded.{c}" for c in cols if c != "coupang_order_no")
    sql = (
        f"INSERT INTO orders ({', '.join(cols)}) VALUES ({placeholders}) "
        f"ON CONFLICT(coupang_order_no) DO UPDATE SET {updates}"
    )
    with get_conn() as conn:
        conn.execute(sql, list(data.values()))


def update_order_status(coupang_order_no: str, status: str, **kwargs):
    """注文ステータス更新"""
    kwargs["status"] = status
    kwargs["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sets = ", ".join(f"{k}=?" for k in kwargs)
    with get_conn() as conn:
        conn.execute(
            f"UPDATE orders SET {sets} WHERE coupang_order_no=?",
            list(kwargs.values()) + [coupang_order_no]
        )


def get_orders_by_status(status: str) -> list:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM orders WHERE status=? ORDER BY created_at",
            [status]
        ).fetchall()
    return [dict(r) for r in rows]


def get_unregistered_haniro_orders() -> list:
    """HANIRO CSV未登録の購入済み注文"""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM orders WHERE status='purchased' AND haniro_registered=0 ORDER BY purchased_at"
        ).fetchall()
    return [dict(r) for r in rows]


def upsert_product(data: dict):
    data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cols = list(data.keys())
    placeholders = ", ".join(["?"] * len(cols))
    updates = ", ".join(f"{c}=excluded.{c}" for c in cols if c != "asin")
    sql = (
        f"INSERT INTO products ({', '.join(cols)}) VALUES ({placeholders}) "
        f"ON CONFLICT(asin) DO UPDATE SET {updates}"
    )
    with get_conn() as conn:
        conn.execute(sql, list(data.values()))


def log_price(asin, amazon_price, coupang_price, fx, margin, action):
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO price_log (asin,checked_at,amazon_price_jpy,coupang_price_krw,fx_rate,margin_pct,action_taken) "
            "VALUES (?,datetime('now','localtime'),?,?,?,?,?)",
            [asin, amazon_price, coupang_price, fx, margin, action]
        )


def log_purchase(coupang_order_no, asin, amazon_account, amazon_order_no, cost_jpy, status, note=""):
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO purchase_log (coupang_order_no,asin,amazon_account,amazon_order_no,cost_jpy,status,note) "
            "VALUES (?,?,?,?,?,?,?)",
            [coupang_order_no, asin, amazon_account, amazon_order_no, cost_jpy, status, note]
        )


def get_stats() -> dict:
    with get_conn() as conn:
        stats = {}
        for status in ["new", "purchased", "haniro_registered", "shipped", "settled"]:
            cnt = conn.execute(
                "SELECT COUNT(*) FROM orders WHERE status=?", [status]
            ).fetchone()[0]
            stats[status] = cnt
        stats["total_products"] = conn.execute("SELECT COUNT(*) FROM products").fetchone()[0]
        return stats


if __name__ == "__main__":
    init_db()
    stats = get_stats()
    print("DB統計:", stats)
