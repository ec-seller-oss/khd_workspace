# -*- coding: utf-8 -*-
"""
🤖 EC完全自動化デーモン (01_経営管理 CFO)
─────────────────────────────────────────────────────────
【このスクリプトが管理する全自動化レイヤー】

  Layer1: 価格監視  (2時間毎)  → ec_price_monitor.py
  Layer2: 受注監視  (30分毎)   → ec_order_monitor.py
  Layer3: HANIRO CSV (毎朝9時) → haniro_csv_generator.py
  Layer4: 精算照合  (毎月25日) → run_monthly_ec_pipeline.sh
  Layer5: 追跡更新  (4時間毎)  → (ec_tracking_updater.py / 将来実装)

【ゆーしが今までやっていた作業との対応】
  Wing DL (月1回)     → wing_auto_download.py     ← 実装済み(月25日cron)
  新規注文確認 (毎日)  → ec_order_monitor.py        ← 実装済み(30分毎)
  Amazon発注 (注文毎)  → ec_order_monitor.py内      ← 実装済み(注文検知と連動)
  HANIRO登録 (週次)   → haniro_csv_generator.py    ← 実装済み(毎朝9時)
  追跡更新 (都度)     → ec_tracking_updater.py     ← 将来実装

【実行方法】
  # デーモン起動（バックグラウンド常駐）
  python3 scripts/ec_full_auto_daemon.py &

  # フォアグラウンド実行（ログ確認しながら）
  python3 scripts/ec_full_auto_daemon.py --foreground

  # 起動確認
  python3 scripts/ec_full_auto_daemon.py --status

  # 停止
  kill $(cat ~/01_honbu_docs_automation/ec_daemon.pid)

【初回セットアップ完了後】
  launchd (macOS LaunchAgent) で自動起動を設定済み
  → ~/Library/LaunchAgents/com.khd.ec-daemon.plist
"""

import os, sys, time, subprocess, signal, argparse
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(__file__))

try:
    from apscheduler.schedulers.blocking import BlockingScheduler
    from apscheduler.schedulers.background import BackgroundScheduler
    from apscheduler.triggers.cron import CronTrigger
    from apscheduler.triggers.interval import IntervalTrigger
    HAS_APScheduler = True
except ImportError:
    HAS_APScheduler = False

BASE      = Path.home() / "01_honbu_docs_automation"
SCRIPTS   = BASE / "scripts"
LOG_FILE  = SCRIPTS / "logs" / "ec_daemon.log"
PID_FILE  = BASE / "ec_daemon.pid"

(SCRIPTS / "logs").mkdir(parents=True, exist_ok=True)


def ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str):
    line = f"[{ts()}] {msg}"
    print(line, flush=True)
    with open(LOG_FILE, "a") as f:
        f.write(line + "\n")


def notify_safe(msg: str, urgent: bool = False):
    try:
        from ec_notify import notify
        notify(msg, urgent=urgent)
    except Exception:
        log(f"[notify] {msg}")


def run_script(script_name: str, args: list = None, timeout: int = 300) -> bool:
    """スクリプトを子プロセスで実行。成功=True"""
    cmd = [sys.executable, str(SCRIPTS / script_name)] + (args or [])
    log(f"▶ {script_name} {' '.join(args or '')}")
    try:
        r = subprocess.run(cmd, timeout=timeout, capture_output=True, text=True)
        if r.stdout: log(r.stdout.strip()[:500])
        if r.returncode != 0:
            log(f"❌ {script_name} 失敗 (exit {r.returncode}): {r.stderr[:200]}")
            return False
        log(f"✅ {script_name} 完了")
        return True
    except subprocess.TimeoutExpired:
        log(f"⚠️ {script_name} タイムアウト ({timeout}s)")
        return False
    except Exception as e:
        log(f"❌ {script_name} 例外: {e}")
        return False


# ─────────────────────────────────────────────────────────
# 各ジョブ定義
# ─────────────────────────────────────────────────────────

def job_price_monitor():
    """Layer1: 価格監視（2時間毎）→ 赤字商品の自動価格更新"""
    log("=== 価格監視 開始 ===")
    ok = run_script("ec_price_monitor.py", ["--check-once"], timeout=600)
    if not ok:
        notify_safe("⚠️ 価格監視 失敗 — ec_price_monitor.py を確認してください", urgent=True)


def job_order_monitor():
    """Layer2: 受注監視（30分毎）→ 新規注文検知 → 販売管理表追記 → Amazon発注"""
    log("=== 受注監視 開始 ===")
    ok = run_script("ec_order_monitor.py", ["--check-once"], timeout=300)
    if not ok:
        notify_safe("⚠️ 受注監視 失敗 — ec_order_monitor.py を確認してください", urgent=True)


def job_haniro_csv():
    """Layer3: HANIRO CSV生成（毎朝9時）→ 代行業者への一括登録CSV"""
    log("=== HANIRO CSV生成 開始 ===")
    ok = run_script("haniro_csv_generator.py", ["--from-mgmt"], timeout=120)
    if ok:
        # 生成ファイルのパスをLINEで通知
        haniro_dir = Path.home() / "Library/CloudStorage/GoogleDrive-ec-seller@kikuchi-hd.net/共有ドライブ/01_個人/2025_帳票、明細/韓国輸出売上/_HANIRO登録CSV"
        csvs = sorted(haniro_dir.glob("HANIRO_batch_*.csv")) if haniro_dir.exists() else []
        if csvs:
            latest = csvs[-1]
            notify_safe(
                f"📦 HANIRO CSV生成完了\n"
                f"ファイル: {latest.name}\n"
                f"→ HANIROにCSVをアップロードしてください"
            )
    else:
        notify_safe("⚠️ HANIRO CSV生成 失敗", urgent=True)


def job_daily_summary():
    """毎日20時: 当日サマリーをLINE通知"""
    log("=== 日次サマリー ===")
    try:
        from ec_automation_db import get_stats
        stats = get_stats()
        msg = (
            f"📊 EC日次サマリー ({datetime.now().strftime('%m/%d')})\n"
            f"新規: {stats['new']}件\n"
            f"仕入済: {stats['purchased']}件\n"
            f"HANIRO登録済: {stats['haniro_registered']}件\n"
            f"出荷済: {stats['shipped']}件\n"
            f"登録商品: {stats['total_products']}品"
        )
        notify_safe(msg)
    except Exception as e:
        log(f"日次サマリー失敗: {e}")


# ─────────────────────────────────────────────────────────
# スケジューラー設定
# ─────────────────────────────────────────────────────────

SCHEDULE = [
    # (job関数, 実行間隔・cron, 説明)
    (job_order_monitor, dict(minutes=30),                    "受注監視 30分毎"),
    (job_price_monitor, dict(hours=2),                       "価格監視 2時間毎"),
    (job_haniro_csv,    dict(hour=9, minute=0),              "HANIRO CSV 毎朝9時"),
    (job_daily_summary, dict(hour=20, minute=0),             "日次サマリー 毎晩20時"),
]


def start_daemon(foreground: bool = False):
    if not HAS_APScheduler:
        print("❌ APScheduler 未インストール: pip3 install apscheduler")
        sys.exit(1)

    log("🚀 EC完全自動化デーモン 起動")
    log(f"  PID: {os.getpid()}")

    # PIDファイル保存
    PID_FILE.write_text(str(os.getpid()))

    scheduler = BlockingScheduler(timezone="Asia/Tokyo")

    for fn, trigger_args, desc in SCHEDULE:
        if "minutes" in trigger_args or "hours" in trigger_args:
            scheduler.add_job(fn, IntervalTrigger(**trigger_args), id=fn.__name__, name=desc,
                              misfire_grace_time=300)
        else:
            scheduler.add_job(fn, CronTrigger(**trigger_args), id=fn.__name__, name=desc,
                              misfire_grace_time=300)
        log(f"  ✅ スケジュール登録: {desc}")

    def shutdown(sig, _):
        log("🛑 シャットダウン中...")
        scheduler.shutdown(wait=False)
        PID_FILE.unlink(missing_ok=True)
        sys.exit(0)

    signal.signal(signal.SIGTERM, shutdown)
    signal.signal(signal.SIGINT, shutdown)

    notify_safe(f"🚀 EC自動化デーモン 起動しました\n全{len(SCHEDULE)}ジョブ稼働中")

    log("デーモン稼働中... (Ctrl+C で停止)")
    try:
        # 起動直後に全ジョブを1回実行
        log("初回実行中...")
        job_order_monitor()
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        pass
    finally:
        PID_FILE.unlink(missing_ok=True)
        log("デーモン停止")


def show_status():
    """実行状態確認"""
    if PID_FILE.exists():
        pid = PID_FILE.read_text().strip()
        print(f"🟢 デーモン稼働中 (PID: {pid})")
        print(f"   ログ: {LOG_FILE}")
        print(f"   停止: kill {pid}")
    else:
        print("🔴 デーモン停止中")
        print(f"   起動: python3 {__file__} &")
    print()
    print("設定済みスケジュール:")
    for _, _, desc in SCHEDULE:
        print(f"  • {desc}")

    # DB統計
    try:
        from ec_automation_db import get_stats
        stats = get_stats()
        print(f"\nDB状態:")
        for k, v in stats.items():
            print(f"  {k}: {v}")
    except Exception:
        pass


def setup_launchd():
    """macOS LaunchAgent plistを設定（自動起動）"""
    plist_path = Path.home() / "Library/LaunchAgents/com.khd.ec-daemon.plist"
    py = sys.executable
    script = Path(__file__).resolve()

    plist = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.khd.ec-daemon</string>

  <key>ProgramArguments</key>
  <array>
    <string>{py}</string>
    <string>{script}</string>
    <string>--foreground</string>
  </array>

  <key>RunAtLoad</key>
  <true/>

  <key>KeepAlive</key>
  <true/>

  <key>StandardOutPath</key>
  <string>{LOG_FILE}</string>
  <key>StandardErrorPath</key>
  <string>{SCRIPTS}/logs/ec_daemon_err.log</string>

  <key>EnvironmentVariables</key>
  <dict>
    <key>HOME</key>
    <string>{Path.home()}</string>
    <key>PATH</key>
    <string>/usr/local/bin:/usr/bin:/bin</string>
  </dict>
</dict>
</plist>"""

    plist_path.write_text(plist)
    print(f"✅ LaunchAgent作成: {plist_path}")
    print("   有効化: launchctl load ~/Library/LaunchAgents/com.khd.ec-daemon.plist")
    print("   無効化: launchctl unload ~/Library/LaunchAgents/com.khd.ec-daemon.plist")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="EC完全自動化デーモン")
    parser.add_argument("--foreground", action="store_true", help="フォアグラウンド実行")
    parser.add_argument("--status",     action="store_true", help="稼働状態確認")
    parser.add_argument("--setup",      action="store_true", help="LaunchAgent設定")
    parser.add_argument("--run-once",   action="store_true", help="全ジョブを1回だけ実行して終了")
    args = parser.parse_args()

    os.chdir(str(BASE))

    if args.status:
        show_status()
    elif args.setup:
        setup_launchd()
    elif args.run_once:
        log("=== 全ジョブ 1回実行 ===")
        job_order_monitor()
        job_price_monitor()
        job_haniro_csv()
        log("=== 完了 ===")
    else:
        start_daemon(foreground=args.foreground)
