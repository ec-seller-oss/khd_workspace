#!/bin/bash
# ═══════════════════════════════════════════════════════════
# 🤖 月次EC粗利パイプライン 完全自動実行スクリプト
# 毎月25日 cron から呼ばれる（または手動実行）
# ═══════════════════════════════════════════════════════════
#
# 実行順:
#   1. wing_auto_download.py   ← ゆーしの仕事を代替
#   2. ec_profit_pipeline.py   ← 販売管理表集計
#   3. ec_product_strategy.py  ← ASIN別4区分
#   4. ec_settlement_recon_v2.py ← 精算実額照合
#   5. ec_dashboard_build.py   ← ダッシュボード更新
#   6. ec_dashboard_recon_patch.py ← 実精算タブ更新
#   7. ec_handoff_03.py        ← 03申し送り生成
#   ★ LINE通知（成功 or 失敗）
#
# 手動実行:
#   bash scripts/run_monthly_ec_pipeline.sh
#   bash scripts/run_monthly_ec_pipeline.sh --skip-download  # Wing DLをスキップ
# ═══════════════════════════════════════════════════════════

export HOME=/Users/kikuchikenta
export PATH=/usr/local/bin:/usr/bin:/bin:$HOME/.local/bin:$HOME/.npm-global/bin

SCRIPT_DIR="$HOME/01_honbu_docs_automation/scripts"
WORK_DIR="$HOME/01_honbu_docs_automation"
LOG_FILE="$SCRIPT_DIR/logs/monthly_ec_pipeline.log"
TODAY=$(date '+%Y-%m-%d %H:%M:%S')
SKIP_DOWNLOAD=false

# 引数パース
for arg in "$@"; do
  case $arg in
    --skip-download) SKIP_DOWNLOAD=true ;;
  esac
done

# ログディレクトリ確認
mkdir -p "$SCRIPT_DIR/logs"

log() {
  echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" | tee -a "$LOG_FILE"
}

# LINE通知（失敗時のみ緊急フラグ）
notify() {
  python3 "$SCRIPT_DIR/ec_notify.py" "$1" 2>/dev/null || true
}

notify_urgent() {
  python3 -c "
import sys
sys.path.insert(0, '$SCRIPT_DIR')
from ec_notify import notify
notify('''$1''', urgent=True)
" 2>/dev/null || true
}

cd "$WORK_DIR"

log "═══ 月次EC粗利パイプライン 開始 ═══"
log "作業ディレクトリ: $WORK_DIR"
FAILED_STEPS=""
STEP_RESULTS=""

# ─────────────────────────────────────────────────────────
# STEP 1: Wing 自動ダウンロード
# ─────────────────────────────────────────────────────────
if [ "$SKIP_DOWNLOAD" = "false" ]; then
  log "▶ STEP1: Wing 自動DL"
  if python3 "$SCRIPT_DIR/wing_auto_download.py" >> "$LOG_FILE" 2>&1; then
    log "✅ STEP1 完了"
    STEP_RESULTS="${STEP_RESULTS}①Wing DL:✅\n"
  else
    log "❌ STEP1 失敗（Wing DL）→ スキップして続行"
    FAILED_STEPS="$FAILED_STEPS STEP1(Wing-DL)"
    STEP_RESULTS="${STEP_RESULTS}①Wing DL:❌\n"
    # Wing DL失敗でも集計は前回分で続行（致命的ではない）
  fi
else
  log "⏭ STEP1: Wing DL スキップ (--skip-download)"
  STEP_RESULTS="${STEP_RESULTS}①Wing DL:スキップ\n"
fi

# ─────────────────────────────────────────────────────────
# STEP 2: 粗利集計パイプライン
# ─────────────────────────────────────────────────────────
log "▶ STEP2: ec_profit_pipeline.py"
if python3 "$SCRIPT_DIR/ec_profit_pipeline.py" >> "$LOG_FILE" 2>&1; then
  log "✅ STEP2 完了"
  STEP_RESULTS="${STEP_RESULTS}②粗利集計:✅\n"
else
  log "❌ STEP2 失敗"
  FAILED_STEPS="$FAILED_STEPS STEP2(profit-pipeline)"
  STEP_RESULTS="${STEP_RESULTS}②粗利集計:❌\n"
fi

# ─────────────────────────────────────────────────────────
# STEP 3: 商品戦略分析
# ─────────────────────────────────────────────────────────
log "▶ STEP3: ec_product_strategy.py"
if python3 "$SCRIPT_DIR/ec_product_strategy.py" >> "$LOG_FILE" 2>&1; then
  log "✅ STEP3 完了"
  STEP_RESULTS="${STEP_RESULTS}③商品戦略:✅\n"
else
  log "❌ STEP3 失敗"
  FAILED_STEPS="$FAILED_STEPS STEP3(product-strategy)"
  STEP_RESULTS="${STEP_RESULTS}③商品戦略:❌\n"
fi

# ─────────────────────────────────────────────────────────
# STEP 4: 精算実額照合
# ─────────────────────────────────────────────────────────
log "▶ STEP4: ec_settlement_recon_v2.py"
if python3 "$SCRIPT_DIR/ec_settlement_recon_v2.py" >> "$LOG_FILE" 2>&1; then
  log "✅ STEP4 完了"
  STEP_RESULTS="${STEP_RESULTS}④精算照合:✅\n"
else
  log "❌ STEP4 失敗"
  FAILED_STEPS="$FAILED_STEPS STEP4(recon-v2)"
  STEP_RESULTS="${STEP_RESULTS}④精算照合:❌\n"
fi

# ─────────────────────────────────────────────────────────
# STEP 5+6: ダッシュボード更新
# ─────────────────────────────────────────────────────────
log "▶ STEP5: ec_dashboard_build.py"
if python3 "$SCRIPT_DIR/ec_dashboard_build.py" >> "$LOG_FILE" 2>&1; then
  log "✅ STEP5 完了"
  STEP_RESULTS="${STEP_RESULTS}⑤DB更新:✅\n"
  # ダッシュボードが成功した場合のみパッチ実行
  log "▶ STEP5b: ec_dashboard_recon_patch.py"
  if python3 "$SCRIPT_DIR/ec_dashboard_recon_patch.py" >> "$LOG_FILE" 2>&1; then
    log "✅ STEP5b 完了"
    STEP_RESULTS="${STEP_RESULTS}⑤b実精算タブ:✅\n"
  else
    log "⚠️ STEP5b 失敗（実精算タブ）"
    FAILED_STEPS="$FAILED_STEPS STEP5b(recon-patch)"
    STEP_RESULTS="${STEP_RESULTS}⑤b実精算タブ:❌\n"
  fi
else
  log "❌ STEP5 失敗"
  FAILED_STEPS="$FAILED_STEPS STEP5(dashboard)"
  STEP_RESULTS="${STEP_RESULTS}⑤DB更新:❌\n"
fi

# ─────────────────────────────────────────────────────────
# STEP 7: 03申し送り更新
# ─────────────────────────────────────────────────────────
log "▶ STEP7: ec_handoff_03.py"
if python3 "$SCRIPT_DIR/ec_handoff_03.py" >> "$LOG_FILE" 2>&1; then
  log "✅ STEP7 完了"
  STEP_RESULTS="${STEP_RESULTS}⑦03申し送り:✅\n"
else
  log "❌ STEP7 失敗"
  FAILED_STEPS="$FAILED_STEPS STEP7(handoff)"
  STEP_RESULTS="${STEP_RESULTS}⑦03申し送り:❌\n"
fi

# ─────────────────────────────────────────────────────────
# 結果集計 → LINE通知
# ─────────────────────────────────────────────────────────
RESULTS_TEXT=$(echo -e "$STEP_RESULTS")

if [ -z "$FAILED_STEPS" ]; then
  log "═══ 全ステップ完了 ═══"
  MSG="✅ 月次EC粗利パイプライン 完了 ($(date '+%m月'))
${RESULTS_TEXT}
▶ ダッシュボード確認:
https://docs.google.com/spreadsheets/d/1QjyPPOto7J1HiqA_Zb9-UIOe_FQZyqAGSn321R37Tzo/edit"
  notify "$MSG"
else
  log "═══ 一部ステップ失敗: $FAILED_STEPS ═══"
  MSG="⚠️ 月次EC粗利パイプライン 一部失敗 ($(date '+%m月'))
${RESULTS_TEXT}
失敗: $FAILED_STEPS
▶ ログ: $LOG_FILE"
  notify_urgent "$MSG"
  exit 1
fi

log "完了"
