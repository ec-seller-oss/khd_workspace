#!/bin/bash
# KHD 2台Mac同期: 作業フォルダ + 記憶(memory) を Drive上のgit金庫とpull/push する
# 使い方: khd_git_sync.sh pull   (入口=金庫から最新を取り出す)
#         khd_git_sync.sh push   (出口=金庫へ預ける)
# ※ 入れ子gitリポ等で失敗してもセッションを止めないよう、エラーは握りつぶして通知だけ出す
set -uo pipefail
MODE="${1:-pull}"
WORKSPACE="/Users/kikuchikenta/01_honbu_docs_automation"
MEMORY="$HOME/.claude/projects/-Users-kikuchikenta-01-honbu-docs-automation/memory"

sync_repo() {
  local dir="$1" label="$2"
  [ -d "$dir/.git" ] || { echo "[$label] gitリポなし→スキップ"; return 0; }
  cd "$dir" 2>/dev/null || return 0
  if [ "$MODE" = "pull" ]; then
    if git pull --rebase --autostash origin main >/dev/null 2>&1; then
      echo "[$label] ✅ pull(最新化)済"
    else
      git rebase --abort >/dev/null 2>&1  # 競合時は中断して壊れた状態を残さない
      echo "[$label] ⚠️ pull失敗(競合の可能性)→安全に中断。git statusで手動確認を"
    fi
  else
    git add -A 2>/dev/null
    if git diff --cached --quiet 2>/dev/null; then
      echo "[$label] 変更なし(push不要)"
    else
      git commit -m "auto-sync($(date '+%Y-%m-%d %H:%M')) [$label]" >/dev/null 2>&1
      if ! git pull --rebase --autostash origin main >/dev/null 2>&1; then
        git rebase --abort >/dev/null 2>&1
        echo "[$label] ⚠️ push前pullで競合→中断。commitは手元に残存。手動確認を"
        return 0
      fi
      if git push origin main >/dev/null 2>&1; then
        echo "[$label] ✅ commit+push(金庫へ)済"
      else
        echo "[$label] ⚠️ push失敗→手動確認を(2台同時pushは禁止)"
      fi
    fi
  fi
}

echo "=== KHD同期 ($MODE) $(date '+%Y-%m-%d %H:%M') ==="
sync_repo "$WORKSPACE" "作業フォルダ"
sync_repo "$MEMORY" "記憶"
