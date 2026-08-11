#!/bin/bash
# tmux_khd.sh — tmux窓と会話セッションの混同を防ぐ命名規約ヘルパー(伊東式)
# 窓名を khd-<用途> に統一し、1窓=1会話セッションを守る。
#
# 使い方:
#   ./ito_style/tmux_khd.sh new <用途>   # 窓 "khd-<用途>" を作成して切替
#   ./ito_style/tmux_khd.sh ls           # khd- 窓の一覧
#   ./ito_style/tmux_khd.sh doctor       # 規約違反・空窓の検出
set -u

SESSION="khd"

usage() { grep '^#   ' "$0" | sed 's/^#   //'; exit 1; }

ensure_session() {
  tmux has-session -t "$SESSION" 2>/dev/null || tmux new-session -d -s "$SESSION" -n "khd-本部"
}

case "${1:-}" in
  new)
    [ -n "${2:-}" ] || usage
    NAME="khd-$2"
    ensure_session
    if tmux list-windows -t "$SESSION" -F '#{window_name}' | grep -qx "$NAME"; then
      echo "窓 $NAME は既に存在します。切り替えます。"
    else
      tmux new-window -t "$SESSION" -n "$NAME"
      echo "窓 $NAME を作成しました。この窓で claude を起動してください(1窓=1会話)。"
    fi
    tmux select-window -t "$SESSION:$NAME"
    [ -z "${TMUX:-}" ] && tmux attach -t "$SESSION"
    ;;
  ls)
    tmux list-windows -a -F '#{session_name}:#{window_name}  (#{pane_current_command})' 2>/dev/null \
      | grep 'khd-' || echo "khd- 窓はありません"
    ;;
  doctor)
    echo "[命名規約違反の窓(khd- で始まらない)]"
    tmux list-windows -a -F '#{session_name}:#{window_name}' 2>/dev/null \
      | grep -v ':khd-' || echo "  なし"
    echo "[空窓(shellのまま放置)]"
    tmux list-windows -a -F '#{session_name}:#{window_name} #{pane_current_command}' 2>/dev/null \
      | awk '$2 ~ /^(bash|zsh|sh)$/ && $1 ~ /khd-/ {print "  " $1 "  ← 未使用。閉じるか用途を割り当てる"}' \
      || true
    ;;
  *) usage ;;
esac
