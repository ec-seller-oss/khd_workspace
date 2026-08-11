#!/bin/bash
# winq.sh — KHD作業フォルダの全体進捗を1コマンドで俯瞰する(伊東式)
# 使い方: ./ito_style/winq.sh
set -u

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT" || exit 1
TODAY=$(date +%Y-%m-%d)

hr() { printf '%s\n' "----------------------------------------"; }

echo "== KHD winq  ($(date '+%Y-%m-%d %H:%M')) =="

hr
echo "[git]"
BRANCH=$(git rev-parse --abbrev-ref HEAD 2>/dev/null)
if [ -n "$BRANCH" ]; then
  echo "  ブランチ: $BRANCH"
  DIRTY=$(git status --porcelain | wc -l | tr -d ' ')
  echo "  未コミット: ${DIRTY}件"
  UPSTREAM=$(git rev-parse --abbrev-ref --symbolic-full-name '@{u}' 2>/dev/null)
  if [ -n "$UPSTREAM" ]; then
    AHEAD=$(git rev-list --count '@{u}..HEAD' 2>/dev/null || echo 0)
    echo "  未プッシュ: ${AHEAD}コミット (upstream: $UPSTREAM)"
  else
    echo "  未プッシュ: upstream未設定"
  fi
else
  echo "  gitリポジトリではありません"
fi

hr
echo "[棚卸しフォルダ 残件]"
FOUND=0
for d in _削除候補_*/ _要ファイリング_*/; do
  [ -d "$d" ] || continue
  FOUND=1
  N=$(find "$d" -type f | wc -l | tr -d ' ')
  echo "  $d ${N}件  ← レビュー待ち"
done
[ "$FOUND" -eq 0 ] && echo "  なし(棚卸しは file_triage.py で開始)"

hr
echo "[秘書室 未完了タスク]"
if [ -d ".company/secretary/todos" ]; then
  MATCHES=$(grep -rn '\- \[ \]' .company/secretary/todos 2>/dev/null | head -15)
  if [ -n "$MATCHES" ]; then
    echo "$MATCHES" | sed 's/^/  /'
  else
    echo "  未完了タスクなし"
  fi
else
  echo "  .company/secretary/todos なし"
fi

hr
echo "[肥大ファイル TOP10(直下)]"
ls -S -l . 2>/dev/null | awk '$1 ~ /^-/ {printf "  %8.1fMB  %s\n", $5/1048576, $NF}' | head -10

hr
echo "[本日生成・更新されたファイル]"
TODAY_FILES=$(find . -maxdepth 1 -type f -newermt "$TODAY" 2>/dev/null | sed 's|^\./||' | head -20)
if [ -n "$TODAY_FILES" ]; then
  echo "$TODAY_FILES" | sed 's/^/  /'
else
  echo "  なし"
fi

hr
echo "[tmux 窓(khd-)]"
if command -v tmux >/dev/null 2>&1 && tmux ls >/dev/null 2>&1; then
  tmux list-windows -a -F '  #{session_name}:#{window_name}  (#{pane_current_command})' 2>/dev/null | grep 'khd-' || echo "  khd- 窓なし"
else
  echo "  tmux未起動"
fi
