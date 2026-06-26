#!/bin/bash
# daily_digest.sh — 毎朝5:57に前日の学びをAuto memoryに反映する

export HOME=/Users/kikuchikenta
export PATH=/Users/kikuchikenta/.npm-global/bin:/usr/local/bin:/usr/bin:/bin

LOG_FILE="$HOME/01_honbu_docs_automation/scripts/logs/daily_digest.log"
MEMORY_DIR="$HOME/.claude/projects/-Users-kikuchikenta-01-honbu-docs-automation/memory"
SECRETARY_NOTES="$HOME/01_honbu_docs_automation/.company/secretary/notes"
TODAY=$(date '+%Y-%m-%d')

echo "[$(date '+%Y-%m-%d %H:%M:%S')] daily-digest 開始" >> "$LOG_FILE"

PROMPT="【daily-digest】今日の日付は${TODAY}です。

以下の作業を順番に実行してください:

1. ${MEMORY_DIR}/ のファイルを全て読む
2. 最近のClaude Codeセッションから重要な情報を抽出:
   - 新しい意思決定（何かを決めたこと）
   - 新しい学び・気づき
   - 繰り返し出てきたパターン
   - 更新が必要な既存の判断軸

3. 抽出した内容をもとに:
   a. 新しい情報があれば ${MEMORY_DIR}/ の既存ファイルを更新 or 新規追加
   b. MEMORY.md のインデックスも更新
   c. ${SECRETARY_NOTES}/${TODAY}-learnings.md に今日の学びを追記（ファイルがあれば追記、なければ新規）

4. 実行した内容を3行以内でサマリ表示

実行に際して必要なファイル操作はすべて行ってください。"

/Users/kikuchikenta/.npm-global/bin/claude \
  --print \
  --dangerously-skip-permissions \
  -p "$PROMPT" \
  >> "$LOG_FILE" 2>&1

echo "[$(date '+%Y-%m-%d %H:%M:%S')] daily-digest 完了" >> "$LOG_FILE"
