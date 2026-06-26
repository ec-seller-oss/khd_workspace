#!/bin/bash
# property_inbox_check.sh — 毎朝8:07に物件受け箱を自動チェック

export HOME=/Users/kikuchikenta
export PATH=/Users/kikuchikenta/.npm-global/bin:/usr/local/bin:/usr/bin:/bin

LOG_FILE="$HOME/01_honbu_docs_automation/scripts/logs/property_inbox.log"
TODAY=$(date '+%Y-%m-%d')

echo "[$(date '+%Y-%m-%d %H:%M:%S')] property-inbox-check 開始" >> "$LOG_FILE"

PROMPT="【物件受け箱チェック】今日の日付は${TODAY}です。

/property-inbox-check スキルの手順に従って以下を実行してください:

1. Google Drive 物件受け箱フォルダ（ID: 1duT8j7Q_ogmZDLZrSbuOR07YmuMZV3YX）の内容を確認し、ファイル名に「✅処理済み_」が付いていない新着ファイルを特定する

2. Gmail MCP が使える場合: 件名または本文に「物件」「売り」「不動産」「利回り」「㎡」「万円」を含む過去24時間のメールを確認する

3. 新着がある場合: 各ファイル/メールの内容を読み取り、玉川式で初期評価（CF率・土地値割合・融資可否）を行い、01_物件検討DB（ID: 1-mf4JxVXLyghcDcyfxOnh3kNYp7Wrfgv4neGHjyEOlw）に追記する

4. 処理結果を以下フォーマットで出力する:
   新着ありの場合:
   「📨 物件受け箱チェック完了
   新着: X件処理 / DBに追記: X件
   [物件名] [所在地] 利回り X.X% → [判断]」

   新着なしの場合:
   「📭 物件受け箱: 新着なし」

MCP（Drive/Gmail）が使えない場合は「スキップ」と表示して処理を続けてください。"

/Users/kikuchikenta/.npm-global/bin/claude \
  --print \
  --dangerously-skip-permissions \
  -p "$PROMPT" \
  >> "$LOG_FILE" 2>&1

echo "[$(date '+%Y-%m-%d %H:%M:%S')] property-inbox-check 完了" >> "$LOG_FILE"
