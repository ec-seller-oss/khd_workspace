#!/bin/bash
# 秘書ノート/ルール/メモリファイルへの書き込みを検知し、khd-logでのNotion記録漏れを防ぐリマインダーhook。
# PostToolUse(Write|Edit)で発火。対象パス以外は無出力(何もしない)。
set -uo pipefail
f=$(jq -r ".tool_input.file_path // empty")
case "$f" in
  *.company/secretary/CLAUDE.md|*.company/CLAUDE.md|*.company/secretary/notes/*.md|*/memory/*.md)
    echo '{"hookSpecificOutput":{"hookEventName":"PostToolUse","additionalContext":"🔔リマインダー: 秘書ノート/ルール/メモリを更新した。この単位の作業をkhd-logでNotion作業ログDBにも記録したか確認せよ(まだなら今のうちにkhd-log実行。使い方: khd-log タスク名 カテゴリ 内容 [URL])。"}}'
    ;;
esac
