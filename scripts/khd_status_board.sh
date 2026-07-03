#!/bin/bash
# セッション開始時に「現在地ボード」を頭に表示する（SessionStartフック用）
# どのMac・どのセッションでも、頭を見れば各タスクの進捗が分かるようにするため。
BOARD="/Users/kikuchikenta/01_honbu_docs_automation/.company/secretary/現在地ボード.md"
if [ -f "$BOARD" ]; then
  echo "───────── 📍 現在地ボード（このセッション頭のタスク進捗）─────────"
  cat "$BOARD"
  echo "──────────────────────────────────────────────"
  echo "※タスクを進めたら 現在地ボード.md の状態と次アクションを更新すること"
else
  echo "(現在地ボード.md が見つかりません)"
fi
