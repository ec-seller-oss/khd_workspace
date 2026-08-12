#!/bin/bash
# setup_khd.sh — 伊東さん提供のREINS巡回ツール(reins-patrol)をKHDのMacに導入する
#
# 前提:
#   - このMacでGoogle Drive(ec-seller@kikuchi-hd.net)がデスクトップ同期されている
#   - キット一式がマイドライブの「reins-patrol_知人PC向け_20260813」フォルダにある
#   - claude CLI がログイン済み / Homebrew あり
#
# 使い方:  bash reins_patrol/setup_khd.sh
set -euo pipefail

DRIVE_BASE="$HOME/Library/CloudStorage/GoogleDrive-ec-seller@kikuchi-hd.net/マイドライブ"
KIT_DIR=$(find "$DRIVE_BASE" -maxdepth 3 -type d -name "reins-patrol_知人PC向け*" 2>/dev/null | head -1)
DEST="$HOME/reins-patrol"

if [ -z "$KIT_DIR" ]; then
  echo "❌ Drive上にキットフォルダ(reins-patrol_知人PC向け*)が見つかりません。"
  echo "   Drive同期が済んでいるか、フォルダ名を確認してください。"
  exit 1
fi
echo "✅ キット発見: $KIT_DIR"

# --- 導入前ゲート(伊東さんの引き継ぎメモ1章。ここを飛ばさない) ---
cat <<'GATE'
==================================================================
⚠️ 導入前の3確認(引き継ぎメモ1章) — yと答えない限り進みません
  (1) REINS利用規程は自動アクセスを禁止。業者ID停止・行政処分等の
      リスクを自分(と、IDがTAW名義なら福井さん)が引き受ける同意があるか
  (2) 使うのは自分がログイン権限を持つREINS業者IDか(他人のID共有は不可)
  (3) 機構は東日本レインズか(伊東さんと同じ=セレクタがそのまま使える見込み)
==================================================================
GATE
read -r -p "3点すべて確認済みですか? [y/N]: " ANS
if [ "${ANS:-n}" != "y" ]; then
  echo "中止しました。確認が取れてから再実行してください。"
  exit 1
fi

# --- コピー(もらってよいファイルだけ) ---
mkdir -p "$DEST"
for f in patrol.py requirements.txt run_with_jitter.sh README.md; do
  cp "$KIT_DIR/$f" "$DEST/"
done
# .env.example は env.example.txt の場合もある
if [ -f "$KIT_DIR/.env.example" ]; then
  cp "$KIT_DIR/.env.example" "$DEST/.env.example"
elif [ -f "$KIT_DIR/env.example.txt" ]; then
  cp "$KIT_DIR/env.example.txt" "$DEST/.env.example"
fi
chmod +x "$DEST/run_with_jitter.sh"
echo "✅ コピー完了: $DEST"

# --- KHDパッチ(引き継ぎメモ5章のハードコード潰し) ---
python3 - "$DEST/patrol.py" <<'PYEOF'
import re, sys
p = sys.argv[1]
src = open(p, encoding="utf-8").read()

# (5-1) pdftoppm を PATH から自動解決(環境変数 PDFTOPPM_BIN があれば優先)
src = src.replace(
    'PDFTOPPM_BIN = "/opt/homebrew/bin/pdftoppm"',
    'import shutil as _shutil\n'
    'PDFTOPPM_BIN = os.environ.get("PDFTOPPM_BIN") or _shutil.which("pdftoppm") '
    'or "/opt/homebrew/bin/pdftoppm"',
)
# (5-2) 提供元の私物ツールのデフォルトパスを無効化(未設定ならスキップされる側に倒す)
src = src.replace(
    '"GMAIL_SENDER_BIN", "/Users/ItoSeiichi/tools/gmail-sender/.venv/bin/gmail-send"',
    '"GMAIL_SENDER_BIN", ""',
)
src = src.replace('"GMAIL_SENDER_ACCOUNT", "ichifuji_0317"', '"GMAIL_SENDER_ACCOUNT", ""')
src = src.replace(
    '"DRIVE_UPLOADER_BIN", "/Users/ItoSeiichi/tools/drive-uploader/.venv/bin/drive-upload"',
    '"DRIVE_UPLOADER_BIN", ""',
)
open(p, "w", encoding="utf-8").write(src)
print("✅ patrol.py にKHDパッチ適用(pdftoppm自動解決・私物パス無効化)")
PYEOF

# --- .env 生成(初回のみ。上書きしない) ---
if [ ! -f "$DEST/.env" ]; then
  REPO_DIR=$(cd "$(dirname "$0")" && pwd)
  cp "$REPO_DIR/env.khd.example" "$DEST/.env"
  echo "✅ .env を作成しました → ID/PW/保存済み検索名を記入してください: $DEST/.env"
else
  echo "ℹ️ 既存の .env は保持しました"
fi

# --- 依存導入 ---
command -v pdftoppm >/dev/null || brew install poppler
cd "$DEST"
python3 -m venv .venv
./.venv/bin/pip install -q -r requirements.txt
./.venv/bin/python -m playwright install chromium
echo ""
echo "==================================================================="
echo "セットアップ完了。次の手順(1つずつ・まとめて流さない):"
echo "  1. REINSのWeb画面で保存済み検索を作る(条件設計が成果を決める)"
echo "  2. $DEST/.env にID/PW/検索名を記入"
echo "  3. ./.venv/bin/python patrol.py --headed   ← ログインが通るかだけ確認"
echo "  4. 引き継ぎメモ6章の順で1ステップずつ確認"
echo "  5. 通ったら crontab は1日2回から(勝手に増やさない)"
echo "==================================================================="
