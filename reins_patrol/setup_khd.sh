#!/bin/bash
# setup_khd.sh — KHD自社実装のREINS巡回ツールを母艦Macに導入する
#
# 変更点(8/13): 外部提供キット(Drive)への依存を廃止。このリポジトリのファイルだけで完結する。
#               コードは自社実装なので、外部への持ち出し・共有は一切していない。
#
# 前提: Homebrew / python3 / claude CLI(任意・AI補完に使う)
# 使い方: bash reins_patrol/setup_khd.sh
set -euo pipefail

SRC=$(cd "$(dirname "$0")" && pwd)
DEST="${REINS_HOME:-$HOME/reins-patrol}"

# --- 導入前ゲート（3確認は8/13にクリア済み。再確認としてここに残す） ---
cat <<'GATE'
==================================================================
⚠️ 導入前の3確認 — yと答えない限り進みません
  (1) REINS利用規程は自動アクセスを禁止している。業者ID停止等の
      リスクを自分が引き受ける、という同意があるか
  (2) 使うのは自分がログイン権限を持つ自社の業者IDか（他人のID共有は不可）
  (3) 機構は東日本レインズか（セレクタは東日本の画面前提）
==================================================================
GATE
read -r -p "3点すべて確認済みですか? [y/N]: " ANS
if [ "${ANS:-n}" != "y" ]; then
  echo "中止しました。"
  exit 1
fi

# --- 配置 ---
mkdir -p "$DEST"
cp "$SRC/patrol.py" "$SRC/requirements.txt" "$SRC/run_with_jitter.sh" "$DEST/"
chmod +x "$DEST/run_with_jitter.sh"
mkdir -p "$DEST/downloads" "$DEST/output" "$DEST/logs" "$DEST/state"
echo "✅ 配置完了: $DEST"

# --- .env（初回のみ。既存は絶対に上書きしない＝IDが消えると再入力になる） ---
if [ ! -f "$DEST/.env" ]; then
  cp "$SRC/env.khd.example" "$DEST/.env"
  chmod 600 "$DEST/.env"
  echo "✅ .env を作成 → ID/PW/保存済み検索名を記入してください: $DEST/.env"
else
  echo "ℹ️ 既存の .env は保持しました"
fi

# --- 依存 ---
cd "$DEST"
[ -d .venv ] || python3 -m venv .venv
./.venv/bin/pip install -q --upgrade pip
./.venv/bin/pip install -q -r requirements.txt
./.venv/bin/python -m playwright install chromium

cat <<EOS

===================================================================
セットアップ完了。次は必ずこの順で、1つずつ。まとめて流さない。

  1. REINSのWeb画面で「保存済み検索」を作る（条件設計が成果を決める）
  2. $DEST/.env に ID / PW / 検索の表示名を記入
     ※検索名は画面の表示をそのままコピペ（一字違うと止まる）
  3. cd $DEST && ./.venv/bin/python patrol.py --headed --pause-after-login
     ← 初回はこれ。各ステップで止まるので画面を目視で確認する
  4. 途中で止まったら logs/ の最新スクショをClaudeに見せる
     → patrol.py 冒頭の SEL 辞書だけ直せば直る作りにしてある
  5. 通しで動いたら --headed 無しで1回。そのあと cron は1日2回から
     crontab の書式は run_with_jitter.sh の先頭コメント参照
===================================================================
EOS
