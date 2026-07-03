#!/bin/bash
# ============================================================================
#  KHD_設定をそろえる.command
#  新Mac / 2台目Mac 用 ワンクリック環境セットアップ（冪等・非破壊）
#
#  使い方: 新Macで下記をクローン後、Finderでこのファイルをダブルクリック:
#     git clone git@github.com:ec-seller-oss/khd_workspace.git ~/01_honbu_docs_automation
#
#  やること: KHDのClaude開発環境の「設定をそろえる」
#    ①npm設定(prefix/cache) ②PATH(~/.zshrc) ③自作CLIのsymlink
#    ④Python依存 ⑤uv ⑥Claude Code ⑦memory金庫clone ⑧検証
#
#  ⚠️ 秘密情報(~/.claude.json のトークン)・各種ログインは【書きません】。
#     最後にチェックリストで「手でやる残り」を表示します。
#  何度実行しても安全（既にある物はスキップ）。
# ============================================================================

# スクリプト自身の場所＝リポジトリ直下へ移動（ダブルクリック対応）
cd "$(dirname "$0")" || exit 1
REPO="$(pwd)"
set -u

# ---- 見た目ヘルパ ----
B=$'\033[1m'; G=$'\033[32m'; Y=$'\033[33m'; R=$'\033[31m'; C=$'\033[36m'; N=$'\033[0m'
PASS=0; WARN=0
ok()   { echo "  ${G}✅ $1${N}"; PASS=$((PASS+1)); }
warn() { echo "  ${Y}⚠️  $1${N}"; WARN=$((WARN+1)); }
info() { echo "  ${C}·${N} $1"; }
sec()  { echo ""; echo "${B}${C}$1${N}"; }

echo "${B}============================================================${N}"
echo "${B} KHD 設定をそろえる  —  Claude開発環境セットアップ${N}"
echo "${B} リポジトリ: ${REPO}${N}"
echo "${B}============================================================${N}"

# ── 0. 前提チェック ─────────────────────────────────────────
sec "0. 前提チェック"
if [ "$(uname)" != "Darwin" ]; then warn "macOS専用スクリプトです。中断します。"; exit 1; fi
if xcode-select -p >/dev/null 2>&1; then ok "Command Line Tools 入り"; else
  warn "Command Line Tools 未導入 → ターミナルで  xcode-select --install  を実行後、再度このファイルを実行"; fi
if command -v git >/dev/null 2>&1; then ok "git: $(git --version | awk '{print $3}')"; else
  warn "git が見つからない（CLT導入で入ります）"; fi

# ── 1. npm / Node 設定 ─────────────────────────────────────
sec "1. Node / npm 設定"
if command -v node >/dev/null 2>&1 && command -v npm >/dev/null 2>&1; then
  ok "node $(node -v) / npm $(npm -v)"
  npm config set prefix "$HOME/.npm-global"   >/dev/null 2>&1 && ok "npm prefix = ~/.npm-global"
  npm config set cache  "$HOME/.npm-cache-khd" >/dev/null 2>&1 && ok "npm cache  = ~/.npm-cache-khd"
else
  warn "Node.js 未導入 → https://nodejs.org でLTSを入れる（or  nvm install --lts）。入れてから再実行"
fi

# ── 2. PATH を ~/.zshrc に追記（冪等・マーカー管理）──────────
sec "2. PATH（~/.zshrc）"
ZSHRC="$HOME/.zshrc"; MARK="# >>> KHD environment (KHD_設定をそろえる) >>>"
touch "$ZSHRC"
if grep -qF "$MARK" "$ZSHRC" 2>/dev/null; then
  ok "KHD PATHブロックは設定済（スキップ）"
else
  {
    echo ""
    echo "$MARK"
    echo 'export PATH="$HOME/.npm-global/bin:$PATH"'
    echo 'export NPM_CONFIG_CACHE="$HOME/.npm-cache-khd"'
    echo '[ -f "$HOME/.local/bin/env" ] && . "$HOME/.local/bin/env"'
    echo 'export PATH="$HOME/.local/bin:$PATH"'
    echo "# <<< KHD environment <<<"
  } >> "$ZSHRC"
  ok "~/.zshrc にKHD PATHブロックを追記"
fi
# このシェルでも即有効化
export PATH="$HOME/.npm-global/bin:$HOME/.local/bin:$PATH"

# ── 3. 自作CLI（リポジトリ同梱）を ~/.local/bin へ symlink ───
sec "3. 自作CLI（khd-log / khd-scan / khd-yt-upload）"
mkdir -p "$HOME/.local/bin"
if [ -d "$REPO/scripts/cli" ]; then
  for f in "$REPO"/scripts/cli/*; do
    [ -f "$f" ] || continue
    name="$(basename "$f")"
    chmod +x "$f" 2>/dev/null
    ln -sf "$f" "$HOME/.local/bin/$name"
    ok "$name → ~/.local/bin/$name（symlink）"
  done
else
  warn "scripts/cli/ が無い（リポジトリが古い可能性）。git pull で最新化を"
fi

# ── 4. uv / uvx ────────────────────────────────────────────
sec "4. uv / uvx（Python実行環境・MCP用）"
if command -v uv >/dev/null 2>&1; then ok "uv: $(uv --version 2>/dev/null | awk '{print $2}')"; else
  info "uv未導入 → インストールします"
  if curl -LsSf https://astral.sh/uv/install.sh 2>/dev/null | sh >/dev/null 2>&1; then
    ok "uv インストール完了"; else warn "uv インストール失敗（ネット環境を確認して再実行）"; fi
fi

# ── 5. Claude Code ─────────────────────────────────────────
sec "5. Claude Code"
if command -v claude >/dev/null 2>&1; then
  ok "claude: $(claude --version 2>/dev/null | head -1)"
elif command -v npm >/dev/null 2>&1; then
  info "claude未導入 → npm で導入します"
  if npm install -g @anthropic-ai/claude-code >/dev/null 2>&1; then
    ok "Claude Code 導入完了"; else warn "Claude Code 導入失敗（npm設定/ネットを確認して再実行）"; fi
else
  warn "npm が無いため Claude Code を導入できません（先にNode.jsを）"
fi

# ── 6. Python 依存（KHDスクリプト用）────────────────────────
sec "6. Python 依存ライブラリ"
if command -v python3 >/dev/null 2>&1; then
  PKGS="python-pptx openpyxl pandas requests google-api-python-client google-auth google-auth-oauthlib google-auth-httplib2"
  info "python3 $(python3 -V 2>&1 | awk '{print $2}') に不足分を導入"
  if python3 -m pip install --user --quiet --disable-pip-version-check $PKGS >/dev/null 2>&1; then
    ok "Python依存 OK（pptx/openpyxl/pandas/google系）"
  else
    warn "一部のPython依存が入らず（$ python3 -m pip install --user $PKGS を手動で）"
  fi
else
  warn "python3 が無い（CLT導入で入ります）"
fi

# ── 7. memory 金庫（Claudeの記憶）の clone ──────────────────
sec "7. memory 金庫（Claudeの記憶）"
MEM="$HOME/.claude/projects/-Users-kikuchikenta-01-honbu-docs-automation/memory"
MEM_REMOTE="git@github.com:ec-seller-oss/khd_memory.git"
if [ -d "$MEM/.git" ]; then
  ok "memory金庫は既にclone済"
else
  mkdir -p "$(dirname "$MEM")"
  if git clone "$MEM_REMOTE" "$MEM" >/dev/null 2>&1; then
    ok "memory金庫を clone（GitHub: ec-seller-oss/khd_memory）"
  else warn "memory金庫のcloneに失敗（GitHubのSSH鍵設定を確認して再実行）"; fi
fi

# ── 8. 秘密情報・ログインの検証（書かない・確認だけ）─────────
sec "8. 秘密情報・接続の検証（※ここは自動で書きません）"
CJ="$HOME/.claude.json"
if [ -f "$CJ" ] && grep -q "NOTION_TOKEN" "$CJ" 2>/dev/null; then
  ok "~/.claude.json あり・NOTION_TOKEN 設定済（MCP接続の土台OK）"
else
  warn "~/.claude.json のトークン未設定 → 旧Mac/1Passwordから ~/.claude.json を移すこと（平文でここに書かない）"
fi

# ── 検証サマリ ─────────────────────────────────────────────
sec "セットアップ結果"
echo "  ${G}✅ 完了 ${PASS}件${N}   ${Y}⚠️ 要対応 ${WARN}件${N}"
echo ""
echo "${B}── このあと手でやる3つ（Claude導入）──${N}"
echo "  1) 新しいターミナルを開く（PATH反映のため）→  ${B}claude${N}  と打つ"
echo "  2) 画面のリンクで Claude / 各MCP(Notion等) にログイン・許可"
echo "  3) Claudeに  ${B}/company${N}  → 秘書が起動。「同期して」で最新化"
echo ""
if [ "$WARN" -gt 0 ]; then
  echo "  ${Y}⚠️ の項目を上から片付けて、もう一度このファイルをダブルクリックすればOK（何度でも安全）${N}"
else
  echo "  ${G}設定はそろいました。ターミナルで claude を起動してください。${N}"
fi
echo ""
echo "（このウィンドウは閉じてOK）"
