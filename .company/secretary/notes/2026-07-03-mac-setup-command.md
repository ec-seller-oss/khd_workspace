---
date: "2026-07-03"
type: 成果物_ツール
本部: "秘書室"
related: "[[feedback_multi_mac_session]] [[2026-06-22-新mac移行マニュアル]]"
---

# 新Mac onboarding を3ステップに畳む：KHD_設定をそろえる.command

## 菊池指示
「Clone → KHD_設定をそろえる.commandダブルクリック → Claude導入」の3ステップにしたい＝その.commandを作れ。

## 背景（従来との違い）
- 従来＝Apple移行アシスタントで丸ごとコピー＋再ログイン3点（2026-06-22 移行マニュアル）。
- 新方式＝**①git clone ②`.command`ダブルクリック ③claude起動** の再現可能フロー。名前どおり「設定をそろえる」＝冪等トップアップ（移行アシスタント後でも素の状態でも動く）。

## 作ったもの
1. **`KHD_設定をそろえる.command`**（リポジトリ直下・実行権限付き・ダブルクリック起動）
   - 0前提(CLT/git) →1 Node/npm(prefix=~/.npm-global, cache=~/.npm-cache-khd) →2 PATH(~/.zshrcにマーカー付きブロック冪等追記) →3 自作CLIをsymlink →4 uv →5 Claude Code →6 Python依存(pptx/openpyxl/pandas/google系) →7 memory金庫clone →8 秘密情報の検証。
   - **秘密情報は絶対に書かない**：~/.claude.json のトークンは検証のみ。無ければ「旧Mac/1Passwordから移せ」と警告表示（平文禁止ルール遵守）。
   - 冪等・非破壊・非fatal（1項目失敗しても続行し最後に⚠️件数表示）。最後に「手でやる3つ(claude起動/ログイン/company)」を提示。
2. **`scripts/cli/`** に自作CLI3本を同梱（khd-log/khd-scan/khd-yt-upload）。
   - これまで~/.local/binのみ＝cloneで復元されなかった穴を塞いだ。3本とも実トークン値なし（token類は~/.claude.json・~/.config/khd/の外部参照）を厳密スキャンで確認済。

## 検証（実機テスト）
- `bash -n` 構文OK。実機フル実行＝**14項目✅ / 警告0**。
- バグ1件修正：ヘルパー関数`head()`が`head`コマンドを上書き→Step5のバージョン表示が"-1"化。関数を`sec()`にリネームで解消（`✅ claude: 2.1.185 (Claude Code)`）。
- 2回目実行でzshrcブロックは「設定済スキップ」＝冪等確認。

## 新Macでの使い方（3ステップ）
1. `git clone <Drive>/KHD_git_remote/khd_workspace.git ~/01_honbu_docs_automation`
2. Finderで `KHD_設定をそろえる.command` をダブルクリック（⚠️初回はGatekeeperで「開けない」→右クリック→開く、で許可）
3. 新ターミナルで `claude` → ログイン/MCP許可 → `/company`

## 残る手動ピース（設計上わざと自動化しない）
- ~/.claude.json のトークン（秘密）＝旧Mac/1Passwordから移す。
- Claude本体ログイン・Google各アカウント(Drive/Gmail/Calendar)ログイン。
- ~/.claude/ のグローバル設定(CLAUDE.md/agents/skills/plugins)＝現状は移行アシスタント/別途同期。**今後の宿題**：これもリポジトリ同梱 or 別金庫にすれば完全再現になる。
