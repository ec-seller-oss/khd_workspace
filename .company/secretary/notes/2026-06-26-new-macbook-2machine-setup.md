# 新MacBook 2台運用セットアップ案（2026-06-26 作成）

## 依頼
明日届く新MacBookのセットアップ。2台でClaudeを動かす／新Mac=フル家用(持ち運ばず)の使い分けは可能か、具体案をスライドで。→スライド(show_widget)で納品済み。

## 結論
- **2台同時運用は可能**。Claude Code/デスクトップは同アカウントで複数台インストール可、同時起動OK。
- 役割分担：**新Mac=母艦(家固定・フル装備・データ正本・重作業)／現Mac=機動(持ち運び・同期コピー・軽作業)**。
- 唯一の注意＝同じフォルダの2台同時編集による競合。→「書き手は母艦1台」＋同期設計で回避。

## 明日の手順（約1時間）
1. **移行アシスタント**で現Mac→新Macを丸ごとクローン（Thunderbolt直結が最速）。`~/.claude.json`(MCP/トークン)・`~/.local/bin`(khd-log)・Googleアカウントも移る。
2. **Claude再接続**：`claude`ログイン／MCP再認証(Drive/Calendar/Gmail/Tasks/Notion/LINE/TradingView)／Notion(npx)はNode必須`node -v`／Drive Desktopログイン同期。
3. **2台同期の設計**：共有する脳=①`~/01_honbu_docs_automation`(.company含) ②Claudeメモリ ③`~/.claude/CLAUDE.md`。
   - 推奨：プライベートGit（開始pull/終了push・履歴で消失ゼロ＝[[project_bysell_internalization]]の"複利の源泉"を守る）。
   - 簡易：Google Drive Desktop同期（小ファイル多数は競合・遅延注意）。
   - 鉄則：書き手1台（母艦マスター）。

## 区分
内務（環境構築）。新ハード到着の一度きり設定＝凍結対象外だが90分上限で決め切り。

## 次アクション候補
- Git同期の具体手順（リポジトリ作成→pull/push）を出す。
- 本スライドをGoogleスライド化してDriveリンクで渡す（[[feedback_deliverable_gslides_rule]]）。
