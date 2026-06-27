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

## ★git化 実行ログ（2026-06-26 完了・リモート=B採用）
- ローカル化：`git init`+初回コミット(28cdbba)→最新2653d55。git 2.50.1、user=菊池研太/ec-seller。
- .gitignore：既存(秘密情報/ログ/Python)に .DS_Store・line-harness-oss/(入れ子repo)・*.tmp を追記。追跡1,131ファイル/.git 61MB/.company 226件。
- **リモート=B：自分のGoogle Drive内bareリポジトリ**（GitHub不採用＝顧客/財務/家族データを新たな第三者に出さない）。
  - 場所：`~/Library/CloudStorage/GoogleDrive-ec-seller@kikuchi-hd.net/マイドライブ/KHD_git_remote/khd_workspace.git`
  - origin登録→`git push -u origin main`成功。ls-remote一致。
- 運用ルール：開始=git pull／終了=git push（Claude代行）＋書き手1台（母艦マスター）。
- ⚠️Bの注意：①push後Driveが雲へ上げ切るまで待つ ②2・3台目はKHD_git_remoteフォルダを「オフラインで使用可能」にしてからclone ③同期ラグあり＝同時pushしない。
- 未了：Claudeメモリ(~/.claude/.../memory 1MB)は作業フォルダ外→別bareリポジトリで同期する(次ステップ)。

## ★メモリ同期＆3台運用 手順書（2026-06-26 完了）
- **メモリ同期 完了**：`~/.claude/projects/-Users-kikuchikenta-01-honbu-docs-automation/memory` をgit化(144ファイル)→Drive bare `KHD_git_remote/khd_memory.git` へpush(ref 4f0229d)。作業＋脳の2リポジトリ体制が完成。
- **3台命名**：Mac1=現メイン／Mac2=ゆーしMac(既存・サブ)／Mac3=新ハイスペック(母艦昇格・Mac1代替)。
- **Mac1環境実査**：iCloud=kemkemsp@yahoo.co.jp(菊池研太)サインイン済・iCloud Drive有効／Node v24.14.1・npm11／claude CLI 2.1.185(~/.npm-global/bin)／Claude・Cursor・Chrome・Drive・Office・LINE導入済／**Homebrew未導入**。
- **Apple連携の答え**：①初期コピー=移行アシスタント(ほぼボタン一発・同一Apple IDで可) ②Handoff/ユニバーサルクリップボード/AirDrop可 ③ただし作業フォルダ&メモリの継続同期はiCloud対象外＝gitで実施(設定済)。
- **手順書スライド作成済(9枚・show_widget)**：Mac3=移行アシスタント経路／Mac2=clone経路／各Phase(移行→サインイン→KHD_git_remoteオフライン化→pull/clone→claude→MCP→/company)＋日々の運用(開始pull/終了push・書き手1台)。
- ⚠️Mac2/3で必須：マイドライブ/KHD_git_remoteを「オフラインで使用可能」にしてからclone/pull。
- 未了：手順書のGoogleスライド化(Driveリンク)＝[[feedback_deliverable_gslides_rule]]。グローバル~/.claude/CLAUDE.mdの3台同期(Mac3は移行で来る／Mac2は手動コピー)。
