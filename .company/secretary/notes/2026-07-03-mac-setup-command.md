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

---

## ★★ 金庫をGitHubへ移行（2026-07-03・恒久対策）
### 事故
この日、Google Driveの金庫フォルダ `マイドライブ/KHD_git_remote` が**丸ごとローカル未同期になり消失**（09:27には書けていたが10:15にはstat不可・プレースホルダも無し）。自動pushウォッチャーも20分待って戻らず＝Drive×git金庫の構造的脆さが再々発。**データ損失はゼロ**（各ローカルリポが全履歴を持つ完全複製）。

### 対処＝GitHub非公開リポへ移行（菊池判断）
- アカウント **`ec-seller-oss`**（このMacのSSH鍵 id_ed25519＝"khd-report" が登録済・`ssh -T git@github.com`で認証OK）。gh/トークン不要、SSHで直接push。
- 菊池がWebで**非公開リポ2つ作成**：`khd_workspace` / `khd_memory`。
- 両ローカルリポの remote を Drive→GitHub へ付替（旧Driveは remote名 `drive` で退避＝消さず保持）。全履歴push成功：workspace=131b3f0 / memory=4977f5a（local=github一致を検証）。
- `scripts/khd_git_sync.sh` は `origin main` 参照なので**無改修でGitHub運用に**。pull/push実走で✅確認。
- `.command` のクローン先・memory clone先も GitHub URL に更新済。

### 新Macクローン（更新後）
`git clone git@github.com:ec-seller-oss/khd_workspace.git ~/01_honbu_docs_automation`
（memoryは`.command`が `git@github.com:ec-seller-oss/khd_memory.git` から自動clone）
→ 前提＝新MacのSSH公開鍵をGitHub(ec-seller-oss)に登録しておくこと。

### 恒久メモ
- **金庫＝GitHub(ec-seller-oss/khd_workspace, khd_memory)が正**。Drive金庫(`KHD_git_remote`)は廃止・信頼しない。`drive` remoteは当面の保険。
- secretary/CLAUDE.mdの旧「Drive金庫/`main 2`不正ref」節は歴史的記述に降格（GitHubでは発生しない）。
