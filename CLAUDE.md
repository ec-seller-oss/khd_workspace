# 01_honbu_docs_automation — プロジェクト方針

## このフォルダの目的

KHD本部の業務自動化・ドキュメント生成・AI活用ツールの開発拠点

---

## 主なツール・スクリプト

| ファイル | 用途 |
|----------|------|
| `tools/slides/build_slides_minimal.py` | python-pptx でセミナースライドを生成 |
| `tools/gas/line_notify_gas.js` | 毎日朝5時 WBS未完了タスク → LINE通知（GASに貼り付けて使用） |
| `.company/` | 仮想秘書室（secretary/todos / notes / inbox） |
| `ファイル置き場ルール.md` | どこに何を置くかの判断基準（2026-07-10 全面監査で制定） |

---

## 開発ルール

- Python スクリプトは `python3` で実行（venv 不要）
- **ファイルの置き場は `ファイル置き場ルール.md` に従う**（道具→`tools/`、完成物→`成果物/<本部>/`、迷ったら`_保留/`）
- 生成ファイルの出力先は `成果物/<該当本部>/`（旧「フォルダ直下」ルールは2026-07-10廃止。ルート直下はEC自動化の生データ等の凍結リストのみ）
- ⚠️ `scripts/`・ルート直下の `ec_*.csv`/`ec_automation.db` は自動化(launchd/hook)が絶対パス参照＝**移動禁止**
- Google Drive へのアップロードは Desktop sync フォルダ経由
  (`~/Library/CloudStorage/GoogleDrive-ec-seller@kikuchi-hd.net/マイドライブ/`)
- スライド生成後は必ず `ls -lh` でファイルサイズを確認

---

## 利用中の外部サービス

- Google Drive MCP（ファイル検索・作成・読み取り）
- Google Calendar MCP（予定確認・作成）
- LINE Notify API（WBS朝報送信）

---

## 秘書室ルール（.company/ 配下）

- `.company/CLAUDE.md` を参照してルールに従う
- 同日ファイルは追記のみ（上書き禁止）
- ファイル操作前に必ず `date` で今日の日付を確認する
