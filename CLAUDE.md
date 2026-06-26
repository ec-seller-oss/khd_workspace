# 01_honbu_docs_automation — プロジェクト方針

## このフォルダの目的

KHD本部の業務自動化・ドキュメント生成・AI活用ツールの開発拠点

---

## 主なツール・スクリプト

| ファイル | 用途 |
|----------|------|
| `build_slides_minimal.py` | python-pptx でセミナースライドを生成 |
| `line_notify_gas.js` | 毎日朝5時 WBS未完了タスク → LINE通知（GASに貼り付けて使用） |
| `.company/` | 仮想秘書室（secretary/todos / notes / inbox） |

---

## 開発ルール

- Python スクリプトは `python3` で実行（venv 不要）
- 生成ファイルの出力先はこのフォルダ直下
- Google Drive へのアップロードは Desktop sync フォルダ経由
  (`~/Library/CloudStorage/GoogleDrive-ec-seller@kikuchi-hd.net/マイドライブ/`)
- スライド生成後は必ず `ls -lh *.pptx` でファイルサイズを確認

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
