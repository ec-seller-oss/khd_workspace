# Routine設定 手順書（インタラクティブセッション用・2026-07-13作成）

## 前提
headlessセッション（今回のような自動処理型）ではcreate_triggerが「要承認」で通らない。
承認ダイアログが出せる場所＝**スマホ/PCのClaude Codeアプリ、またはWeb（claude.ai/code）**でこの`khd_workspace`セッションを開き、下記をそのまま貼り付けて送るだけ。

## 手順
1. スマホまたはPCでClaude Codeアプリ（or claude.ai/code）を開く
2. `khd_workspace`のセッション（このリポジトリ）を開く／新規セッションでも可
3. 下の「貼り付け用メッセージ」をそのままコピペして送信
4. 「Routineを作成しますか」の承認ダイアログが出るので**許可をタップ**（2回、Routineが2つあるため）
5. 「Routineを2つ作成しました」的な返答が来たら完了

---

## 貼り付け用メッセージ（このまま送る）

```
以下の2つのRoutineをcreate_triggerで作成して。

【Routine 1】
name: 現在地ボード日次更新
cron_expression: 0 12 * * *
create_new_session_on_fire: true
notifications: {"push": false, "email": false}
prompt:
あなたはKHD本部の仮想秘書室セッションです。/home/user/khd_workspace の .company/CLAUDE.md と .company/secretary/CLAUDE.md を読み、そのルールに従って行動してください（同日ファイルは追記のみ、上書き禁止、ファイル操作前に date で日付確認、等）。

## タスク：現在地ボードの最新化
1. `git -C /home/user/khd_workspace log --since="26 hours ago" --oneline` 等で直近の作業内容を把握する
2. `.company/secretary/todos/` の最新日次ファイル、`.company/secretary/notes/` の直近ファイルも確認する
3. `.company/secretary/現在地ボード.md` の「🔴いま進行中/待ち」「✅直近完了」を実態に合わせて書き換える。完了した項目は「直近完了」へ移し、古い完了項目（1週間超）は間引く。「最終更新」行を今日の日時と「自動更新セッション」に書き換える
4. 変更が実質的にあれば commit する。fast-forwardで完了するpushは確認不要で自動実行してよい（.company/secretary/CLAUDE.md のpush運用ルール準拠）。pushでコンテンツ競合が起きた場合は無理に解決せず中断し、そのままレポートに残す
5. 昨日から実質的な変化が無ければボードは書き換えず、その旨だけ記録して終了する（無理に作業を作らない）
6. 何をしたか `.company/secretary/notes/YYYY-MM-DD-decisions.md`（今日の日付、既存があれば追記）に1〜2行残す

作業が終わったら短く結果を報告して終了してください（ユーザーへの質問・確認は不要）。

【Routine 2】
name: Gmail下書き自動巡回
cron_expression: 0 0-12 * * *
create_new_session_on_fire: true
notifications: {"push": true, "email": false}
prompt:
あなたはKHD本部の仮想秘書室セッションです。/home/user/khd_workspace の .company/CLAUDE.md と .company/secretary/CLAUDE.md を読み、特に「送信前3秒チェック」「8大ブロッカー」「顧客マスター自動メンテ」「賃貸追客 顧客別運用ルール」のセクションに従って行動してください。

## タスク：Gmail下書き巡回
1. ToolSearch で "Gmail" を検索し、mcp__Gmail系（search_threads等）ツールを使う。直近1〜2時間の未読・新着メールを確認する
2. 自動配信/no-reply/私的メール（マリオット・愛育ベビー・Coupang・MF・銀行通知等）は除外する
3. 残ったメールを仕分ける：
   - 営業直結（顧客・取引先とのやり取り）→ 下書き作成対象
   - それ以外（内務・雑多）→ 対象外、記録のみ
4. 営業直結メールについて、「送信前3秒チェック」5点（①長文を半分に削る・3文ルール ②冒頭は問いで自分語りしない ③提案を我慢し相手の合格点を聞く形 ④未確定の数字/合否を書かない ⑤絵文字・謝罪は2個以内）を満たす返信下書きを作成し、create_draft で保存する。**絶対に自動送信しない**
5. 日程調整が絡む場合はGoogle Calendar系ツール（suggest_time等）で空き時間を確認し、下書きに候補日時を入れる
6. 処理結果（下書き件数・宛先・要件を1行ずつ）を `.company/secretary/notes/YYYY-MM-DD-decisions.md`（今日の日付、既存なら追記）に記録する
7. 顧客マスター未登録っぽい新規の取引相手がいれば、その旨だけメモに残す（マスターへの直書きはしない）
8. 対象メールが無ければ何もせず終了してよい（無理に作業を作らない）

作業が終わったら「✅下書きN件」の形で短く結果を報告して終了してください（ユーザーへの質問・確認は不要。下書きは必ずユーザーの最終確認後に送信される前提）。
```

## 補足
- Routine 1（ボード更新）は毎日12:00 UTC（=21:00 JST）に発火。
- Routine 2（Gmail巡回）はJST9:00〜21:00の間、毎時発火（cron最小粒度が1時間のため）。下書きができた時だけスマホにpush通知が飛ぶ設定。
- 一度作れば以降は完全放置でOK。止めたい/変更したいときは「Routine一覧見せて」「◯◯を止めて」と言えば`list_triggers`/`delete_trigger`で対応可能。
