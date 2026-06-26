# 02_作業DB 安全クリーン再構築（Plan A）2026-06-16 12:20

- 本部: 02資金調達/操縦席入力ハブ。起点=菊池「02クリーン再構築」→ 対象=02_作業DBタブ（AskUserで確定）。
- 仕上げ範囲: A 安全クリーン（列位置不動・他タブ無傷）をAskUserで選択。B 完全再構築（列並べ替え＋8参照貼り直し）は別枠・TB/福井後に保留。

## ライブ実読で確定した「散らかり」（read_db02 / db02_headers）
1. ゴミ尾: 実データは上部数十行のみ。F列「—」で row1155 まで埋まり肥大化。
2. 重複・残骸ヘッダ3つ: 「内容」が2列／「実所要分」と「実所要(分)」併存／意味不明の「報告」列。→ append_db02(列名対応)誤爆リスク。
3. 🔴 配線断線: 今朝の列インサート(確度ランク+自動4列)で 報告項目 O→現S・報告値 P→現T にズレ、03売上見込みF/Gの確度/着金参照が空振り。
4. 波及: '02_作業DB'! を列文字参照するGAS 8本(_dashboard/_board01/_kpi03/_kpi03b/_logictree/_trend/_dazn01/_unify03・Jun-7生成)も道連れでズレている可能性大。→ だから列を動かさないA採用。

## 成果物
- scripts/clean_db02_safe.gs（115行・冪等・バックアップ先取り）。cleanDb02Safe 実行。
- 動作: 0)backup複製 1)内容重複 名寄せ 2)報告→結果 名寄せ 3)実所要分 無効化 4)ゴミ尾削除 5)03確度/着金 配線を現列へ修復。
- 列削除/移動なし＝（未使用）はスロット残置で他タブ参照保護。

## 状態
- [x] 診断・GAS作成・クリップボードコピー済。
- [ ] 菊池が Apps Script で cleanDb02Safe 実行（バックアップ自動）。
- [ ] 実行後 db02_headers で検証（重複解消/尾削除/配線S・T化を実読確認）。

関連: project_cockpit_rebuild / reference_khd_finance_files / feedback_no_blind_formula / feedback_spreadsheet_visual_rule / feedback_destructive_ops
