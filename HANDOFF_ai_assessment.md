# 引き継ぎ書：KHD AI査定エンジン（新しい窓用）2026-06-03

新セッションはこれを読めば続行できる。

## いま頼みたいこと（未完の依頼）
**「使い方をイメージできるよう“スクショ入り”のマニュアルスライドを作る」**
- 既に文字版マニュアルデッキは完成：`build_manual_deck.py`（全8枚）→ out_screener/KHD_AI査定エンジン_マニュアル.pptx
- 追加：実画面のスクショ（macOS `screencapture -x`）を撮って各スライドに埋め込む。
- 撮る画面：①Notion🤖自動査定結果DB ②Googleタスク ③Drive_査定受け箱 ④判定コンソール/融資資料デッキ
- スクショ手順：`open <URL or path>` → sleep数秒 → `screencapture -x shots/xx.png` → python-pptxで add_picture。

## システム概要
物件を渡す→自動査定→🟢なら融資資料デッキまで全自動→Notion/Googleタスク通知。バイセル(榎本×SALT)のAI査定をTTPした自社版。**完成・稼働中**。

## 全スクリプト（~/01_honbu_docs_automation/scripts/）
- property_intake.py：マイソクPDF抽出（NFKC・athome対応・年賃料/公簿/容積）
- reinfolib_client.py：路線価自動(reinfolib API)・住所分解parse_address
- property_screener.py：玉川式判定（土地値割合/利回り/CF/粗利＋積算割合/B＝参考）
- screen_property.py：統合ランナー（--address/--url/--deck/--db/--bank）
- kenbiya_fetch.py / rakumachi_fetch.py：健美家/楽待 詳細自動取得
- loan_deck.py / build_loan_deck_277.py：融資資料デッキ
- sheets_db.py：物DB(SSoT)自動蓄積（BF積算割合/BG融資カバー率Bは計算式）
- pick_good.py：良物件オートピック
- queue_assess.py：業者メール件名→一次査定
- auto_pipeline.py：無人本査定（--sheet/--drive/--all）→🟢はfinish_greenでデッキ生成+物DB+Notion+Googleタスク
- gtasks.py/auth_tasks.py：Googleタスク作成（tasks_token.pickle）

## 重要ID・パス
- 物件管理マスターDB(物タブ gid=1649873874)：1XTPXFxvJtaoEKVlEaigP3U1VdYfG-IHa_9pqOiZ1-hA
- 🤖自動査定結果DB(Notion)：f3df9251e34c49bb9e493fb990325bf0 / data_source 2f7e824a-8711-4065-b7a4-620c28c121fe
- 正本ページ(Notion)：3737d27fe29581618c4de89a27686c0b（使い方+全体設計）
- 物件マッチング一覧スプシ：1a0w6K-fi_BpTGGAVmB1lHqAJYPjnM4M8fw8Rs25ghnc（O列=物件URL）
- Drive受け箱：マイドライブ/_査定受け箱（PDF入れると毎朝7時自動査定）/ デッキ出力=_査定結果デッキ
- reinfolib APIキー：scripts/.env（REINFOLIB_API_KEY）
- launchd：~/Library/LaunchAgents/com.khd.autopipeline.plist（毎朝7時・python直起動）
- sheets書込：scripts/sheets_token.pickle

## 🟢の定義（玉川式・判定の真実源）
- 収益：実質利回り8%↑/CF率1.5%↑/CCR15%↑/土地値割合0.4↑
- 再販：粗利率20%↑/土地値割合0.4↑
- 総合：🔴1つで見送り/🟡欠損残りは要検討/全🟢で買い
- **積算割合・融資カバー率B＝参考表示のみ（判定に効かせない）**※菊池決定2026-06-03
  - 土地値割合A＝路線価×面積÷価格（投資・玉川）
  - 積算割合＝(土地積算 路線価×面積＋建物積算 構造別再調達単価×延床×残存÷耐用)÷価格（銀行・融資根拠）
  - B＝仕入×0.8÷(面積×路線価)（旧式・融資カバー率の簡易版）
  - 物DB列：BF=積算割合(計算式)/BG=B(計算式)。築古木造は建物積算0が当然＝最低ライン。

## 使い方（コマンド）
- PDF：`python3 scripts/screen_property.py 物件.pdf --bank 岩手銀行 --db --deck`
- 住所：`... --address "岩手県盛岡市高松2丁目34-5" --price 2500万 --land 335 --nenchin 352万 --db --deck`
- URL：`... --url "https://www.kenbiya.com/.../re_xxxx/" --db --deck`
- 無人：`python3 scripts/auto_pipeline.py --all`（launchdが毎朝実行）
- ピック：`python3 scripts/pick_good.py`

## 残ワンタイム（菊池のブラウザ操作）
- ① launchctl load ✅済 / ② Tasks API有効化+認証 ✅済
- ③ GAS：gas_property_url_extract.gs を script.google.com の既存GASに貼付→setupUrlColumn実行→backfillPropertyUrlsをトリガー（業者メール自動収集）。※任意。やると業者メールも手入力ゼロ。

## 状態：完成・稼働中。次の窓では「スクショ入りマニュアルスライド」を仕上げる。
</parameter>
</invoke>
