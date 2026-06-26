---
date: "2026-06-22"
type: deploy_done
案件: "京橋クリニック 本番LINE（LINE Harness）デプロイ完了"
本部: "04_コンサル"
related: "[[line_harness_jissouan]] [[line_kaisetsu_kit]] [[line_reminder_3hon]] [[line_faq_chatbot]]"
---

# 京橋クリニック 本番LINE（LINE Harness）デプロイ完了 2026-06-22

## 何ができたか
LINE Harness(OSS) を Cloudflare(菊池アカウント) に実デプロイ。京橋クリニック（テスト）公式LINEのCRM基盤が本番稼働。

## 重要URL・ID
- 公式アカウント: 京橋クリニック（テスト） @769qhcpr
- Messaging API Channel ID: 2010475261
- LINE Login Channel ID: 2010475364
- LIFF ID: 2010475364-DQhEoGNt
- Worker(本体): https://kyobashi-clinic.ec-seller.workers.dev
- 管理画面: https://kyobashi-clinic-admin-6dfb2db8.pages.dev
- Webhook URL: https://kyobashi-clinic.ec-seller.workers.dev/webhook
- LIFFエンドポイント: https://kyobashi-clinic.ec-seller.workers.dev?liffId=2010475364-DQhEoGNt
- 友だち追加URL(共有用): https://kyobashi-clinic.ec-seller.workers.dev/auth/line?ref=setup
- ⚠️ API Key: 別途パスワードマネージャに保存（このファイルには平文で書かない）

## 残りの手動設定（LINEコンソール側・菊池が実施）
1. 応答設定：チャット=オフ／あいさつ=オフ／Webhook=オン／応答メッセージ=オフ（OA Manager→設定→応答設定）
2. Webhook URL設定＋Webhook利用ON（OA Manager→Messaging API）
3. LINE Loginチャネル：リンクされた公式アカウント選択／友だち追加オプションOn(aggressive)／Callback URL登録（…/auth/callback）
4. LIFFエンドポイントURLを上記に更新（?liffId=必須）
5. 友だち追加は ⑤のURL経由で共有

## 次（コンテンツ投入＝Claudeが担当）
- リッチメニュー（作成済PNG）設置
- 再診リマインド3本（CPAP定期/予約確認/前日）をステップ配信に
- FAQ自動応答10件＋エスカレーションを応答ルールに
- 予約・Web問診フォームをLIFFに
（Harness API or 同梱MCP経由で投入。①②の応答設定・Webhookが前提）
