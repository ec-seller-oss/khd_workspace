#!/usr/bin/env bash
# 京橋クリニック 公式LINE 自動セットアップ（LINE Messaging API）
# 使い方: LINE_TOKEN='＜チャネルアクセストークン＞' bash setup_line.sh
#   ※トークンはこのファイルに書かない（環境変数で渡す）。実行後は菊池側で再発行=無効化可。
set -euo pipefail

: "${LINE_TOKEN:?トークン未設定。 LINE_TOKEN='xxxx' bash setup_line.sh の形で渡してください}"
IMG="/Users/kikuchikenta/01_honbu_docs_automation/kyobashi_richmenu.png"
API="https://api.line.me"; APID="https://api-data.line.me"
AUTH="Authorization: Bearer ${LINE_TOKEN}"

echo "▶ 0) 接続確認（Botの情報取得）"
curl -s "${API}/v2/bot/info" -H "${AUTH}" | jq . || { echo "✗ トークンが無効かも。チャネルアクセストークン（長期）を確認"; exit 1; }

echo "▶ 1) リッチメニュー作成（6タイル）"
RICHMENU_JSON='{
  "size": {"width": 2500, "height": 1686},
  "selected": true,
  "name": "kyobashi-main",
  "chatBarText": "メニュー",
  "areas": [
    {"bounds":{"x":0,"y":0,"width":833,"height":843},"action":{"type":"message","text":"順番"}},
    {"bounds":{"x":833,"y":0,"width":834,"height":843},"action":{"type":"message","text":"予約"}},
    {"bounds":{"x":1667,"y":0,"width":833,"height":843},"action":{"type":"message","text":"問診"}},
    {"bounds":{"x":0,"y":843,"width":833,"height":843},"action":{"type":"message","text":"診療時間"}},
    {"bounds":{"x":833,"y":843,"width":834,"height":843},"action":{"type":"message","text":"アクセス"}},
    {"bounds":{"x":1667,"y":843,"width":833,"height":843},"action":{"type":"message","text":"FAQ"}}
  ]
}'
RID=$(curl -s -X POST "${API}/v2/bot/richmenu" -H "${AUTH}" -H "Content-Type: application/json" -d "${RICHMENU_JSON}" | jq -r .richMenuId)
echo "  richMenuId = ${RID}"

echo "▶ 2) リッチメニュー画像アップロード"
curl -s -X POST "${APID}/v2/bot/richmenu/${RID}/content" -H "${AUTH}" -H "Content-Type: image/png" --data-binary @"${IMG}" -o /dev/null -w "  HTTP %{http_code}\n"

echo "▶ 3) 全友だちの既定メニューに設定"
curl -s -X POST "${API}/v2/bot/user/all/richmenu/${RID}" -H "${AUTH}" -o /dev/null -w "  HTTP %{http_code}\n"

echo "▶ 4) テスト配信（全友だちへ）"
curl -s -X POST "${API}/v2/bot/message/broadcast" -H "${AUTH}" -H "Content-Type: application/json" \
  -d '{"messages":[{"type":"text","text":"【テスト】京橋クリニック公式LINEです。リッチメニューを設置しました。下のメニューから予約・順番・問診をご利用いただけます。"}]}' \
  -o /dev/null -w "  HTTP %{http_code}\n"

echo "✅ 完了。スマホでこの公式アカウントを友だち追加 → リッチメニューとテスト配信が出ます。"
