#!/usr/bin/env python3
"""LINE送信テスト: 環境変数を設定してから実行"""
import os, sys, requests

token   = os.environ.get("LINE_CHANNEL_TOKEN", "")
user_id = os.environ.get("LINE_USER_ID", "")

if not token or not user_id:
    print("使い方: LINE_CHANNEL_TOKEN=xxx LINE_USER_ID=Uyyy python3 test_line_send.py")
    sys.exit(1)

resp = requests.post(
    "https://api.line.me/v2/bot/message/push",
    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    json={"to": user_id, "messages": [{"type": "text", "text": "✅ KHD WBS 日報テスト送信OK"}]},
)
print(f"Status: {resp.status_code}")
print(resp.text)
