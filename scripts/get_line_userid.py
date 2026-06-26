#!/usr/bin/env python3
"""
LINE Messaging API の自分の userId を取得するヘルパー。

使い方:
  1. LINE Developers Console でチャンネルアクセストークン（長期）を発行
  2. 下記コマンドで起動:
       LINE_CHANNEL_TOKEN=<トークン> python3 get_line_userid.py
  3. ブラウザで表示されるQRコードからLINE公式アカウントを友だち追加
  4. 友だちになった後、LINEからそのアカウントに任意のメッセージを送る
  5. ターミナルに userId が表示される
"""
import json
import os
import sys
from http.server import BaseHTTPRequestHandler, HTTPServer

CHANNEL_TOKEN = os.environ.get("LINE_CHANNEL_TOKEN", "")
PORT = 8765


class WebhookHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length))
        self.send_response(200)
        self.end_headers()

        for event in body.get("events", []):
            user_id = event.get("source", {}).get("userId")
            if user_id:
                print(f"\n✅ userId を取得しました: {user_id}")
                print(f"\n   plist に設定するコマンド:")
                print(f'   export LINE_USER_ID="{user_id}"')
                print()

    def log_message(self, *args):
        pass


def main():
    if not CHANNEL_TOKEN:
        print("ERROR: LINE_CHANNEL_TOKEN を環境変数に設定してから実行してください。", file=sys.stderr)
        sys.exit(1)

    print(f"Webhook サーバーを起動中 (port {PORT})...")
    print("次の手順:")
    print("  1. ngrok などで外部公開: ngrok http 8765")
    print("  2. 表示された https://xxxx.ngrok.io を LINE Developers の Webhook URL に設定")
    print("  3. LINEアプリからそのアカウントにメッセージを送信")
    print("  4. ここに userId が表示されます\n")
    print("Ctrl+C で終了\n")

    server = HTTPServer(("0.0.0.0", PORT), WebhookHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n終了しました。")


if __name__ == "__main__":
    main()
