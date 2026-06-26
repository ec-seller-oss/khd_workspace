#!/usr/bin/env python3
"""
LINE Messaging API セットアップ & launchd 登録を一括実行。
webhook.site を使って userId を自動取得する。

使い方:
  python3 setup_line.py <チャンネルアクセストークン>
"""
import json
import os
import subprocess
import sys
import time
import uuid

import requests

PLIST_PATH = os.path.expanduser(
    "~/Library/LaunchAgents/net.kikuchi-hd.line-wbs-daily.plist"
)
LABEL = "net.kikuchi-hd.line-wbs-daily"
LINE_API = "https://api.line.me/v2/bot"


def verify_token(token):
    resp = requests.get(f"{LINE_API}/info", headers={"Authorization": f"Bearer {token}"})
    if resp.status_code != 200:
        print(f"  トークンが無効です: {resp.status_code} {resp.text}")
        sys.exit(1)
    return resp.json().get("displayName", "不明")


def create_webhook_site():
    """webhook.site の一時URLを作成して返す。"""
    resp = requests.post("https://webhook.site/token", json={"default_status": 200})
    resp.raise_for_status()
    uid = resp.json()["uuid"]
    url = f"https://webhook.site/{uid}"
    api  = f"https://webhook.site/token/{uid}/requests"
    return url, api


def set_line_webhook(token, url):
    resp = requests.put(
        f"{LINE_API}/channel/webhook/endpoint",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"webhookEndpointUrl": url},
    )
    return resp.status_code == 200


def poll_for_user_id(api_url, timeout=180):
    """webhook.site をポーリングして LINE の userId を取得する。"""
    seen = set()
    for _ in range(timeout // 3):
        time.sleep(3)
        try:
            resp = requests.get(api_url, params={"sorting": "newest", "per_page": 5})
            for req in resp.json().get("data", []):
                rid = req.get("uuid")
                if rid in seen:
                    continue
                seen.add(rid)
                try:
                    body = json.loads(req.get("content", "{}"))
                    for event in body.get("events", []):
                        uid = event.get("source", {}).get("userId")
                        if uid:
                            return uid
                except Exception:
                    pass
        except Exception:
            pass
    return None


def update_plist(token, user_id):
    with open(PLIST_PATH, "r") as f:
        content = f.read()
    content = content.replace("__REPLACE_WITH_CHANNEL_ACCESS_TOKEN__", token)
    content = content.replace("__REPLACE_WITH_YOUR_USER_ID__", user_id)
    with open(PLIST_PATH, "w") as f:
        f.write(content)
    print(f"  plist 更新完了: {PLIST_PATH}")


def reload_launchd():
    subprocess.run(["launchctl", "unload", PLIST_PATH], capture_output=True)
    result = subprocess.run(["launchctl", "load", PLIST_PATH], capture_output=True)
    if result.returncode == 0:
        print("  launchd 登録完了 → 毎朝5:00に自動実行されます")
    else:
        print(f"  launchd エラー: {result.stderr.decode()}")


def clear_line_webhook(token):
    """webhook.site の一時URLを解除する。"""
    requests.put(
        f"{LINE_API}/channel/webhook/endpoint",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"webhookEndpointUrl": ""},
    )


def send_test(token, user_id):
    resp = requests.post(
        f"{LINE_API}/message/push",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"to": user_id, "messages": [{"type": "text", "text": "✅ KHD WBS 日報セットアップ完了！\n毎朝5:00に未完了タスクをお届けします。"}]},
    )
    if resp.status_code == 200:
        print("  テストメッセージ送信OK → LINEを確認してください")
    else:
        print(f"  送信失敗: {resp.status_code} {resp.text}")


def main():
    if len(sys.argv) < 2:
        print("使い方: python3 setup_line.py <チャンネルアクセストークン>")
        sys.exit(1)

    token = sys.argv[1].strip()

    print("\n[1/5] トークンを確認中...")
    bot_name = verify_token(token)
    print(f"  OK: ボット名 = {bot_name}")

    print("\n[2/5] 一時Webhookを作成中 (webhook.site)...")
    webhook_url, poll_api = create_webhook_site()
    print(f"  Webhook URL: {webhook_url}")

    print("\n[3/5] LINEチャンネルにWebhookを設定中...")
    if set_line_webhook(token, webhook_url):
        print("  設定完了")
    else:
        print("  ⚠️ 自動設定失敗。LINE Developers Console で手動設定してください。")
        print(f"  Webhook URL: {webhook_url}")

    print("\n[4/5] userId を待機中...")
    print("  → LINEアプリでこのボットに「こんにちは」とメッセージを送ってください")
    print("    （まだ友だち追加していない場合は、QRコードから追加してください）")
    print("  最大3分待ちます...\n")

    user_id = poll_for_user_id(poll_api, timeout=180)

    if not user_id:
        print("  タイムアウト。LINEからメッセージを送信してから再実行してください。")
        clear_line_webhook(token)
        sys.exit(1)

    print(f"  userId 取得: {user_id}")

    print("\n[5/5] plist更新 & launchd登録 & テスト送信...")
    update_plist(token, user_id)
    reload_launchd()
    clear_line_webhook(token)
    send_test(token, user_id)

    print("\n✅ 全完了！毎朝5:00にWBS未完了タスクが届きます。")


if __name__ == "__main__":
    main()
