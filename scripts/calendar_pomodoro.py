#!/usr/bin/env python3
"""
KHD カレンダー × ポモドーロ通知
25分ごとに実行。直近25分以内に終わったカレンダー予定をLINEに通知。
"""
import os, pickle, requests
from datetime import datetime, timedelta, timezone
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

SCRIPTS_DIR  = os.path.dirname(__file__)
TOKEN_PATH   = os.path.join(SCRIPTS_DIR, "sheets_token.pickle")
LINE_TOKEN   = os.environ.get("LINE_CHANNEL_TOKEN", "sBuUa7G/fy88sHuXZ11aPUzs2uf2FOCoN/D/jG0kJC73urwZZuRNVb1M07+8hJcZTTmELnEuYVAybD5JeceD/eABSOZxSGqEo3aXoTCVaavbtzHzB1EvfW77f6UTEB1KT9EcD/PVVPQPLrhzjO9l2wdB04t89/1O/w1cDnyilFU=")
APPS_URL     = "https://script.google.com/macros/s/AKfycbw1uDeXqnzVvAMJ5Rt2BnCUNKdNqjpgS5rQPTr8Dpujzwqo1JgPlfZNGSQNmzB0S7KrlA/exec"
SECRET_TOKEN = "khd_report_2026"
WINDOW_MIN   = 26  # 25分 + 1分バッファ


def get_creds():
    with open(TOKEN_PATH, "rb") as f:
        creds = pickle.load(f)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_PATH, "wb") as f:
            pickle.dump(creds, f)
    return creds


def get_ending_events(creds):
    now = datetime.now(timezone.utc)
    since = now - timedelta(minutes=WINDOW_MIN)

    service = build("calendar", "v3", credentials=creds, cache_discovery=False)
    result = service.events().list(
        calendarId="primary",
        timeMin=since.isoformat(),
        timeMax=now.isoformat(),
        singleEvents=True,
        orderBy="startTime",
    ).execute()

    ending = []
    for ev in result.get("items", []):
        if ev.get("status") == "cancelled": continue
        end_str = ev.get("end", {}).get("dateTime")
        if not end_str: continue  # 終日イベントはスキップ
        end_dt = datetime.fromisoformat(end_str.replace("Z", "+00:00"))
        if since <= end_dt <= now:
            ending.append({
                "title": ev.get("summary", "（タイトルなし）"),
                "end":   end_dt.astimezone().strftime("%H:%M"),
            })
    return ending


def broadcast(text):
    requests.post(
        "https://api.line.me/v2/bot/message/broadcast",
        headers={"Authorization": f"Bearer {LINE_TOKEN}", "Content-Type": "application/json"},
        json={"messages": [{"type": "text", "text": text}]},
        timeout=10,
    )


def get_next_task():
    r = requests.post(APPS_URL, json={"token": SECRET_TOKEN, "action": "get_next_task"}, timeout=20)
    d = r.json()
    if d.get("ok") and d.get("result", {}).get("suggestion"):
        return d["result"]["suggestion"]
    return None


def log_conversation(content, action_taken="カレンダー通知"):
    requests.post(APPS_URL, json={
        "token": SECRET_TOKEN,
        "action": "log_conversation",
        "direction": "送信",
        "content": content,
        "action_taken": action_taken,
    }, timeout=10)


def main():
    if not os.path.exists(TOKEN_PATH):
        print("トークンなし。先に auth_calendar.py を実行してください")
        return

    try:
        creds = get_creds()
    except Exception as e:
        print(f"認証エラー: {e}")
        return

    events = get_ending_events(creds)
    if not events:
        print("終了したイベントなし")
        return

    for ev in events:
        msg = (
            f"📅 「{ev['title']}」が終わりました（{ev['end']}）\n\n"
            f"いかがでしたか？気軽に返信してください。\n"
            f"例:「完了」「対応中」「まだ」「〇〇で詰まった」"
        )
        broadcast(msg)
        print(f"通知送信: {ev['title']}")
        log_conversation(f"カレンダー終了通知: {ev['title']}")

    # 次タスク提案
    next_task = get_next_task()
    if next_task:
        broadcast(next_task)


if __name__ == "__main__":
    main()
