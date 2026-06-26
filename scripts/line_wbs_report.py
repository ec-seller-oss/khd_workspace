#!/usr/bin/env python3
"""
毎朝5時に KHD全社WBSの未完了タスクをLINEに送信するスクリプト。
"""
import os
import pickle
import sys
import requests
from datetime import date, datetime
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# ─── 設定 ───────────────────────────────────────────────
SPREADSHEET_ID = "1OLNAcbPIMpdHFr8b-nYQosVuOKzzhXuuGDPsEwSk2i4"
SHEET_GID = 1170356369             # URL の gid= から取得
TOKEN_PATH = os.path.join(os.path.dirname(__file__), "sheets_token.pickle")
LINE_CHANNEL_TOKEN = os.environ.get("LINE_CHANNEL_TOKEN", "")

# ステータスが以下に含まれるものを「未完了」とみなす
INCOMPLETE_STATUSES = {"未着手", "対応中", "計画中", "最優先", "新規立ち上げ", "新規構築中", "段階拡大中"}

# ─── Google Sheets 読み取り ──────────────────────────────
def get_credentials():
    if not os.path.exists(TOKEN_PATH):
        print(f"ERROR: トークンが見つかりません。先に auth_sheets.py を実行してください。", file=sys.stderr)
        sys.exit(1)
    with open(TOKEN_PATH, "rb") as f:
        creds = pickle.load(f)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_PATH, "wb") as f:
            pickle.dump(creds, f)
    return creds


def get_sheet_name(service):
    """GIDからシート名を自動検出する。"""
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in meta.get("sheets", []):
        if sheet["properties"]["sheetId"] == SHEET_GID:
            return sheet["properties"]["title"]
    raise ValueError(f"シートID {SHEET_GID} が見つかりませんでした。")


def fetch_wbs_tasks():
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)
    sheet_name = get_sheet_name(service)
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:H",
    ).execute()
    rows = result.get("values", [])
    return rows


def parse_tasks(rows):
    """
    ヘッダー行を探し、タスクリストを返す。
    列: 工程 | 狙い・完了条件 | 担当本部長 | 担当者 | From | TO | 進捗率 | ステータス
    """
    header_idx = None
    for i, row in enumerate(rows):
        if len(row) >= 7 and row[0] == "工程" and "ステータス" in row:
            header_idx = i
            break

    if header_idx is None:
        return []

    header = rows[header_idx]
    try:
        col_task   = header.index("工程")
        col_head   = header.index("担当本部長")
        col_person = header.index("担当者")
        col_from   = header.index("From")
        col_to     = header.index("TO")
        col_prog   = header.index("進捗率")
        col_status = header.index("ステータス")
    except ValueError:
        return []

    tasks = []
    for row in rows[header_idx + 1:]:
        if len(row) <= col_status:
            continue
        status = row[col_status].strip()
        if status not in INCOMPLETE_STATUSES:
            continue
        task_name = row[col_task].strip() if len(row) > col_task else ""
        if not task_name:
            continue

        tasks.append({
            "task":   task_name,
            "head":   row[col_head].strip()   if len(row) > col_head   else "",
            "person": row[col_person].strip() if len(row) > col_person else "",
            "from":   row[col_from].strip()   if len(row) > col_from   else "",
            "to":     row[col_to].strip()     if len(row) > col_to     else "",
            "prog":   row[col_prog].strip()   if len(row) > col_prog   else "",
            "status": status,
        })
    return tasks


def is_overdue(to_str):
    if not to_str:
        return False
    try:
        d = datetime.strptime(to_str, "%Y/%m/%d").date()
        return d < date.today()
    except ValueError:
        return False


# ─── メッセージ整形 ──────────────────────────────────────
def build_message(tasks):
    today = date.today().strftime("%Y/%m/%d")
    overdue = [t for t in tasks if is_overdue(t["to"])]
    active  = [t for t in tasks if not is_overdue(t["to"])]

    lines = [f"\n📋 KHD未完了タスク [{today}]"]
    lines.append(f"計 {len(tasks)}件（うち期限超過 {len(overdue)}件）")
    lines.append("─" * 20)

    if overdue:
        lines.append("🔴 期限超過")
        for t in overdue:
            lines.append(f"▶ {t['task']}")
            lines.append(f"   {t['head']} / {t['person']} | 〜{t['to']} | {t['prog']} | {t['status']}")

    if active:
        lines.append("🟡 進行中・未着手")
        for t in active:
            to_label = f"〜{t['to']}" if t["to"] else "期限なし"
            lines.append(f"▶ {t['task']}")
            lines.append(f"   {t['head']} | {to_label} | {t['prog']} | {t['status']}")

    return "\n".join(lines)


# ─── LINE Messaging API 送信（broadcast） ────────────────
def send_line(message):
    if not LINE_CHANNEL_TOKEN:
        print("ERROR: LINE_CHANNEL_TOKEN が未設定。", file=sys.stderr)
        sys.exit(1)

    chunks = [message[i:i+4900] for i in range(0, len(message), 4900)]
    for chunk in chunks:
        resp = requests.post(
            "https://api.line.me/v2/bot/message/broadcast",
            headers={
                "Authorization": f"Bearer {LINE_CHANNEL_TOKEN}",
                "Content-Type": "application/json",
            },
            json={"messages": [{"type": "text", "text": chunk}]},
        )
        if resp.status_code != 200:
            print(f"LINE送信失敗: {resp.status_code} {resp.text}", file=sys.stderr)
            sys.exit(1)
    print(f"LINE送信完了 ({len(chunks)}チャンク)")


# ─── エントリポイント ────────────────────────────────────
def main():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] WBS取得中...")
    rows  = fetch_wbs_tasks()
    tasks = parse_tasks(rows)
    print(f"未完了タスク: {len(tasks)}件")

    if not tasks:
        message = f"\n✅ KHD全社WBS：本日の未完了タスクなし [{date.today()}]"
    else:
        message = build_message(tasks)

    send_line(message)


if __name__ == "__main__":
    main()
