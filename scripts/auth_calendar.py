#!/usr/bin/env python3
"""Google Calendar 認証スクリプト（1回だけ実行）"""
import os, pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/calendar.readonly",
]
CLIENT_SECRETS = os.path.join(os.path.dirname(__file__), "client_secrets.json")
TOKEN_PATH     = os.path.join(os.path.dirname(__file__), "sheets_token.pickle")

def main():
    creds = None
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "rb") as f:
            creds = pickle.load(f)

    # Calendar スコープが含まれていなければ再認証
    if creds and creds.valid and all(s in (creds.scopes or []) for s in SCOPES):
        print("トークンは既に有効です（Calendar含む）")
        return

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            print("トークンをリフレッシュしました")
        except Exception:
            creds = None

    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS, SCOPES)
        creds = flow.run_local_server(port=0)

    with open(TOKEN_PATH, "wb") as f:
        pickle.dump(creds, f)
    print(f"認証完了。トークンを保存しました: {TOKEN_PATH}")

if __name__ == "__main__":
    main()
