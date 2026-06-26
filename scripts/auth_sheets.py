#!/usr/bin/env python3
"""
一回だけ実行するGoogle Sheets認証スクリプト。
sheets_token.pickle を生成します。
"""
import os
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]  # 読み書き両方
CLIENT_SECRETS = os.path.join(os.path.dirname(__file__), "client_secrets.json")
TOKEN_PATH = os.path.join(os.path.dirname(__file__), "sheets_token.pickle")

def main():
    creds = None
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "rb") as f:
            creds = pickle.load(f)

    if creds and creds.valid:
        print("トークンは既に有効です。")
        return

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS, SCOPES)
        creds = flow.run_local_server(port=0)

    with open(TOKEN_PATH, "wb") as f:
        pickle.dump(creds, f)
    print(f"認証完了。トークンを保存しました: {TOKEN_PATH}")

if __name__ == "__main__":
    main()
