#!/usr/bin/env python3
"""一回だけ実行：Google Tasks 書込トークン(tasks_token.pickle)を生成。
   ブラウザが開くので承認 → 以降 cron からタスク作成可。"""
import os, pickle
from google_auth_oauthlib.flow import InstalledAppFlow
SCOPES = ["https://www.googleapis.com/auth/tasks"]
CS = os.path.join(os.path.dirname(__file__), "client_secrets.json")
TOK = os.path.join(os.path.dirname(__file__), "tasks_token.pickle")
def main():
    flow = InstalledAppFlow.from_client_secrets_file(CS, SCOPES)
    creds = flow.run_local_server(port=0)
    with open(TOK, "wb") as f: pickle.dump(creds, f)
    print("認証完了:", TOK)
if __name__ == "__main__": main()
