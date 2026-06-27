#!/usr/bin/env python3
"""ペット確認結果・内見キャンセルをスプシに反映"""
import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SPREADSHEET_ID = '1rJc23QUSUG4hihEudFb8YWZ0Qmc_0R_hNDC0KYPqkwQ'
TOKEN_FILE = os.path.expanduser('~/.config/gspread/tanaka_token.json')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

UPDATES = [
    # H列(ペット): 電話確認結果を反映
    {
        'range': 'H2:H6',
        'values': [
            ["🔴不可（電話確認済・脱落）"],       # 西小中台
            ["要確認（未完）"],                    # 西千葉GH
            ["🔴不可（マイソク確認済・脱落）"],    # 高洲
            ["🟢可（管理規約）← 確定"],            # 稲毛PH
            ["要確認（未完）"],                    # 真砂
        ]
    },
    # O列(6/28内見): 台風でキャンセル
    {
        'range': 'O2:O6',
        'values': [
            ["🌀台風キャンセル"],
            ["🌀台風キャンセル"],
            ["🌀台風キャンセル"],
            ["🌀台風キャンセル"],
            ["🌀台風キャンセル"],
        ]
    },
    # W列(ペット詳細)も更新
    {
        'range': 'W2:W6',
        'values': [
            ["🔴不可（6/26電話確認）"],
            ["要確認"],
            ["🔴不可（マイソク明記）"],
            ["🟢可（管理規約）← 確定"],
            ["要確認"],
        ]
    },
    # X列(田中の決め手)を更新
    {
        'range': 'X2:X6',
        'values': [
            ["❌脱落（ペット不可）"],
            ["ペット要確認 / 価格◎690万 / 3LDK◎ → 可なら最有力"],
            ["❌脱落（ペット不可）"],
            ["🟢有力（ペット可・3LDK・1280万）← 現状唯一の候補"],
            ["ペット要確認 / 2LDK・告知事項あり → 可なら検討"],
        ]
    },
]

def main():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    service = build('sheets', 'v4', credentials=creds)
    body = {
        'valueInputOption': 'RAW',
        'data': [{'range': u['range'], 'values': u['values']} for u in UPDATES]
    }
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    print(f"完了: {result.get('totalUpdatedCells')} セル更新")

if __name__ == '__main__':
    main()
