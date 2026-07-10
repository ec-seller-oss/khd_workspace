#!/usr/bin/env python3
"""田中送付URLをQ列に追加"""
import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SPREADSHEET_ID = '1rJc23QUSUG4hihEudFb8YWZ0Qmc_0R_hNDC0KYPqkwQ'
TOKEN_FILE = os.path.expanduser('~/.config/gspread/tanaka_token.json')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

URL_DATA = [
    ["田中送付URL"],
    ["（メール記載なし）"],
    ["https://www.homes.co.jp/mansion/b-1309750000974/"],
    ["https://www.homes.co.jp/mansion/b-1442100003872/"],
    ["https://suumo.jp/ms/chuko/chiba/sc_chibashiinage/nc_21125497/"],
    ["https://www.homes.co.jp/mansion/b-1442100003960/"],
]

def main():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    service = build('sheets', 'v4', credentials=creds)
    body = {
        'valueInputOption': 'RAW',
        'data': [{'range': 'Q1', 'values': URL_DATA}]
    }
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    print(f"完了: {result.get('totalUpdatedCells')} セル更新")

if __name__ == '__main__':
    main()
