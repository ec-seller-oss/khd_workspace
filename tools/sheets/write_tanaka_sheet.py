#!/usr/bin/env python3
"""田中案件 物件リストをGoogle Sheetsに書き込む"""
import os, sys, warnings
warnings.filterwarnings('ignore')

SPREADSHEET_ID = '1rJc23QUSUG4hihEudFb8YWZ0Qmc_0R_hNDC0KYPqkwQ'
CREDENTIALS_FILE = os.path.expanduser('~/credentials.json')
TOKEN_FILE = os.path.expanduser('~/.config/gspread/tanaka_token.json')

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

DATA = [
    ["No","物件名","住所","価格","間取り","面積","築年月","ペット","仲介業者","仲介TEL","定休日","問合番号","物元業者","物元TEL","6/28土14時内見OK","備考"],
    ["1","西小中台団地 2-24-204","千葉市花見川区西小中台2-24-204","930万円","1LDK","内法46.77㎡","1972年7月","要確認","ネクストリンク","043-312-3114","水曜","—","","","","リフォーム済・2階・JR新検見川徒歩16分"],
    ["2","西千葉グリーンハイツ 1棟","千葉市稲毛区作草部町","690万円","3LDK","77.11㎡","1978年11月","要確認","大成有楽不動産","0037-633-07941","—","793619","","","","5件中最安・最広・田中が安い物件希望"],
    ["3","高洲一丁目住宅 第1-17-4号棟","千葉市美浜区高洲1丁目","1288万円","2LDK","48.9㎡","1976年3月","要確認","?","0037-633-25303","—","074107","","","","④と同一問合番号・1本で2件確認可"],
    ["4","稲毛パークハウスG棟","千葉市稲毛区園生町1223","1280万円","3LDK","68.99㎡","1974年1月","◎確認済","—","043-497-2977","—","—","","","","田中がSUUMOで確認済みペット可確定"],
    ["5","真砂1丁目団地 11-10","千葉市美浜区真砂1丁目","1280万円","2LDK","51.7㎡","1976年3月","要確認","?","0037-633-25303","—","317638","","","","③と同一問合番号"],
]

def main():
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    import json

    os.makedirs(os.path.dirname(TOKEN_FILE), exist_ok=True)

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as f:
            f.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    body = {'valueInputOption': 'RAW', 'data': [{'range': 'A1', 'values': DATA}]}
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    print(f"書き込み完了: {result.get('totalUpdatedCells')} セル更新")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")

if __name__ == '__main__':
    main()
