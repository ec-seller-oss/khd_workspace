#!/usr/bin/env python3
"""田中案件スプシ v2: マイソクリンク・物元・ヒアリング・決め手列を追加"""
import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SPREADSHEET_ID = '1rJc23QUSUG4hihEudFb8YWZ0Qmc_0R_hNDC0KYPqkwQ'
TOKEN_FILE = os.path.expanduser('~/.config/gspread/tanaka_token.json')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# A-Q列（既存）はそのまま維持し、R列以降に追加
# R: マイソクLink
# S: 管理費+修繕積立（月合計）
# P列(備考)とH列(ペット)も詳細に更新

UPDATES = [
    # R列: マイソクリンク (R1:R6)
    {
        'range': 'R1',
        'values': [
            ["マイソクLink"],
            ["https://drive.google.com/file/d/1ZYPi3liPP8Cbk4HiBHG_0eFaCvZDQbgW/view"],
            ["https://drive.google.com/file/d/1voG3UsMsc_BASHFlsfsx21aHrWAKO0AK/view"],
            ["https://drive.google.com/file/d/1w0IMqjmuGOl9Z6OY5RhE8yr08zQxrPdD/view"],
            ["https://drive.google.com/file/d/1RZSWV9CBwh-_Tll9k6_yvdyIwVjzrE9f/view"],
            ["https://drive.google.com/file/d/1EGcQgSvPZj_BSTPqfjX0g-hJOOJdqn3b/view"],
        ]
    },
    # S列: 管理費+修繕積立（月額合計）
    {
        'range': 'S1',
        'values': [
            ["管理費+修繕積立(月)"],
            ["要確認"],          # 西小中台（マイソク記載なし）
            ["13,550円"],        # 西千葉GH: 5,500+7,500+550
            ["18,550円"],        # 高洲: 5,000+13,550
            ["19,100円"],        # 稲毛PH: 7,900+11,200（駐車場除く）
            ["15,000円"],        # 真砂: 6,000+9,000
        ]
    },
    # T列: 最寄り駅・徒歩
    {
        'range': 'T1',
        'values': [
            ["最寄り駅・徒歩"],
            ["JR新検見川 徒歩16分"],
            ["千葉都市モノレール作草部 徒歩12分 / 西千葉バス+徒歩5分"],
            ["JR稲毛海岸 徒歩7分"],
            ["JR稲毛 徒歩19分 / 千葉モノ穴川 徒歩16分"],
            ["JR稲毛海岸 徒歩15分 / JR検見川浜 徒歩13分"],
        ]
    },
    # U列: EV有無
    {
        'range': 'U1',
        'values': [
            ["EV"],
            ["要確認"],
            ["無"],
            ["無"],
            ["要確認"],
            ["要確認"],
        ]
    },
    # V列: リフォーム状況
    {
        'range': 'V1',
        'values': [
            ["リフォーム状況"],
            ["リフォーム済（詳細不明）"],
            ["未リフォーム（現況渡し）"],
            ["新規内装全面リフォーム（2025年完了）"],
            ["大規模リノベ済・大規模修繕2024済"],
            ["新規内装リフォーム（2025年11月完了）"],
        ]
    },
    # W列: ペット詳細（マイソク記載）
    {
        'range': 'W1',
        'values': [
            ["ペット詳細（マイソク記載）"],
            ["記載なし→要電話確認"],
            ["記載なし→要電話確認"],
            ["🔴ペット飼育不可（マイソク明記）"],
            ["🟢ペット飼育OK（管理規約による）"],
            ["記載なし→要電話確認"],
        ]
    },
    # X列: 田中の希望との合致度
    {
        'range': 'X1',
        'values': [
            ["田中の決め手チェック"],
            ["ペット:要確認 / 価格:◎安い / 間取り:△1LDK(狭め) / 総合:保留"],
            ["ペット:要確認 / 価格:◎最安690万 / 間取り:◎3LDK広い / 総合:最有力候補"],
            ["ペット:🔴不可=脱落 / 価格:△ / 間取り:△ / 総合:内見不要検討"],
            ["ペット:🟢確認済 / 価格:△1280万 / 間取り:◎3LDK / 総合:有力"],
            ["ペット:要確認 / 価格:△1280万 / 間取り:△2LDK / 総合:保留"],
        ]
    },
    # Y列: 特記事項
    {
        'range': 'Y1',
        'values': [
            ["特記事項"],
            ["2階・定休水曜・仲介ネクストリンク"],
            ["EV無・5階建4階・高台・南西向き・物件No.AAA14954"],
            ["🔴ペット不可で田中条件NG・EV無・京葉線稲毛海岸7分"],
            ["南向き1階・三菱地所管理・大規模修繕2024済・564戸"],
            ["🔴告知事項あり・売主指定司法書士・買取物件"],
        ]
    },
]

# H列(ペット)を更新: より詳細に
PET_UPDATE = {
    'range': 'H2:H6',
    'values': [
        ["要電話確認（水曜休み→本日可）"],
        ["要電話確認"],
        ["🔴不可（マイソク確認済）"],
        ["🟢可（管理規約）← 田中確認済"],
        ["要電話確認"],
    ]
}

def main():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    service = build('sheets', 'v4', credentials=creds)

    all_updates = UPDATES + [PET_UPDATE]
    body = {
        'valueInputOption': 'RAW',
        'data': [{'range': u['range'], 'values': u['values']} for u in all_updates]
    }
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    print(f"完了: {result.get('totalUpdatedCells')} セル更新")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")

if __name__ == '__main__':
    main()
