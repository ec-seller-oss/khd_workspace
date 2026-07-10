#!/usr/bin/env python3
"""スプシv3: マイソク精査結果を反映 + 新規物件追加"""
import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SPREADSHEET_ID = '1rJc23QUSUG4hihEudFb8YWZ0Qmc_0R_hNDC0KYPqkwQ'
TOKEN_FILE = os.path.expanduser('~/.config/gspread/tanaka_token.json')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

UPDATES = [
    # No.6（590万・最上階）: 現況「居住中」=OC 赤フラグ
    {
        'range': 'O3',
        'values': [["🔴OC=居住中→送付禁止"]]
    },
    {
        'range': 'X3',
        'values': [["🔴OC（現況:居住中）=実需NG・リスト除外"]]
    },
    {
        'range': 'Y3',
        'values': [["⚠️徒歩10分表記あり（他は18-19分）・居住中＝OC・送ってはダメ"]]
    },
    # No.7（890万・1F専用庭）: 空室確認、管理費詳細を修正
    {
        'range': 'S4',
        'values': [["管理費4,200+修繕積立10,600+専用庭使用料1,100=月15,900円"]]
    },
    {
        'range': 'O4',
        'values': [["🟢空室・即可（確認済）"]]
    },
    {
        'range': 'X4',
        'values': [["🟢有力：ペット可・空室即可・890万・3DK・専用庭39㎡でわんちゃん最適"]]
    },
    # No.8（検見川アートホームズ980万・5F）: 空室確認
    {
        'range': 'O5',
        'values': [["🟢空室・即可（確認済）"]]
    },
    {
        'range': 'S5',
        'values': [["管理費8,760+修繕積立11,200+自治会費200=月20,160円"]]
    },
    {
        'range': 'X5',
        'values': [["ペット確認中/空室即可/980万/3LDK76㎡(最広) → ペット可なら最有力"]]
    },
]

# 新規2件追加（No.10: アートホームズリノベ済、No.11: アートホームズ3F750万）
NEW_ROWS = [
    # No.10 検見川アートホームズ 9号棟203（リノベ全面完了2025年5月）
    ["10","検見川アートホームズ 9号棟203 リノベ済","千葉市美浜区真砂2丁目23","980万円","3LDK","76.21㎡(壁芯)","1974年3月","🔵要確認","検見川リフォーム工業（株）","043-256-2220","—","—","","","🟢空室・即可",
     "2025年5月リノベ完了（水回り全交換・内装全面）",
     "—（レインズ掲載）",
     "管理費9,340+修繕積立11,680+自治会費200=月21,220円",
     "JR検見川浜 徒歩15分 / JR新検見川 徒歩16分","無","◎全面リノベ2025年5月完了（水回り全交換）",
     "🔵ペット要確認（記載なし）",
     "ペット確認中/リノベ済◎/空室即可/980万/3LDK/76㎡",
     "水回り全設備交換済・給湯給水管更新・空室即可・2階"],
    # No.11 検見川アートホームズ 3F（居住中・相談）
    ["11","検見川アートホームズ 3F角部屋","千葉市美浜区真砂2丁目","750万円","3LDK","76.21㎡(壁芯)","1974年3月","🔵要確認","ロイヤルハウジング販売（株）","043-279-1211","—","bsj0000004426","","","🟡居住中・引渡相談",
     "2線2駅利用可・角部屋・南西向き",
     "—（レインズ掲載）",
     "管理費8,760+修繕積立11,200+自治会費200=月20,160円",
     "JR検見川浜 徒歩15分 / JR新検見川 徒歩16分","無","未リフォーム",
     "🔵ペット要確認",
     "ペット確認中/750万◎安い/3LDK/角部屋 → 現況「居住中」=OCか自己居住か要確認",
     "⚠️現況「居住中」=要確認（OCの可能性）・引渡相談・角部屋"],
]

def main():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    service = build('sheets', 'v4', credentials=creds)

    all_data = [{'range': u['range'], 'values': u['values']} for u in UPDATES]
    all_data.append({'range': 'A12', 'values': NEW_ROWS})

    body = {'valueInputOption': 'RAW', 'data': all_data}
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    print(f"完了: {result.get('totalUpdatedCells')} セル更新")

if __name__ == '__main__':
    main()
