#!/usr/bin/env python3
"""田中新規4件をスプシに追加（No.6〜9）"""
import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SPREADSHEET_ID = '1rJc23QUSUG4hihEudFb8YWZ0Qmc_0R_hNDC0KYPqkwQ'
TOKEN_FILE = os.path.expanduser('~/.config/gspread/tanaka_token.json')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# A-Y列のヘッダーに合わせて追加
# A:No B:物件名 C:住所 D:価格 E:間取り F:面積 G:築年月 H:ペット I:仲介業者 J:仲介TEL
# K:定休日 L:問合番号 M:物元業者 N:物元TEL O:6/28内見OK P:備考 Q:田中送付URL
# R:マイソクLink S:管理費+修繕積立 T:最寄り駅・徒歩 U:EV V:リフォーム状況
# W:ペット詳細 X:田中決め手チェック Y:特記事項

NEW_ROWS = [
    # No.6 メゾンドール検見川（最上階）
    ["6","メゾンドール検見川（5F最上階）","千葉市美浜区真砂5丁目","590万円","3DK","55.68㎡(壁芯)","不明","🟡ペット相談","新日本住販(株)","要確認","—","—","","","—","最上階・前面棟なし・南西向き",
     "https://www.pethomeweb.com/mansion/b-6990459598/",
     "—","管理費4,200円（修繕積立不明）","JR検見川浜 徒歩19分","不明","不明",
     "🟡相談可（詳細要確認）",
     "ペット相談/5F最上階/価格◎最安590万/3DK◎ → 管理会社確認が先決",
     "自主管理・OC除外確認済・駐車場2台"],
    # No.7 メゾンドール検見川（1F専用庭）
    ["7","メゾンドール検見川（1F専用庭）","千葉市美浜区真砂5丁目","890万円","3DK","55.68㎡(壁芯)","不明","🟡ペット相談","新日本住販(株)","要確認","—","—","","","—","1階・専用庭付き・前面公園",
     "https://www.pethomeweb.com/mansion/b-6990476449/",
     "—","管理費4,200円（修繕積立不明）","JR検見川浜 徒歩18分","不明","不明",
     "🟡相談可（詳細要確認）",
     "ペット相談/1F専用庭/価格○890万/3DK◎ → 庭でペット飼いやすい",
     "専用庭・前面公園・平置き駐車場・自主管理"],
    # No.8 検見川アートホームズ
    ["8","検見川アートホームズ（5F最上階）","千葉市美浜区真砂2丁目","980万円","3LDK","76.01㎡","1974年3月","🔵要確認","三井不動産リハウス","要確認","—","FK3ABA13","","","—","3LDK76㎡と最広・最上階・南東向き",
     "https://www.rehouse.co.jp/buy/mansion/bkdetail/FK3ABA13/",
     "—","管理費8,760円+修繕積立11,200円=19,960円","JR検見川浜 徒歩16分","不明","空室・即内覧可",
     "🔵記載なし（要確認）",
     "ペット要確認/5F最上階/価格○980万/3LDK最広76㎡◎ → ペット確認次第で有力",
     "空室即内覧可・穴吹ハウジングサービス管理・前面公園"],
    # No.9 ルツクハイツ新検見川（田中が「1番良い！」）
    ["9","ルツクハイツ新検見川 ★田中推し","千葉市美浜区磯辺5","1,150万円","3LDK","68.01㎡","1979年2月","🟡ペット相談","三井不動産リハウス","要確認","—","—","","","—","★田中が「1番良いです！」",
     "https://suumo.jp/ms/chuko/chiba/sc_chibashimihama/nc_21133664/",
     "—","管理費6,440円+修繕積立12,280円=18,720円","JR検見川浜 徒歩5分（最近）","不明","内装リフォーム予定（2024年9月）",
     "🟡相談（詳細要確認）",
     "★田中推し/ペット相談/駅5分◎最近/価格△1,150万/3LDK◎ → 最優先で管理会社確認",
     "即引渡可・4階・南東向き・内装リフォーム予定"],
]

def main():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    service = build('sheets', 'v4', credentials=creds)
    body = {
        'valueInputOption': 'RAW',
        'data': [{'range': 'A8', 'values': NEW_ROWS}]
    }
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    print(f"完了: {result.get('totalUpdatedCells')} セル更新")

if __name__ == '__main__':
    main()
