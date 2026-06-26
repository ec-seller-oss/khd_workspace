# -*- coding: utf-8 -*-
"""
オーロラ大元スプシ（宮崎と共有）に「板橋オーナー ヒアリング記録(6/4)」タブを追加。
先輩経営者の一次情報(本音)を取る→回答欄を埋めれば今日の報告書になる一覧表。
宮崎合意の聞き方＝A直球(最優先)／B-Eざっくり。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID = "18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB = "16_板橋ヒアリング記録(6_4)"
N = 6

ROWS = [
    ("SEC", "■ 面談メタ（一次情報＝先輩オーナーの本音を取る場）"),
    ("KV", "日時・場所", "2026-06-04（木）13:00／表参道スタバ（※板橋でなく表参道で対面）"),
    ("KV", "相手", "オーロラ板橋店オーナー（運営＝洗心メディカルサポート? 冒頭で『直営/加盟店どちらか』を確認）"),
    ("KV", "同席", "宮崎さん（流動・行けそうなら一報）"),
    ("KV", "狙い／スタンス", "本部の公式見解でなく『現場オーナーのリアル』を取る。Aだけ直球で先に取り切る→B〜Eは『ぶっちゃけ◯◯どう？』で雑談的に本音を回収。これを今日の報告書にする。"),

    ("SEC", "■ ヒアリング一覧（回答欄＝黄色に当日記入→そのまま報告書）"),
    ("THDR", ["#", "テーマ", "聞き方（宮崎合意の方向性）", "取りたい一次情報（狙い）", "板橋オーナーの回答【当日記入】", "所感・本部公称との差"]),
    ("ROW", ["A", "同意書を早く回す【直球・最優先】",
             "「徳洲会さんは出るけど遅いと聞いてて、早く回すコツや、早く出してくれる先（診療科・クリニック）の見つけ方ってどうしてます？」「申請→実際に算定できるまで板橋だとどのくらい？」",
             "律速＝同意書スピードの外し方。早く出す診療所の探し方。算定までの実日数。", "", ""]),
    ("ROW", ["B", "本部って実際どう【ざっくり】",
             "「ぶっちゃけ、オーロラの本部ってどうですか？付き合ってみての本音で、サポートの良し悪し」",
             "採用支援70名・集客サポートの実効性。本部の当たり外れ・対応の速さ。", "", ""]),
    ("ROW", ["C", "この事業どう【ざっくり】",
             "「この事業、実際やってみてどうですか？想定と違ったとこ、しんどいとこは？」",
             "立ち上げの詰まり（利用者獲得 vs 採用）。本音の苦労・続けられる感触。", "", ""]),
    ("ROW", ["D", "立地の決め方【半ざっくり】",
             "「お店の場所と訪問エリア、どうやって決めました？」",
             "在宅密度の見極め方。同意書が出るエリアの選び方。高根（団地×徳洲会）の妥当性の裏取り。", "", ""]),
    ("ROW", ["E", "採用の実際【半ざっくり】",
             "「人ってどうやって採ってます？本部紹介って実際回りますか？」",
             "本部紹介の再現性（船橋/高根で何名）。Indeedの要否。歩合での定着。", "", ""]),
    ("ROW", ["F", "収益のリアル【流れで】",
             "「黒字化までどのくらいかかりました？手残りって正直どんな感じですか？」",
             "利用者数の推移・1人あたり件数/売上・ロイヤリティ等控除後の手残り実額・BEP月。", "", ""]),
    ("ROW", ["G", "競業の扱い【さらっと】",
             "「前職の絡みがある人を採るときって、どうしてます？」",
             "顧客持ち込みのリスク・本部の考え・トラブル経験。", "", ""]),
    ("ROW", ["H", "ここだけの一次情報【自由】",
             "「他に、加盟前に知っておくべきことってありますか？」",
             "想定外の落とし穴・リアルな一言・人脈的な裏話。", "", ""]),

    ("SEC", "■ 今日の報告書（結論）※面談後に記入"),
    ("NOTE", "【総括】Go判断3点セットへの示唆＝①採用（本部紹介は船橋で回るか）：__　②収益（単店BEP・手残り）：__　③集客＝同意書見込み（高根で早く出るか）：__　→ 所感／次アクション：__"),
]

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
YEL={"red":1.0,"green":0.97,"blue":0.80}; WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}

def main():
    svc = build("sheets", "v4", credentials=creds(), cache_discovery=False)
    info = svc.spreadsheets().get(spreadsheetId=SID, fields="sheets.properties").execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in info["sheets"]}
    if TAB in existing:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": [{"deleteSheet": {"sheetId": existing[TAB]}}]}).execute()
    res = svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": [{"addSheet": {"properties": {
        "title": TAB, "gridProperties": {"rowCount": len(ROWS)+4, "columnCount": N, "frozenRowCount": 1}}}}]}).execute()
    gid = res["replies"][0]["addSheet"]["properties"]["sheetId"]

    values = [["板橋オーナー ヒアリング記録（一次情報）｜ 2026-06-04 13:00＠表参道スタバ ／ 回答を埋めれば今日の報告書"]+[""]*(N-1)]
    meta=[]
    for r in ROWS:
        i=len(values); k=r[0]
        if k=="SEC": values.append([r[1]]+[""]*(N-1)); meta.append((i,"SEC"))
        elif k=="KV": values.append([r[1], r[2]]+[""]*(N-2)); meta.append((i,"KV"))
        elif k=="THDR": values.append(r[1]); meta.append((i,"THDR"))
        elif k=="ROW": values.append(r[1]); meta.append((i,"ROW"))
        elif k=="NOTE": values.append([r[1]]+[""]*(N-1)); meta.append((i,"NOTE"))
    svc.spreadsheets().values().update(spreadsheetId=SID, range=f"'{TAB}'!A1",
        valueInputOption="USER_ENTERED", body={"values": values}).execute()

    reqs=[]
    for c,w in enumerate([34,150,300,250,300,200]):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":c,"endIndex":c+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":len(values)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"TOP","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":12},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":38},"fields":"pixelSize"}})
    row_idx=[]
    for (i,kind) in meta:
        if kind=="SEC":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":11},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
        elif kind=="KV":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":CARD,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
        elif kind=="THDR":
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
        elif kind=="ROW":
            row_idx.append(i)
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"fontSize":11,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
            # 回答欄(列4)を黄色
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":4,"endColumnIndex":5},"cell":{"userEnteredFormat":{"backgroundColor":YEL}},"fields":"userEnteredFormat.backgroundColor"}})
        elif kind=="NOTE":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":YEL,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat"}})
            reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":56},"fields":"pixelSize"}})
    if row_idx:
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":row_idx[0],"endIndex":row_idx[-1]+1},"properties":{"pixelSize":74},"fields":"pixelSize"}})
        reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":row_idx[0]-1,"endRowIndex":row_idx[-1]+1,"startColumnIndex":0,"endColumnIndex":N},"innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":reqs}).execute()
    print(f"DONE gid={gid}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
