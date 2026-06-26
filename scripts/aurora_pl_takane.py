# -*- coding: utf-8 -*-
"""
オーロラ大元スプシに「単店PL_高根」タブ追加。本部FCシミュ(P28)に高根家賃・ロイヤリティ13.2%を当て込み、
6/12/24ヶ月の営業利益・BEP・投資回収を算出。17タブの収益条件(状態E14)も更新。2026-06-04。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID="18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB="20_単店PL_高根(6_4)"
N=4
RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
YEL={"red":1.0,"green":0.97,"blue":0.80}; WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}

ROWS=[
 ("SEC","■ 前提（本部FCシミュP28 × 高根の実情）"),
 ("KV","ベースモデル","本部FCシミュ(P28)＝平均客単価3.5万円・利用者数ランプ(6ヶ月15-20名/12ヶ月45-50名/24ヶ月90-100名)。"),
 ("KV","高根の調整","ロイヤリティ13.2%税込(本部表示は12%)／家賃8万(高根の小事務所・想定※要実額確認)／駐車1.2万/台。人件費・車両は本部準拠。"),
 ("KV","稼働の前提","石原は1日9-10件・土日祝休→集客ランプは本部標準〜やや遅め。本表は本部の利用者ランプ(楽観側)で試算。"),
 ("SEC","■ 月次PL（高根調整・円/月）"),
 ("THDR",["項目","開業6ヶ月","12ヶ月","24ヶ月"]),
 ("ROW",["利用者数","15〜20名","45〜50名","90〜100名"]),
 ("ROW",["売上（単価3.5万）","700,000","1,700,000","3,400,000"]),
 ("ROW",["ロイヤリティ13.2%","92,400","224,400","448,800"]),
 ("ROW",["事務手数料3%","21,000","51,000","102,000"]),
 ("ROW",["広告分担金","10,000","15,000","20,000"]),
 ("ROW",["家賃（高根・想定8万）","80,000","80,000","80,000"]),
 ("ROW",["人件費","300,000(1名)","600,000(2名)","1,200,000(4名)"]),
 ("ROW",["駐車場(1.2万/台)","12,000","24,000","48,000"]),
 ("ROW",["光熱費","5,000","5,000","5,000"]),
 ("ROW",["車両維持費","45,000","90,000","180,000"]),
 ("ROW",["通信費","7,000","7,000","10,000"]),
 ("ROW",["システム利用料","5,000","15,000","15,000"]),
 ("ROW",["雑費","―","―","20,000"]),
 ("ROW",["■ 営業利益（高根調整）","+122,600","+588,600","+1,271,200"]),
 ("ROW",["(参考)本部モデル営業利益","+173,000","+653,000","+1,388,000"]),
 ("SEC","■ BEP・投資回収"),
 ("KV","損益分岐(BEP)","営業利益0＝売上 約55万/月＝利用者 約16名。→ 開業6ヶ月目標(15-20名)で到達圏。家賃8万込みでも早期黒字化が射程。"),
 ("KV","年換算 営業利益","12ヶ月時点 約708万/年／24ヶ月時点 約1,524万/年。アッパー2,000万/エリアは利用者増 or 複数店で射程。"),
 ("KV","投資回収","加盟・初期 約590万＋運転 約150万＝約740万。1年目累計 約250万→2年目で回収完了＝回収は約1.5〜2年が堅い。"),
 ("SEC","■ 結論（李牧）"),
 ("NOTE","高根家賃8万を入れてもBEPは利用者16名前後＝立ち上げ早期に黒字化が射程。ただし『1年で回収』は集客が本部モデル通り早く立つ最速シナリオ。同意書(徳洲会が遅い)＋石原の稼働限定を踏まえると回収は1.5〜2年が現実的。→ 17タブの収益条件『高根家賃でのBEPが見えた(利用者16名)』は満たせる見込み。最後の要確認＝高根の事務所家賃の実額。"),
]

def creds():
    with open(TOKEN,"rb") as f: c=pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN,"wb") as f: pickle.dump(c,f)
    return c

def main():
    svc=build("sheets","v4",credentials=creds(),cache_discovery=False)
    info=svc.spreadsheets().get(spreadsheetId=SID,fields="sheets.properties").execute()
    ex={s["properties"]["title"]:s["properties"]["sheetId"] for s in info["sheets"]}
    if TAB in ex:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":[{"deleteSheet":{"sheetId":ex[TAB]}}]}).execute()
    res=svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":[{"addSheet":{"properties":{"title":TAB,"gridProperties":{"rowCount":len(ROWS)+4,"columnCount":N,"frozenRowCount":1}}}}]}).execute()
    gid=res["replies"][0]["addSheet"]["properties"]["sheetId"]
    values=[["単店PL｜オーロラ高根店（本部シミュ×高根家賃）2026-06-04 ※家賃は想定・要実額確認"]+[""]*(N-1)]
    meta=[]
    for r in ROWS:
        i=len(values); k=r[0]
        if k=="SEC": values.append([r[1]]+[""]*(N-1)); meta.append((i,"SEC"))
        elif k=="KV": values.append([r[1],r[2]]+[""]*(N-2)); meta.append((i,"KV"))
        elif k=="THDR": values.append(r[1]); meta.append((i,"THDR"))
        elif k=="ROW": values.append(r[1]); meta.append((i,"ROW",r[1][0]))
        elif k=="NOTE": values.append([r[1]]+[""]*(N-1)); meta.append((i,"NOTE"))
    svc.spreadsheets().values().update(spreadsheetId=SID,range=f"'{TAB}'!A1",valueInputOption="USER_ENTERED",body={"values":values}).execute()
    reqs=[]
    for c,w in enumerate([230,180,180,180]):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":c,"endIndex":c+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":len(values)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":12}}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":38},"fields":"pixelSize"}})
    for m in meta:
        i=m[0]; kind=m[1]
        if kind=="SEC":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":11}}},"fields":"userEnteredFormat"}})
        elif kind=="KV":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":CARD,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
        elif kind=="THDR":
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"horizontalAlignment":"LEFT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
        elif kind=="ROW":
            lab=m[2]
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
            if "営業利益" in lab and "参考" not in lab:
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":YEL,"textFormat":{"bold":True,"fontSize":11,"foregroundColor":REDD}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
            else:
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
        elif kind=="NOTE":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"wrapStrategy":"WRAP","verticalAlignment":"TOP"}},"fields":"userEnteredFormat"}})
            reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":78},"fields":"pixelSize"}})
    # PLテーブル罫線
    thdr=[m[0] for m in meta if m[1]=="THDR"][0]
    last_row=[m[0] for m in meta if m[1]=="ROW"][-1]
    reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":thdr,"endRowIndex":last_row+1,"startColumnIndex":0,"endColumnIndex":N},"innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":reqs}).execute()

    # 17タブの収益条件(状態 E14)を更新
    try:
        svc.spreadsheets().values().update(spreadsheetId=SID,range="'17_船橋GO決定プロセス(6_4)'!E14",
            valueInputOption="USER_ENTERED",body={"values":[["一次完了→20タブ(BEP利用者16名)"]]}).execute()
    except Exception as e:
        print("17更新スキップ:",e)
    print(f"DONE gid={gid}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
