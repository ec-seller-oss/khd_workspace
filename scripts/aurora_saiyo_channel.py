# -*- coding: utf-8 -*-
"""
オーロラ大元スプシに「採用チャネル比較」タブを追加（2026-06-03 ヒトイキ紹介を受けて）。
論点＝本部紹介(0円) vs エージェント(高コスト)。本部モデルの施術者人件費 月24〜28万との比較も明示。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID = "18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB = "12_採用チャネル比較(6_3)"

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

HDR = ["チャネル", "採用コスト（1名あたり）", "質・スピード", "メリット", "デメリット・注意", "現状の使い方／判定"]
ROWS = [
    ["Indeed（先出し・自走）",
     "無料掲載〜有料クリック課金。※サングローブの運用代行は180万営業→見送り済",
     "露出 週IMP約30＝不足。自分で母集団形成",
     "無料で母集団テスト・先出し法で加盟前にリスク検証できる",
     "露出が弱い。有料化やエージェント営業が来る",
     "◎ テスト継続（パート主軸に再設計）"],
    ["本部紹介（オーロラ）",
     "紹介料 0円（本部の強み）",
     "学校紹介＋エリア違い応募者の紹介／直営2年で70名採用実績",
     "コスト0・ミスマッチ低・本部審査つき",
     "船橋での再現性が未検証（＝最大論点）",
     "◎ 最優先。6/4板橋で“船橋で何名回せるか”を確認"],
    ["人材紹介エージェント（例：ヒトイキ）",
     "成功報酬＝要確認。相場は想定年収の30〜35%（≒90〜120万/人）。※他社紹介料の実例50〜70万（本部資料P19）",
     "即戦力を直接提案・スピード速い（今回：50代鍼灸師・市川・即戦力）",
     "すぐ会える・スクリーニング済の人材",
     "高コスト。施術者人件費 月24〜28万の“数ヶ月分”に相当し得る。0円の本部紹介と比べ割高",
     "△ 費用対効果を要確認。まず紹介料の金額を確認→本部紹介を優先"],
    ["運用代行（サングローブ等）",
     "180万（営業）→見送り",
     "成果保証/運用代行か曖昧",
     "—",
     "高額×成果曖昧。Go判断前に払うのは順番が逆",
     "✕ 見送り（記録済）"],
]
NOTE = ("【菊池の論点】本部モデルでも施術者の人件費は月24〜28万円（本部資料P15）。エージェント紹介料がこれを上回るなら割高。"
        "原則＝①Indeed/本部紹介（0円）で採る → ②どうしても採れない時だけ、紹介料を確認のうえ有料を限定検討。")
ZANMU = ("【残務（採用チャネル）】① ヒトイキ有馬さんへ返信（鍼治療可否・訪問エリアの2点回答）＋紹介料を確認　"
         "② 本部紹介ルートの実態を6/4板橋で確認　③ Indeedをパート主軸に再設計　④ 候補者M00122307の面接可否を判断（面接可能6月中）")

def main():
    svc = build("sheets", "v4", credentials=creds(), cache_discovery=False)
    info = svc.spreadsheets().get(spreadsheetId=SID, fields="sheets.properties").execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in info["sheets"]}
    if TAB in existing:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": [{"deleteSheet": {"sheetId": existing[TAB]}}]}).execute()
    res = svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": [{"addSheet": {"properties": {
        "title": TAB, "gridProperties": {"rowCount": 30, "columnCount": 6, "frozenRowCount": 3}}}}]}).execute()
    gid = res["replies"][0]["addSheet"]["properties"]["sheetId"]

    title = ["採用チャネル比較 ｜ 本部紹介(0円) vs エージェント(高コスト)　2026-06-03"] + [""]*5
    values = [title, [NOTE]+[""]*5, HDR] + ROWS + [[""]*6, [ZANMU]+[""]*5]
    svc.spreadsheets().values().update(spreadsheetId=SID, range=f"'{TAB}'!A1",
        valueInputOption="USER_ENTERED", body={"values": values}).execute()

    RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
    CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
    WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
    data_start=3; data_end=data_start+len(ROWS); note_bottom=len(values)-1
    reqs=[]
    for i,w in enumerate([170,250,230,250,260,200]):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":6},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":13},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":40},"fields":"pixelSize"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":1,"endRowIndex":2,"startColumnIndex":0,"endColumnIndex":6},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":1,"endRowIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"foregroundColor":REDD,"bold":True,"fontSize":10},"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":1,"endIndex":2},"properties":{"pixelSize":54},"fields":"pixelSize"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":2,"endRowIndex":3},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":11},"wrapStrategy":"WRAP","horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":data_start,"endRowIndex":data_end},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"TOP","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    for k in range(len(ROWS)):
        bg = CARD if k%2 else WHT
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":data_start+k,"endRowIndex":data_start+k+1},"cell":{"userEnteredFormat":{"backgroundColor":bg}},"fields":"userEnteredFormat.backgroundColor"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":data_start,"endRowIndex":data_end,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat.textFormat"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":data_start,"endRowIndex":data_end,"startColumnIndex":5,"endColumnIndex":6},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":data_start,"endIndex":data_end},"properties":{"pixelSize":78},"fields":"pixelSize"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":note_bottom,"endRowIndex":note_bottom+1,"startColumnIndex":0,"endColumnIndex":6},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":note_bottom,"endRowIndex":note_bottom+1},"cell":{"userEnteredFormat":{"backgroundColor":CARD,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK},"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":note_bottom,"endIndex":note_bottom+1},"properties":{"pixelSize":50},"fields":"pixelSize"}})
    reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":2,"endRowIndex":data_end,"startColumnIndex":0,"endColumnIndex":6},"innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":reqs}).execute()
    print(f"DONE gid={gid}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
