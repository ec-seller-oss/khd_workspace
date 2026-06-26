# -*- coding: utf-8 -*-
"""
オーロラ大元スプシに「石原さん面談記録(6/2対面)」タブを追加。
2026-06-02 Indeed経由・ドトール船橋で対面、9項目＋競合/立地＋総合所感を記録。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID = "18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB = "13_石原面談記録(6_2)"

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

HDR = ["#", "項目", "石原さんの回答（6/2対面）", "所感・要注意"]
ROWS = [
    ["1", "転職理由",
     "現職＝津田沼周辺の独自FC（FC化なし店舗）。2025/1入社。半年で営業マンがオーナーと揉め同年9月退職→新規が激減。自身も片手間で営業する程度で時間を割けず、現給与が脅かされる懸念から転職検討。",
     "◎ 環境要因・他責でなく建設的。動機は明確"],
    ["2", "希望条件（給与/稼働/歩合）",
     "現給与38万＋交通費実費（内訳：基本26万＋管理施術者手当3万＋固定残業代）。賞与は未確認（恐らく無し／メールでいつでも確認可）。希望＝土日祝休み・直行直帰・8〜17時（16時上がりも多い）・基本チャリ移動・近場。",
     "現給与38万が基準線。妻(内縁)の介護で安定・近場志向。稼働は限定的"],
    ["3", "訪問件数の実績 ★中身確認済",
     "履歴書「1日30人以上」は1人10分の大型店舗＝来店型の話。訪問は1日9〜10人が現実的。",
     "★来店型と判明。訪問実数9〜10件＝本部標準11件よりやや下だが安定はする"],
    ["4", "集客力（ケアマネ営業）",
     "今もちょこちょこ程度。",
     "営業は強くない・片手間。集客はKHD/本部で補完する前提"],
    ["5", "いつから動けるか",
     "8月から。",
     "候補シート（7〜8月）と整合。退職段取りは現実的"],
    ["6", "現職在籍中の状況",
     "在職中。2025/1入社→半年で営業退職→新規減で給与不安。条件はメールでいつでも確認可能。",
     "退職動機が明確。並行検討を正直に話す＝誠実さ◎"],
    ["7", "マネジメント経験／将来の役割",
     "妻の体調等でリスクは取れないが、独立したい気持ちはある。オーロラのノウハウも知りたい様子（菊池の推察）。",
     "⚠️独立志向＝将来の自走 or 競業の両面。拠点長より『安定＋ノウハウ吸収』狙いの可能性。要見極め"],
    ["8", "体力・稼働ペース（53歳）",
     "自転車移動で十分いけるとのこと。そのエリア（高根方面）は任せられそう。",
     "9〜10件/日を自走で回せる。エリアの土地勘あり＝立ち上げ戦力"],
    ["9", "競業リスク（現職＝津田沼・船橋近接）",
     "現職と近接。『少しなら顧客をそのまま引き継げる可能性』をチラッと仄めかす。",
     "🔴持ち出しの示唆＝コンプラ赤信号。採用時は線引き（持ち出しNG）を明確化必須"],
    ["—", "競合・立地情報（石原提供＝あながち信頼できそう）",
     "リボン（船橋駅前）／フレアス（北習志野・前原 ※習志野は船橋と医師会が異なり同意書が出ない。3箇所出す医師会に背くクリニックは把握済）／アシスト（高根?・江戸川）／ケイロウ（西船橋駅 ※団地少なく未開拓かも）／津田沼（石原の現職）。空き仮説＝高根公団（徳洲会あり）・高根木戸（新京成沿線）。",
     "★立地は高根エリアが有力候補。同意書の医師会差（習志野で出ない）が収益の急所＝船橋市側の実態を要確認"],
]
SOUKAN = ("【総合所感】即戦力で誠実・エリアの土地勘と競合情報は価値大。一方①稼働は限定的（土日祝休・実質8h・訪問9-10件）"
          "②安定/近場志向で攻めの拡大役ではない③競業の持ち出し示唆は赤信号。"
          "→ 採用するなら『高根エリアの現場リード＋線引き明確化』。拠点長フル委任より、本部紹介で複線採用しつつ石原は安定戦力に。"
          "急所は『船橋(高根)の同意書通過率』と『本部紹介が船橋で何名回るか』＝6/4板橋で詰める。")

def main():
    svc = build("sheets", "v4", credentials=creds(), cache_discovery=False)
    info = svc.spreadsheets().get(spreadsheetId=SID, fields="sheets.properties").execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in info["sheets"]}
    if TAB in existing:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": [{"deleteSheet": {"sheetId": existing[TAB]}}]}).execute()
    res = svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": [{"addSheet": {"properties": {
        "title": TAB, "gridProperties": {"rowCount": 30, "columnCount": 4, "frozenRowCount": 2}}}}]}).execute()
    gid = res["replies"][0]["addSheet"]["properties"]["sheetId"]
    title = ["石原 洋 さん 面談記録 ｜ 2026-06-02 ドトール船橋・対面（Indeed経由）"] + [""]*3
    values = [title, HDR] + ROWS + [[""]*4, [SOUKAN]+[""]*3]
    svc.spreadsheets().values().update(spreadsheetId=SID, range=f"'{TAB}'!A1",
        valueInputOption="USER_ENTERED", body={"values": values}).execute()

    RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
    CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
    WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
    ds=2; de=ds+len(ROWS); nb=len(values)-1
    reqs=[]
    for i,w in enumerate([40,210,470,300]):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":4},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":13},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":40},"fields":"pixelSize"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":1,"endRowIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":11},"wrapStrategy":"WRAP","horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ds,"endRowIndex":de},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"TOP","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    for k in range(len(ROWS)):
        bg = CARD if k%2 else WHT
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ds+k,"endRowIndex":ds+k+1},"cell":{"userEnteredFormat":{"backgroundColor":bg}},"fields":"userEnteredFormat.backgroundColor"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ds,"endRowIndex":de,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"fontSize":11}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ds,"endRowIndex":de,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat.textFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":ds,"endIndex":de},"properties":{"pixelSize":76},"fields":"pixelSize"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":nb,"endRowIndex":nb+1,"startColumnIndex":0,"endColumnIndex":4},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":nb,"endRowIndex":nb+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":nb,"endIndex":nb+1},"properties":{"pixelSize":74},"fields":"pixelSize"}})
    reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":1,"endRowIndex":de,"startColumnIndex":0,"endColumnIndex":4},"innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":reqs}).execute()
    print(f"DONE gid={gid}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
