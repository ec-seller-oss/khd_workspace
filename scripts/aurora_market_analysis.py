# -*- coding: utf-8 -*-
"""
オーロラ大元スプシ（宮崎と共有）に「競合・市場分析」タブを追加。2026-06-03 / 04李牧。
徳洲会=同意書出るが遅い を織込み。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID = "18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB = "14_競合市場分析(6_3)"

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

SEC = "§"  # セクション見出しの目印
LINES = [
    (SEC, "■ 結論"),
    ("1", "勝ち筋は明確。船橋の成否は『立地 × 同意書スピード』のほぼ1点に集約。"),
    ("2", "徳洲会(高根)は同意書『出るが遅い』→事業は成立するがランプは緩やか。CFは本部立替で耐える。律速は『採用』から『同意書スピード×利用者獲得』へ移った。"),
    ("3", "立地は高根(公団・木戸)でほぼ確定の筋。採用は本部紹介(0円)主軸＋石原+1〜2名の複線。"),
    ("4", "BEP=利用者15-20名(≒6ヶ月)。12ヶ月45-50名で営業利益65万、24ヶ月90-100名で138万(本部シミュ・ロイヤリティ12%)。"),
    (SEC, "■ 競合ポジショニング"),
    ("・", "駅前/主要動線は各社が押さえ済：リボン(船橋駅前)/ケイロウ(西船橋)/フレアス(北習志野・前原)/アシスト(高根?・江戸川)/津田沼(石原の現職)。"),
    ("・", "訪問マッサージは『駅商圏』でなく『在宅密度 × 同意書が出る医師会』で決まる。ケイロウ西船橋は団地少=在宅密度低、高根は団地密度高く筋が良い。"),
    ("・", "フレアスは習志野で同意書が出ず船橋側へ食い込みたいはず→高根を先に押さえる先行者利益が大きい。"),
    (SEC, "■ 市場・収益"),
    ("・", "船橋市 後期高齢者8万人超。利用率1-2%でも潜在800-1,600人＝1店は黒字15-20名/軌道90-100名で競合5社いても枯渇しない。"),
    ("・", "ストック型・季節変動なし・価格競争なし(医療保険9割)＝予測しやすい。"),
    ("・", "石原制約(9-10件/日・土日祝休)→1名の月売上は本部標準80-100万より下(推定60-75万/仮定)。単店BEPは射程だが拡大には複線採用が必須。"),
    ("・", "徳洲会『遅い』の含意：初回課金が後ろ倒し→初月CFが薄い(本部立替でカバー)/同意書更新を定型化/早く出る近隣診療所も並行開拓し導線分散。"),
    (SEC, "■ 課題（優先順）"),
    ("1", "🔴同意書スピード（最優先・律速）：徳洲会の遅さを、早く出る近隣診療所の複数開拓＋同意書更新フロー定型化で吸収。"),
    ("2", "立地確定：高根(公団/木戸)を在宅密度・競合不在・徳洲会導線で裏取り。"),
    ("3", "採用複線：石原1名依存回避。本部紹介で+1〜2名(高根で回るか)。ヒトイキは保険・紹介料確認後。"),
    ("4", "競業ガバナンス：石原の前職顧客持ち込みを雇用条件で禁止明文化。"),
    ("5", "稼働設計：石原の土日祝休・9-10件前提に、もう1名と曜日/エリアを補完(土日カバー)。"),
    ("6", "収益検証：高根の家賃・駐車場で本部シミュに当て込み、6/12/24ヶ月の利益とBEP月を確定。"),
    (SEC, "■ 6/4 板橋で取りに行く（更新）"),
    ("A", "同意書を『早く回す』実務：徳洲会の遅さの吸収法・早く出る診療科/クリニックの見つけ方。"),
    ("B", "立地の決め方：在宅密度の見極め、競合薄×同意書出るの交差点。"),
    ("C", "採用：本部紹介で船橋(高根)に何名・どの学校から。"),
    ("D", "収益：施術者1名あたり件数・売上・手残り、単店BEP。"),
    ("E", "競業：採用者が前職顧客を持ち込むリスクと本部の考え（トラブル経験）。"),
]

def main():
    svc = build("sheets", "v4", credentials=creds(), cache_discovery=False)
    info = svc.spreadsheets().get(spreadsheetId=SID, fields="sheets.properties").execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in info["sheets"]}
    if TAB in existing:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": [{"deleteSheet": {"sheetId": existing[TAB]}}]}).execute()
    res = svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": [{"addSheet": {"properties": {
        "title": TAB, "gridProperties": {"rowCount": 40, "columnCount": 2, "frozenRowCount": 1}}}}]}).execute()
    gid = res["replies"][0]["addSheet"]["properties"]["sheetId"]

    title = ["オーロラ船橋 競合・市場 深掘り分析 ｜ 2026-06-03（04李牧）※徳洲会=同意書出るが遅い 反映"]
    values = [title] + [["" if a == SEC else a, b] for (a, b) in LINES]
    svc.spreadsheets().values().update(spreadsheetId=SID, range=f"'{TAB}'!A1",
        valueInputOption="USER_ENTERED", body={"values": values}).execute()

    RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
    CARD={"red":0.945,"green":0.925,"blue":0.882}; WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
    reqs=[]
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":0,"endIndex":1},"properties":{"pixelSize":42},"fields":"pixelSize"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":1,"endIndex":2},"properties":{"pixelSize":900},"fields":"pixelSize"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":2},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":12},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":38},"fields":"pixelSize"}})
    # 全体 wrap
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":1,"endRowIndex":len(values)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"TOP","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    # セクション見出し行
    for idx,(a,b) in enumerate(LINES, start=1):
        if a==SEC:
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":idx,"endRowIndex":idx+1,"startColumnIndex":0,"endColumnIndex":2},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":idx,"endRowIndex":idx+1},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":11},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
        else:
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":idx,"endRowIndex":idx+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":reqs}).execute()
    print(f"DONE gid={gid}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
