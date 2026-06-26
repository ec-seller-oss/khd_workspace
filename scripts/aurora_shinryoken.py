# -*- coding: utf-8 -*-
"""
オーロラ大元スプシ（宮崎と共有）に「診療圏調査_訪問マッサージ(高根台)」タブを追加。
2026-06-03。Web実データ＋石原ヒアリングで競合・診療圏・同意書(医師会)を深掘り。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID = "18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB = "15_診療圏調査_訪問MA(6_3)"
N = 6

# 行定義： ("SEC",text) / ("KV",label,value) / ("THDR",[6]) / ("ROW",[6]) / ("NOTE",text)
ROWS = [
    ("SEC", "Ⅰ. 調査概要"),
    ("KV", "目的", "オーロラFC船橋の出店候補『高根台エリア』の診療圏・競合・同意書環境を、訪問マッサージ版の診療圏調査として評価する。"),
    ("KV", "候補地（中心）", "船橋市 高根台2丁目近辺（千葉徳洲会病院＝高根台2-11-1 の徒歩圏）／新京成『高根公団』『高根木戸』駅圏"),
    ("KV", "調査日", "2026-06-03（Web資料＋石原氏ヒアリング6/2）"),
    ("KV", "事業形態", "訪問鍼灸マッサージ（医療保険・医師同意書ベース）。移動＝電動自転車/バイク"),

    ("SEC", "Ⅱ. 診療圏（訪問圏）設定"),
    ("KV", "訪問圏", "高根公団・高根木戸を中心に、チャリ圏（半径約2〜3km）。高根台団地＋松が丘・高根町・西習志野・前原・習志野台を内包。"),
    ("KV", "立地仮説の核", "①在宅高齢者密度（老朽UR団地）②船橋市医師会側で同意書が出る③主要駅は競合が押さえ済→高根台が空白、の3条件が重なる。"),

    ("SEC", "Ⅲ. 競合施設一覧（実在確認＋石原ヒアリング）"),
    ("THDR", ["No", "事業者・院名", "種別", "所在地（判明分）", "商圏内の位置", "同意書・備考"]),
    ("ROW", ["1", "千葉徳洲会病院", "病院（同意書元）", "船橋市高根台2-11-1", "圏の中心", "同意書は出るが『遅い』(石原)＝立ち上げの律速。近接は強み"]),
    ("ROW", ["2", "レイス治療院 船橋（フレアス系）", "訪問マ・鍼灸FC", "船橋市西習志野1-13-13", "圏 西（習志野寄り）", "習志野市医師会側で同意書ハンデ＝伸び悩む構造。差別化余地"]),
    ("ROW", ["3", "リボン", "訪問マ", "船橋駅前（石原情報）", "圏外 南西", "駅前商圏。高根とは別エリア"]),
    ("ROW", ["4", "ケイロウ(KEiROW)", "訪問マFC", "西船橋駅（石原情報）", "圏外 西", "団地少で手薄との情報"]),
    ("ROW", ["5", "アシスト", "訪問マ", "高根?・江戸川（石原情報）", "圏内の可能性", "高根に居る可能性＝要現地確認（最重要）"]),
    ("ROW", ["6", "津田沼の独自FC（石原の現職）", "訪問マ", "津田沼周辺", "圏外 南", "新規が減少中。石原の競業元"]),
    ("ROW", ["7", "京葉訪問マッサージセンター", "訪問マ・鍼灸", "船橋/習志野/市川/鎌ヶ谷 広域", "広域", "広域分散型"]),
    ("ROW", ["8", "リカバリー 訪問マッサージ船橋", "訪問マ", "船橋市", "広域", "広域分散型"]),
    ("ROW", ["9", "匠たくみ治療院", "訪問鍼灸マ", "船橋市坪井町135-4", "圏 北（八千代寄り）", "北部中心"]),
    ("NOTE", "※来院型の あマ指/鍼灸 施術所は船橋市『はり・きゅう・マッサージ費用助成 登録施術所一覧』に多数。訪問専業の直接競合は上記。アシストの高根所在は現地確認の最優先事項。"),

    ("SEC", "Ⅳ. 診療圏人口・需要推定"),
    ("KV", "船橋市 全体", "人口 642,907人（2020国勢調査）／高齢化率 24.3%／65歳以上 約15.6万人／75歳以上（後期高齢者）約8万人 ※後期は要実数確認"),
    ("KV", "高根台エリアの特性", "高根台団地＝1961年入居開始の大規模UR団地。老朽化で建替中、高齢者施設も新設＝高齢化が進行＝在宅医療マッサージ需要が構造的に高い。"),
    ("KV", "需要推定（概算）", "圏内75歳以上を仮に8,000〜12,000人とし、訪問マッサージ利用率1.0〜1.5%なら潜在利用者 約80〜180人。1店のBEPは利用者15〜20名＝需要は十分（要：町丁別実数で精緻化）。"),
    ("KV", "次段（精緻化）", "船橋市『町丁別・年齢別人口』Excelで 高根台/松が丘/高根町/西習志野/習志野台 の75歳以上を取得し母数を確定する。"),

    ("SEC", "Ⅴ. 同意書（医師会）マップ ★収益の急所"),
    ("KV", "船橋市医師会（高根台側）", "同意書は出る。ただし徳洲会(高根台2)は『出るが遅い』＝初回課金が後ろ倒し→ランプ緩やか（CFは本部立替で耐える）。"),
    ("KV", "習志野市医師会（西習志野・北習志野）", "同意書が出ない。レイス/フレアスが西習志野に居ても伸びにくい構造的理由＝裏取れた。"),
    ("KV", "打ち手", "徳洲会(遅い)に依存せず、高根台周辺の整形外科・内科クリニックで『早く出す先』を複数開拓し導線を分散。同意書更新フローを定型化。"),

    ("SEC", "Ⅵ. 総合結論"),
    ("NOTE", "高根台（徳洲会近接・新京成 高根公団/高根木戸）は、『在宅高齢者密度（老朽団地）× 船橋市医師会で同意書が出る × 主要駅競合の空白』が重なる最有力候補。律速は徳洲会の同意書スピード→早く出る近隣診療所の複線で吸収。レイス(西習志野)は習志野医師会ハンデで伸び悩む構造＝高根台側を取れば明確に差別化できる。→ 6/4板橋では『同意書を早く回す実務』『アシストの高根所在』『本部紹介で高根に何名』を最優先で詰める。"),

    ("SEC", "Ⅶ. 出典・ご注意"),
    ("NOTE", "出典：船橋市統計(人口642,907・高齢化率24.3%／city.funabashi.lg.jp)、千葉徳洲会病院=高根台2-11-1(chibatoku.or.jp)、レイス治療院船橋/各競合公式、高根台団地(Wikipedia/UR)、石原氏ヒアリング(6/2)。町丁別75歳以上の実数とアシストの所在は未確定＝現地/Excelで要裏取り。数値は概算で将来を保証しない。"),
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
WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}

def main():
    svc = build("sheets", "v4", credentials=creds(), cache_discovery=False)
    info = svc.spreadsheets().get(spreadsheetId=SID, fields="sheets.properties").execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in info["sheets"]}
    if TAB in existing:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": [{"deleteSheet": {"sheetId": existing[TAB]}}]}).execute()
    res = svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": [{"addSheet": {"properties": {
        "title": TAB, "gridProperties": {"rowCount": len(ROWS)+4, "columnCount": N, "frozenRowCount": 1}}}}]}).execute()
    gid = res["replies"][0]["addSheet"]["properties"]["sheetId"]

    values = [["診療圏調査（訪問マッサージ版）｜ オーロラFC船橋・高根台エリア　2026-06-03"]+[""]*(N-1)]
    meta = []  # (rowidx0based, kind)
    for r in ROWS:
        i = len(values)
        k = r[0]
        if k == "SEC":
            values.append([r[1]]+[""]*(N-1)); meta.append((i,"SEC"))
        elif k == "KV":
            values.append([r[1], r[2]]+[""]*(N-2)); meta.append((i,"KV"))
        elif k == "THDR":
            values.append(r[1]); meta.append((i,"THDR"))
        elif k == "ROW":
            values.append(r[1]); meta.append((i,"ROW"))
        elif k == "NOTE":
            values.append([r[1]]+[""]*(N-1)); meta.append((i,"NOTE"))
    svc.spreadsheets().values().update(spreadsheetId=SID, range=f"'{TAB}'!A1",
        valueInputOption="USER_ENTERED", body={"values": values}).execute()

    reqs=[]
    for c,w in enumerate([40,230,140,210,150,360]):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":c,"endIndex":c+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    # 全体wrap
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":len(values)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"TOP","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    # タイトル
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":12},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":38},"fields":"pixelSize"}})
    for (i,kind) in meta:
        if kind=="SEC":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":11},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
        elif kind=="KV":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":CARD,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
        elif kind=="THDR":
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
        elif kind=="NOTE":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
    # 競合テーブルの罫線（THDR～最後のROW）
    thdr_rows=[i for (i,k) in meta if k=="THDR"]
    row_rows=[i for (i,k) in meta if k=="ROW"]
    if thdr_rows and row_rows:
        reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":thdr_rows[0],"endRowIndex":row_rows[-1]+1,"startColumnIndex":0,"endColumnIndex":N},"innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":reqs}).execute()
    print(f"DONE gid={gid}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
