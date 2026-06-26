# -*- coding: utf-8 -*-
"""
オーロラ大元スプシに「市場リサーチ_高根&岩手」タブ追加（課題②⑧のAI市場リサーチ結果）。2026-06-04。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID="18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB="19_市場リサーチ_高根&岩手(6_4)"
N=2
RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}

ROWS=[
 ("SEC","■ 高根エリア（船橋・出店候補）"),
 ("KV","船橋市 全体","人口 642,907（2020国勢）／高齢化率 24.3%／75歳以上 約8万人。"),
 ("KV","高根台の特性","高根台団地＝1961年入居の大規模UR、老朽建替中で高齢者施設も新設＝高齢化進行＝在宅マッサージ需要が構造的に厚い。"),
 ("KV","町丁別75歳+（要Excel）","高根台/高根町/松が丘/西習志野の75歳+実数は船橋市『町丁別・年齢別人口』Excel（p029445）で取得＝次段で確定。"),
 ("KV","★競合アシストの実態","『あしすと訪問リハビリ鍼灸マッサージ院』は江戸川区が拠点で船橋も訪問対象エリア。高根に物理拠点なし＝石原の『高根?』は拠点でなく訪問範囲の話。→ 高根の訪問競合密度はむしろ低い＝追い風。"),
 ("KV","同意書（再掲）","千葉徳洲会(高根台2-11-1)＝出るが遅い。習志野市側(西習志野/北習志野)＝出ない(レイス/フレアスのハンデ)。船橋市医師会側の高根が筋。"),
 ("KV","◎ 高根の結論","在宅高齢者密度(老朽団地)×船橋市医師会で同意書出る×訪問競合の物理拠点が薄い＝出店地としての妥当性がさらに補強された。"),
 ("SEC","■ 岩手（盛岡）展開の前提"),
 ("KV","盛岡市 人口・高齢化","人口 約28〜29万／老年人口(65歳+)28.4%(2020・上昇中、2050年に約40%見通し)。75歳+ 推定 約4〜4.5万（要実数）。"),
 ("KV","★オーロラ未進出","オーロラ全国 約60拠点(宮城・北海道・沖縄等)に岩手は含まれず＝岩手はterritory空き＝先行者で取れる可能性大。"),
 ("KV","競合","レイス治療院 盛岡(フレアス系)が既存。地方は移動＝車前提・広域。"),
 ("KV","留意","地方は人材(あマ指/鍼灸)の母数が薄い可能性＝採用が船橋以上のボトルネックに。盛岡の医師会・同意書事情も別途確認要。"),
 ("KV","△ 岩手の結論","territory空きは魅力だが、船橋で型(採用・同意書・収益)を完成→横展開が順序。岩手単独先行はリスク。"),
 ("SEC","■ 出典・要確認"),
 ("NOTE","出典：船橋市統計(city.funabashi.lg.jp)／千葉徳洲会病院(高根台2-11-1)／あしすと訪問院(houmon-care.com)／オーロラ公式(aurora-houmon.com・全国約60拠点)／レイス治療院盛岡(leis.jp/iwate/morioka)／盛岡市人口ビジョン(city.morioka.iwate.jp)。要実数確認＝高根町丁別75歳+／盛岡75歳+の確定値／盛岡の医師会同意書事情。数値は概算で将来を保証しない。"),
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
    values=[["市場リサーチ｜高根(船橋) & 岩手(盛岡)　2026-06-04（Claude調べ・一次調査）"]+[""]*(N-1)]
    meta=[]
    for r in ROWS:
        i=len(values); k=r[0]
        if k=="SEC": values.append([r[1]]+[""]*(N-1)); meta.append((i,"SEC"))
        elif k=="KV": values.append([r[1],r[2]]); meta.append((i,"KV"))
        elif k=="NOTE": values.append([r[1]]+[""]*(N-1)); meta.append((i,"NOTE"))
    svc.spreadsheets().values().update(spreadsheetId=SID,range=f"'{TAB}'!A1",valueInputOption="USER_ENTERED",body={"values":values}).execute()
    reqs=[]
    for c,w in enumerate([220,760]):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":c,"endIndex":c+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":len(values)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"TOP","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":12},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":38},"fields":"pixelSize"}})
    for (i,kind) in meta:
        if kind=="SEC":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":11},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
        elif kind=="KV":
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":CARD,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
        elif kind=="NOTE":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":9,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
            reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":56},"fields":"pixelSize"}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":reqs}).execute()
    print(f"DONE gid={gid}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
