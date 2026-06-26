# -*- coding: utf-8 -*-
"""
オーロラ大元スプシ（宮崎と共有）に2タブ追加。
17_船橋GO決定プロセス：本部へ船橋確定連絡までの逆算ロードマップ（穴埋め式）。
18_戦略再設計_課題と相談先：3軸×展開×学び＋課題/落とし穴/相談先(福井・宮崎・狩野・AI)マトリクス。
2026-06-04 狩野さん面談（投資判断GO）後。auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID = "18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"

RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
YEL={"red":1.0,"green":0.97,"blue":0.80}; WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

def build_tab(svc, title_tab, headline, rows, widths, status_col):
    N = len(widths)
    info = svc.spreadsheets().get(spreadsheetId=SID, fields="sheets.properties").execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in info["sheets"]}
    if title_tab in existing:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":[{"deleteSheet":{"sheetId":existing[title_tab]}}]}).execute()
    res = svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":[{"addSheet":{"properties":{
        "title":title_tab,"gridProperties":{"rowCount":len(rows)+4,"columnCount":N,"frozenRowCount":1}}}}]}).execute()
    gid = res["replies"][0]["addSheet"]["properties"]["sheetId"]
    values=[[headline]+[""]*(N-1)]; meta=[]
    for r in rows:
        i=len(values); k=r[0]
        if k=="SEC": values.append([r[1]]+[""]*(N-1)); meta.append((i,"SEC"))
        elif k=="KV": values.append([r[1], r[2]]+[""]*(N-2)); meta.append((i,"KV",r[2]))
        elif k=="THDR": values.append(r[1]); meta.append((i,"THDR"))
        elif k=="ROW": values.append(r[1]); meta.append((i,"ROW"))
        elif k=="NOTE": values.append([r[1]]+[""]*(N-1)); meta.append((i,"NOTE"))
    svc.spreadsheets().values().update(spreadsheetId=SID, range=f"'{title_tab}'!A1",
        valueInputOption="USER_ENTERED", body={"values":values}).execute()
    reqs=[]
    for c,w in enumerate(widths):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":c,"endIndex":c+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":len(values)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"TOP","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":12},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":38},"fields":"pixelSize"}})
    row_idx=[]
    for m in meta:
        i=m[0]; kind=m[1]
        if kind=="SEC":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":11},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
        elif kind=="KV":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":CARD,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
            if "穴埋め" in str(m[2]):
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"cell":{"userEnteredFormat":{"backgroundColor":YEL,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
        elif kind=="THDR":
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
        elif kind=="ROW":
            row_idx.append(i)
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat.textFormat"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":status_col,"endColumnIndex":status_col+1},"cell":{"userEnteredFormat":{"backgroundColor":YEL}},"fields":"userEnteredFormat.backgroundColor"}})
        elif kind=="NOTE":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
            reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":58},"fields":"pixelSize"}})
    if row_idx:
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":row_idx[0],"endIndex":row_idx[-1]+1},"properties":{"pixelSize":50},"fields":"pixelSize"}})
        reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":row_idx[0]-1,"endRowIndex":row_idx[-1]+1,"startColumnIndex":0,"endColumnIndex":N},"innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":reqs}).execute()
    return gid

T17 = [
    ("SEC","■ 概要・ゴール（6/4 狩野さん面談で投資GO）"),
    ("KV","投資判断","医療で着実・5年は安定成長／1年で回収できるリターン大と判断。"),
    ("KV","アッパー目標","2,000万円／エリア（採用と営業を強化して到達を狙う上限目標）。"),
    ("KV","展開構想","船橋でスタート→田舎(岩手)でも展開できる型に。医療AIコンサル×不動産×訪問マッサージFCの3軸で強化。"),
    ("KV","学びの置き方","1人オーナーで店舗運営は難しい→『人の採用マネジメント』を現場で足を使って学ぶフェーズと位置づけ。狩野さんと協力体制を強化。"),
    ("KV","★本部へ船橋確定連絡 目標日","【穴埋めで決定】　推奨＝6/24のGo判断と同時。※先行者利益（フレアスが習志野で苦戦＝高根を早く押さえる）の観点では前倒しも可"),
    ("SEC","■ 逆算ロードマップ（状態を黄色欄に穴埋め）"),
    ("THDR",["フェーズ/期間","タスク","担当","期限","状態【穴埋め】","完了条件・メモ"]),
    ("ROW",["P0 直近 6/4-6/10","狩野さんと協力体制を合意（採用・営業の連携／ノウハウの扱い）","菊池→狩野","6/10","","協力の枠組みを口頭合意→書面化へ"]),
    ("ROW",["P0","ヒトイキ候補(M00122307)の面接実施","菊池","6/10","","面接日確定→実施"]),
    ("ROW",["P0","Indeedをパート主軸に再設計＋掲載露出改善","菊池","6/10","","掲載更新"]),
    ("ROW",["P0","高根周辺の診療所へ同意書の出やすさを事前確認(徳洲会は遅い前提)","菊池","6/10","","2〜3件に感触ヒアリング"]),
    ("ROW",["P1 検証 6/11-6/20","高根の事務所候補・家賃で単店PL(6/12/24ヶ月)とBEPを確定","菊池","6/18","","本部シミュに当て込み"]),
    ("ROW",["P1","加盟金590万の資金段取り(公庫／自己資金600万)","菊池→橋本/公庫","6/18","","公庫打診の可否"]),
    ("ROW",["P1","妻の最終承認","菊池","6/20","","承認"]),
    ("ROW",["P1","採用の確度を上げる(内定1名 or 複線の見込み)","菊池","6/20","","採用○の根拠"]),
    ("ROW",["P2 GO判断 6/24目安","3点セット判定(採用○×収益○×同意書見込み○)","菊池","6/24","","Go/No-go"]),
    ("ROW",["P2","★本部(フライハイト)へ船橋エリア確定の連絡","菊池→本部","【穴埋め】","","ここが意思決定の確定点"]),
    ("ROW",["P3 加盟・開業 7月〜","加盟契約・初期費用支払・研修・開業準備","菊池","7月","","開業へ"]),
    ("SEC","■ 決定の条件（これが揃えば本部連絡）"),
    ("NOTE","①採用＝本部紹介 or Indeedパート or ヒトイキで『高根で人が採れる』確度／②収益＝高根家賃での単店BEPが見えている／③同意書＝高根近隣で早く出る先の目処。3つ揃った時点で本部へ船橋確定連絡（目標日を上の黄色欄に記入）。"),
]

T18 = [
    ("SEC","■ 戦略再設計の前提（3軸×展開×学び）"),
    ("KV","3軸","医療AIコンサル(フック)×不動産(キャッシュ化)×訪問マッサージFC(医療ストック収益)。"),
    ("KV","展開","船橋で型を作る→岩手(田舎)でも再現。territory承認・同意書・人口で横展開の可否を判断。"),
    ("KV","狩野さんの位置づけ","オーロラ経験豊富な協力者。ノウハウ販売は『協力体制を強化してから』。採用・営業・同意書の実務を学ぶ相手。"),
    ("KV","菊池の弱点(自覚)","総花でリスク管理が甘くなりがち→選択と集中、書面化、先に最悪想定で潰す。"),
    ("SEC","■ 課題・落とし穴・相談先マトリクス（状態を穴埋め）"),
    ("THDR",["#","論点・課題","想定される落とし穴","検証/聞くこと","相談先","期限","状態【穴埋め】"]),
    ("ROW",["1","同意書スピード(高根・徳洲会・近隣診療所)","徳洲会が遅く立ち上げが鈍る／エリアで出ない","早く出す診療科・先の見つけ方／算定までの日数","狩野・福井・現地診療所","6/10",""]),
    ("ROW",["2","立地(高根の在宅密度・町丁別75歳+)","団地でも実需要が薄い／競合(アシスト)が高根に居る","町丁別75歳+の実数／アシスト所在／在宅密度","AI(Claude/Gemini)・狩野","6/12",""]),
    ("ROW",["3","採用の複線化(本部紹介の再現性)","石原1名依存／船橋で本部紹介が回らない","本部紹介で高根に何名／学校／Indeed要否","狩野・本部・宮崎","6/14",""]),
    ("ROW",["4","単店収益・BEP(高根家賃で)","稼働限定(石原)で黒字に届かない","1人あたり件数／売上／手残り／BEP月","本部シミュ・狩野・宮崎","6/18",""]),
    ("ROW",["5","競業(石原の顧客持ち込み)","コンプラ赤信号でトラブル化","持ち込み禁止の明文化／本部の考え","本部・狩野","6/14",""]),
    ("ROW",["6","資金(加盟590万・公庫)","自己資金600万を圧迫／融資不調","公庫打診の可否／返済計画","橋本・公庫","6/18",""]),
    ("ROW",["7","1人オーナーで店舗が回るか","採用・営業・現場で破綻","採用マネジメントの実地学習(狩野に同行/相談)","狩野","継続",""]),
    ("ROW",["8","岩手展開の前提","需要／同意書／人材が薄く再現できない","岩手の人口・後期高齢者・territory・医師会","AI(Claude/Gemini)・将来","7月以降",""]),
    ("ROW",["9","3軸シナジーの設計","総花で集中が崩れる(菊池の弱点)","フック→キャッシュ化の導線設計","宮崎・福井・自分","6/20",""]),
    ("ROW",["10","狩野さんとの協力体制・ノウハウの扱い","関係が曖昧で空中分解／書面化なし","協力範囲・対価・ノウハウ販売の枠組み","狩野と直接","6/12",""]),
    ("SEC","■ 次アクション"),
    ("NOTE","落とし穴回避の原則＝『書面化(狩野/協力体制)』『先に最悪を想定(損切りライン)』『一次情報(狩野/現地)×市場データ(AI)で裏取り』。これらを潰したうえで、具体的な事業計画を作り直す。"),
]

def main():
    svc = build("sheets","v4",credentials=creds(),cache_discovery=False)
    g17 = build_tab(svc, "17_船橋GO決定プロセス(6_4)",
                    "船橋GO 決定プロセス ｜ 本部へ船橋確定連絡までの逆算（穴埋め式）2026-06-04",
                    T17, [140,300,130,80,150,260], status_col=4)
    g18 = build_tab(svc, "18_戦略再設計_課題と相談先(6_4)",
                    "戦略再設計 ｜ 課題・落とし穴・相談先(福井/宮崎/狩野/AI) 2026-06-04",
                    T18, [34,200,230,250,170,80,150], status_col=6)
    print(f"DONE 17 gid={g17}")
    print(f"DONE 18 gid={g18}")
    print(f"URL17: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={g17}")
    print(f"URL18: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={g18}")

if __name__=="__main__":
    main()
