# -*- coding: utf-8 -*-
"""
オーロラ高根店 事業計画 叩き台（営業ファネル型）。
営業訪問→接触→問合せ→CV→利用者ストック(解約4%/月)→施術者キャパ律速→PL/CPA/回収。
本部標準ランプでなく『足で営業』の実態でシビアに設計。2026-06-04。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID="18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB="21_事業計画叩き台_営業ファネル(6_4)"

# ── 前提（可変） ──
VISITS_DAY=30          # 営業訪問/日(9-11,16-18の合間/雨天はケアマネに会える)
SALES_DAYS=10          # 営業日/月
CONTACT=0.50           # 接触率(キーマン在席)
CV=0.55                # 問合せ→利用開始(同意書取得込)
CHURN=0.04             # 解約率/月
PRICE=38000            # 利用者あたり月単価(3.5-4万の中間)
CAP=22                 # 施術者1人あたり利用者上限
ROY=0.132; OFFICE=0.03; THER=0.30  # ロイヤリティ/事務手数料/施術者人件費(売上比)
SALES_COST=250000      # 営業人件費/月(専任1名・6人以上で2名=+25万)
RENT=80000; OTHER=60000  # 家賃/その他固定(光熱通信システム広告雑費)
PARK=12000; VEH=45000  # 施術者1人あたり 駐車/車両
INIT=7400000           # 加盟+初期590万+運転150万

def inquiry_rate(m):   # 問合せ率(関係構築でランプ)
    if m<=3: return 0.05
    if m<=9: return 0.09
    return 0.13
def new_users(m):
    return VISITS_DAY*SALES_DAYS*CONTACT*inquiry_rate(m)*CV

def simulate():
    rows=[]; u=0.0; ther=1; cum=-INIT
    for m in range(1,25):
        # 施術者キャパに達したら翌月増員
        if u > ther*CAP*0.85: ther+=1
        nu=new_users(m)
        u=min(u*(1-CHURN)+nu, ther*CAP)
        sales=u*PRICE
        scost=SALES_COST*(2 if ther>=6 else 1)
        cost=sales*(ROY+OFFICE+THER)+scost+RENT+OTHER+PARK*ther+VEH*ther
        op=sales-cost
        cum+=op
        rows.append(dict(m=m,users=round(u),ther=ther,new=round(nu,1),sales=round(sales),op=round(op),cum=round(cum)))
    return rows

S=simulate()
def at(m): return next(r for r in S if r["m"]==m)
def yen(x): return f"{x:,}"
# 回収月
payback=next((r["m"] for r in S if r["cum"]>=0), ">24")
cpa=round(SALES_COST/new_users(12))
ltv_gross=round(PRICE/CHURN)
ltv_margin=round(PRICE*(1-ROY-OFFICE-THER)/CHURN)
months=[3,6,9,12,18,24]

ROWS=[
 ("SEC","■ 前提（営業ファネル・可変セル）※叩き台。実績で随時上書き"),
 ("KV","営業活動","営業訪問 30件/日 × 営業日 10日/月 ＝ 300訪問/月（9-11・16-18時の合間／雨天はケアマネに会いやすい）。居宅キーマン(ケアマネ)約40事業所を2週で巡回＋施設・病院は自力開拓（本部は紹介しない）。"),
 ("KV","ファネル率","接触率50%／問合せ率 5%(〜3ヶ月)→9%(〜9ヶ月)→13%(成熟)／CV率(同意書取得込)55%。"),
 ("KV","ストック","解約率 4%/月。利用者単価 月3.8万。施術者1人あたり利用者上限 22名(超える前に増員)。"),
 ("KV","コスト","ロイヤリティ13.2%／事務手数料3%／施術者人件費=売上の30%／営業人件費25万(6人以上で2名)／家賃8万(高根想定)／駐車1.2万・車両4.5万/施術者／その他6万。"),
 ("KV","初期投資","加盟+初期 590万 + 運転 150万 = 740万。"),

 ("SEC","■ 月次シミュレーション（叩き台）"),
 ("THDR",["項目"]+[f"{m}ヶ月" for m in months]),
 ("ROW",["新規利用者/月"]+[str(at(m)["new"]) for m in months]),
 ("ROW",["利用者数(累計・解約後)"]+[str(at(m)["users"]) for m in months]),
 ("ROW",["施術者数"]+[str(at(m)["ther"]) for m in months]),
 ("ROW",["売上"]+[yen(at(m)["sales"]) for m in months]),
 ("ROW",["■ 営業利益/月"]+[("+" if at(m)["op"]>=0 else "")+yen(at(m)["op"]) for m in months]),
 ("ROW",["累計CF(初期-740万から)"]+[("+" if at(m)["cum"]>=0 else "")+yen(at(m)["cum"]) for m in months]),

 ("SEC","■ CPA・LTV・回収"),
 ("KV","CPA(顧客獲得単価)",f"営業人件費25万 ÷ 新規利用者 ≒ {yen(cpa)}円/人（成熟期）。"),
 ("KV","LTV(継続価値)",f"平均継続 25ヶ月(=1/解約4%)。売上ベース {yen(ltv_gross)}円／粗利ベース(ロイヤリティ・事務・施術人件費控除後) {yen(ltv_margin)}円。CPAの約{round(ltv_margin/max(cpa,1))}倍＝健全。"),
 ("KV","投資回収",f"累計CFがプラス転換＝約 {payback}ヶ月目（営業を足で回す前提）。本部標準ランプより速いが、施術者の採用が律速。"),
 ("KV","BEP",f"営業利益0は 利用者 約16名（売上約60万/月）。早ければ4〜6ヶ月で到達。"),

 ("SEC","■ 結論（叩き台の読み）"),
 ("NOTE","勝ち筋＝『足で営業(ケアマネ40件×反復)→解約4%でストック積み上げ』。本部標準(12ヶ月45-50名)よりシビアに営業すれば12ヶ月で利用者"+str(at(12)["users"])+"名・営業利益 月"+yen(at(12)["op"])+"円も射程。ただし利用者の伸びは"+"『施術者の採用』が律速（"+str(at(24)["ther"])+"名規模が必要）＝採用が事業の上限を決める。リスク＝①営業が回らないと新規が立たない②採用が追いつかないと頭打ち③同意書(徳洲会の遅れ)でCVが後ろ倒し。次＝この前提を狩野さん/本部の実数(接触率・問合せ率・CPA)で上書きして精緻化。"),
]

RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
YEL={"red":1.0,"green":0.97,"blue":0.80}; WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
N=7

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
    values=[["事業計画 叩き台｜オーロラ高根店（営業ファネル型・解約4%）2026-06-04 ※前提は可変・実績で上書き"]+[""]*(N-1)]
    meta=[]
    for r in ROWS:
        i=len(values); k=r[0]
        if k=="SEC": values.append([r[1]]+[""]*(N-1)); meta.append((i,"SEC"))
        elif k=="KV": values.append([r[1],r[2]]+[""]*(N-2)); meta.append((i,"KV"))
        elif k=="THDR": values.append(r[1]+[""]*(N-len(r[1]))); meta.append((i,"THDR"))
        elif k=="ROW": values.append(r[1]+[""]*(N-len(r[1]))); meta.append((i,"ROW",r[1][0]))
        elif k=="NOTE": values.append([r[1]]+[""]*(N-1)); meta.append((i,"NOTE"))
    svc.spreadsheets().values().update(spreadsheetId=SID,range=f"'{TAB}'!A1",valueInputOption="USER_ENTERED",body={"values":values}).execute()
    reqs=[]
    for c,w in enumerate([180,110,110,110,110,110,110]):
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
        elif kind=="ROW":
            lab=m[2]
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
            if "営業利益" in lab or "累計CF" in lab:
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":YEL,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
        elif kind=="NOTE":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"wrapStrategy":"WRAP","verticalAlignment":"TOP"}},"fields":"userEnteredFormat"}})
            reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":92},"fields":"pixelSize"}})
    thdr=[m[0] for m in meta if m[1]=="THDR"][0]; lastrow=[m[0] for m in meta if m[1]=="ROW"][-1]
    reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":thdr,"endRowIndex":lastrow+1,"startColumnIndex":0,"endColumnIndex":N},"innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":reqs}).execute()
    print(f"DONE gid={gid}")
    print("payback=",payback,"users12=",at(12)["users"],"op12=",at(12)["op"],"users24=",at(24)["users"],"op24=",at(24)["op"],"ther24=",at(24)["ther"])
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
