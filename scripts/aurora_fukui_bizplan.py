# -*- coding: utf-8 -*-
"""
福井式 事業計画書フォーマットにオーロラ高根を落とし込み（5年損益）＋抜け漏れチェック。
営業ファネル型(tab21)の前提を60ヶ月に延長して年度別に集計。2026-06-05。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID="18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB="22_事業計画書_福井式_高根(6_5)"

# 前提（tab21と同一）
VISITS_DAY=30; SALES_DAYS=10; CONTACT=0.50; CV=0.55; CHURN=0.04; PRICE=38000; CAP=22
ROY=0.132; OFFICE=0.03; THER=0.30; SALES_COST=250000; RENT=80000; OTHER=60000; PARK=12000; VEH=45000
LOAN=4400000; LOAN_RATE=0.018; LOAN_YEARS=10; TAX=0.25

def inquiry(m): return 0.05 if m<=3 else (0.09 if m<=9 else 0.13)
def newu(m): return VISITS_DAY*SALES_DAYS*CONTACT*inquiry(m)*CV

def sim(months=60):
    out=[]; u=0.0; ther=1
    for m in range(1,months+1):
        if u>ther*CAP*0.85: ther+=1
        u=min(u*(1-CHURN)+newu(m), ther*CAP)
        sales=u*PRICE
        scost=SALES_COST*(2 if ther>=6 else 1)
        op=sales-(sales*(ROY+OFFICE+THER)+scost+RENT+OTHER+PARK*ther+VEH*ther)
        out.append(dict(m=m,u=u,ther=ther,sales=sales,op=op))
    return out
S=sim(60)
# 公庫 元利均等 月返済
r=LOAN_RATE/12; n=LOAN_YEARS*12
pmt=LOAN*r*(1+r)**n/((1+r)**n-1)
def year_rows(y):
    ms=[x for x in S if (y-1)*12 < x["m"] <= y*12]
    sales=sum(x["sales"] for x in ms); op=sum(x["op"] for x in ms)
    # 残高から年利息概算
    bal=LOAN-(LOAN/n)*((y-1)*12)
    interest=max(bal,0)*LOAN_RATE
    pre=op-interest
    tax=max(pre,0)*TAX
    aftertax=pre-tax
    repay_principal=min(LOAN/LOAN_YEARS, max(bal,0))
    cf=aftertax-repay_principal
    return dict(uend=round(ms[-1]["u"]),ther=ms[-1]["ther"],sales=round(sales),op=round(op),
                interest=round(interest),pre=round(pre),tax=round(tax),aftertax=round(aftertax),
                repay=round(repay_principal),cf=round(cf))
Y=[year_rows(y) for y in range(1,6)]
def yen(x): return f"{x:,}"
def sg(x): return ("+" if x>=0 else "")+yen(x)

ROWS=[
 ("SEC","■ 1. 事業概要"),
 ("KV","事業・院名","オーロラ高根店（訪問鍼灸マッサージ・FC加盟＝株式会社フライハイト）"),
 ("KV","形態／開業区分","FC加盟・新規開業（医療保険適用の訪問マッサージ・はり灸／消費税非課税）"),
 ("KV","開業予定","2026年8〜9月（本部への船橋確定連絡後）"),
 ("KV","経営者","菊池 研太（KHD／テナントアシスト・ウイン）"),
 ("KV","拠点","船橋市高根台 近辺（新京成 高根公団／高根木戸）・小規模事務所"),
 ("SEC","■ 2. 診療圏（→15・19タブ 診療圏調査）"),
 ("KV","診療圏","高根台／松が丘／高根町／西習志野 等、新京成沿線のチャリ圏（半径約2〜3km）。"),
 ("KV","需要母数","船橋市75歳以上 約8万人。圏内の潜在利用者 約80〜180名（町丁別実数は要取得）。"),
 ("KV","同意書(医療特有)","船橋市医師会＝出る／千葉徳洲会(高根台2-11-1)＝出るが遅い／習志野市側＝出ない。"),
 ("SEC","■ 3. 人員・単価・稼働（福井式インプット）"),
 ("KV","施術者(あマ指/鍼灸)","開業1名→24ヶ月で8名（利用者22名/人で増員＝採用が律速）。"),
 ("KV","営業","専任1名（ケアマネ約40事業所を2週で巡回＋施設・病院は自力。本部は紹介しない）。"),
 ("KV","単価・稼働","利用者単価 月3.8万円。営業ファネル＝訪問300/月×接触50%×問合せ5→13%×CV55%。"),
 ("KV","解約率","4%/月（平均継続25ヶ月）。"),
 ("SEC","■ 4. 年度別損益（オーロラ高根・自社試算／円）"),
 ("THDR",["項目","1年目","2年目","3年目","4年目","5年目"]),
 ("ROW",["利用者数(年末)"]+[str(y["uend"]) for y in Y]),
 ("ROW",["施術者数(年末)"]+[str(y["ther"]) for y in Y]),
 ("ROW",["売上(年)"]+[yen(y["sales"]) for y in Y]),
 ("ROW",["営業利益(年)"]+[sg(y["op"]) for y in Y]),
 ("ROW",["支払利息"]+[yen(y["interest"]) for y in Y]),
 ("ROW",["税引前損益"]+[sg(y["pre"]) for y in Y]),
 ("ROW",["法人税等(25%)"]+[yen(y["tax"]) for y in Y]),
 ("ROW",["税引後利益"]+[sg(y["aftertax"]) for y in Y]),
 ("ROW",["借入元金返済"]+[yen(y["repay"]) for y in Y]),
 ("ROW",["■ 返済後CF"]+[sg(y["cf"]) for y in Y]),
 ("SEC","■ 5. 資金調達・回収"),
 ("KV","所要資金","加盟・初期 590万 ＋ 運転 150万 ＝ 740万。"),
 ("KV","調達(案)",f"自己資金 300万 ＋ 公庫借入 440万（{LOAN_YEARS}年・年{LOAN_RATE*100:.1f}%・元利均等＝月返済 約{round(pmt):,}円）。※借入額・期間は要相談。"),
 ("KV","投資回収","営業利益ベースで約17ヶ月（tab21）。返済後CFは2年目から黒字化の見込み。"),
 ("SEC","■ 6. 福井さんに見てほしい抜け漏れ・論点"),
 ("ROW",["①","資金調達・返済","加盟590万+運転150万の最適な自己資金/公庫バランス・期間・利率は？運転資金150万で足りるか。"]),
 ("ROW",["②","5年損益の現実性","利用者は『ケアマネ営業→紹介→同意書→解約4%でストック』。この積み上げ前提は甘くないか。"]),
 ("ROW",["③","同意書の実務(本領)","徳洲会は出るが遅い。同意書取得・更新で立ち上げを早めるコツ／取りやすい診療科・先の見極めは？"]),
 ("ROW",["④","労務・税務","施術者を雇用/業務委託どちらで設計すべきか（社保・労災・源泉）。医療保険非課税と課税事業の区分。"]),
 ("ROW",["⑤","減価償却・リース","車両・施術機材・サイネージのリース/償却の扱い。"]),
 ("ROW",["⑥","開業手続き","施術所開設届・出張施術業務開始届(保健所)の段取り・期間。"]),
 ("ROW",["⑦","採用・離職","施術者8名規模の労務・離職対策(本部離職率4%の実効性)。"]),
 ("ROW",["⑧","物件・契約","高根の小事務所の用途・契約・原状回復。FC契約の不利条項・中途解約・テリトリー。"]),
 ("SEC","■ 7. 相談サマリ"),
 ("NOTE","オーロラ高根の事業計画を福井式フォーマットに落とした叩き台。投資判断はGO方向。プロ目線で①資金・返済②5年損益の現実性③同意書の実務④労務税務を中心に、抜け漏れ・甘い前提・落とし穴をご指摘いただきたい。診療圏は15・19タブ、営業ファネルの根拠は21タブ参照。"),
]

RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
YEL={"red":1.0,"green":0.97,"blue":0.80}; WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
N=6

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
    values=[["事業計画書（福井式）｜オーロラ高根店　2026-06-05 ※福井さんレビュー用 叩き台"]+[""]*(N-1)]
    meta=[]
    for rr in ROWS:
        i=len(values); k=rr[0]
        if k=="SEC": values.append([rr[1]]+[""]*(N-1)); meta.append((i,"SEC"))
        elif k=="KV": values.append([rr[1],rr[2]]+[""]*(N-2)); meta.append((i,"KV"))
        elif k=="THDR": values.append(rr[1]+[""]*(N-len(rr[1]))); meta.append((i,"THDR"))
        elif k=="ROW":
            row=rr[1]
            if len(row)==3: row=[row[0],row[1],row[2],"","",""]  # 抜け漏れ行(論点を C 以降結合)
            values.append(row+[""]*(N-len(row))); meta.append((i,"ROW",row[0],len(rr[1])))
        elif k=="NOTE": values.append([rr[1]]+[""]*(N-1)); meta.append((i,"NOTE"))
    svc.spreadsheets().values().update(spreadsheetId=SID,range=f"'{TAB}'!A1",valueInputOption="USER_ENTERED",body={"values":values}).execute()
    reqs=[]
    for c,w in enumerate([150,140,140,140,140,150]):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":c,"endIndex":c+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":len(values)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":12}}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":36},"fields":"pixelSize"}})
    for mm in meta:
        i=mm[0]; kind=mm[1]
        if kind=="SEC":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":11}}},"fields":"userEnteredFormat"}})
        elif kind=="KV":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":CARD,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
        elif kind=="THDR":
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat"}})
        elif kind=="ROW":
            lab=mm[2]; ncols=mm[3]
            if ncols==3:  # 抜け漏れ行：B論点 太字、C以降結合
                reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":2,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
            else:  # 損益行
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
                if "営業利益" in lab or "返済後CF" in lab:
                    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":YEL,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
                    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
        elif kind=="NOTE":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"wrapStrategy":"WRAP","verticalAlignment":"TOP"}},"fields":"userEnteredFormat"}})
            reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":72},"fields":"pixelSize"}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":reqs}).execute()
    print(f"DONE gid={gid}")
    for i,y in enumerate(Y,1): print(f"Y{i}: 利用者{y['uend']} 売上{y['sales']:,} 営業利益{y['op']:,} 返済後CF{y['cf']:,}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
