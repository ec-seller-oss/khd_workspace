# -*- coding: utf-8 -*-
"""
福井さん提出用の【独立した新スプレッドシート】を作成。
Tab1=診療圏調査 / Tab2=事業計画書(福井式・5年損益＋抜け漏れ) / Tab3=相談文。
2026-06-05。auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
TITLE="オーロラ高根_福井さん提出_診療圏調査＆事業計画書(6_5)"

# ── 事業計画 前提（tab21/22と同一）──
VISITS_DAY=30; SALES_DAYS=10; CONTACT=0.50; CV=0.55; CHURN=0.04; PRICE=38000; CAP=22
ROY=0.132; OFFICE=0.03; THER=0.30; SALES_COST=250000; RENT=80000; OTHER=60000; PARK=12000; VEH=45000
LOAN=4400000; LOAN_RATE=0.018; LOAN_YEARS=10; TAX=0.25
def inq(m): return 0.05 if m<=3 else (0.09 if m<=9 else 0.13)
def newu(m): return VISITS_DAY*SALES_DAYS*CONTACT*inq(m)*CV
def sim(n=60):
    out=[];u=0.0;ther=1
    for m in range(1,n+1):
        if u>ther*CAP*0.85: ther+=1
        u=min(u*(1-CHURN)+newu(m),ther*CAP)
        sales=u*PRICE; scost=SALES_COST*(2 if ther>=6 else 1)
        op=sales-(sales*(ROY+OFFICE+THER)+scost+RENT+OTHER+PARK*ther+VEH*ther)
        out.append(dict(m=m,u=u,ther=ther,sales=sales,op=op))
    return out
S=sim(60)
def yp(y):
    ms=[x for x in S if (y-1)*12<x["m"]<=y*12]
    sales=sum(x["sales"] for x in ms);op=sum(x["op"] for x in ms)
    bal=LOAN-(LOAN/(LOAN_YEARS*12))*((y-1)*12); interest=max(bal,0)*LOAN_RATE
    pre=op-interest; tax=max(pre,0)*TAX; aftertax=pre-tax
    repay=min(LOAN/LOAN_YEARS,max(bal,0)); cf=aftertax-repay
    return dict(uend=round(ms[-1]["u"]),ther=ms[-1]["ther"],sales=round(sales),op=round(op),pre=round(pre),aftertax=round(aftertax),cf=round(cf),interest=round(interest),tax=round(tax))
Y=[yp(y) for y in range(1,6)]
def yen(x): return f"{x:,}"
def sg(x): return ("+" if x>=0 else "")+yen(x)
r=LOAN_RATE/12;n=LOAN_YEARS*12; pmt=LOAN*r*(1+r)**n/((1+r)**n-1)

# ── 各タブの行定義 ──
DIAG=[
 ("SEC","■ 1. 調査概要"),
 ("KV","目的","オーロラFC船橋の出店候補『高根台エリア』の診療圏・競合・同意書環境を訪問マッサージ版の診療圏調査として評価。"),
 ("KV","候補地","船橋市 高根台2丁目近辺（千葉徳洲会病院＝高根台2-11-1 の徒歩圏）／新京成『高根公団』『高根木戸』駅圏。"),
 ("KV","事業形態","訪問鍼灸マッサージ（医療保険・医師同意書ベース・消費税非課税）。移動＝電動自転車/バイク。"),
 ("SEC","■ 2. 診療圏（訪問圏）"),
 ("KV","訪問圏","高根公団・高根木戸を中心にチャリ圏（半径約2〜3km）。高根台団地＋松が丘・高根町・西習志野・前原を内包。"),
 ("KV","立地仮説","①在宅高齢者密度(老朽UR団地) ②船橋市医師会側で同意書が出る ③主要駅は競合が押さえ済→高根台が空白、の3条件が重なる。"),
 ("SEC","■ 3. 診療圏人口・需要"),
 ("KV","船橋市 全体","人口642,907(2020国勢)／高齢化率24.3%／75歳以上 約8万人。"),
 ("KV","高根台の特性","高根台団地＝1961年入居の大規模UR、老朽建替中で高齢者施設も新設＝高齢化進行＝在宅需要が構造的に厚い。"),
 ("KV","需要推定","圏内75歳+を8,000〜12,000人と仮定、利用率1.0〜1.5%で潜在利用者 約80〜180人。1店BEPは利用者16名＝需要は十分（町丁別実数は要取得）。"),
 ("SEC","■ 4. 競合（実在確認＋一次情報）"),
 ("KV","徳洲会","千葉徳洲会病院(高根台2-11-1)＝同意書元。出るが『遅い』＝立ち上げの律速。近接が強み。"),
 ("KV","レイス治療院船橋(フレアス系)","西習志野1-13-13。習志野市医師会側で同意書ハンデ＝伸び悩む構造。＝高根台側を取れば差別化。"),
 ("KV","その他競合","リボン(船橋駅前)/ケイロウ(西船橋・団地少)/アシスト(江戸川拠点で船橋は訪問範囲のみ＝高根に物理拠点なし)/京葉/匠(坪井・北部)。"),
 ("KV","◎含意","高根は訪問競合の物理拠点が薄い＝追い風。"),
 ("SEC","■ 5. 同意書（医師会）マップ ★収益の急所"),
 ("KV","船橋市医師会(高根台側)","出る。ただし徳洲会は遅い→初回課金が後ろ倒し(CFは本部立替で耐える)。"),
 ("KV","習志野市医師会(西習志野等)","出ない＝レイス/フレアスのハンデ。"),
 ("KV","打ち手","徳洲会に依存せず、高根台周辺の整形/内科で『早く出す先』を複数開拓し導線分散。更新フロー定型化。"),
 ("SEC","■ 6. 総合結論"),
 ("NOTE","高根台は『在宅高齢者密度(老朽団地)×船橋市医師会で同意書出る×主要駅競合の空白』が重なる最有力候補。律速＝徳洲会の同意書スピード→近隣診療所の複線で吸収。出典：船橋市統計/千葉徳洲会病院/各競合公式/石原氏ヒアリング。数値は概算で将来を保証しない。"),
]

BIZ=[
 ("SEC","■ 1. 事業概要"),
 ("KV","事業・院名","オーロラ高根店（訪問鍼灸マッサージ・FC加盟＝株式会社フライハイト）"),
 ("KV","形態／開業区分","FC加盟・新規開業（医療保険適用・消費税非課税）"),
 ("KV","開業予定","2026年8〜9月（本部への船橋確定連絡後）"),
 ("KV","経営者","菊池 研太（KHD／テナントアシスト・ウイン）"),
 ("KV","拠点","船橋市高根台 近辺（新京成 高根公団／高根木戸）・小規模事務所"),
 ("SEC","■ 2. 人員・単価・稼働（福井式インプット）"),
 ("KV","施術者(あマ指/鍼灸)","開業1名→24ヶ月で8名（利用者22名/人で増員＝採用が律速）。"),
 ("KV","営業","専任1名（ケアマネ約40事業所を2週で巡回＋施設・病院は自力。本部は紹介しない）。"),
 ("KV","単価・稼働","利用者単価 月3.8万円。営業ファネル＝訪問300/月×接触50%×問合せ5→13%×CV55%。"),
 ("KV","解約率","4%/月（平均継続25ヶ月）。"),
 ("SEC","■ 3. 年度別損益（自社試算／円）"),
 ("THDR",["項目","1年目","2年目","3年目","4年目","5年目"]),
 ("ROW",["利用者数(年末)"]+[str(y["uend"]) for y in Y]),
 ("ROW",["施術者数(年末)"]+[str(y["ther"]) for y in Y]),
 ("ROW",["売上(年)"]+[yen(y["sales"]) for y in Y]),
 ("ROW",["営業利益(年)"]+[sg(y["op"]) for y in Y]),
 ("ROW",["支払利息"]+[yen(y["interest"]) for y in Y]),
 ("ROW",["税引前損益"]+[sg(y["pre"]) for y in Y]),
 ("ROW",["税引後利益"]+[sg(y["aftertax"]) for y in Y]),
 ("ROW",["■ 返済後CF"]+[sg(y["cf"]) for y in Y]),
 ("SEC","■ 4. 資金調達・回収"),
 ("KV","所要資金","加盟・初期 590万 ＋ 運転 150万 ＝ 740万。"),
 ("KV","調達(案)",f"自己資金 300万 ＋ 公庫借入 440万（{LOAN_YEARS}年・年{LOAN_RATE*100:.1f}%・元利均等＝月返済 約{round(pmt):,}円）。※借入額・期間は要相談。"),
 ("KV","BEP・回収","BEP＝利用者16名(売上約60万/月)。営業利益ベースの投資回収は約17ヶ月。"),
 ("SEC","■ 5. 福井さんに見てほしい抜け漏れ・論点"),
 ("ROW3",["①","資金調達・返済","加盟590万+運転150万の最適な自己資金/公庫バランス・期間・利率は？運転150万で足りるか。"]),
 ("ROW3",["②","5年損益の現実性","利用者は『ケアマネ営業→紹介→同意書→解約4%でストック』。この積み上げ前提は甘くないか。"]),
 ("ROW3",["③","同意書の実務(本領)","徳洲会は出るが遅い。同意書取得・更新で立ち上げを早めるコツ／取りやすい診療科・先の見極めは？"]),
 ("ROW3",["④","労務・税務","施術者を雇用/業務委託どちらで設計すべきか（社保・労災・源泉）。医療保険非課税と課税事業の区分。"]),
 ("ROW3",["⑤","減価償却・リース","車両・施術機材・サイネージのリース/償却の扱い。"]),
 ("ROW3",["⑥","開業手続き","施術所開設届・出張施術業務開始届(保健所)の段取り・期間。"]),
 ("ROW3",["⑦","採用・離職","施術者8名規模の労務・離職対策(本部離職率4%の実効性)。"]),
 ("ROW3",["⑧","物件・契約","高根の小事務所の用途・契約。FC契約の不利条項・中途解約・テリトリー。"]),
 ("SEC","■ 6. 総括"),
 ("NOTE","投資判断はGO方向。プロ目線で①資金・返済②5年損益の現実性③同意書の実務④労務税務を中心に、抜け漏れ・甘い前提・落とし穴をご指摘いただきたい。"),
]

SOUDAN_TEXT=("福井さん\n\nお世話になります、菊池です。\n"
 "先日ご相談したオーロラ（訪問鍼灸マッサージFC）の船橋・高根での開業、投資判断はGO方向で固めつつあります。"
 "福井さんの開業ノウハウ（診療圏調査・事業計画書のフォーマット）に、こちらの事業を落とし込んでみました。\n\n"
 "医療開業のプロの目線で、抜け漏れ・前提の甘さ・落とし穴を見ていただきたく、本ファイルを共有します。\n"
 "・1_診療圏調査タブ\n・2_事業計画書(福井式・5年損益)タブ\n\n"
 "特に相談したいのは下記4点です：\n"
 "①資金調達・返済：加盟590万+運転150万=740万。自己資金と公庫借入のバランス・期間・利率の最適は？運転150万で足りるか。\n"
 "②5年損益の現実性：利用者は『ケアマネ営業→紹介→同意書→解約4%でストック』で積み上げる前提。強気すぎないか。\n"
 "③同意書の実務（福井さんの本領）：徳洲会(高根台)は出るが遅いと聞きました。取得・更新で立ち上げを早めるコツ、取りやすい診療科・先の見極めは？\n"
 "④労務・税務：施術者を雇用/業務委託どちらで設計すべきか（社保・税の観点）。\n\n"
 "お手すきの際に15分ほど、ざっくり所感をいただけると助かります。資料のどこからでも、気になる点を潰していただければ。\nよろしくお願いします。")

RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
YEL={"red":1.0,"green":0.97,"blue":0.80}; WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}

def creds():
    with open(TOKEN,"rb") as f: c=pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN,"wb") as f: pickle.dump(c,f)
    return c

def build_tab(svc, sid, gid, title_head, rows, widths):
    N=len(widths)
    values=[[title_head]+[""]*(N-1)]; meta=[]
    for rr in rows:
        i=len(values); k=rr[0]
        if k=="SEC": values.append([rr[1]]+[""]*(N-1)); meta.append((i,"SEC"))
        elif k=="KV": values.append([rr[1],rr[2]]+[""]*(N-2)); meta.append((i,"KV"))
        elif k=="THDR": values.append(rr[1]+[""]*(N-len(rr[1]))); meta.append((i,"THDR"))
        elif k=="ROW": values.append(rr[1]+[""]*(N-len(rr[1]))); meta.append((i,"ROW",rr[1][0]))
        elif k=="ROW3": values.append([rr[1][0],rr[1][1],rr[1][2]]+[""]*(N-3)); meta.append((i,"ROW3"))
        elif k=="NOTE": values.append([rr[1]]+[""]*(N-1)); meta.append((i,"NOTE"))
    svc.spreadsheets().values().update(spreadsheetId=sid,range=f"'{title_head_tab(svc,sid,gid)}'!A1",valueInputOption="USER_ENTERED",body={"values":values}).execute()
    return values,meta,N

def title_head_tab(svc,sid,gid):
    info=svc.spreadsheets().get(spreadsheetId=sid,fields="sheets(properties(sheetId,title))").execute()
    for s in info["sheets"]:
        if s["properties"]["sheetId"]==gid: return s["properties"]["title"]
    return "Sheet1"

def fmt(svc,sid,gid,values,meta,N,widths):
    reqs=[]
    for c,w in enumerate(widths):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":c,"endIndex":c+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":len(values)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":12}}},"fields":"userEnteredFormat"}})
    reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":36},"fields":"pixelSize"}})
    for mm in meta:
        i=mm[0];kind=mm[1]
        if kind=="SEC":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":11}}},"fields":"userEnteredFormat"}})
        elif kind=="KV":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":CARD,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
        elif kind=="THDR":
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat"}})
        elif kind=="ROW":
            lab=mm[2]
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
            if "営業利益" in lab or "返済後CF" in lab:
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":YEL,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
                reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
        elif kind=="ROW3":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":2,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
        elif kind=="NOTE":
            reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"wrapStrategy":"WRAP","verticalAlignment":"TOP"}},"fields":"userEnteredFormat"}})
            reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":80},"fields":"pixelSize"}})
    svc.spreadsheets().batchUpdate(spreadsheetId=sid,body={"requests":reqs}).execute()

def main():
    svc=build("sheets","v4",credentials=creds(),cache_discovery=False)
    ss=svc.spreadsheets().create(body={"properties":{"title":TITLE},"sheets":[
        {"properties":{"title":"1_診療圏調査","gridProperties":{"frozenRowCount":1}}},
        {"properties":{"title":"2_事業計画書_福井式","gridProperties":{"frozenRowCount":1}}},
        {"properties":{"title":"3_福井さんへの相談文","gridProperties":{"frozenRowCount":0}}},
    ]},fields="spreadsheetId,sheets.properties").execute()
    sid=ss["spreadsheetId"]
    g1=ss["sheets"][0]["properties"]["sheetId"]; g2=ss["sheets"][1]["properties"]["sheetId"]; g3=ss["sheets"][2]["properties"]["sheetId"]
    w1=[230,760]; w2=[150,140,140,140,140,150]
    v1,m1,N1=build_tab(svc,sid,g1,"診療圏調査（訪問マッサージ版）｜オーロラ高根店　2026-06-05",DIAG,w1); fmt(svc,sid,g1,v1,m1,N1,w1)
    v2,m2,N2=build_tab(svc,sid,g2,"事業計画書（福井式）｜オーロラ高根店　2026-06-05 ※福井さんレビュー用",BIZ,w2); fmt(svc,sid,g2,v2,m2,N2,w2)
    # Tab3 相談文
    svc.spreadsheets().values().update(spreadsheetId=sid,range="'3_福井さんへの相談文'!A1",valueInputOption="USER_ENTERED",body={"values":[["福井さんへの相談文（コピーして送付）"],[SOUDAN_TEXT]]}).execute()
    svc.spreadsheets().batchUpdate(spreadsheetId=sid,body={"requests":[
        {"updateDimensionProperties":{"range":{"sheetId":g3,"dimension":"COLUMNS","startIndex":0,"endIndex":1},"properties":{"pixelSize":900},"fields":"pixelSize"}},
        {"repeatCell":{"range":{"sheetId":g3,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":12}}},"fields":"userEnteredFormat"}},
        {"repeatCell":{"range":{"sheetId":g3,"startRowIndex":1,"endRowIndex":2},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"TOP","textFormat":{"fontSize":11,"foregroundColor":INK}}},"fields":"userEnteredFormat"}},
        {"updateDimensionProperties":{"range":{"sheetId":g3,"dimension":"ROWS","startIndex":1,"endIndex":2},"properties":{"pixelSize":520},"fields":"pixelSize"}},
    ]}).execute()
    print("DONE sid=",sid)
    print(f"URL: https://docs.google.com/spreadsheets/d/{sid}/edit")

if __name__=="__main__":
    main()
