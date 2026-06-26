# -*- coding: utf-8 -*-
"""
福井さんの2フォーマット（診療圏調査／事業計画書）を忠実に再現し、オーロラ高根の実データで埋める。
独立スプシ 1DuyP0d7 の Tab1/Tab2 を作り直す。町丁別人口は船橋市R8.4実データ。2026-06-05。
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID="1DuyP0d7-JHYz73wn6UKgOooI8q8dWCOV5uI-gy6fgNo"

# ── 事業計画 前提・シミュ ──
VD=30;SD=10;CON=0.50;CV=0.55;CH=0.04;PRICE=38000;CAP=22
ROY=0.132;OFF=0.03;THC=0.30;SC=250000;RENT=80000;OTH=60000;PK=12000;VH=45000
LOAN=4400000;LR=0.018;LY=10;TAX=0.25
def inq(m):return 0.05 if m<=3 else(0.09 if m<=9 else 0.13)
def nu(m):return VD*SD*CON*inq(m)*CV
def sim(n=60):
    o=[];u=0.0;t=1
    for m in range(1,n+1):
        if u>t*CAP*0.85:t+=1
        u=min(u*(1-CH)+nu(m),t*CAP);s=u*PRICE
        op=s-(s*(ROY+OFF+THC)+SC*(2 if t>=6 else 1)+RENT+OTH+PK*t+VH*t)
        o.append(dict(m=m,u=u,t=t,s=s,op=op))
    return o
S=sim(60)
def yp(y):
    ms=[x for x in S if (y-1)*12<x["m"]<=y*12]
    s=sum(x["s"] for x in ms);op=sum(x["op"] for x in ms)
    bal=LOAN-(LOAN/(LY*12))*((y-1)*12);intr=max(bal,0)*LR
    pre=op-intr;tax=max(pre,0)*TAX;at=pre-tax;rep=min(LOAN/LY,max(bal,0))
    return dict(ue=round(ms[-1]["u"]),t=ms[-1]["t"],s=round(s),op=round(op),intr=round(intr),pre=round(pre),tax=round(tax),at=round(at),rep=round(rep),cf=round(at-rep),bal=round(max(bal-LOAN/LY,0)))
Y=[yp(i) for i in range(1,6)]
def yn(x):return f"{x:,}"
def sg(x):return ("+" if x>=0 else "")+yn(x)

# ── 町丁別 実データ(船橋市R8.4) (町丁,範囲%,計,75+計,75+男,75+女) ──
CHO=[("高根台１丁目","100",2542,669,244,425),("高根台２丁目","100",1649,491,165,326),
("高根台３丁目","100",3268,579,227,352),("高根台４丁目","100",1230,270,98,172),
("高根台５丁目","100",1269,356,135,221),("高根台６丁目","100",2409,408,153,255),
("高根台７丁目","100",1741,395,151,244),("松が丘１丁目","100",2527,695,281,414),
("松が丘２丁目","100",1869,488,189,299),("松が丘３丁目","100",3109,848,335,513),
("松が丘４丁目","100",2688,584,229,355),("松が丘５丁目","100",2335,636,252,384),
("高根町","80",2278,197,77,120),
("西習志野１丁目（補助）","50",2623,550,210,340),("西習志野２丁目（補助）","50",2116,402,152,250),
("西習志野３丁目（補助）","50",2709,467,190,277),("西習志野４丁目（補助）","50",1425,256,98,158)]
core=[c for c in CHO if "西習志野" not in c[0]]
core_pop=sum(c[2] for c in core);core_75=sum(c[3] for c in core)
all_pop=sum(c[2] for c in CHO);all_75=sum(c[3] for c in CHO)

RED={"red":0.667,"green":0.180,"blue":0.149};REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882};REDBG={"red":0.957,"green":0.894,"blue":0.886}
YEL={"red":1.0,"green":0.97,"blue":0.80};WHT={"red":1,"green":1,"blue":1};INK={"red":0.1,"green":0.1,"blue":0.1}

DIAG=[
 ("SEC","Ⅰ. 調査概要"),
 ("KV2","事業・院名","オーロラ高根店（訪問鍼灸マッサージ・FC加盟＝(株)フライハイト）"),
 ("KV2","所在地（候補）","千葉県船橋市高根台2丁目 近辺"),
 ("KV2","最寄駅","新京成線 高根公団／高根木戸（チャリ圏 半径約2〜3km）"),
 ("KV2","サービス内容","訪問マッサージ・はり灸（医療保険・医師同意書ベース・消費税非課税）"),
 ("KV2","開業区分／予定","FC加盟・新規開業／2026年8〜9月"),
 ("KV2","管轄保健所","船橋市保健所"),
 ("KV2","調査日／担当","2026-06-05／菊池（テナントアシスト・ウイン）"),
 ("SEC","Ⅱ. 競合施設一覧（実在確認＋一次情報）"),
 ("TH",["No","事業者・院名","種別","所在地","エリア","同意書・備考"]),
 ("R",["1","千葉徳洲会病院","病院(同意書元)","高根台2-11-1","圏中心","出るが遅い＝律速。近接が強み"]),
 ("R",["2","レイス治療院船橋(フレアス系)","訪問マ・鍼灸FC","西習志野1-13-13","圏西","習志野医師会で同意書ハンデ＝伸び悩む"]),
 ("R",["3","リボン","訪問マ","船橋駅前","圏外南西","駅前商圏"]),
 ("R",["4","ケイロウ(KEiROW)","訪問マFC","西船橋駅","圏外西","団地少で手薄"]),
 ("R",["5","アシスト","訪問マ","江戸川区(船橋は訪問範囲)","圏外","高根に物理拠点なし＝追い風"]),
 ("R",["6","京葉/リカバリー/匠 ほか","訪問マ・鍼灸","船橋市広域/坪井(北部)","広域","分散型"]),
 ("SEC","Ⅲ. 診療圏 地区別人口（船橋市 住民基本台帳 令和8年4月）"),
 ("TH",["町丁名","対象範囲%","人口(計)","75歳以上(計)","75歳以上(男)","75歳以上(女)"]),
]+[("R",[c[0],c[1],yn(c[2]),yn(c[3]),yn(c[4]),yn(c[5])]) for c in CHO]+[
 ("R",["核診療圏 小計(船橋市医師会側)","—",yn(core_pop),yn(core_75),"—","—"]),
 ("R",["圏全体 合計(西習志野含む)","—",yn(all_pop),yn(all_75),"—","—"]),
 ("SEC","Ⅳ. 潜在利用者数 算出（訪問マッサージ）"),
 ("KV2","算出式","潜在利用者 ＝ 75歳以上人口 × 訪問マッサージ利用率（1.0〜1.5%）"),
 ("KV2","核診療圏(同意書出る)",f"75歳以上 {core_75:,}人 × 1.0〜1.5% ＝ 潜在 約{round(core_75*0.01)}〜{round(core_75*0.015)}人"),
 ("KV2","圏全体",f"75歳以上 {all_75:,}人 × 1.0〜1.5% ＝ 潜在 約{round(all_75*0.01)}〜{round(all_75*0.015)}人"),
 ("KV2","高齢化率(75+)",f"核診療圏 {core_75/core_pop*100:.1f}%（全国平均 約15%を大きく上回る＝在宅需要が厚い）"),
 ("KV2","判定","1店のBEPは利用者16名。核診療圏だけで潜在66〜99人＝1店は十分に成立。"),
 ("SEC","Ⅴ. ご注意"),
 ("NOTE","人口は船橋市『町丁別・年齢別人口（令和8年4月）』実データ。徳洲会=高根台2-11-1。競合は各社公式＋石原氏ヒアリング。西習志野は習志野市医師会側で同意書が出にくいため対象範囲50%・補助扱い。利用率は地域差があり概算。将来を保証するものではない。"),
]

BIZ=[
 ("SEC","Ⅰ. 基本情報"),
 ("KV2","事業名","オーロラ高根店（訪問鍼灸マッサージ・FC加盟＝(株)フライハイト）"),
 ("KV2","所在地（候補）","船橋市高根台2丁目 近辺（新京成 高根公団/高根木戸）"),
 ("KV2","経営者／連絡先","菊池 研太（KHD／テナントアシスト・ウイン）"),
 ("KV2","開業予定／形態","2026年8〜9月／FC加盟・新規開業（医療保険・消費税非課税）"),
 ("SEC","Ⅱ. 人員構成（年度末）"),
 ("TH",["職種","勤務形態","1年目","3年目","5年目","月額給与(目安)"]),
 ("R",["施術者(あマ指/鍼灸)","常勤/委託",str(Y[0]['t']),str(Y[2]['t']),str(Y[4]['t']),"売上の30%(歩合/固定)"]),
 ("R",["営業","常勤","1","1","2","25万"]),
 ("R",["医療事務(レセプト)","本部代行","—","—","—","事務手数料3%に内包"]),
 ("SEC","Ⅲ. 利用者数想定・単価"),
 ("TH",["区分","1年目末","2年目末","3年目末","5年目末","単価"]),
 ("R",["利用者数(年末)",str(Y[0]['ue']),str(Y[1]['ue']),str(Y[2]['ue']),str(Y[4]['ue']),"月3.8万円"]),
 ("KV2","ランプの前提","営業ファネル＝訪問300/月×接触50%×問合せ5→13%×CV55%。解約率4%/月(継続25ヶ月)。施術者1人22名で増員(採用が律速)。"),
 ("SEC","Ⅳ. 概算事業費・資金調達"),
 ("KV2","設備資金","加盟金・研修・備品・初期システム ＝ 約590万円"),
 ("KV2","運転資金","立上げ人件費・採用・予備 ＝ 約150万円"),
 ("KV2","所要資金 合計","740万円"),
 ("KV2","資金調達(案)",f"自己資金 300万 ／ 公庫借入 440万（{LY}年・年{LR*100:.1f}%・元利均等・据置無）"),
 ("SEC","Ⅴ. 年度別損益計算（円）"),
 ("TH",["項目","1年目","2年目","3年目","4年目","5年目"]),
 ("R",["①売上(利用者×単価)"]+[yn(y['s']) for y in Y]),
 ("R",["②施術者人件費(売上30%)"]+[yn(round(y['s']*THC)) for y in Y]),
 ("R",["③ロイヤリティ13.2%"]+[yn(round(y['s']*ROY)) for y in Y]),
 ("R",["④事務手数料3%"]+[yn(round(y['s']*OFF)) for y in Y]),
 ("R",["⑤営業・家賃・車両等"]+[yn(round(y['s']-y['op']-y['s']*(THC+ROY+OFF))) for y in Y]),
 ("R",["■営業利益(①-②-③-④-⑤)"]+[sg(y['op']) for y in Y]),
 ("R",["⑥支払利息"]+[yn(y['intr']) for y in Y]),
 ("R",["税引前損益"]+[sg(y['pre']) for y in Y]),
 ("R",["法人税等(25%)"]+[yn(y['tax']) for y in Y]),
 ("R",["税引後利益"]+[sg(y['at']) for y in Y]),
 ("R",["借入元金返済"]+[yn(y['rep']) for y in Y]),
 ("R",["■返済後CF"]+[sg(y['cf']) for y in Y]),
 ("R",["借入金残高"]+[yn(y['bal']) for y in Y]),
 ("SEC","Ⅵ. 福井さんに見てほしい抜け漏れ・論点"),
 ("R3",["①","資金・返済","自己資金/公庫のバランス・期間・利率の最適は？運転150万で足りるか。"]),
 ("R3",["②","5年損益の現実性","ケアマネ営業→紹介→同意書→解約4%でストック。この積み上げ前提は甘くないか。"]),
 ("R3",["③","同意書の実務(本領)","徳洲会は出るが遅い。早めるコツ／取りやすい診療科・先の見極めは？"]),
 ("R3",["④","労務・税務","雇用/業務委託どちらで設計すべきか（社保・労災・源泉・課税区分）。"]),
 ("R3",["⑤","開業手続き等","施術所開設届・出張施術届(保健所)、減価償却/リース、FC契約の不利条項。"]),
]

def creds():
    with open(TOKEN,"rb") as f:c=pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN,"wb") as f:pickle.dump(c,f)
    return c

def build(svc,gid,head,rows,widths):
    N=len(widths)
    vals=[[head]+[""]*(N-1)];meta=[]
    for rr in rows:
        i=len(vals);k=rr[0]
        if k=="SEC":vals.append([rr[1]]+[""]*(N-1));meta.append((i,"SEC"))
        elif k=="KV2":vals.append([rr[1],rr[2]]+[""]*(N-2));meta.append((i,"KV"))
        elif k=="TH":vals.append(rr[1]+[""]*(N-len(rr[1])));meta.append((i,"TH"))
        elif k=="R":vals.append(rr[1]+[""]*(N-len(rr[1])));meta.append((i,"R",rr[1][0]))
        elif k=="R3":vals.append([rr[1][0],rr[1][1],rr[1][2]]+[""]*(N-3));meta.append((i,"R3"))
        elif k=="NOTE":vals.append([rr[1]]+[""]*(N-1));meta.append((i,"NOTE"))
    info=svc.spreadsheets().get(spreadsheetId=SID,fields="sheets(properties(sheetId,title))").execute()
    tt=[s["properties"]["title"] for s in info["sheets"] if s["properties"]["sheetId"]==gid][0]
    svc.spreadsheets().values().update(spreadsheetId=SID,range=f"'{tt}'!A1",valueInputOption="USER_ENTERED",body={"values":vals}).execute()
    rq=[]
    for c,w in enumerate(widths):
        rq.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":c,"endIndex":c+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":len(vals)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE","textFormat":{"fontSize":9,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    rq.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
    rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":12}}},"fields":"userEnteredFormat"}})
    for mm in meta:
        i=mm[0];k=mm[1]
        if k=="SEC":
            rq.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":RED,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":10}}},"fields":"userEnteredFormat"}})
        elif k=="KV":
            rq.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":CARD,"textFormat":{"bold":True,"fontSize":9,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}})
        elif k=="TH":
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":9,"foregroundColor":REDD},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat"}})
        elif k=="R":
            lab=mm[2]
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":2,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":9,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
            if "■" in lab or "小計" in lab or "合計" in lab:
                rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":YEL,"textFormat":{"bold":True,"fontSize":9,"foregroundColor":REDD}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
                rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":2,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
        elif k=="R3":
            rq.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":2,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":9,"foregroundColor":REDD}}},"fields":"userEnteredFormat.textFormat"}})
        elif k=="NOTE":
            rq.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":9,"foregroundColor":REDD},"wrapStrategy":"WRAP","verticalAlignment":"TOP"}},"fields":"userEnteredFormat"}})
            rq.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":70},"fields":"pixelSize"}})
    # テーブル罫線（TH〜次のSEC手前）
    th_idx=[m[0] for m in meta if m[1]=="TH"]
    sec_idx=[m[0] for m in meta if m[1]=="SEC"]+[len(vals)]
    for th in th_idx:
        end=min(s for s in sec_idx if s>th)
        rq.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":th,"endRowIndex":end,"startColumnIndex":0,"endColumnIndex":N},"innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":rq}).execute()

def main():
    svc=__import__("googleapiclient.discovery",fromlist=["build"]).build("sheets","v4",credentials=creds(),cache_discovery=False)
    info=svc.spreadsheets().get(spreadsheetId=SID,fields="sheets(properties(sheetId,title))").execute()
    cur={s["properties"]["title"]:s["properties"]["sheetId"] for s in info["sheets"]}
    # 旧Tab1/2を削除して作り直し
    reqs=[]
    for t in ["1_診療圏調査","2_事業計画書_福井式"]:
        if t in cur: reqs.append({"deleteSheet":{"sheetId":cur[t]}})
    reqs.append({"addSheet":{"properties":{"title":"1_診療圏調査(実データ)","index":0,"gridProperties":{"frozenRowCount":1}}}})
    reqs.append({"addSheet":{"properties":{"title":"2_事業計画書_福井式","index":1,"gridProperties":{"frozenRowCount":1}}}})
    res=svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":reqs}).execute()
    new={r["addSheet"]["properties"]["title"]:r["addSheet"]["properties"]["sheetId"] for r in res["replies"] if "addSheet" in r}
    build(svc,new["1_診療圏調査(実データ)"],"診療圏調査｜オーロラ高根店（船橋市実データ）2026-06-05",DIAG,[200,90,90,110,90,200])
    build(svc,new["2_事業計画書_福井式"],"事業計画書（福井式）｜オーロラ高根店 2026-06-05",BIZ,[180,120,120,120,120,130])
    print("DONE")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit")
    for i,y in enumerate(Y,1):print(f"Y{i} 利用者{y['ue']} 売上{y['s']:,} 営業益{y['op']:,}")

if __name__=="__main__":
    main()
