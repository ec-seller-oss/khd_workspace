# -*- coding: utf-8 -*-
"""
福井式に倣い【1項目=1タブ】でPDF資料集になるよう、診療圏調査/事業計画書を多タブ化。
既存の独立スプシ A=15G26X.../ B=1ruE... を多タブで作り直す（URL維持）。2026-06-05。
"""
import pickle
from googleapiclient.discovery import build as gbuild
from google.auth.transport.requests import Request
TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID_A="15G26X_62MQZD3AtYXZLGK_Jp78yKj9VBsoR2tAV5N5A"
SID_B="1ruE_rE8i_OKpO0JMV-82AJDMqLPf8S-r4HxJIEgyL8U"

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
    b0=LOAN-(LOAN/(LY*12))*((y-1)*12);intr=max(b0,0)*LR
    pre=op-intr;tax=max(pre,0)*TAX;at=pre-tax;rep=min(LOAN/LY,max(b0,0))
    return dict(ue=round(ms[-1]["u"]),t=ms[-1]["t"],s=round(s),op=round(op),intr=round(intr),pre=round(pre),tax=round(tax),at=round(at),rep=round(rep),cf=round(at-rep),bal=round(max(b0-LOAN/LY,0)))
Y=[yp(i) for i in range(1,6)]
def yn(x):return f"{x:,}"
def sg(x):return ("+" if x>=0 else "")+yn(x)
r=LR/12;n=LY*12;pmt=LOAN*r*(1+r)**n/((1+r)**n-1)
CHO=[("高根台１丁目","100",2542,669),("高根台２丁目","100",1649,491),("高根台３丁目","100",3268,579),
("高根台４丁目","100",1230,270),("高根台５丁目","100",1269,356),("高根台６丁目","100",2409,408),
("高根台７丁目","100",1741,395),("松が丘１丁目","100",2527,695),("松が丘２丁目","100",1869,488),
("松が丘３丁目","100",3109,848),("松が丘４丁目","100",2688,584),("松が丘５丁目","100",2335,636),
("高根町","80",2278,197),("西習志野１丁目（補助）","50",2623,550),("西習志野２丁目（補助）","50",2116,402),
("西習志野３丁目（補助）","50",2709,467),("西習志野４丁目（補助）","50",1425,256)]
core=[c for c in CHO if "西習志野" not in c[0]]
cpop=sum(c[2] for c in core);c75=sum(c[3] for c in core);apop=sum(c[2] for c in CHO);a75=sum(c[3] for c in CHO)
AB={'0-14':3254,'15-64':15811,'65-74':3233,'75-84':4165,'85+':2451,'計':28914}
e65=AB['65-74']+AB['75-84']+AB['85+'];e75=AB['75-84']+AB['85+']
RED={"red":0.667,"green":0.180,"blue":0.149};REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882};REDBG={"red":0.957,"green":0.894,"blue":0.886}
YEL={"red":1.0,"green":0.97,"blue":0.80};WHT={"red":1,"green":1,"blue":1};INK={"red":0.1,"green":0.1,"blue":0.1}

# ─ 各タブ＝(タブ名, 見出し, rows, widths) ─
def diag_tabs():
    w2=[230,720];w6=[40,210,120,170,90,170];wp=[230,90,110,110,100,120];w3=[210,150,360]
    return [
    ("1_調査概要","① 調査概要｜オーロラ高根店 訪問マッサージ",[
      ("KV","事業・院名","オーロラ高根店（訪問鍼灸マッサージ・FC加盟＝(株)フライハイト）"),
      ("KV","所在地（候補）","千葉県船橋市高根台2丁目 近辺"),
      ("KV","最寄駅","新京成線 高根公団・高根木戸（チャリ圏 半径約2〜3km）"),
      ("KV","サービス内容","訪問マッサージ・はり灸（医療保険・医師同意書・消費税非課税）"),
      ("KV","開業区分／予定","FC加盟・新規開業／2026年8〜9月"),
      ("KV","管轄保健所","船橋市保健所"),
      ("KV","調査日／担当","2026-06-05／菊池（テナントアシスト・ウイン）"),
      ("KV","本資料の構成","②競合 ③診療圏人口 ④年齢別構成 ⑤潜在利用者 ⑥出典 の各タブで構成。"),
    ],w2),
    ("2_競合施設一覧","② 競合施設一覧（実在確認＋一次情報）",[
      ("TH",["No","事業者・院名","種別","所在地","エリア","同意書・備考"]),
      ("R",["1","千葉徳洲会病院","病院(同意書元)","高根台2-11-1","圏中心","出るが遅い＝律速・近接が強み"]),
      ("R",["2","レイス治療院船橋(フレアス系)","訪問マ鍼灸FC","西習志野1-13-13","圏西","習志野医師会で同意書ハンデ"]),
      ("R",["3","リボン","訪問マ","船橋駅前","圏外南西","駅前商圏"]),
      ("R",["4","ケイロウ","訪問マFC","西船橋駅","圏外西","団地少で手薄"]),
      ("R",["5","アシスト","訪問マ","江戸川区(船橋は訪問範囲)","圏外","高根に物理拠点なし＝追い風"]),
      ("R",["6","京葉/リカバリー/匠 ほか","訪問マ鍼灸","船橋広域/坪井(北部)","広域","分散型"]),
      ("NOTE","主要駅は競合が押さえ済だが、訪問は『在宅密度×同意書が出る医師会』で決まる。高根に訪問の物理拠点は薄く、船橋市医師会側＝差別化余地。"),
    ],w6),
    ("3_診療圏人口","③ 診療圏 地区別人口（船橋市 住民基本台帳 令和8年4月・実データ）",[
      ("TH",["町丁名","対象範囲%","人口(計)","75歳以上","75+比率",""]),
    ]+[("R",[c[0],c[1],yn(c[2]),yn(c[3]),f"{c[3]/c[2]*100:.1f}%",""]) for c in CHO]+[
      ("R",["核診療圏 小計(船橋市医師会側)","—",yn(cpop),yn(c75),f"{c75/cpop*100:.1f}%",""]),
      ("R",["圏全体 合計","—",yn(apop),yn(a75),f"{a75/apop*100:.1f}%",""]),
    ],wp),
    ("4_年齢別構成","④ 年齢別人口構成（核診療圏・実データ）",[
      ("TH",["年齢区分","人口","構成比"]),
      ("R",["0〜14歳",yn(AB['0-14']),f"{AB['0-14']/AB['計']*100:.1f}%"]),
      ("R",["15〜64歳",yn(AB['15-64']),f"{AB['15-64']/AB['計']*100:.1f}%"]),
      ("R",["65〜74歳",yn(AB['65-74']),f"{AB['65-74']/AB['計']*100:.1f}%"]),
      ("R",["75〜84歳",yn(AB['75-84']),f"{AB['75-84']/AB['計']*100:.1f}%"]),
      ("R",["85歳以上",yn(AB['85+']),f"{AB['85+']/AB['計']*100:.1f}%"]),
      ("R",["計",yn(AB['計']),"100%"]),
      ("R",["65歳以上(高齢化率)",yn(e65),f"{e65/AB['計']*100:.1f}%"]),
      ("R",["75歳以上(後期高齢)",yn(e75),f"{e75/AB['計']*100:.1f}%"]),
      ("NOTE","高齢化率34.1%(全国≒29%)・75歳以上22.9%(全国≒15%)＝在宅マッサージ需要が構造的に厚い。"),
    ],w3),
    ("5_潜在利用者","⑤ 潜在利用者数 算出",[
      ("KV","算出式","潜在利用者 ＝ 75歳以上人口 × 訪問マッサージ利用率(1.0〜1.5%)"),
      ("KV","核診療圏",f"75歳以上 {c75:,}人 × 1.0〜1.5% ＝ 潜在 約{round(c75*0.01)}〜{round(c75*0.015)}人"),
      ("KV","圏全体",f"75歳以上 {a75:,}人 × 1.0〜1.5% ＝ 潜在 約{round(a75*0.01)}〜{round(a75*0.015)}人"),
      ("KV","判定","1店BEPは利用者16名。核診療圏だけで潜在66〜99人＝1店は十分成立。"),
    ],w2),
    ("6_出典","⑥ 出典・ご注意",[
      ("KV","人口データ","船橋市『町丁別・年齢別人口（令和8年4月）』2026_R08_04nennreibetu.xlsx（Drive保存・添付）"),
      ("KV","施設","千葉徳洲会病院=高根台2-11-1／各競合の公式サイト／石原氏ヒアリング(6/2)"),
      ("NOTE","西習志野は習志野市医師会側で同意書が出にくいため対象範囲50%・補助扱い。利用率は地域差あり概算。本資料は第三者的検討であり将来を保証しない。"),
    ],w2),
    ]

def biz_tabs():
    w2=[230,720];w6=[180,120,120,120,120,120];wp=[200,120,120,120,120,130];w3=[210,150,360]
    return [
    ("1_基本情報","① 基本情報",[
      ("KV","事業名","オーロラ高根店（訪問鍼灸マッサージ・FC加盟＝(株)フライハイト）"),
      ("KV","所在地（候補）","船橋市高根台2丁目 近辺（新京成 高根公団/高根木戸）"),
      ("KV","経営者","菊池 研太（KHD／テナントアシスト・ウイン）"),
      ("KV","開業予定／形態","2026年8〜9月／FC加盟・新規開業（医療保険・消費税非課税）"),
      ("KV","本資料の構成","②人員 ③利用者想定 ④収入額算定 ⑤概算事業費 ⑥資金調達 ⑦年度別損益 ⑧返済明細 ⑨抜け漏れ の各タブ。"),
    ],w2),
    ("2_人員構成","② 人員構成（年度末）",[
      ("TH",["職種","勤務形態","1年目","3年目","5年目","月額(目安)"]),
      ("R",["施術者(あマ指/鍼灸)","常勤/委託",str(Y[0]['t']),str(Y[2]['t']),str(Y[4]['t']),"売上の30%"]),
      ("R",["営業","常勤","1","1","2","25万"]),
      ("R",["医療事務(レセプト)","本部代行","—","—","—","事務手数料3%に内包"]),
      ("NOTE","利用者の伸びは施術者の採用が律速（1人あたり利用者22名で増員）。"),
    ],w6),
    ("3_利用者数想定","③ 利用者数想定・単価",[
      ("TH",["区分","1年目末","2年目末","3年目末","4年目末","5年目末"]),
      ("R",["利用者数"]+[str(y['ue']) for y in Y]),
      ("KV","単価","利用者単価 月3.8万円（保険9割・消費税非課税）"),
      ("KV","ランプ前提","営業ファネル=訪問300/月×接触50%×問合せ5→13%×CV55%。解約4%/月(継続25ヶ月)。"),
    ],wp),
    ("4_収入額算定表","④ 収入額算定表（年次＝利用者×単価×12）",[
      ("TH",["項目","1年目","2年目","3年目","4年目","5年目"]),
      ("R",["利用者(年末)"]+[str(y['ue']) for y in Y]),
      ("R",["年間売上"]+[yn(y['s']) for y in Y]),
    ],w6),
    ("5_概算事業費","⑤ 概算事業費 明細（本部資料P27ベース）",[
      ("KV","設備：加盟金","150万円"),
      ("KV","設備：開業初期費用","約370〜420万円（研修・備品・車両・サイネージ・初期システム等／内訳は本部確認）"),
      ("KV","運転資金","約150万円"),
      ("KV","所要資金 合計","約670〜720万円（≒740万で設計）"),
    ],w2),
    ("6_資金調達","⑥ 資金調達",[
      ("TH",["区分","金額","条件","",""," "]),
      ("R",["自己資金","300万円","—","","",""]),
      ("R",["公庫借入","440万円",f"{LY}年・年{LR*100:.1f}%・元利均等・据置無","","",""]),
      ("R",["合計","740万円",f"月返済 約{round(pmt):,}円","","",""]),
      ("NOTE","自己資金/借入のバランス・期間・利率は要相談（福井さん論点①）。"),
    ],w6),
    ("7_年度別損益","⑦ 年度別損益計算（円）",[
      ("TH",["項目","1年目","2年目","3年目","4年目","5年目"]),
      ("R",["①売上"]+[yn(y['s']) for y in Y]),
      ("R",["②施術者人件費(売上30%)"]+[yn(round(y['s']*THC)) for y in Y]),
      ("R",["③ロイヤリティ13.2%"]+[yn(round(y['s']*ROY)) for y in Y]),
      ("R",["④事務手数料3%"]+[yn(round(y['s']*OFF)) for y in Y]),
      ("R",["⑤営業/家賃/車両/他"]+[yn(round(y['s']-y['op']-y['s']*(THC+ROY+OFF))) for y in Y]),
      ("R",["■営業利益"]+[sg(y['op']) for y in Y]),
      ("R",["⑥支払利息"]+[yn(y['intr']) for y in Y]),
      ("R",["税引前損益"]+[sg(y['pre']) for y in Y]),
      ("R",["法人税等(25%)"]+[yn(y['tax']) for y in Y]),
      ("R",["税引後利益"]+[sg(y['at']) for y in Y]),
      ("R",["借入元金返済"]+[yn(y['rep']) for y in Y]),
      ("R",["■返済後CF"]+[sg(y['cf']) for y in Y]),
    ],w6),
    ("8_借入金返済明細","⑧ 借入金返済明細（年次）",[
      ("TH",["項目","1年目","2年目","3年目","4年目","5年目"]),
      ("R",["元金返済"]+[yn(y['rep']) for y in Y]),
      ("R",["支払利息"]+[yn(y['intr']) for y in Y]),
      ("R",["借入金残高"]+[yn(y['bal']) for y in Y]),
      ("NOTE",f"借入440万・{LY}年・年{LR*100:.1f}%・元利均等で試算（条件は要相談）。"),
    ],w6),
    ("9_抜け漏れ","⑨ 福井さんに見てほしい抜け漏れ・論点",[
      ("R3",["①","資金・返済","自己資金/公庫バランス・期間・利率の最適は？運転150万で足りるか。"]),
      ("R3",["②","5年損益の現実性","ケアマネ営業→紹介→同意書→解約4%でストック。積み上げ前提は甘くないか。"]),
      ("R3",["③","同意書の実務(本領)","徳洲会は出るが遅い。早めるコツ／取りやすい診療科・先は？"]),
      ("R3",["④","労務・税務","雇用/業務委託どちらで設計すべきか（社保・労災・源泉・課税区分）。"]),
      ("R3",["⑤","手続き等","施術所開設届・出張施術届(保健所)、減価償却/リース、FC契約の不利条項。"]),
    ],w3),
    ]

def creds():
    with open(TOKEN,"rb") as f:c=pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN,"wb") as f:pickle.dump(c,f)
    return c

def fill(svc,sid,gid,tt,head,rows,widths):
    N=len(widths)
    vals=[[head]+[""]*(N-1)];meta=[]
    for rr in rows:
        i=len(vals);k=rr[0]
        if k=="KV":vals.append([rr[1],rr[2]]+[""]*(N-2));meta.append((i,"KV"))
        elif k=="TH":vals.append(rr[1]+[""]*(N-len(rr[1])));meta.append((i,"TH"))
        elif k=="R":vals.append(rr[1]+[""]*(N-len(rr[1])));meta.append((i,"R",rr[1][0]))
        elif k=="R3":vals.append([rr[1][0],rr[1][1],rr[1][2]]+[""]*(N-3));meta.append((i,"R3"))
        elif k=="NOTE":vals.append([rr[1]]+[""]*(N-1));meta.append((i,"NOTE"))
    svc.spreadsheets().values().update(spreadsheetId=sid,range=f"'{tt}'!A1",valueInputOption="USER_ENTERED",body={"values":vals}).execute()
    rq=[]
    for c,w in enumerate(widths):
        rq.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":c,"endIndex":c+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":len(vals)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    rq.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
    rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":13},"verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat"}})
    rq.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":42},"fields":"pixelSize"}})
    for mm in meta:
        i=mm[0];k=mm[1]
        if k=="KV":
            rq+=[{"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"mergeType":"MERGE_ALL"}},
                 {"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":CARD,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat"}}]
        elif k=="TH":
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat"}})
        elif k=="R":
            lab=mm[2]
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":N},"cell":{"userEnteredFormat":{"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat.horizontalAlignment"}})
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK},"horizontalAlignment":"LEFT"}},"fields":"userEnteredFormat(textFormat,horizontalAlignment)"}})
            if "■" in lab or "小計" in lab or "合計" in lab or "高齢化率" in lab or "後期" in lab:
                rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":YEL,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
        elif k=="R3":
            rq+=[{"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":2,"endColumnIndex":N},"mergeType":"MERGE_ALL"}},
                 {"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD}}},"fields":"userEnteredFormat.textFormat"}}]
        elif k=="NOTE":
            rq+=[{"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}},
                 {"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":9,"foregroundColor":REDD},"wrapStrategy":"WRAP","verticalAlignment":"TOP"}},"fields":"userEnteredFormat"}},
                 {"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":64},"fields":"pixelSize"}}]
    th=[m[0] for m in meta if m[1]=="TH"]
    if th:
        end=len(vals)
        for t in th:
            rq.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":t,"endRowIndex":end,"startColumnIndex":0,"endColumnIndex":N},"innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=sid,body={"requests":rq}).execute()

def rebuild(svc,sid,tabs):
    info=svc.spreadsheets().get(spreadsheetId=sid,fields="sheets(properties(sheetId,title))").execute()
    old=[s["properties"]["sheetId"] for s in info["sheets"]]
    # 新タブ追加
    add=[{"addSheet":{"properties":{"title":t[0],"index":i,"gridProperties":{"frozenRowCount":1}}}} for i,t in enumerate(tabs)]
    res=svc.spreadsheets().batchUpdate(spreadsheetId=sid,body={"requests":add}).execute()
    gids={r["addSheet"]["properties"]["title"]:r["addSheet"]["properties"]["sheetId"] for r in res["replies"]}
    for (name,head,rows,widths) in tabs:
        fill(svc,sid,gids[name],name,head,rows,widths)
    # 旧タブ削除
    svc.spreadsheets().batchUpdate(spreadsheetId=sid,body={"requests":[{"deleteSheet":{"sheetId":g}} for g in old]}).execute()

def main():
    svc=gbuild("sheets","v4",credentials=creds(),cache_discovery=False)
    rebuild(svc,SID_A,diag_tabs())
    rebuild(svc,SID_B,biz_tabs())
    print("DONE")
    print(f"A: https://docs.google.com/spreadsheets/d/{SID_A}/edit")
    print(f"B: https://docs.google.com/spreadsheets/d/{SID_B}/edit")

if __name__=="__main__":
    main()
