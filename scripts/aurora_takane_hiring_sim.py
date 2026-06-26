# -*- coding: utf-8 -*-
"""
開業準備ワークブック(1QH8F2…)に「03_採用シミュレーター」タブを追加。
利用者＝MIN(送客需要, 施術者キャパ22名/人)。本部紹介での採用ペース(=ブラックボックス)を
入力1つで変えると、利用者天井・売上・人件費構造・営業利益・回収月が連動する。
上/中/下シナリオの静的比較も冒頭に提示（採用可否で何が決まるかを一目で）。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID="1QH8F2So98kGUUHeNhbATkvnKD0486-kiZ8b7ziOuoYM"
TAB="03_採用シミュレーター"

RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
INBG={"red":0.999,"green":0.969,"blue":0.792}; TIERA={"red":0.886,"green":0.949,"blue":0.905}
SEC={"red":0.357,"green":0.357,"blue":0.357}; BLUEBG={"red":0.812,"green":0.886,"blue":0.953}
HI={"red":1.0,"green":0.898,"blue":0.6}

# ── 前提（静的シナリオ計算用＝ライブ初期値と一致）──
P=dict(unit=38000, cap=22, churn=0.04, roy=0.132, mgmt=100000, fee=0.03, fixoth=150000,
       init=5900000, emp_pay=320000, com_rate=0.5, pt_pay=120000, sokyaku=5)

def sim(emp0,com0,pt0,demp,dcom,dpt,months=18):
    cust=0; cum=-P["init"]; out=[]
    for m in range(1,months+1):
        emp=emp0+demp*(m-1); com=com0+dcom*(m-1); pt=pt0+dpt*(m-1); tot=emp+com+pt
        capn=tot*P["cap"]
        demand = P["sokyaku"] if m==1 else cust*(1-P["churn"])+P["sokyaku"]
        cust=min(demand,capn)
        rev=cust*P["unit"]
        lc_emp=emp*P["emp_pay"]; lc_com=(rev*(com/tot)*P["com_rate"]) if tot>0 else 0; lc_pt=pt*P["pt_pay"]
        roy=rev*P["roy"]+P["mgmt"]+rev*P["fee"]; op=rev-lc_emp-lc_com-lc_pt-roy-P["fixoth"]
        cum+=op; out.append(dict(m=m,tot=tot,cust=cust,rev=rev,op=op,cum=cum))
    return out

def payback(rows):
    for r in rows:
        if r["cum"]>0: return r["m"]
    return None

SCEN = {  # (emp0,com0,pt0, dEmp,dCom,dPt)
 "上（本部協力◎・採れる）":(0,1,1, 0,0.3,0.2),
 "中（標準）":(0,1,0, 0,0.2,0.1),
 "下（本部紹介乏しく増えない）":(0,1,0, 0,0,0),
}
def yen(v):
    s="▲" if v<0 else ""; v=abs(round(v))
    if v>=1_000_000: return f"{s}¥{v/1_000_000:.2f}M"
    return f"{s}¥{v:,}"

def creds():
    with open(TOKEN,"rb") as f:c=pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN,"wb") as f:pickle.dump(c,f)
    return c
def mr(g,r,c0,nc):return {"mergeCells":{"range":{"sheetId":g,"startRowIndex":r,"endRowIndex":r+1,"startColumnIndex":c0,"endColumnIndex":nc},"mergeType":"MERGE_ALL"}}
def fr(g,r,bg,fg,b,sz,wrap=True,h="LEFT"):return {"repeatCell":{"range":{"sheetId":g,"startRowIndex":r,"endRowIndex":r+1},"cell":{"userEnteredFormat":{"backgroundColor":bg,"textFormat":{"foregroundColor":fg,"bold":b,"fontSize":sz},"wrapStrategy":"WRAP" if wrap else "OVERFLOW_CELL","verticalAlignment":"MIDDLE","horizontalAlignment":h}},"fields":"userEnteredFormat"}}
def rh(g,r,px,r2=None):return {"updateDimensionProperties":{"range":{"sheetId":g,"dimension":"ROWS","startIndex":r,"endIndex":(r2 or r)+1},"properties":{"pixelSize":px},"fields":"pixelSize"}}
def cw(g,ws):return [{"updateDimensionProperties":{"range":{"sheetId":g,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":w},"fields":"pixelSize"}} for i,w in enumerate(ws)]

def main():
    svc=build("sheets","v4",credentials=creds(),cache_discovery=False)
    info=svc.spreadsheets().get(spreadsheetId=SID,fields="sheets.properties").execute()
    ex={s["properties"]["title"]:s["properties"]["sheetId"] for s in info["sheets"]}
    if TAB in ex:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":[{"deleteSheet":{"sheetId":ex[TAB]}}]}).execute()
    NC=16
    add=svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":[{"addSheet":{"properties":{"title":TAB,"gridProperties":{"rowCount":75,"columnCount":NC,"frozenRowCount":2}}}}]}).execute()
    gid=add["replies"][0]["addSheet"]["properties"]["sheetId"]

    # 静的シナリオ計算
    scen_rows=[]
    for name,(e0,c0,p0,de,dc,dp) in SCEN.items():
        r=sim(e0,c0,p0,de,dc,dp,18); m12=r[11]; m18=r[17]; pb=payback(r)
        scen_rows.append([name,
            f"初期 委{c0}/パ{p0}/社{e0}・月次 委+{dc}/パ+{dp}/社+{de}",
            f"{m12['tot']:.1f}名", f"{m12['cust']:.0f}名", yen(m12['rev']), yen(m12['op']),
            yen(m18['cum']), (f"{pb}ヶ月" if pb else "18M内 未回収"),
        ])

    R=[]; A=lambda r:R.append(r+[""]*(NC-len(r)))
    A(["🧮 採用シミュレーター｜利用者＝MIN(送客需要, 施術者キャパ22名/人)。採用(本部紹介)が“成長の天井”と人件費を決める"])
    A(["結論の出し方＝下の入力②『月次採用(本部紹介)』を狩野/本部の実数に変える→未来が連動。社員は固定費リスク／業務委託は変動(歩合)／パートは半固定。"])
    A([""])
    A(["■ 採用シナリオ比較（上/中/下＝本部紹介ブラックボックスの見込み別）★ここが結論"]); sec1=len(R)-1
    A(["シナリオ","採用前提(初期/月次)","12M 施術者","12M 利用者","12M 売上/月","12M 営業利益/月","18M 累積CF(初期投資590万込)","回収月"]); h1=len(R)-1
    s1=len(R)
    for sr in scen_rows: A(sr)
    e1=len(R)
    A([""])
    A(["■ 結論：採用可否で“何が”決まるのか"]); sec2=len(R)-1
    A(["① 初期(〜数ヶ月)は『送客＝営業』が律速。利用者がまだ少なく、施術者キャパに余裕があるため。"])
    A(["② 利用者が施術者キャパ(22名/人)に達した瞬間から『採用』が律速に切替。採れないと売上が天井で固定＝事業が伸びず、加盟金590万を回収できない(下シナリオ)。"])
    A(["③ ゆえに“本部紹介で船橋・高根に何名回るか”が事業の成否を直接決める。これを狩野さん/本部に必ず数字で確認し、下の入力②に入れる。"])
    A(["④ 初期配置の推奨＝固定費リスクを抑え『パート(安定供給)＋業務委託(変動/歩合)』中心。社員(固定給)は利用者が積み上がってから。板橋も『パート◎/歩合は初期グリップ難』。"])
    A([""])
    A(["■ 入力①：基本前提（黄=自由に変更）"]); sec3=len(R)-1
    base=len(R)+1  # 1-idx of first 入力① row
    A(["単価(円/月/人)",P["unit"]]);           r_unit=len(R)
    A(["担当上限(名/施術者)",P["cap"]]);        r_cap=len(R)
    A(["解約率/月",P["churn"]]);                 r_chn=len(R)
    A(["ロイヤリティ率(売上%)",P["roy"]]);       r_roy=len(R)
    A(["運営委託費(月固定)",P["mgmt"]]);         r_mgmt=len(R)
    A(["事務手数料率(売上%)",P["fee"]]);         r_fee=len(R)
    A(["その他固定費/月(家賃等)",P["fixoth"]]);  r_fix=len(R)
    A(["初期投資(加盟金+研修+運転)",P["init"]]); r_init=len(R)
    A([""])
    A(["■ 入力②：採用・人員配置　★本部紹介＝ここに実数を入れると未来が動く"]); sec4=len(R)-1
    A(["開業時 社員",0]);            r_e0=len(R)
    A(["開業時 業務委託",1]);        r_c0=len(R)
    A(["開業時 パート(常勤換算)",0]);r_p0=len(R)
    A(["★月次採用 社員/月",0]);      r_de=len(R)
    A(["★月次採用 業務委託/月(本部紹介＋自力)",0.2]); r_dc=len(R)
    A(["★月次採用 パート/月",0.1]);  r_dp=len(R)
    A([""])
    A(["■ 入力③：コスト構造（雇用形態別＝人件費ブラックボックスの正体）"]); sec5=len(R)-1
    A(["社員 月給(社保込/人)",P["emp_pay"]]);       r_ep=len(R)
    A(["業務委託 歩合率(担当売上%)",P["com_rate"]]);r_cr=len(R)
    A(["パート 月コスト(/人)",P["pt_pay"]]);         r_pp=len(R)
    A([""])
    A(["■ 入力④：送客（営業ファネルの結果＝需要）"]); sec6=len(R)-1
    A(["月次 新規利用者(送客)/月",P["sokyaku"]]); r_sk=len(R)
    A([""])
    A(["■ 月次シミュレーション（18ヶ月・上記入力で自動計算）"]); sec7=len(R)-1
    A(["月","社員","業委","パート","施術者計","キャパ(名)","送客需要","利用者(名)","売上","人件費 社員","人件費 業委","人件費 パート","ロイヤリティ等","その他固定","営業利益","累積CF"]); eh=len(R)-1
    es=len(R)
    def B(r):return f"$B${r}"
    for m in range(1,19):
        r=len(R)+1  # 1-idx of this month row
        if m==1:
            A([1, f"={B(r_e0)}", f"={B(r_c0)}", f"={B(r_p0)}",
               f"=B{r}+C{r}+D{r}", f"=E{r}*{B(r_cap)}", f"={B(r_sk)}", f"=MIN(G{r},F{r})",
               f"=H{r}*{B(r_unit)}", f"=B{r}*{B(r_ep)}", f"=IF(E{r}=0,0,I{r}*(C{r}/E{r})*{B(r_cr)})", f"=D{r}*{B(r_pp)}",
               f"=I{r}*{B(r_roy)}+{B(r_mgmt)}+I{r}*{B(r_fee)}", f"={B(r_fix)}",
               f"=I{r}-J{r}-K{r}-L{r}-M{r}-N{r}", f"=O{r}-{B(r_init)}"])
        else:
            p=r-1
            A([f"=A{p}+1", f"=B{p}+{B(r_de)}", f"=C{p}+{B(r_dc)}", f"=D{p}+{B(r_dp)}",
               f"=B{r}+C{r}+D{r}", f"=E{r}*{B(r_cap)}", f"=H{p}*(1-{B(r_chn)})+{B(r_sk)}", f"=MIN(G{r},F{r})",
               f"=H{r}*{B(r_unit)}", f"=B{r}*{B(r_ep)}", f"=IF(E{r}=0,0,I{r}*(C{r}/E{r})*{B(r_cr)})", f"=D{r}*{B(r_pp)}",
               f"=I{r}*{B(r_roy)}+{B(r_mgmt)}+I{r}*{B(r_fee)}", f"={B(r_fix)}",
               f"=I{r}-J{r}-K{r}-L{r}-M{r}-N{r}", f"=P{p}+O{r}"])
    ee=len(R)
    m12row=es+12; m18row=es+18  # 1-idx
    A([""])
    A(["■ サマリ"]); sec8=len(R)-1
    A(["12ヶ月時点", f"利用者 =H{m12row}", f"売上/月 =I{m12row}", f"営業利益/月 =O{m12row}", f"累積CF =P{m12row}"])
    A(["18ヶ月時点", f"利用者 =H{m18row}", f"売上/月 =I{m18row}", f"営業利益/月 =O{m18row}", f"累積CF =P{m18row}"])
    # 数式を =H.. にするためA1で直書きできないので、サマリは値参照式に
    R[-2]=["12ヶ月時点", f"=H{m12row}", f"=I{m12row}", f"=O{m12row}", f"=P{m12row}"]+[""]*(NC-5)
    R[-1]=["18ヶ月時点", f"=H{m18row}", f"=I{m18row}", f"=O{m18row}", f"=P{m18row}"]+[""]*(NC-5)
    sumrow1=len(R)-2; sumrow2=len(R)-1

    svc.spreadsheets().values().update(spreadsheetId=SID,range=f"'{TAB}'!A1",valueInputOption="USER_ENTERED",body={"values":R}).execute()

    # 書式
    reqs=cw(gid,[150,150,90,90,110,120,150,110,110,110,110,110,120,100,120,130])
    reqs+=[mr(gid,0,0,NC),fr(gid,0,REDD,WHT,True,13),rh(gid,0,42),
           mr(gid,1,0,NC),fr(gid,1,REDBG,REDD,True,9),rh(gid,1,46)]
    for s in (sec1,sec2,sec3,sec4,sec5,sec6,sec7,sec8):
        reqs+=[mr(gid,s,0,NC),fr(gid,s,SEC,WHT,True,11),rh(gid,s,26)]
    # シナリオ表
    reqs.append(fr(gid,h1,RED,WHT,True,10,True,"CENTER"))
    for i in range(s1,e1):
        bg=TIERA if i==s1 else (HI if i==e1-1 else WHT)
        reqs.append(fr(gid,i,bg,INK,False,10,True,"LEFT")); reqs.append(rh(gid,i,40))
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True}}},"fields":"userEnteredFormat.textFormat"}})
    # 結論テキスト行
    for i in range(sec2+1,sec3):
        reqs+=[mr(gid,i,0,NC),fr(gid,i,WHT,INK,False,10),rh(gid,i,30)]
    # 入力ラベル+値（A=ラベル, B=黄入力）
    for s,e in [(sec3+1,r_init),(sec4+1,r_dp),(sec5+1,r_pp),(sec6+1,r_sk)]:
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s-1,"endRowIndex":e,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10},"horizontalAlignment":"LEFT"}},"fields":"userEnteredFormat(textFormat,horizontalAlignment)"}})
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s-1,"endRowIndex":e,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":INBG,"horizontalAlignment":"RIGHT","textFormat":{"bold":True,"fontSize":10}}},"fields":"userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"}})
    # ★本部紹介の月次採用(業委)を特に強調
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":r_dc-1,"endRowIndex":r_dc,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":HI}},"fields":"userEnteredFormat.backgroundColor"}})
    # エンジン
    reqs.append(fr(gid,eh,RED,WHT,True,9,True,"CENTER")); reqs.append(rh(gid,eh,34))
    for i in range(es,ee):
        reqs.append(fr(gid,i,CARD if (i-es)%2 else WHT,INK,False,9,False,"RIGHT")); reqs.append(rh(gid,i,20))
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":es,"endRowIndex":ee,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    # 利用者列H・営業利益O・累積P を強調
    for col in (7,14,15):
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":es,"endRowIndex":ee,"startColumnIndex":col,"endColumnIndex":col+1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True}}},"fields":"userEnteredFormat.textFormat"}})
    # 数値フォーマット（売上以降を ¥/整数, 率は%）
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":es,"endRowIndex":ee,"startColumnIndex":8,"endColumnIndex":16},"cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"#,##0;▲#,##0"}}},"fields":"userEnteredFormat.numberFormat"}})
    # サマリ
    for i in (sumrow1,sumrow2):
        reqs.append(fr(gid,i,TIERA if i==sumrow1 else HI,INK,True,11,False,"LEFT")); reqs.append(rh(gid,i,28))
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":5},"cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"#,##0;▲#,##0"},"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat(numberFormat,horizontalAlignment)"}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":reqs}).execute()
    print("DONE tab=",TAB,"gid=",gid)
    print("=== 静的シナリオ(検証) ===")
    for sr in scen_rows: print(sr[0],"| 12M利用者",sr[3],"| 12M利益/月",sr[5],"| 18M累積",sr[6],"| 回収",sr[7])
    print("URL:",f"https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
