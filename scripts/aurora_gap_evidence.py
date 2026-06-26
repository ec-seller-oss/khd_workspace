# -*- coding: utf-8 -*-
"""
事業計画書スプシ(SID_B)に【10_福井様式比較_抜け漏れ】【11_証票セット_裏付け】タブを追加。2026-06-05。
"""
import pickle
from googleapiclient.discovery import build as gbuild
from google.auth.transport.requests import Request
TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID_B="1ruE_rE8i_OKpO0JMV-82AJDMqLPf8S-r4HxJIEgyL8U"
RED={"red":0.667,"green":0.180,"blue":0.149};REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882};REDBG={"red":0.957,"green":0.894,"blue":0.886}
YEL={"red":1.0,"green":0.97,"blue":0.80};WHT={"red":1,"green":1,"blue":1};INK={"red":0.1,"green":0.1,"blue":0.1}

GAP=[
 ("TH",["福井様式の項目","状態","不足点・対応・証票"]),
 ("R",["基本情報／診療圏／競合","○","記入済（診療圏は別スプシ・実データ）"]),
 ("R",["人員構成（職種/人数/給与）","○","施術者・営業を年次で記入。証票=本部資料P15/P24"]),
 ("R",["患者数想定・単価","○","利用者ramp＋単価3.8万。証票=本部P15(4,000〜4,500円/回)"]),
 ("R",["収入額算定表（5年）","○","利用者×単価×12で記入済"]),
 ("R",["物件情報(用途地域/面積/賃料)","△","家賃8万のみ＝高根の事務所が未確定。物件決定後に賃貸借条件を記入【要取得】"]),
 ("R",["医療原価（消耗品）","△","施術人件費に内包＝消耗品原価を別出し（売上の数%）。証票=本部資料"]),
 ("R",["法定福利費（社保）","✕","雇用なら給与の約15%を加算。設計は福井相談（雇用/委託）【要確定】"]),
 ("R",["減価償却費","✕","車両・施術機材・サイネージ。証票=本部初期費用P27"]),
 ("R",["リース明細（車両）","✕","車両リース 月約2.8万。証票=本部P24(名東 リース1台28,000円)"]),
 ("R",["経費明細(光熱/通信/広告/保険/修繕/消耗品)","△","『その他』に一括→項目別に分解。証票=本部FCシミュP28"]),
 ("R",["損益分岐点（固定費÷限界利益率）","△","利用者16名と概算→正式計算を追加【要追記】"]),
 ("R",["顧客推移×稼働日数","△","利用者数のみ→1人1日件数×稼働日で裏取り。証票=石原9〜10件/日"]),
 ("R",["可処分所得・運転資金 推移","✕","返済後CFはあるが、可処分所得・運転資金・累計の推移を追加【要追記】"]),
 ("R",["借入金返済明細（月次/元利内訳）","△","年次サマリのみ→月次or元利内訳を追加【要追記】"]),
 ("NOTE","総括＝福井様式に対し『物件・法定福利費・減価償却・リース・経費明細・損益分岐点の正式計算・可処分所得/運転資金推移・月次返済明細』が不足。物件と労務/税務の前提が固まれば埋まる。参考資料が限られたため、数字の裏付けは下記『11_証票セット』で提出する。"),
]
EVID=[
 ("TH",["No","証票（裏付け資料）","内容（何を裏付けるか）","所在"]),
 ("R",["①","本部 FC加盟店資料2026最新.pdf","単価4,000〜4,500円/回・施術者人件費・ロイヤリティ12%・FCシミュ(P28)・初期費用(P27)・名東/一宮実績","Drive(本部資料)"]),
 ("R",["②","板橋オーナー(狩野)ヒアリング記録","採用(本部紹介)・同意書(徳洲会は遅い)・黒字化(3→5月)・移動(自転車/バイク)","大元スプシ16タブ"]),
 ("R",["③","石原氏 面談記録(6/2)","訪問は1日9〜10件・競合マップ・高根エリアの土地勘","大元スプシ13タブ"]),
 ("R",["④","船橋市 町丁別年齢別人口(R8.4)Excel","診療圏の75歳以上＝需要母数(核診療圏 6,616人)","Drive 287/01_概要＋添付済"]),
 ("R",["⑤","ヒトイキ紹介条件＋候補者キャリアシート","採用コスト(常勤=年収35%/委託60万)・即戦力鍼灸師の人材","Gmail(ヒトイキ往復)"]),
 ("R",["⑥","競合各社公式／千葉徳洲会病院","競合所在・同意書元(高根台2-11-1)","各公式サイト"]),
 ("R",["⑦","本部 加盟請求書","加盟金の実額","Drive(本部資料)"]),
 ("NOTE","裏付けが足りず【要取得/要確定】の点：(a)高根の事務所賃料の実額(物件未確定) (b)同意書の通過率・所要日数の実数 (c)本部の標準解約率/離職率の実データ (d)施術者の歩合率の実数 (e)法定福利費・税務の前提(雇用/委託)。→ (b)(c)(d)は狩野さん、(e)は福井さんに確認。"),
]

def creds():
    with open(TOKEN,"rb") as f:c=pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN,"wb") as f:pickle.dump(c,f)
    return c
def fill(svc,sid,gid,tt,head,rows,widths):
    N=len(widths);vals=[[head]+[""]*(N-1)];meta=[]
    for rr in rows:
        i=len(vals);k=rr[0]
        if k=="TH":vals.append(rr[1]+[""]*(N-len(rr[1])));meta.append((i,"TH"))
        elif k=="R":vals.append(rr[1]+[""]*(N-len(rr[1])));meta.append((i,"R",rr[1][1] if len(rr[1])>1 else ""))
        elif k=="NOTE":vals.append([rr[1]]+[""]*(N-1));meta.append((i,"NOTE"))
    svc.spreadsheets().values().update(spreadsheetId=sid,range=f"'{tt}'!A1",valueInputOption="USER_ENTERED",body={"values":vals}).execute()
    rq=[]
    for c,w in enumerate(widths):
        rq.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":c,"endIndex":c+1},"properties":{"pixelSize":w},"fields":"pixelSize"}})
    rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":len(vals)},"cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    rq.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}})
    rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},"cell":{"userEnteredFormat":{"backgroundColor":REDD,"textFormat":{"foregroundColor":WHT,"bold":True,"fontSize":13}}},"fields":"userEnteredFormat"}})
    rq.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":42},"fields":"pixelSize"}})
    for mm in meta:
        i=mm[0];k=mm[1]
        if k=="TH":
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":10,"foregroundColor":REDD},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat"}})
        elif k=="R":
            st=mm[2]
            rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
            if st=="✕":
                rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":{"red":0.98,"green":0.85,"blue":0.85},"horizontalAlignment":"CENTER","textFormat":{"bold":True,"foregroundColor":RED}}},"fields":"userEnteredFormat"}})
            elif st=="△":
                rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":YEL,"horizontalAlignment":"CENTER","textFormat":{"bold":True}}},"fields":"userEnteredFormat"}})
            elif st=="○":
                rq.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"foregroundColor":{"red":0.1,"green":0.5,"blue":0.2}}}},"fields":"userEnteredFormat"}})
        elif k=="NOTE":
            rq+=[{"mergeCells":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1,"startColumnIndex":0,"endColumnIndex":N},"mergeType":"MERGE_ALL"}},
                 {"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":REDBG,"textFormat":{"bold":True,"fontSize":9,"foregroundColor":REDD},"wrapStrategy":"WRAP","verticalAlignment":"TOP"}},"fields":"userEnteredFormat"}},
                 {"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":78},"fields":"pixelSize"}}]
    th=[m[0] for m in meta if m[1]=="TH"]
    if th:
        rq.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":th[0],"endRowIndex":len(vals),"startColumnIndex":0,"endColumnIndex":N},"innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=sid,body={"requests":rq}).execute()

def main():
    svc=gbuild("sheets","v4",credentials=creds(),cache_discovery=False)
    info=svc.spreadsheets().get(spreadsheetId=SID_B,fields="sheets(properties(sheetId,title))").execute()
    have={s["properties"]["title"]:s["properties"]["sheetId"] for s in info["sheets"]}
    add=[]
    for t in ["10_福井様式比較_抜け漏れ","11_証票セット_裏付け"]:
        if t in have: svc.spreadsheets().batchUpdate(spreadsheetId=SID_B,body={"requests":[{"deleteSheet":{"sheetId":have[t]}}]}).execute()
        add.append({"addSheet":{"properties":{"title":t,"gridProperties":{"frozenRowCount":1}}}})
    res=svc.spreadsheets().batchUpdate(spreadsheetId=SID_B,body={"requests":add}).execute()
    g={r["addSheet"]["properties"]["title"]:r["addSheet"]["properties"]["sheetId"] for r in res["replies"]}
    fill(svc,SID_B,g["10_福井様式比較_抜け漏れ"],"10_福井様式比較_抜け漏れ","⑩ 福井様式との比較・抜け漏れチェック",GAP,[230,70,560])
    fill(svc,SID_B,g["11_証票セット_裏付け"],"11_証票セット_裏付け","⑪ 証票セット（数字の裏付け）＋要取得",EVID,[40,260,330,190])
    print("DONE")
    print(f"B: https://docs.google.com/spreadsheets/d/{SID_B}/edit")

if __name__=="__main__":
    main()
