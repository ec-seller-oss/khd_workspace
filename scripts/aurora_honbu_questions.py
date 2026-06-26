# -*- coding: utf-8 -*-
"""
開業用スプシ「04_事業計画書（オーロラ船橋店）」に
「25_本部質問リスト(フライハイト)」タブを追加。
そのまま送れるカバー文＋本部回答を書き込める質問表（用途・優先つき）。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID="1ruE_rE8i_OKpO0JMV-82AJDMqLPf8S-r4HxJIEgyL8U"
TAB="25_本部質問リスト(フライハイト)"

RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
INBG={"red":0.999,"green":0.984,"blue":0.882}; TIERA={"red":0.886,"green":0.949,"blue":0.905}
SEC={"red":0.357,"green":0.357,"blue":0.357}; HI={"red":1.0,"green":0.898,"blue":0.6}
BLUEBG={"red":0.812,"green":0.886,"blue":0.953}

COVER=("お世話になっております。株式会社KHDの菊池です。\n"
 "千葉県船橋市（高根）での加盟・開業に向け、事業計画と日本政策金融公庫の融資準備を進めております。\n"
 "つきましては下記についてご確認をお願いできますでしょうか。特に【★】は公庫の申込（6月下旬予定）に直結するため、優先的にご回答いただけますと大変助かります。\n"
 "また、お手数ですが ①FC加盟契約書（案） ②初期費用のお見積書 を事前にご共有いただけますと、融資申込の添付資料として使わせていただきたく存じます。\n"
 "どうぞよろしくお願いいたします。　菊池研太／ec-seller@kikuchi-hd.net")

Q=[ # 分類, 質問, 用途(何が埋まる/動く), 優先
 ["★資金/契約","加盟金・研修費・保証金・備品・機材・初期システム等の【お見積書】をいただけますか（設備資金の内訳・根拠として）","公庫の「設備資金600万」の根拠書類／創業計画書8の内訳が埋まる","◎"],
 ["★資金/契約","加盟金等の【支払期日を融資の実行後（着金後）】に設定いただけますか。難しい場合の最短可能時期は","融資の鉄則=支払済み資金は対象外。開業スケジュールの逆算が確定","◎"],
 ["★資金/契約","FC加盟契約書（案）を事前にご提供いただけますか","公庫提出の証憑／不利条項の事前チェック","◎"],
 ["資金/契約","ロイヤリティは13.2%（税込）で確定でしょうか（資料により12%表記もあり）。運営委託費・広告分担金の月額も","収益計画(7_損益)の固定費が確定","○"],
 ["資金/契約","船橋テリトリーの保護範囲・契約期間・更新条件・中途解約条件","撤退/継続のリスク評価。契約前提の確定","○"],
 ["★採用","本部紹介で船橋・高根エリアに有資格者を何名／どの頻度で紹介可能でしょうか（学校紹介・エリア違い応募者の実態）","採用シミュの最重要入力。事業の成否を直接左右","◎"],
 ["★採用","「採用支援70名」は全店合計／1店あたり年何名のどちらでしょうか。支援の具体内容（求人代行・人材紹介・媒体費）も","採用見込みの数字根拠／配置設計","◎"],
 ["採用","開業初期の理想の人員配置（社員／業務委託／パートの比率）の本部推奨は","人件費3パターン(15タブ)の確定・初期固定費リスク","○"],
 ["★保険請求","受領委任の【施術管理者】は誰が担う想定でしょうか（実務1年＋指定研修）。要件充足者を本部が紹介できますか","保険請求の生命線。退職=請求停止の単一障害点の回避","◎"],
 ["★保険請求","レセプト代行の範囲・手数料、療養費の【入金サイト（請求→入金まで何ヶ月）】","運転資金額の根拠(16タブ:2年の谷)／資金繰りの確定","◎"],
 ["★単価/モデル","訪問1回の【保険診療単価】と、施術者1人あたりの標準モデル（1日訪問件数・稼働日数・担当上限）","売上計画(3・4タブ)と創業計画書9の見通しが埋まる","◎"],
 ["同意書","医師の同意書取得について本部の支援内容。エリアで取りやすい/にくい傾向のデータがあれば","CVの律速(同意書)対策。船橋の事前確認の精度UP","○"],
 ["手続き","施術所開設届/出張施術業務開始届（保健所）・受領委任登録の【標準スケジュール】（所要日数）","開業日の逆算が可能に(現状ブランク)","○"],
 ["手続き","必要な保険（施術賠償責任・事業用車両）の本部推奨・概算","5_概算事業費への計上漏れを防ぐ","○"],
 ["手続き","車両はリース推奨と伺っていますが、中古車購入でも問題ないでしょうか（税制・費用面で検討中）","費用設計の確定(狩野=中古車有利の助言と突合)","△"],
 ["集客","開業直後の集客支援（ポスティング・LP・本部リスト等）の具体内容とCPA実績","立ち上げ初月の利用者獲得計画の精度UP","○"],
]

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
    NC=6
    res=svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":[{"addSheet":{"properties":{"title":TAB,"gridProperties":{"rowCount":40,"columnCount":NC,"frozenRowCount":2}}}}]}).execute()
    gid=res["replies"][0]["addSheet"]["properties"]["sheetId"]

    R=[];A=lambda r:R.append(r+[""]*(NC-len(r)))
    A(["📨 本部（フライハイト）質問リスト｜面談/メールでそのまま使用。回答は黄色欄に記入→計画書・融資準備に直結"])
    A(["優先◎=公庫申込(6月下旬)に直結で最優先。回答が来たら 創業計画書・7損益・16資金繰り・15人件費・採用シミュ に反映する。"])
    A([""])
    A(["■ ① 送付カバー文（このまま送れる）"]); sec1=len(R)-1
    A([COVER]); rcov=len(R)-1
    A([""])
    A(["■ ② 質問リスト（回答を黄色欄に記入）"]); sec2=len(R)-1
    A(["#","分類","質問","用途（何が埋まる/動く）","優先","本部回答"]); hdr=len(R)-1
    s=len(R)
    for i,q in enumerate(Q,1): A([str(i),q[0],q[1],q[2],q[3],""])
    e=len(R)
    A([""])
    A(["■ ③ 回答が来たら反映する先（チェック）"]); sec3=len(R)-1
    A(["見積・単価・入金サイト→創業計画書(様式)＋3/4/7/16タブ ／ 採用紹介数→採用シミュ(03)＋15人件費 ／ 契約書案→不利条項チェック＋公庫添付 ／ 支払期日→23タブの順序確定"]); rref=len(R)-1

    svc.spreadsheets().values().update(spreadsheetId=SID,range=f"'{TAB}'!A1",valueInputOption="USER_ENTERED",body={"values":R}).execute()

    reqs=cw(gid,[30,110,360,300,46,300])
    reqs+=[mr(gid,0,0,NC),fr(gid,0,REDD,WHT,True,13),rh(gid,0,42),
           mr(gid,1,0,NC),fr(gid,1,REDBG,REDD,True,9),rh(gid,1,40)]
    for sc in (sec1,sec2,sec3):
        reqs+=[mr(gid,sc,0,NC),fr(gid,sc,SEC,WHT,True,11),rh(gid,sc,26)]
    reqs+=[mr(gid,rcov,0,NC),
           {"repeatCell":{"range":{"sheetId":gid,"startRowIndex":rcov,"endRowIndex":rcov+1},"cell":{"userEnteredFormat":{"backgroundColor":BLUEBG,"wrapStrategy":"WRAP","verticalAlignment":"TOP","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat"}},
           rh(gid,rcov,150)]
    reqs.append(fr(gid,hdr,RED,WHT,True,10,True,"CENTER")); reqs.append(rh(gid,hdr,30))
    for i in range(s,e):
        bg=HI if R[i][4]=="◎" else (CARD if (i-s)%2 else WHT)
        reqs.append(fr(gid,i,bg,INK,False,10,True,"LEFT")); reqs.append(rh(gid,i,52))
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s,"endRowIndex":e,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s,"endRowIndex":e,"startColumnIndex":4,"endColumnIndex":5},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"fontSize":12}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s,"endRowIndex":e,"startColumnIndex":5,"endColumnIndex":6},"cell":{"userEnteredFormat":{"backgroundColor":INBG}},"fields":"userEnteredFormat.backgroundColor"}})
    reqs+=[mr(gid,rref,0,NC),fr(gid,rref,WHT,INK,False,9),rh(gid,rref,40)]
    reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":hdr,"endRowIndex":e,"startColumnIndex":0,"endColumnIndex":NC},
        "innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},
        "innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":reqs}).execute()
    print("DONE tab=",TAB,"gid=",gid)
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
