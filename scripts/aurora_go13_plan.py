# -*- coding: utf-8 -*-
"""
開業用スプシ「04_事業計画書（オーロラ船橋店）」に
「23_GO前13日プラン(6_11-6_24)」タブを追加。
0_前提の検証結果（修正箇所・不足点・対策）＋日割りプランを1タブに集約。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID="1ruE_rE8i_OKpO0JMV-82AJDMqLPf8S-r4HxJIEgyL8U"
TAB="23_GO前13日プラン(6_11-6_24)"

RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
INBG={"red":0.999,"green":0.984,"blue":0.882}; TIERA={"red":0.886,"green":0.949,"blue":0.905}
SEC={"red":0.357,"green":0.357,"blue":0.357}; HI={"red":1.0,"green":0.898,"blue":0.6}
BLUEBG={"red":0.812,"green":0.886,"blue":0.953}

FIX=[ # 種別, 内容, 対策, 優先, 期限
 ["🔧修正","資金調達額がタブ間で不一致：0_前提=公庫1,100万(計1,400万)⇄ 6_資金調達/16=公庫440万(計740万)。16の『✗不足』判定は740万前提のまま","どちらを正にするか今日決定→6/16タブを1,400万で再計算。C型必要資金1,356万に対し1,400万だと残44万(3%)のみ","◎","6/11"],
 ["🔧修正","施術者人件費率0.5(C型blend)は初年度楽観：立上げはパート採用が間に合わず業務委託(歩合60%＝狩野相場)から始まる公算","初年度のみ0.60→2年目以降0.50の2段に。17_採用シナリオと整合させる","◎","6/13"],
 ["🔧修正","3_利用者数想定が年平均のみで月次カーブなし＋注記に旧ファネル(接触50%)が残存（0_前提は狩野反映で30%）","注記を0_前提参照に修正。月次は『1事業所から送客まで2-4ヶ月』(狩野)を織り込んだ立上げカーブで年平均32の根拠を確認","○","6/18"],
 ["⚠️不足","レセプト入金サイト未反映：療養費は請求→入金まで2〜4ヶ月遅れ。営業利益ベースの16タブより実際の資金の谷は深い","16タブに入金遅延2ヶ月でのCF行を追加。バッファ44万では吸収不能→公庫増額 or 自己資金の厚み確認","◎","6/15"],
 ["⚠️不足","受領委任の施術管理者要件：実務経験1年＋指定研修が必要。雇った施術者＝施術管理者だと退職で保険請求が止まる(単一障害点)","本部へ質問：誰が施術管理者になるか／要件充足者の紹介可否／医療事務代行の範囲","◎","6/16"],
 ["⚠️不足","保険の段取り：施術賠償責任保険・事業用車両保険(中古車購入方針)が計画に未計上","本部推奨の保険を確認→5_概算事業費へ計上","○","6/16"],
 ["⚠️不足","狩野さんとの協力体制が口頭のみ(6/17再アポ予定)","枠組み(採用連携・ノウハウの扱い・対価)を1枚書面化して6/17に持参","○","6/17"],
 ["⚠️不足","開設手続きの実日程：施術所開設届/出張施術届(保健所)・受領委任登録の所要日数が未確認＝開業日が逆算できない","本部に標準スケジュールを確認(9_抜け漏れ⑤と統合)","○","6/18"],
]

PLAN=[ # 日付(テキスト固定), 曜, やること, ゴール/完了条件
 ["6/11","木","【資金の正を確定】公庫1,100万申込で統一→0_前提/6/16タブ/創業計画書(自己600万のまま)を一致させる","タブ間・計画書の数字が1つに揃う"],
 ["6/12","金","☎公庫 事業資金相談ダイヤル(0120-154-505)：法人窓口の管轄確認(上野でいけるか/船橋支店か)＋法人初回の必要書類確定／ヒトイキ候補 面接(予定済)／Indeedパート主軸 再設計","公庫の窓口・書類が確定／採用複線①前進"],
 ["6/13","土","AM:同意書ヒア1〜2件(高根周辺診療所・徳洲会は遅い前提)＋人件費率2段(60→50%)を17タブに反映","同意書の感触取得・修正②完了"],
 ["6/14","日","家族(死守)","—"],
 ["6/15","月","レセプト入金遅延2ヶ月をCF表に反映(16タブ)→真の資金の谷を確定→運転資金の申込額に反映","『谷がいくらか』確定＝運転資金の根拠"],
 ["6/16","火","本部へ質問送付：施術管理者要件/受領委任/保険/開設日程＋★『支払期日を融資実行後にできるか・契約書(案)の事前提供』を必ず確認","不足③④⑤＋支払タイミングが回答待ちに"],
 ["6/17","水","狩野さん再アポ：協力体制の書面たたき合意＋0_前提の最終レビュー依頼","協力枠組み合意"],
 ["6/18","木","単店PL/BEP確定(P1)／創業計画書の空欄を24タブの記入値で埋め完成／3タブ月次カーブ確認","収益○の根拠＋計画書 完成"],
 ["6/19","金","採用確度の棚卸し(内定or複線見込み)／妻プレゼン資料準備","採用○の根拠整理"],
 ["6/20","土","栄町売却の登記進捗確認(決済は登記律速・6月末〜7月見込み)／空き枠で公庫面談の想定問答練習","栄町の現在地把握・面談準備"],
 ["6/21","日","家族＋妻の最終承認の場","妻承認"],
 ["6/22","月","3点セット仮判定(採用○×収益○×同意書○)＋本部回答の反映","仮判定が出る"],
 ["6/23","火","最終確認：必要書類スキャン準備(登記簿/決算書/通帳/見積)・残課題ゼロチェック","申込ボタンを押せる状態"],
 ["6/24","水","★Go/No-go確定 → GOなら本部へ船橋確定連絡＋同日中に公庫オンライン申込(契約書案・見積添付)","意思決定＋申込完了"],
 ["6/25-7月","—","公庫面談(7月上旬目安)→融資実行・着金(申込から3週間〜1ヶ月)→★着金後に加盟金支払い→研修→開業準備","※先に自己資金で払うと融資対象外。支払いは必ず着金後"],
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
    res=svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":[{"addSheet":{"properties":{"title":TAB,"gridProperties":{"rowCount":60,"columnCount":NC,"frozenRowCount":2}}}}]}).execute()
    gid=res["replies"][0]["addSheet"]["properties"]["sheetId"]

    R=[];A=lambda r:R.append(r+[""]*(NC-len(r)))
    A(["🛡 GO前13日プラン（6/11→6/24）｜“ノーリスク”でなく“既知・有界・ゲート管理”にする"])
    A(["判定＝①致命3リスク(採用/同意書/資金の谷)はGo/No-goゲートで管理済み＝構造は正しい ②下の修正3＋不足5を潰せば『計算できるリスクだけで判断できる』状態になる ③6/24に3点セット(採用○×収益○×同意書○)で確定"])
    A([""])
    A(["■ 修正箇所・不足点と対策（潰し込みリスト）"]); sec1=len(R)-1
    A(["種別","内容","対策","優先","期限","状態"]); h1=len(R)-1
    s1=len(R)
    for f in FIX: A(f[:4]+["'"+f[4],""])  # 期限はテキスト固定(日付自動変換を防ぐ)
    e1=len(R)
    A([""])
    A(["■ 日割りプラン（6/20栄町・日曜家族を考慮済み）"]); sec2=len(R)-1
    A(["日付","曜","やること（オーロラ）","ゴール/完了条件","","状態"]); h2=len(R)-1
    s2=len(R)
    for p in PLAN: A(["'"+p[0],p[1],p[2],p[3],"",""])  # 日付はテキスト固定
    e2=len(R)
    A([""])
    A(["■ 6/24 GO判定の3点セット（17_船橋GO決定プロセスと同一）"]); sec3=len(R)-1
    A(["①採用○＝本部紹介 or Indeedパート or ヒトイキで『高根で人が採れる』根拠／②収益○＝1,400万資金でC型の谷を越えBEP利用者16名が現実的／③同意書○＝高根周辺で取れる感触。3つ揃わなければ加盟金は払わない(No-go/保留)。"])

    svc.spreadsheets().values().update(spreadsheetId=SID,range=f"'{TAB}'!A1",valueInputOption="USER_ENTERED",body={"values":R}).execute()

    reqs=cw(gid,[64,40,420,330,56,70])
    reqs+=[mr(gid,0,0,NC),fr(gid,0,REDD,WHT,True,14),rh(gid,0,42),
           mr(gid,1,0,NC),fr(gid,1,REDBG,REDD,True,9),rh(gid,1,50)]
    for s in (sec1,sec2,sec3):
        reqs+=[mr(gid,s,0,NC),fr(gid,s,SEC,WHT,True,11),rh(gid,s,26)]
    for h in (h1,h2):
        reqs.append(fr(gid,h,RED,WHT,True,10,True,"CENTER")); reqs.append(rh(gid,h,30))
    # 潰し込みリスト
    for i in range(s1,e1):
        bg=CARD if (i-s1)%2 else WHT
        reqs.append(fr(gid,i,bg,INK,False,9,True,"LEFT")); reqs.append(rh(gid,i,58))
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s1,"endRowIndex":e1,"startColumnIndex":3,"endColumnIndex":5},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"fontSize":10}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s1,"endRowIndex":e1,"startColumnIndex":5,"endColumnIndex":6},"cell":{"userEnteredFormat":{"backgroundColor":INBG,"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,horizontalAlignment)"}})
    # 日割り
    for i in range(s2,e2):
        d=R[i][0].lstrip("'")
        bg = HI if d in("6/24","6/25-7月") else (BLUEBG if d in("6/20",) else (TIERA if R[i][1] in("土","日") else (CARD if (i-s2)%2 else WHT)))
        reqs.append(fr(gid,i,bg,INK,False,9,True,"LEFT")); reqs.append(rh(gid,i,36))
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s2,"endRowIndex":e2,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"fontSize":10}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s2,"endRowIndex":e2,"startColumnIndex":5,"endColumnIndex":6},"cell":{"userEnteredFormat":{"backgroundColor":INBG,"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,horizontalAlignment)"}})
    reqs.append(mr(gid,sec3+1,0,NC)); reqs.append(fr(gid,sec3+1,HI,INK,True,10)); reqs.append(rh(gid,sec3+1,46))
    reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":h1,"endRowIndex":e2,"startColumnIndex":0,"endColumnIndex":NC},
        "innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},
        "innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":reqs}).execute()
    print("DONE tab=",TAB,"gid=",gid)
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
