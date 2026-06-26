# -*- coding: utf-8 -*-
"""
新規スプレッドシートを作成：オーロラ高根 開業準備 司令塔＋目標(上中下)＋実行WBS。
情報過多を断ち切り「何を見るか／現状から上中下／日々やること(優先度)」に集約。
auth: scripts/sheets_token.pickle (scope=spreadsheets ※create可)
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
TITLE = "オーロラ高根_開業準備 司令塔＋目標(上中下)＋実行WBS_2026"

RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
INBG={"red":0.999,"green":0.984,"blue":0.882}; TIERA={"red":0.886,"green":0.949,"blue":0.905}
SEC={"red":0.357,"green":0.357,"blue":0.357}; BLUEBG={"red":0.812,"green":0.886,"blue":0.953}
HI={"red":1.0,"green":0.949,"blue":0.800}

def creds():
    with open(TOKEN,"rb") as f: c=pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request());
        with open(TOKEN,"wb") as f: pickle.dump(c,f)
    return c

def mr(gid,r,c0,nc): return {"mergeCells":{"range":{"sheetId":gid,"startRowIndex":r,"endRowIndex":r+1,"startColumnIndex":c0,"endColumnIndex":nc},"mergeType":"MERGE_ALL"}}
def fr(gid,r,bg,fg,bold,size,wrap=True,h="LEFT"):
    return {"repeatCell":{"range":{"sheetId":gid,"startRowIndex":r,"endRowIndex":r+1},
        "cell":{"userEnteredFormat":{"backgroundColor":bg,"textFormat":{"foregroundColor":fg,"bold":bold,"fontSize":size},
            "wrapStrategy":"WRAP" if wrap else "OVERFLOW_CELL","verticalAlignment":"MIDDLE","horizontalAlignment":h}},"fields":"userEnteredFormat"}}
def rh(gid,r,px,r2=None): return {"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":r,"endIndex":(r2 or r)+1},"properties":{"pixelSize":px},"fields":"pixelSize"}}
def cw(gid,ws): return [{"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":w},"fields":"pixelSize"}} for i,w in enumerate(ws)]
def body(gid,a,b,wrap=True):
    return {"repeatCell":{"range":{"sheetId":gid,"startRowIndex":a,"endRowIndex":b},
        "cell":{"userEnteredFormat":{"wrapStrategy":"WRAP" if wrap else "OVERFLOW_CELL","verticalAlignment":"MIDDLE","textFormat":{"fontSize":10,"foregroundColor":INK}}},
        "fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}}
def alt(gid,a,b,nc):
    r=[]
    for i in range(a,b):
        r.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":i,"endRowIndex":i+1},"cell":{"userEnteredFormat":{"backgroundColor":CARD if (i-a)%2 else WHT}},"fields":"userEnteredFormat.backgroundColor"}})
    return r
def borders(gid,a,b,nc):
    return {"updateBorders":{"range":{"sheetId":gid,"startRowIndex":a,"endRowIndex":b,"startColumnIndex":0,"endColumnIndex":nc},
        "innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},"innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}}

def put(svc,sid,title,rows):
    svc.spreadsheets().values().update(spreadsheetId=sid,range=f"'{title}'!A1",valueInputOption="USER_ENTERED",body={"values":rows}).execute()


# ───────── 00 司令塔 ─────────
def t_cockpit(svc,sid,gid):
    NC=5
    R=[]; add=lambda r: R.append(r+[""]*(NC-len(r)))
    secs=[]; tblhdr=[]; tbl=[]
    add(["🧭 オーロラ高根 開業準備｜司令塔（現在地・優先・上中下）"])
    add(["北極星＝12ヶ月で利用者75名／営業利益 月91万（高根PL）。それを満たす“送客”と“採用”の地盤を、開業前の今つくる。情報は溢れる→ここに書いた数個だけ見る。"])
    add([""])
    add(["■ 開業準備フェーズで“見るべき”のはこの3つ（先行指標＝今日動かせる）"]); secs.append(len(R)-1)
    add(["#","指標","現状","中(目標)","なぜ見るか／優先"]); tblhdr.append(len(R)-1); s=len(R)
    add(["1","採用（有資格 施術者数）","0名","2名","◎送客を“受ける”供給。最大律速。いないと送客が捌けず信頼を失う"])
    add(["2","ケアマネ接触（40居宅）","0件","40件接触/2巡","◎利用者獲得の蛇口。先行指標で今日動かせる。9-11/16-18・雨天◎"])
    add(["3","同意書が出る地盤（高根周辺の医療機関）","未把握","取りやすい5院を把握","◎CVの律速。徳洲会は遅い前提。エリア差を開業前に潰す"])
    tbl.append((s,len(R)))
    add([""])
    add(["■ 律速の順番（ここを外さない）"]); secs.append(len(R)-1)
    add(["採用　＞　同意書の地盤　＞　ケアマネ接触　＞　（開業後）利用者の積み上げ"])
    add(["※利用者数・売上は“遅行＝結果”。今は先行（採用/接触/同意書）だけ見る。利用者数で一喜一憂しない。"])
    add([""])
    add(["■ 今週の最優先3つ（毎週ここを書き換える）"]); secs.append(len(R)-1)
    add(["① 6/4(木)13:00 狩野さん対面：3タブ＋構造化docを見せ、確認2点（活動量/転換率）＋削ぎ落とし依頼"])
    add(["② 採用：Indeedをパート主軸に再設計／本部紹介ルート（学校・エリア違い応募者が高根に回るか）を確認"])
    add(["③ 同意書：高根周辺の病院・診療所をリスト化→出やすさを事前問い合わせ着手（徳洲会は遅い前提）"])
    add([""])
    add(["■ 連動・出典（数字の本体はここ）"]); secs.append(len(R)-1)
    add(["上中下の目標→「01_目標設定」タブ／日々やること→「02_実行WBS」タブ／ファネル数値・営業台帳・KPI→大元スプシ tab21-24／構造化ドキュメント→Notion"])
    reqs=cw(gid,[34,260,110,150,430])
    reqs+= [mr(gid,0,0,NC),fr(gid,0,REDD,WHT,True,14),rh(gid,0,42),
            mr(gid,1,0,NC),fr(gid,1,REDBG,REDD,True,10),rh(gid,1,52)]
    for si in secs: reqs+=[mr(gid,si,0,NC),fr(gid,si,SEC,WHT,True,11),rh(gid,si,26)]
    for hi in tblhdr: reqs.append(fr(gid,hi,RED,WHT,True,10,True,"CENTER"))
    for (a,b) in tbl:
        reqs.append(body(gid,a,b)); reqs+=alt(gid,a,b,NC); reqs.append(borders(gid,a-1,b,NC)); reqs.append(rh(gid,a,46,b-1))
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":a,"endRowIndex":b,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":a,"endRowIndex":b,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"textFormat":{"bold":True}}},"fields":"userEnteredFormat.textFormat"}})
    # 今週最優先＆律速＆出典の本文行はマージ＋折返し
    for i in range(len(R)):
        if R[i][0] and R[i][1]=="" and i not in secs and i not in [0,1]:
            reqs+=[mr(gid,i,0,NC),fr(gid,i,WHT,INK,False,10),rh(gid,i,30)]
    put(svc,sid,"00_司令塔",R)
    return reqs


# ───────── 01 目標設定(上中下) ─────────
def t_goals(svc,sid,gid):
    HDR=["#","領域","指標","区分","現状","下(最低死守)","中(目標)","上(ストレッチ)","期限","優先","根拠/備考"]
    NC=len(HDR)
    ROWS=[
    ["1","採用","有資格 施術者数","先行","0名","1名","2名","3名","開業前〜M1","◎","本部紹介中心＋パート寄せ(板橋)。送客を受ける供給=最大律速"],
    ["2","採用","採用チャネル確立(本部紹介/養成校/Indeedパート)","先行","未","1ルート","2ルート","3ルート","M1","○","Indeed単独依存しない(板橋はIndeed不使用)"],
    ["3","ケアマネ営業","居宅接触数(40件)","先行","0件","TierA18件","40件","40件×複数接点","M2","◎","9-11/16-18・2週で1巡。台帳22で管理"],
    ["4","ケアマネ営業","グリップ③面談以上","先行","0件","10件","20件","30件","M2","○","主任CM/困難ケース保有を優先"],
    ["5","同意書","出やすい医療機関の把握","基盤","未把握","3院確認","取れる5院＋不可も把握","主要全院マップ","開業前","◎","徳洲会は遅い前提。エリア差を先に把握"],
    ["6","問合せ","新規問合せ/月","中間","0","5","10","15","M3","○","接触50%→問合せ率5→9→13%"],
    ["7","CV","新規利用者/月","遅行","0","3","5","8","M3-6","△","CV55%・同意書が律速"],
    ["8","利用者","利用者数(6ヶ月)","遅行","0","15名","20名","25名","M6","△","tab20/21(要狩野/本部確認)"],
    ["9","利用者","利用者数(12ヶ月)","遅行","0","47名","61名","75名","M12","△","同上。解約4%/月を織込"],
    ["10","採算","月次営業利益(12ヶ月)","遅行","—","30万","65万","91万","M12","△","tab20 高根PL"],
    ["11","資金","運転資金の確保","基盤","—","加盟金+6M(≈590万)","+予備3M","余裕","開業前","◎","公庫創業計画書(作成済)→融資打診"],
    ["12","稼働","施術者1人あたり利用者(上限22名)","管理","0","—","18名で増員判断","22名上限","随時","○","超える前に採用＝律速の同期"],
    ]
    R=[["🎯 目標設定（上中下）｜現状値からの梯子。下=最低死守／中=目標／上=ストレッチ"]+[""]*(NC-1),
       ["主に“先行指標”を管理（採用・接触・同意書）。遅行(利用者/売上)は結果なので追わない。優先◎>○>△。"]+[""]*(NC-1),
       HDR]+ROWS
    s=3; e=3+len(ROWS)
    reqs=cw(gid,[30,96,250,56,96,150,160,150,86,46,250])
    reqs+=[mr(gid,0,0,NC),fr(gid,0,REDD,WHT,True,14),rh(gid,0,42),
           mr(gid,1,0,NC),fr(gid,1,REDBG,REDD,True,9),rh(gid,1,34),
           fr(gid,2,RED,WHT,True,10,True,"CENTER"),rh(gid,2,38),
           body(gid,s,e)]
    reqs+=alt(gid,s,e,NC)
    # 区分で色分け(先行=薄緑/基盤=黄)
    for i,row in enumerate(ROWS):
        if row[3]=="先行" or row[3]=="基盤":
            reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s+i,"endRowIndex":s+i+1,"startColumnIndex":3,"endColumnIndex":4},"cell":{"userEnteredFormat":{"backgroundColor":TIERA if row[3]=="先行" else HI,"horizontalAlignment":"CENTER","textFormat":{"bold":True}}},"fields":"userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"}})
    # 上中下=入力意識で中(F)強調・優先列中央
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s,"endRowIndex":e,"startColumnIndex":6,"endColumnIndex":7},"cell":{"userEnteredFormat":{"backgroundColor":INBG,"textFormat":{"bold":True}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s,"endRowIndex":e,"startColumnIndex":9,"endColumnIndex":10},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"fontSize":12}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s,"endRowIndex":e,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    reqs.append(rh(gid,s,40,e-1)); reqs.append(borders(gid,2,e,NC))
    reqs.append({"setBasicFilter":{"filter":{"range":{"sheetId":gid,"startRowIndex":2,"endRowIndex":e,"startColumnIndex":0,"endColumnIndex":NC}}}})
    put(svc,sid,"01_目標設定(上中下)",R)
    return reqs


# ───────── 02 実行WBS ─────────
def t_wbs(svc,sid,gid):
    HDR=["領域","タスク（日々やること）","優先","頻度","期限/開始","担当","関連目標#","状態","メモ"]
    NC=len(HDR)
    G=[
     ("【6/4 対面｜最重要・単発】",[
        ("ー","狩野さんに3タブ＋構造化docを見せ、確認2点(活動量/転換率)＋削ぎ落とし依頼","◎","単発","6/4","菊池","全体","未","深掘り1枚紙を持参"),
     ]),
     ("【① 採用（最大律速）】",[
        ("採用","Indeed求人をパート主軸に再設計(業務委託フル歩合は初期グリップ難)","◎","単発","〜6/10","菊池/宮崎","1,2","未","板橋:パート◎"),
        ("採用","本部紹介ルートの実態確認(学校紹介/エリア違い応募者が高根に回るか)","◎","単発","6/4・本部面談","菊池","1,2","未","板橋はIndeed不使用=紹介が主"),
        ("採用","近隣の鍼灸/あマ指 養成校をリスト化→求人打診","○","単発","〜M1","宮崎","2","未",""),
        ("採用","Indeed反響(週IMP/応募)を確認","○","毎週","毎週金","菊池","1","未","週IMP約30=露出不足"),
     ]),
     ("【② 同意書の地盤】",[
        ("同意書","高根周辺の病院・診療所をリスト化(東船橋病院/船橋総合/徳洲会/近隣診療所)","◎","単発","〜6/6","菊池","5","未",""),
        ("同意書","出やすさを事前問い合わせ(徳洲会は遅い前提・取りやすい診療科を把握)","◎","単発","6/3〜6/10","菊池/事務","5","未","CVの律速を先に潰す"),
     ]),
     ("【③ ケアマネ営業（蛇口）】",[
        ("営業","居宅TierA18件へ訪問(9-11/16-18・雨天◎)","◎","毎週","開業前後","菊池","3","未","台帳22 Tier A"),
        ("営業","訪問後すぐ台帳22更新(接触日/グリップ段階▼/送客)","◎","毎日","訪問都度","菊池","3,4","未",""),
        ("営業","主任CM/困難ケース保有CMを特定","○","毎週","—","菊池","4","未",""),
        ("営業","プレイブック24の手土産/トークで関係維持・月1報告","○","毎週","—","菊池/施術者","4","未","送客を切らさない"),
     ]),
     ("【④ 数字／レビュー】",[
        ("数字","週次活動ログをtab23に入力→計画対比をレビュー","◎","毎週","毎週月","菊池","3,6","未",""),
        ("数字","先行指標(訪問/接触)だけ毎日チェック(利用者数で一喜一憂しない)","○","毎日","—","菊池","—","未",""),
     ]),
     ("【⑤ 資金／本部】",[
        ("資金","公庫創業計画書を本部確認の数字で更新→融資打診","○","単発","M1","菊池/橋本","11","未","様式作成済"),
        ("本部","本部面談で加盟条件/テリトリー/解約条件を確認","○","単発","本部面談","菊池","11","未",""),
     ]),
    ]
    R=[["✅ 実行WBS（日々やること）｜領域別・優先度つき。頻度=毎日/毎週/単発。状態=未/進行/完"]+[""]*(NC-1),
       ["“今やること”はここだけ見る。優先◎を上から。関連目標#は「01_目標設定」の番号。"]+[""]*(NC-1),
       HDR]
    grp_idx=[]; data_spans=[]
    for gname,tasks in G:
        R.append([gname]+[""]*(NC-1)); grp_idx.append(len(R)-1)
        s=len(R)
        for t in tasks: R.append(list(t))
        data_spans.append((s,len(R)))
    # デイリールーティンの型
    R.append([""]*NC)
    R.append(["■ デイリールーティンの型（開業準備〜初動）"]+[""]*(NC-1)); rt_sec=len(R)-1
    R.append(["朝 9-11時：居宅訪問2-3件 ／ 日中：採用・同意書の単発タスク ／ 夕 16-18時：居宅訪問2-3件 ／ 夜：台帳22更新＋先行指標1分チェック"]+[""]*(NC-1)); rt_body=len(R)-1
    e_all=len(R)
    reqs=cw(gid,[150,330,42,60,110,96,80,56,180])
    reqs+=[mr(gid,0,0,NC),fr(gid,0,REDD,WHT,True,14),rh(gid,0,42),
           mr(gid,1,0,NC),fr(gid,1,REDBG,REDD,True,9),rh(gid,1,30),
           fr(gid,2,RED,WHT,True,10,True,"CENTER"),rh(gid,2,34)]
    for gi in grp_idx: reqs+=[mr(gid,gi,0,NC),fr(gid,gi,SEC,WHT,True,10),rh(gid,gi,24)]
    for (a,b) in data_spans:
        reqs.append(body(gid,a,b)); reqs+=alt(gid,a,b,NC); reqs.append(rh(gid,a,38,b-1))
        # 優先(C=2)中央太字・頻度(D=3)中央・状態(H=7)中央/黄
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":a,"endRowIndex":b,"startColumnIndex":2,"endColumnIndex":4},"cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"fontSize":10}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":a,"endRowIndex":b,"startColumnIndex":7,"endColumnIndex":8},"cell":{"userEnteredFormat":{"backgroundColor":INBG,"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,horizontalAlignment)"}})
    reqs.append(borders(gid,2,e_all-2,NC))
    reqs+=[mr(gid,rt_sec,0,NC),fr(gid,rt_sec,SEC,WHT,True,11),rh(gid,rt_sec,26),
           mr(gid,rt_body,0,NC),fr(gid,rt_body,BLUEBG,INK,True,10),rh(gid,rt_body,40)]
    put(svc,sid,"02_実行WBS",R)
    return reqs


def main():
    svc=build("sheets","v4",credentials=creds(),cache_discovery=False)
    created=svc.spreadsheets().create(body={"properties":{"title":TITLE},
        "sheets":[{"properties":{"title":t,"gridProperties":{"rowCount":80,"columnCount":12,"frozenRowCount":fr}}}
                  for t,fr in [("00_司令塔",2),("01_目標設定(上中下)",3),("02_実行WBS",3)]]}).execute()
    sid=created["spreadsheetId"]
    gids={s["properties"]["title"]:s["properties"]["sheetId"] for s in created["sheets"]}
    reqs=[]
    reqs+=t_cockpit(svc,sid,gids["00_司令塔"])
    reqs+=t_goals(svc,sid,gids["01_目標設定(上中下)"])
    reqs+=t_wbs(svc,sid,gids["02_実行WBS"])
    svc.spreadsheets().batchUpdate(spreadsheetId=sid,body={"requests":reqs}).execute()
    print("DONE  SID=",sid)
    print("URL:",f"https://docs.google.com/spreadsheets/d/{sid}/edit")

if __name__=="__main__":
    main()
