# -*- coding: utf-8 -*-
"""
オーロラ大元スプシの「営業運用クラスタ」を役割別3タブに再設計して構築。
  22_営業台帳_居宅(船橋)        … 営業が毎日使う事業所リスト＋接触/グリップ/送客ログ
  23_営業KPIダッシュボード(計画実績) … 週次活動ログ→自動集計→tab21計画と対比＋グリップ段階ストック
  24_営業プレイブック(役割/トーク)   … グリップ施策・手土産/コンプラ・役割分担・初回トーク
旧「22_船橋居宅ケアマネ営業(6_4)」は削除して置き換え。planning系(0-21)は不変。
auth: scripts/sheets_token.pickle  出典:ハートページ船橋市版(2024/7)
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID = "18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB_LIST = "22_営業台帳_居宅(船橋)"
TAB_DASH = "23_営業KPIダッシュボード(計画実績)"
TAB_PB   = "24_営業プレイブック(役割/トーク)"
OLD_TABS = ["22_船橋居宅ケアマネ営業(6_4)", TAB_LIST, TAB_DASH, TAB_PB]

STAGES = ["①未接触","②挨拶/名刺","③面談","④初回送客","⑤継続送客"]

# 色
RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
INBG={"red":0.999,"green":0.984,"blue":0.882}; TIERA={"red":0.886,"green":0.949,"blue":0.905}
SEC={"red":0.357,"green":0.357,"blue":0.357}; BLUEBG={"red":0.812,"green":0.886,"blue":0.953}

JIGYO = [
    ("A","天空ケアプラニングセンター","高根台","船橋市高根台7-32-15","047-460-9544",""),
    ("A","ケアプランらぼ","高根台","船橋市高根台1-11-7","047-407-2556",""),
    ("A","朝日ケアコンサルタント テレサ会 船橋","高根台","船橋市高根台3-15-5","047-469-3128",""),
    ("A","東船橋病院 居宅介護支援","高根台","船橋市高根台4-29-1","047-401-5051","病院併設=大箱"),
    ("A","ケアプラン高根台","高根台","船橋市高根台2-11-1","047-774-0414",""),
    ("A","ケアプランすみれ","高野台","船橋市高野台4-5-34","047-407-2096",""),
    ("A","ケアプランマザーリーフ","習志野台","船橋市習志野台8-12-10","047-402-6030",""),
    ("A","けあのま","習志野台","船橋市習志野台5-40-21","070-3285-5521",""),
    ("A","みずたま介護ST 北習志野","習志野台","船橋市習志野台3-17-15","047-496-7691","東京海上系・大手"),
    ("A","ケアプラン花輪","習志野台","船橋市習志野台2-71-15","047-462-2320",""),
    ("A","ムーンライトケアマネジャー事務所","習志野台","船橋市習志野台4-58-17","047-404-1582",""),
    ("A","ケアマネ事務所フリーダム","習志野台","船橋市習志野台3-13-25","090-9961-0289",""),
    ("A","ケアプラン天使","習志野台","船橋市習志野台3-12-8","047-494-2218",""),
    ("A","ケアプランゆい","薬円台","船橋市薬円台5-6-1","047-498-9126",""),
    ("A","あおぞらの里 薬円台","薬円台","船橋市薬円台4-14-16","047-456-6500","施設併設"),
    ("A","ペアレント・ホームケア親おもい","薬円台","船橋市薬円台6-11-11","047-496-0127",""),
    ("A","ソラスト薬園台","薬円台","船橋市薬円台6-16-8","047-496-3211","大手・多拠点"),
    ("A","ケアプランセンターえがお","薬円台","船橋市薬円台4-15-1","047-465-2416",""),
    ("B","いけだ居宅介護支援事業所","前原東","船橋市前原東1-6-4","047-472-2310",""),
    ("B","テルウェル東日本 船橋介護センタ","前原西","船橋市前原西6-6-80","047-455-8102","NTT系・大手"),
    ("B","居宅介護支援あおぞら","前原西","船橋市前原西2-11-5","047-470-4420",""),
    ("B","ケアプラン大穴","大穴北","船橋市大穴北7-22-1","047-456-7899",""),
    ("B","あゆみケアプランセンター","大穴南","船橋市大穴南5-1-30","047-401-1021",""),
    ("B","タカサケアサポート船橋","三咲","船橋市三咲2-9-37","047-407-8830",""),
    ("B","三咲在宅介護センター","三咲","船橋市三咲5-32-10","047-448-8511",""),
    ("B","みさき在宅支援センター南生苑","三咲","船橋市三咲4-1-11","047-449-3762","施設併設"),
    ("B","Care Management Office ノーブル","二和東","船橋市二和東5-48-20","047-401-0537",""),
    ("B","ケアプランかりん","二和東","船橋市二和東5-22-16","047-449-2020",""),
    ("B","ケアプランひだまりの丘","二和東","船橋市二和東6-17-3","047-404-1151",""),
    ("B","ケアプランセンター和みの城","松が丘","船橋市松が丘4-18-7","047-456-5855",""),
    ("B","ケアサービス船橋介護支援センター","滝台","船橋市滝台1-12-37","047-465-7712",""),
    ("B","介護屋みらい船橋店","滝台町","船橋市滝台町107-42","047-404-2660",""),
    ("B","ケアプラン八木が谷","咲が丘","船橋市咲が丘3-11-4","047-448-6301",""),
    ("B","ケアプラン初恋","咲が丘","船橋市咲が丘2-18-20","047-499-2809",""),
    ("B","ケアパートナー飯山満","芝山","船橋市芝山3-2-1","047-461-3330",""),
    ("B","ケアプラン・オレンジガーデン","芝山","船橋市芝山7-41-2","047-461-0038",""),
    ("B","こひつじ船橋居宅介護支援センター","夏見台","船橋市夏見台4-8-10","047-406-8777",""),
    ("B","ファミリアケアプランセンター","前貝塚町","船橋市前貝塚町1007-8","047-401-3961",""),
    ("C","船橋総合病院 居宅","北本町","船橋市北本町1-13-1","047-425-1151","病院・地域連携室=自力(本部紹介なし)"),
    ("C","居宅介護支援事業所 星の子船橋","三山","船橋市三山8-29-11","047-473-5356","南部・習志野寄り"),
]

KPI = [
    ["1 訪問","訪問数（居宅＋施設＋病院）","300件/月","30件/日（板橋:1日25-40）。9-11・16-18時／雨天はケアマネ在席で◎","tab21／板橋"],
    ["2 接触","接触率（キーマンに会えた率）","50%","在席時間を狙う。受付で“主任ケアマネ”を確認","tab21／板橋"],
    ["・居宅巡回","キーマンケアマネ（約40事業所）","2週で1巡","40件を2週で回る多忙CMの在席に合わせる","板橋"],
    ["・施設/病院","自力開拓の接触率（※本部は紹介しない）","別管理","病院は地域連携室・MSW、施設は相談員へ","板橋"],
    ["3 問合せ","新規問合せ率（接触→問合せ）","5%→9%→13%","〜3ヶ月5/〜9ヶ月9/成熟13。露出反復で上昇","tab21"],
    ["4 個別相談","具体ケースの相談化","—","“同意書取得・医師連携を代行”でCMの手間を減らす","施策"],
    ["5 CV","利用者CV率（同意書取得込）","55%","相談→利用者化。同意書のエリア通過率に注意","tab21"],
    ["6 依頼/継続","送客（依頼）・解約4%/月","ストック積上","月1の経過報告訪問で継続。単価3.8万・平均継続25ヶ月","tab21"],
]

GRIP = [
    "◆ 会える時間に合わせる：9-11時／16-18時（訪問の前後で在席）。雨の日はケアマネが事務所に居る＝狙い目。",
    "◆ キーマン特定：1事業所に複数CM。困難ケース/在宅復帰/看取りを多く持つCMほど送客が出る。受付で“主任ケアマネ”を聞く。",
    "◆ 若手の印象づけ（差別化3点）：①レス爆速（当日中に訪問可否を返す）②報告の丁寧さ（施術後フィードバック1枚）③清潔感・礼儀。『若いのにちゃんとしてる』を作る。",
    "◆ 手土産（潤滑油・コンプラ配慮）：高額・個人への金品はNG。“事業所全体で使える消耗品”が無難＝個包装菓子／実用ノベルティ（付箋・ボールペン）／夏は冷感タオル。過度な利益供与と取られない範囲で。",
    "◆ 送客が出る入口：『同意書取得・医師連携をこちらが代行』『初回訪問の報告書を丁寧に出す』＝ケアマネの手間を減らす＝また頼みたくなる。",
    "◆ 切らさない継続：①月1定期訪問（ご利用者の経過報告を持参）②サービス担当者会議に積極参加③CMの困りごとを1つ解決（他職種紹介）④紹介御礼は“結果報告”で返す（金品でなく信頼）。",
    "◆ NG（信頼を失う）：売り込み一辺倒／CMへの金品／利用者の囲い込み感。【中核信条＝信頼の対価で送客を得る】",
]

ROLES = [
    ["営業（菊池/施術者兼務 初期）","居宅・施設・病院への訪問と関係構築","台帳22の更新（接触日/グリップ段階/送客）。9-11・16-18時に訪問を寄せる。週次ログ23入力。"],
    ["事務","送客を捌く一次対応・記録・同意書段取り","問合せ→当日訪問可否の連絡→同意書の病院確認・書類管理→レセプト準備。報告書のフォーマット管理。"],
    ["施術者","訪問品質・継続の信頼づくり","初回訪問の丁寧な報告書1枚／サービス担当者会議に参加／利用者の経過をCMへ月1報告（送客継続の生命線）。"],
]

TALK = [
    "【受付突破】お世話になります。訪問鍼灸マッサージのオーロラ高根店です。ケアマネジャーさん（主任さん）に5分だけご挨拶させてください。",
    "【趣旨】高根に開業予定で、保険適用の訪問マッサージです。先生の同意書取得や医師連携はこちらで代行し、初回の報告書も丁寧にお出しします。先生方のお手間を増やさない形でお役に立てればと。",
    "【GIVE】困難ケースや在宅復帰で“体を動かす支援”が要る方がいたら、まず無料で相談に乗ります。合わなければ断ってもらって大丈夫です。",
    "【クロージング】また顔出しさせてください。経過はきちんと報告します。（名刺＋事業所で使える消耗品の手土産）",
]


def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request());
        with open(TOKEN, "wb") as f: pickle.dump(c, f)
    return c

def add_sheet(svc, title, rows, cols, frozen_rows=0, frozen_cols=0):
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":[{"addSheet":{"properties":{
        "title":title,"gridProperties":{"rowCount":rows,"columnCount":cols,
        "frozenRowCount":frozen_rows,"frozenColumnCount":frozen_cols}}}}]}).execute()
    return gid_of(svc, title)

def gid_of(svc, title):
    info=svc.spreadsheets().get(spreadsheetId=SID,fields="sheets.properties").execute()
    return [s["properties"]["sheetId"] for s in info["sheets"] if s["properties"]["title"]==title][0]

def put(svc, title, a1, values):
    svc.spreadsheets().values().update(spreadsheetId=SID, range=f"'{title}'!{a1}",
        valueInputOption="USER_ENTERED", body={"values":values}).execute()

def colwidths(gid, ws):
    return [{"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},
        "properties":{"pixelSize":w},"fields":"pixelSize"}} for i,w in enumerate(ws)]
def mergerow(gid,r,ncol):
    return {"mergeCells":{"range":{"sheetId":gid,"startRowIndex":r,"endRowIndex":r+1,"startColumnIndex":0,"endColumnIndex":ncol},"mergeType":"MERGE_ALL"}}
def fmtrow(gid,r,bg,fg,bold,size,wrap=True,h="LEFT"):
    return {"repeatCell":{"range":{"sheetId":gid,"startRowIndex":r,"endRowIndex":r+1},
        "cell":{"userEnteredFormat":{"backgroundColor":bg,"textFormat":{"foregroundColor":fg,"bold":bold,"fontSize":size},
            "wrapStrategy":"WRAP" if wrap else "OVERFLOW_CELL","verticalAlignment":"MIDDLE","horizontalAlignment":h}},"fields":"userEnteredFormat"}}
def rowheight(gid,r,px,r2=None):
    return {"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":r,"endIndex":(r2 or r)+1},"properties":{"pixelSize":px},"fields":"pixelSize"}}


# ───────────────────────── 22 営業台帳 ─────────────────────────
def build_list(svc):
    HDR=["#","Tier","事業所名","エリア","住所","TEL","担当/主任CM","初回接触","最終接触","グリップ段階","訪問回数","送客累計","直近送客日","次アクション/メモ"]
    NCOL=len(HDR)
    BC=3  # バナーは固定列(A-C)外のD列から結合
    rows=[["","","","営業台帳｜居宅介護支援事業所（船橋・高根店）　★1行=1事業所。接触したら更新。KPIは『23_営業KPIダッシュボード』へ自動連動"]+[""]*(NCOL-4),
          ["","","","使い方=営業が毎日更新：接触したら日付、関係の段階を『グリップ段階▼』で選択（①未接触→⑤継続送客）、送客が出たら件数を加算。Tier A(薄緑)=高根店至近を優先。施設/病院は本部紹介なし=自力(C)。"]+[""]*(NCOL-4),
          HDR]
    hdr_idx=2; ds=3
    for i,(t,n,a,ad,tel,memo) in enumerate(JIGYO,1):
        rows.append([str(i),t,n,a,ad,tel,"","","","",""," ",memo and "" or "",memo or ""])
    # メモは末尾列に。memoがある行は次アクション欄に注記として入れる
    for i,(t,n,a,ad,tel,memo) in enumerate(JIGYO):
        rows[ds+i][13]=memo
        rows[ds+i][9]=STAGES[0]  # 初期=①未接触
    de=ds+len(JIGYO)
    gid=add_sheet(svc, TAB_LIST, de+4, NCOL, frozen_rows=hdr_idx+1, frozen_cols=3)
    put(svc, TAB_LIST, "A1", rows)
    def merge_from(gid,r,c0,ncol):
        return {"mergeCells":{"range":{"sheetId":gid,"startRowIndex":r,"endRowIndex":r+1,"startColumnIndex":c0,"endColumnIndex":ncol},"mergeType":"MERGE_ALL"}}
    reqs=colwidths(gid,[30,42,224,78,188,116,150,80,80,118,72,70,96,210])
    reqs.append(merge_from(gid,0,BC,NCOL)); reqs.append(fmtrow(gid,0,REDD,WHT,True,13)); reqs.append(rowheight(gid,0,40))
    reqs.append(merge_from(gid,1,BC,NCOL)); reqs.append(fmtrow(gid,1,REDBG,REDD,True,9)); reqs.append(rowheight(gid,1,46))
    reqs.append(fmtrow(gid,hdr_idx,RED,WHT,True,10,True,"CENTER")); reqs.append(rowheight(gid,hdr_idx,38))
    # 本体
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ds,"endRowIndex":de},
        "cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE","textFormat":{"fontSize":10,"foregroundColor":INK}}},
        "fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    for i,(t,*_ ) in enumerate(JIGYO):
        bg=TIERA if t=="A" else (CARD if i%2 else WHT)
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ds+i,"endRowIndex":ds+i+1},
            "cell":{"userEnteredFormat":{"backgroundColor":bg}},"fields":"userEnteredFormat.backgroundColor"}})
    # 入力欄(G..M=6..12)薄黄
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ds,"endRowIndex":de,"startColumnIndex":6,"endColumnIndex":13},
        "cell":{"userEnteredFormat":{"backgroundColor":INBG}},"fields":"userEnteredFormat.backgroundColor"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ds,"endRowIndex":de,"startColumnIndex":0,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"fontSize":10}}},"fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ds,"endRowIndex":de,"startColumnIndex":2,"endColumnIndex":3},
        "cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
    # グリップ段階(J=9)にプルダウン
    reqs.append({"setDataValidation":{"range":{"sheetId":gid,"startRowIndex":ds,"endRowIndex":de,"startColumnIndex":9,"endColumnIndex":10},
        "rule":{"condition":{"type":"ONE_OF_LIST","values":[{"userEnteredValue":s} for s in STAGES]},"showCustomUi":True,"strict":False}}})
    reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":hdr_idx,"endRowIndex":de,"startColumnIndex":0,"endColumnIndex":NCOL},
        "innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},
        "innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    reqs.append({"setBasicFilter":{"filter":{"range":{"sheetId":gid,"startRowIndex":hdr_idx,"endRowIndex":de,"startColumnIndex":0,"endColumnIndex":NCOL}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":reqs}).execute()
    return gid, (ds+1, de)  # 1-idx data range


# ───────────────────────── 23 ダッシュボード ─────────────────────────
def build_dash(svc, list_data_range):
    li, lj = list_data_range
    NCOL=8
    rows=[]
    def add(r): rows.append(r+[""]*(NCOL-len(r)))
    add(["営業KPIダッシュボード（計画⇔実績 連動）｜高根店"]);
    add(["毎週『週次活動ログ』を入力→合計が自動集計→tab21の計画と対比。施設病院は本部紹介なし=接触率別管理。狩野さん確認列に◎/△/×を入れてもらう。"])
    add([""])
    add(["■ 週次活動ログ（毎週入力）"]); wlog_sec=len(rows)-1
    add(["週","訪問数","接触数","新規問合せ","個別相談","CV(新規利用者)","送客依頼","メモ"]); wlog_hdr=len(rows)-1
    w_start=len(rows)
    for w in range(1,9): add([f"W{w}"])
    w_end=len(rows)
    a,b=w_start+1,w_end
    add(["合計",f"=SUM(B{a}:B{b})",f"=SUM(C{a}:C{b})",f"=SUM(D{a}:D{b})",f"=SUM(E{a}:E{b})",f"=SUM(F{a}:F{b})",f"=SUM(G{a}:G{b})",""])
    wtot=len(rows)  # 1-idx
    add([""])
    add(["■ 計画⇔実績 対比（ファネル）"]); cmp_sec=len(rows)-1
    add(["段階","計画/月(tab21)","実績(集計)","達成率","狩野さん確認 ◎/△/×"]); cmp_hdr=len(rows)-1
    cs=len(rows)
    add(["訪問数（月）",300,f"=B{wtot}","",""]); r=len(rows); rows[r-1][3]=f"=IFERROR(TEXT(C{r}/B{r},\"0%\"),\"\")"
    add(["接触率","50%",f"=IFERROR(TEXT(C{wtot}/B{wtot},\"0%\"),\"\")","",""])
    add(["新規問合せ（月）",15,f"=D{wtot}","",""]); r=len(rows); rows[r-1][3]=f"=IFERROR(TEXT(C{r}/B{r},\"0%\"),\"\")"
    add(["個別相談（月）","—",f"=E{wtot}","",""])
    add(["CV＝新規利用者（月）","4.1→10.7",f"=F{wtot}","",""])
    add(["送客依頼（継続）","—",f"=G{wtot}","",""])
    ce=len(rows)
    add([""])
    add(["■ グリップ段階ストック（台帳22のJ列を自動集計＝関係構築の進捗）"]); st_sec=len(rows)-1
    add(STAGES+["未記入","計"]) ; st_hdr=len(rows)-1
    cnt=[f"=COUNTIF('{TAB_LIST}'!J{li}:J{lj},\"{s}\")" for s in STAGES]
    cnt.append(f"=COUNTBLANK('{TAB_LIST}'!J{li}:J{lj})")
    cnt.append(f"=COUNTA('{TAB_LIST}'!C{li}:C{lj})")
    add(cnt); st_val=len(rows)-1
    add([""])
    add(["■ tab21計画前提（実測が溜まったらこの前提を上書き→計画を精緻化）"]); n_sec=len(rows)-1
    add(["接触率50%／問合せ率5→9→13%／CV55%(同意書込)／解約4%/月／単価3.8万／施術者1人22名上限／1日30訪問×10日=300/月。出典:21_営業ファネル。"])

    gid=add_sheet(svc, TAB_DASH, len(rows)+4, NCOL, frozen_rows=2)
    put(svc, TAB_DASH, "A1", rows)
    reqs=colwidths(gid,[150,118,110,110,100,120,100,220])
    reqs.append(mergerow(gid,0,NCOL)); reqs.append(fmtrow(gid,0,REDD,WHT,True,14)); reqs.append(rowheight(gid,0,40))
    reqs.append(mergerow(gid,1,NCOL)); reqs.append(fmtrow(gid,1,REDBG,REDD,True,9)); reqs.append(rowheight(gid,1,44))
    for s in (wlog_sec,cmp_sec,st_sec,n_sec):
        reqs.append(mergerow(gid,s,NCOL)); reqs.append(fmtrow(gid,s,SEC,WHT,True,11)); reqs.append(rowheight(gid,s,26))
    reqs.append(fmtrow(gid,wlog_hdr,RED,WHT,True,10,True,"CENTER"))
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":w_start-0,"endRowIndex":w_end,"startColumnIndex":1,"endColumnIndex":8},
        "cell":{"userEnteredFormat":{"backgroundColor":INBG,"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,horizontalAlignment)"}})
    reqs.append(fmtrow(gid,wtot-1,CARD,INK,True,10,False,"CENTER"))
    reqs.append(fmtrow(gid,cmp_hdr,RED,WHT,True,10,True,"CENTER"))
    for i in range(cs,ce):
        reqs.append(fmtrow(gid,i,CARD if (i-cs)%2 else WHT,INK,False,10,True,"LEFT")); reqs.append(rowheight(gid,i,28))
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":cs,"endRowIndex":ce,"startColumnIndex":0,"endColumnIndex":1},
        "cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10}}},"fields":"userEnteredFormat.textFormat"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":cs,"endRowIndex":ce,"startColumnIndex":4,"endColumnIndex":5},
        "cell":{"userEnteredFormat":{"backgroundColor":INBG}},"fields":"userEnteredFormat.backgroundColor"}})
    reqs.append(fmtrow(gid,st_hdr,SEC,WHT,True,10,True,"CENTER"))
    reqs.append(fmtrow(gid,st_val,TIERA,INK,True,12,False,"CENTER")); reqs.append(rowheight(gid,st_val,30))
    reqs.append(mergerow(gid,n_sec+1,NCOL)); reqs.append(fmtrow(gid,n_sec+1,WHT,INK,False,9)); reqs.append(rowheight(gid,n_sec+1,40))
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":reqs}).execute()
    return gid


# ───────────────────────── 24 プレイブック ─────────────────────────
def build_pb(svc):
    NCOL=3
    rows=[]
    def add(r): rows.append(r+[""]*(NCOL-len(r)))
    add(["営業プレイブック｜ケアマネをグリップして“送客を切らさない”（全員必読・新人/事務/施術者も）"])
    add(["中核信条＝売り込まない。信頼の対価で送客を得る。相手(CM)の手間を減らすGIVEが先。"])
    add([""])
    add(["■ グリップ施策（送客を切らさない）"]); g_sec=len(rows)-1
    g_start=len(rows)
    for g in GRIP: add([g])
    g_end=len(rows)
    add([""])
    add(["■ 役割分担（営業・事務・施術者）"]); r_sec=len(rows)-1
    add(["役割","ミッション","具体アクション"]); r_hdr=len(rows)-1
    r_start=len(rows)
    for role in ROLES: add(role)
    r_end=len(rows)
    add([""])
    add(["■ 初回トーク雛形（受付突破→趣旨→GIVE→クロージング）"]); t_sec=len(rows)-1
    t_start=len(rows)
    for t in TALK: add([t])
    t_end=len(rows)
    add([""])
    add(["■ KPIの先行/遅行（迷ったらここを見る）"]); k_sec=len(rows)-1
    add(["先行指標＝訪問数・接触数（今日動かせる）。まずここを毎日埋める。"])
    add(["遅行指標＝新規利用者・売上（2-3ヶ月後の結果）。先行が積めば後から付いてくる。一喜一憂しない。"])

    gid=add_sheet(svc, TAB_PB, len(rows)+4, NCOL, frozen_rows=2)
    put(svc, TAB_PB, "A1", rows)
    reqs=colwidths(gid,[230,300,470])
    reqs.append(mergerow(gid,0,NCOL)); reqs.append(fmtrow(gid,0,REDD,WHT,True,14)); reqs.append(rowheight(gid,0,40))
    reqs.append(mergerow(gid,1,NCOL)); reqs.append(fmtrow(gid,1,REDBG,REDD,True,10)); reqs.append(rowheight(gid,1,30))
    for s in (g_sec,r_sec,t_sec,k_sec):
        reqs.append(mergerow(gid,s,NCOL)); reqs.append(fmtrow(gid,s,SEC,WHT,True,11)); reqs.append(rowheight(gid,s,26))
    for i in range(g_start,g_end):
        reqs.append(mergerow(gid,i,NCOL)); reqs.append(fmtrow(gid,i,CARD if (i-g_start)%2 else WHT,INK,False,10)); reqs.append(rowheight(gid,i,34))
    reqs.append(fmtrow(gid,r_hdr,RED,WHT,True,10,True,"CENTER"))
    for i in range(r_start,r_end):
        reqs.append(fmtrow(gid,i,CARD if (i-r_start)%2 else WHT,INK,False,10)); reqs.append(rowheight(gid,i,52))
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":r_start,"endRowIndex":r_end,"startColumnIndex":0,"endColumnIndex":1},
        "cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10}}},"fields":"userEnteredFormat.textFormat"}})
    for i in range(t_start,t_end):
        reqs.append(mergerow(gid,i,NCOL)); reqs.append(fmtrow(gid,i,BLUEBG if (i-t_start)%2 else WHT,INK,False,10)); reqs.append(rowheight(gid,i,40))
    for i in range(k_sec+1,len(rows)):
        reqs.append(mergerow(gid,i,NCOL)); reqs.append(fmtrow(gid,i,WHT,INK,False,10)); reqs.append(rowheight(gid,i,30))
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests":reqs}).execute()
    return gid


def main():
    svc = build("sheets","v4",credentials=creds(),cache_discovery=False)
    info=svc.spreadsheets().get(spreadsheetId=SID,fields="sheets.properties").execute()
    existing={s["properties"]["title"]:s["properties"]["sheetId"] for s in info["sheets"]}
    for t in OLD_TABS:
        if t in existing:
            svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":[{"deleteSheet":{"sheetId":existing[t]}}]}).execute()
    g1,lr=build_list(svc)
    g2=build_dash(svc,lr)
    g3=build_pb(svc)
    base=f"https://docs.google.com/spreadsheets/d/{SID}/edit#gid="
    print("DONE")
    print("22_営業台帳_居宅:", base+str(g1))
    print("23_KPIダッシュボード:", base+str(g2))
    print("24_プレイブック:", base+str(g3))

if __name__ == "__main__":
    main()
