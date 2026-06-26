# -*- coding: utf-8 -*-
"""
オーロラ大元スプシに「22_船橋居宅ケアマネ営業(6_4)」タブを新設。
船橋市(高根店territory)の居宅介護支援事業所リスト＋ケアマネ・グリップ施策
＋新規営業ファネルKPI(tab21連動・板橋オーナー基準)を1枚に統合。
データ出典: ハートページナビ船橋市版(2024/7) https://www.heartpage.jp/funabashi/list?type=in_home
auth: scripts/sheets_token.pickle (scope=spreadsheets)
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID = "18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB = "22_船橋居宅ケアマネ営業(6_4)"
NCOL = 12

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

# ── 居宅事業所（Tier, 事業所名, エリア, 住所, TEL, メモ）──
# Tier A=高根店至近(高根台/習志野台/薬円台/北習志野) B=近接(前原/大穴/三咲/二和/松が丘/滝台/咲が丘/飯山満/夏見) C=施設病院(自力開拓)
JIGYO = [
    ("A","天空ケアプラニングセンター","高根台","船橋市高根台7-32-15","047-460-9544",""),
    ("A","ケアプランらぼ","高根台","船橋市高根台1-11-7","047-407-2556",""),
    ("A","朝日ケアコンサルタント テレサ会 船橋","高根台","船橋市高根台3-15-5","047-469-3128",""),
    ("A","東船橋病院 居宅介護支援","高根台","船橋市高根台4-29-1","047-401-5051","※病院併設=大箱・連携室にも接触"),
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
    ("C","船橋総合病院 居宅","北本町","船橋市北本町1-13-1","047-425-1151","※病院・地域連携室=自力開拓(本部紹介なし)"),
    ("C","居宅介護支援事業所 星の子船橋","三山","船橋市三山8-29-11","047-473-5356","南部・習志野寄り"),
]
LIST_HDR = ["#","Tier","事業所名","エリア","住所","TEL","CM数(要確認)","初回接触","グリップ段階","手土産/施策メモ","直近送客(件)","次アクション"]

KPI_HDR = ["段階","指標","月間目標","日次・時間帯/補足","根拠"]
KPI = [
    ["1 訪問","訪問数（居宅＋施設＋病院）","300件/月","30件/日（板橋:1日25-40）。9-11・16-18時／雨天はケアマネ在席で◎","tab21／板橋"],
    ["2 接触","接触率（キーマンに会えた率）","50%","在席時間を狙う。受付で“主任ケアマネ”を確認","tab21／板橋"],
    ["・居宅巡回","キーマンケアマネ（約40事業所）","2週で1巡","40件を2週で回る多忙CMの在席に合わせる(=こちらが合わせる)","板橋"],
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

GRIP_STAGE_LEGEND = "グリップ段階の凡例：①未接触 → ②挨拶/名刺 → ③CMと面談 → ④初回送客 → ⑤継続送客(月次)。各行のI列に記入。"


def main():
    svc = build("sheets", "v4", credentials=creds(), cache_discovery=False)
    info = svc.spreadsheets().get(spreadsheetId=SID, fields="sheets.properties").execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in info["sheets"]}
    if TAB in existing:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={
            "requests": [{"deleteSheet": {"sheetId": existing[TAB]}}]}).execute()

    # ── 値を段階的に積む。各セクションの行indexを記録 ──
    values = []
    def add(row): values.append(row + [""]*(NCOL-len(row)))
    fmt = {"title":[], "note":[], "listhdr":0, "list":(0,0), "tierA":[], "tierC":[],
           "kpihdr":0, "kpi":(0,0), "griphdr":0, "grip":(0,0), "sechdr":[]}

    add(["船橋市 居宅ケアマネ営業 ｜ オーロラ高根店 送客チャネル開拓（2026-06-04）"])
    fmt["title"].append(len(values)-1)
    add(["★指標は tab21『事業計画叩き台_営業ファネル(6_4)』と連動。板橋オーナー基準＝1日25-40訪問／会えるのは9-11・16-18時／雨の日／キーマンCM約40事業所を2週で巡回。施設・病院は本部紹介なし=自力開拓。" + GRIP_STAGE_LEGEND])
    fmt["note"].append(len(values)-1)
    add([""])

    # ── 事業所一覧 ──
    add(["■ 船橋市 居宅介護支援事業所リスト（Tier=高根店からの近さ。出典:ハートページ2024/7・全166件から高根周辺40件を抽出）"])
    fmt["sechdr"].append(len(values)-1)
    add(LIST_HDR); fmt["listhdr"] = len(values)-1
    list_start = len(values)
    for i,(tier,name,area,addr,tel,memo) in enumerate(JIGYO,1):
        add([str(i),tier,name,area,addr,tel,"","","",memo,"",""])
    list_end = len(values)
    fmt["list"] = (list_start, list_end)
    add([""])

    # ── 営業ファネルKPI ──
    add(["■ 新規営業ファネルKPI（tab21連動・板橋オーナー基準で再設計＝6_KPIツリーの実戦版）"])
    fmt["sechdr"].append(len(values)-1)
    add(KPI_HDR); fmt["kpihdr"] = len(values)-1
    k_start = len(values)
    for r in KPI: add(r)
    fmt["kpi"] = (k_start, len(values))
    add([""])

    # ── 計画⇔実績 連動レイヤー（tab21と同じ定義で集計→計画対比） ──
    add(["■ 計画⇔実績 連動（営業集計＝tab21ファネルと同じ定義。毎週ここを埋める→計画と自動対比。これを狩野さんに確認依頼）"])
    fmt["sechdr"].append(len(values)-1)
    # 週次活動ログ（入力）
    add(["週次活動ログ（入力）","訪問数","接触数","新規問合せ","個別相談","CV(新規利用者)","送客依頼","メモ"])
    fmt["wloghdr"] = len(values)-1
    w_start = len(values)            # 0-idx 最初のWeek行
    for w in range(1,9): add([f"W{w}"])
    w_end = len(values)             # exclusive 0-idx
    a, b = w_start+1, w_end          # 1-idx 範囲
    add(["合計", f"=SUM(B{a}:B{b})", f"=SUM(C{a}:C{b})", f"=SUM(D{a}:D{b})",
         f"=SUM(E{a}:E{b})", f"=SUM(F{a}:F{b})", f"=SUM(G{a}:G{b})", ""])
    wtot = len(values)               # 1-idx 合計行
    fmt["wlog"] = (w_start, w_end)
    fmt["wtot"] = wtot-1             # 0-idx
    add([""])
    # 計画⇔実績 対比
    add(["計画⇔実績 対比（ファネル）","計画/月(tab21)","実績(集計)","達成率","狩野さん確認 ◎/△/×"])
    fmt["cmphdr"] = len(values)-1
    c_start = len(values)
    def cmprow(stage, plan, real_formula, rate=True):
        r = len(values)+1            # この対比行の1-idx
        rate_f = f"=IF(N({chr(67)}{r})=0,\"\",TEXT(C{r}/B{r},\"0%\"))" if rate else ""
        add([stage, plan, real_formula, rate_f, ""])
    cmprow("訪問数（月）", 300, f"=B{wtot}")
    # 接触率は率なので 実績=接触/訪問
    r_now = len(values)+1
    add(["接触率", "50%", f"=IFERROR(TEXT(C{wtot}/B{wtot},\"0%\"),\"\")", "", ""])
    cmprow("新規問合せ（月）", 15, f"=D{wtot}")
    add(["個別相談（月）", "—", f"=E{wtot}", "", ""])
    add(["CV＝新規利用者（月）", "4.1→10.7", f"=F{wtot}", "", ""])
    add(["送客依頼（継続）", "—", f"=G{wtot}", "", ""])
    c_end = len(values)
    fmt["cmp"] = (c_start, c_end)
    add([""])
    # グリップ段階ストック（リストI列を自動集計）
    add(["グリップ段階ストック（リストのI列を自動集計）"])
    fmt["sechdr"].append(len(values)-1)
    add(["①未接触","②挨拶/名刺","③面談","④初送客","⑤継続送客","未記入"])
    fmt["stockhdr"] = len(values)-1
    li, lj = list_start+1, list_end   # 1-idx リストデータ範囲
    add([f"=COUNTIF(I{li}:I{lj},\"①*\")", f"=COUNTIF(I{li}:I{lj},\"②*\")",
         f"=COUNTIF(I{li}:I{lj},\"③*\")", f"=COUNTIF(I{li}:I{lj},\"④*\")",
         f"=COUNTIF(I{li}:I{lj},\"⑤*\")", f"=COUNTBLANK(I{li}:I{lj})"])
    fmt["stock"] = len(values)-1
    add([""])

    # ── グリップ施策 ──
    add(["■ ケアマネをグリップして“送客を切らさない”施策（手土産・若手の印象づけ・継続）"])
    fmt["sechdr"].append(len(values)-1)
    g_start = len(values)
    for g in GRIP: add([g])
    fmt["grip"] = (g_start, len(values))

    nrows = len(values)
    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={
        "requests": [{"addSheet": {"properties": {
            "title": TAB,
            "gridProperties": {"rowCount": nrows+5, "columnCount": NCOL,
                               "frozenRowCount": fmt["listhdr"]+1}}}}]}).execute()
    gid = [s["properties"]["sheetId"] for s in
           svc.spreadsheets().get(spreadsheetId=SID,fields="sheets.properties").execute()["sheets"]
           if s["properties"]["title"]==TAB][0]

    svc.spreadsheets().values().update(spreadsheetId=SID, range=f"'{TAB}'!A1",
        valueInputOption="USER_ENTERED", body={"values": values}).execute()

    RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
    CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
    WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
    INBG={"red":0.999,"green":0.984,"blue":0.882}
    TIERA={"red":0.886,"green":0.949,"blue":0.905}  # 薄緑=至近
    SEC={"red":0.357,"green":0.357,"blue":0.357}
    reqs=[]
    widths=[30,42,230,80,210,118,86,78,150,210,90,180]
    for i,w in enumerate(widths):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},
            "properties":{"pixelSize":w},"fields":"pixelSize"}})
    def merge(r):
        reqs.append({"mergeCells":{"range":{"sheetId":gid,"startRowIndex":r,"endRowIndex":r+1,
            "startColumnIndex":0,"endColumnIndex":NCOL},"mergeType":"MERGE_ALL"}})
    def cell(r,bg,fg,bold,size,wrap=True,h="LEFT"):
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":r,"endRowIndex":r+1},
            "cell":{"userEnteredFormat":{"backgroundColor":bg,"textFormat":{"foregroundColor":fg,"bold":bold,"fontSize":size},
                "wrapStrategy":"WRAP" if wrap else "OVERFLOW_CELL","verticalAlignment":"MIDDLE","horizontalAlignment":h}},
            "fields":"userEnteredFormat"}})
    def rowh(r,px):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":gid,"dimension":"ROWS","startIndex":r,"endIndex":r+1},
            "properties":{"pixelSize":px},"fields":"pixelSize"}})
    # タイトル
    for r in fmt["title"]:
        merge(r); cell(r,REDD,WHT,True,14); rowh(r,40)
    # note
    for r in fmt["note"]:
        merge(r); cell(r,REDBG,REDD,True,9); rowh(r,58)
    # セクション見出し
    for r in fmt["sechdr"]:
        merge(r); cell(r,SEC,WHT,True,11); rowh(r,28)
    # リストヘッダ
    cell(fmt["listhdr"],RED,WHT,True,10,True,"CENTER")
    # リスト本体：交互色＋Tier A薄緑＋入力欄(H〜L)薄黄
    ls,le=fmt["list"]
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ls,"endRowIndex":le},
        "cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE",
            "textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    for idx in range(ls,le):
        tier=JIGYO[idx-ls][0]
        bg=TIERA if tier=="A" else (CARD if (idx-ls)%2 else WHT)
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":idx,"endRowIndex":idx+1},
            "cell":{"userEnteredFormat":{"backgroundColor":bg}},"fields":"userEnteredFormat.backgroundColor"}})
    # 入力欄 H,I,J(memoは既値),K,L を薄黄 → H(7),I(8),K(10),L(11)
    for col in (7,8,10,11):
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ls,"endRowIndex":le,"startColumnIndex":col,"endColumnIndex":col+1},
            "cell":{"userEnteredFormat":{"backgroundColor":INBG}},"fields":"userEnteredFormat.backgroundColor"}})
    # Tier列・#列 中央
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ls,"endRowIndex":le,"startColumnIndex":0,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{"horizontalAlignment":"CENTER","textFormat":{"bold":True,"fontSize":10}}},
        "fields":"userEnteredFormat(horizontalAlignment,textFormat)"}})
    # 事業所名 太字
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ls,"endRowIndex":le,"startColumnIndex":2,"endColumnIndex":3},
        "cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat.textFormat"}})
    # KPIヘッダ＋本体
    cell(fmt["kpihdr"],RED,WHT,True,10,True,"CENTER")
    ks,ke=fmt["kpi"]
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ks,"endRowIndex":ke},
        "cell":{"userEnteredFormat":{"wrapStrategy":"WRAP","verticalAlignment":"MIDDLE","textFormat":{"fontSize":10,"foregroundColor":INK}}},
        "fields":"userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    for idx in range(ks,ke):
        bg=CARD if (idx-ks)%2 else WHT
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":idx,"endRowIndex":idx+1},
            "cell":{"userEnteredFormat":{"backgroundColor":bg}},"fields":"userEnteredFormat.backgroundColor"}})
        rowh(idx,34)
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ks,"endRowIndex":ke,"startColumnIndex":0,"endColumnIndex":1},
        "cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10}}},"fields":"userEnteredFormat.textFormat"}})
    # 連動：週次活動ログ
    cell(fmt["wloghdr"],RED,WHT,True,10,True,"CENTER")
    ws,we=fmt["wlog"]
    # 入力欄(B〜H)薄黄
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ws,"endRowIndex":we,"startColumnIndex":1,"endColumnIndex":8},
        "cell":{"userEnteredFormat":{"backgroundColor":INBG,"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,horizontalAlignment)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":ws,"endRowIndex":we,"startColumnIndex":0,"endColumnIndex":1},
        "cell":{"userEnteredFormat":{"textFormat":{"bold":True}}},"fields":"userEnteredFormat.textFormat"}})
    # 合計行
    cell(fmt["wtot"],CARD,INK,True,10,False,"CENTER")
    # 計画⇔実績 対比
    cell(fmt["cmphdr"],RED,WHT,True,10,True,"CENTER")
    cs,ce=fmt["cmp"]
    for idx in range(cs,ce):
        bg=CARD if (idx-cs)%2 else WHT
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":idx,"endRowIndex":idx+1},
            "cell":{"userEnteredFormat":{"backgroundColor":bg,"textFormat":{"fontSize":10,"foregroundColor":INK},"verticalAlignment":"MIDDLE"}},
            "fields":"userEnteredFormat(backgroundColor,textFormat,verticalAlignment)"}})
        rowh(idx,30)
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":cs,"endRowIndex":ce,"startColumnIndex":0,"endColumnIndex":1},
        "cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10}}},"fields":"userEnteredFormat.textFormat"}})
    # 狩野確認列(E)薄黄
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":cs,"endRowIndex":ce,"startColumnIndex":4,"endColumnIndex":5},
        "cell":{"userEnteredFormat":{"backgroundColor":INBG}},"fields":"userEnteredFormat.backgroundColor"}})
    # ストック
    cell(fmt["stockhdr"],SEC,WHT,True,10,True,"CENTER")
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":fmt["stock"],"endRowIndex":fmt["stock"]+1},
        "cell":{"userEnteredFormat":{"backgroundColor":TIERA,"textFormat":{"bold":True,"fontSize":11},"horizontalAlignment":"CENTER"}},
        "fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"}})
    # グリップ施策（全幅マージ・左寄せ折返し）
    gs,ge=fmt["grip"]
    for r in range(gs,ge):
        merge(r); cell(r,CARD if (r-gs)%2 else WHT,INK,False,10); rowh(r,34)
    # 枠線（リスト）
    reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":fmt["listhdr"],"endRowIndex":le,
        "startColumnIndex":0,"endColumnIndex":NCOL},
        "innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},
        "innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    # オートフィルタ（リスト）
    reqs.append({"setBasicFilter":{"filter":{"range":{"sheetId":gid,"startRowIndex":fmt["listhdr"],"endRowIndex":le,
        "startColumnIndex":0,"endColumnIndex":NCOL}}}})

    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": reqs}).execute()
    print("DONE tab=",TAB,"gid=",gid,"事業所=",len(JIGYO))
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__ == "__main__":
    main()
