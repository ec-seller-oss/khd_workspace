# -*- coding: utf-8 -*-
"""
オーロラ大元スプシ「医療不動産×オーロラFC_統合事業計画」に
「FC架電ヒアリング（相場軸）」タブを新設する。
先輩オーナー/競合FCへ一斉架電して、採用相場(7項目+移動手段)を
1社1行で即メモ・横比較できる相場マトリクス。
auth: scripts/sheets_token.pickle (scope=spreadsheets)
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID = "18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB = "11_オーロラ加盟店ヒアリング(6_1)"
OLD_TABS = ["11_FC架電ヒアリング(6_1)"]  # 旧・競合FC版があれば掃除

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

# ── ヘッダ（15列 A〜O）──
HDR = [
    "#", "社名／切り口", "連絡先・架電順",
    "架電結果\n(受付突破/本人/不在/フォーム送信)",
    "アポ可否・日時",
    "Q1 採用頻度\n(出して〜有資格1人採るまで 期間/応募数)",
    "Q2 コスト\n(無料/有料・1人あたり採用コスト)",
    "Q3 定着\n(離職率/歩合で食えるか)",
    "Q4 母数\n(あマ指・鍼灸師は集まる職種か)",
    "Q5 本部支援\n(具体に何を／採用支援70名の実態)",
    "Q6 初月の詰まり\n(利用者集め vs 採用 どっちが先)",
    "Q7 同意書\n(医師の同意書スムーズに取れたか)",
    "＋移動手段\n(チャリ困難→バイク/車の生産性影響)",
    "所感・Go/No-go寄与",
    "次アクション",
]

# ── 架電先（1社1行）──。F〜O列は空欄＝電話しながら手入力 ──
def row(n, name, contact):
    return [n, name, contact] + [""] * 12

ROWS = [
    # ── ①千葉県内（同意書・ケアマネ事情が船橋に最も近い＝最優先）──
    row("1",
        "市川浦安店【千葉・最優先】\n切り口:同じ千葉県。同意書/ケアマネ事情・採用相場が船橋にそのまま当てはまる",
        "千葉県浦安市北栄1-6-35\n☎070-9311-3392"),
    row("2",
        "千葉蘇我店【千葉】\n切り口:千葉市の県内相場。あマ指/鍼灸の母数感・同意書の取りやすさ",
        "千葉市中央区南町2-10-1\n☎080-9547-3899"),
    # ── ②事例店（YouTube黒字化／単価実績）──
    row("3",
        "板橋店【YouTube黒字化事例】\n切り口:3月開業→5月黒字化を拝見し連絡。立上げ初月の採用と客集めのリアル ※運営会社=要確認(洗心メディカル?)",
        "東京都板橋区大和町6-8\n☎080-7659-5531"),
    row("4",
        "名東店【単価4万円実績・資料P24】\n切り口:平均単価4万の作り方／セラピストが歩合で食えるまでの早さ",
        "名古屋市名東区高針台1-107\n☎090-8070-1147"),
    # ── ③郊外チャリ困難型（移動手段の仮説をそのまま検証）──
    row("5",
        "越谷店【埼玉郊外＝チャリ困難で船橋類似】\n切り口:坂・移動距離でバイク/車主体か。1日の訪問件数と生産性",
        "埼玉県越谷市神明町2-379\n☎090-5425-0001"),
    row("6",
        "町田店【坂多い郊外＝移動仮説検証】\n切り口:チャリで回れないエリアの代替モデル・1日何軒が現実的か",
        "東京都町田市南大谷5-12-12\n☎090-3844-4972"),
    row("7",
        "所沢店【埼玉郊外】\n切り口:郊外の移動手段別の稼働率。雨天時の稼働ロス",
        "埼玉県所沢市西新井町22-1\n☎070-6578-0608"),
    # ── ④船橋近接・都市型 ──
    row("8",
        "江戸川店【船橋に隣接】\n切り口:すぐ隣のエリア。既存業者との競合感・利用者の集めやすさ",
        "東京都江戸川区一之江5-8-8\n☎080-9664-2485"),
    row("9",
        "足立葛飾店【下町チャリ訪問】\n切り口:密集地のチャリ訪問の実態。1日の回転数",
        "東京都足立区青井1-16-7\n☎080-3025-4934"),
    row("10",
        "練馬店\n切り口:都市型チャリ訪問。採用にかかった期間・応募数",
        "東京都練馬区早宮1-33-8-203\n☎080-2141-3842"),
    # ── ⑤新店（立上げ初期の記憶が新しい）──
    row("11",
        "上野店【2025/1開業の新店】\n切り口:立上げ初月に詰まったのは“利用者集め”か“採用”か。記憶が新しい",
        "東京都台東区北上野1-5-3\n☎080-5839-4037"),
    row("12",
        "中野店【都市密集の移動効率】\n切り口:狭域高密度の回り方。本部の採用支援(70名実績)は具体に何を",
        "東京都中野区江古田4-29-6\n☎080-6679-3337"),
]

NOTE_TOP = ("📞 ヒアリング先＝オーロラの実加盟店オーナー（運営=株式会社フライハイト／代表 猪原健斗・全国約65店）。"
            "競合FCではなく同ブランドだから、本部支援・収益モデル・Indeed採用手法がそのまま比較できる＝相場軸として唯一意味がある。"
            "大義名分『加盟を前向きに検討中。船橋出店で“採用”が一番の不安。先輩オーナーに現場のリアルを10分だけ』。"
            "受付突破はYouTube『フランチャイズ探偵Q』等で事例拝見し連絡、で。"
            "★船橋に効く順=①千葉県内(同意書/ケアマネ事情が同じ)②郊外チャリ困難型(移動仮説の検証)。"
            "聞いた数字→自分のIndeed反響(週IMP約30=露出不足が最大ボトルネック)と突合→6/24 Go/No-go("
            "有資格の内定1名 or 相場と遜色ない反響→GO／両方ダメ→加盟金約590万は払わず保留)。")
NOTE_BOTTOM = ("【相場まとめ（架電後に記入）】　採用に要する期間/応募数の相場：__　／　1人あたり採用コスト相場：__"
               "（外注サングローブ＝180万は“上限”参照点）　／　あマ指・鍼灸師の母数感：__　／　同意書の取りやすさ：__"
               "　／　移動手段の結論(船橋はバイク/車前提か)：__　／　総合：Go・No-go寄与＝__")

def main():
    svc = build("sheets", "v4", credentials=creds(), cache_discovery=False)
    info = svc.spreadsheets().get(spreadsheetId=SID, fields="sheets.properties").execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in info["sheets"]}
    for t in [TAB] + OLD_TABS:
        if t in existing:
            svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={
                "requests": [{"deleteSheet": {"sheetId": existing[t]}}]}).execute()
    NCOL = 15
    res = svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={
        "requests": [{"addSheet": {"properties": {
            "title": TAB,
            "gridProperties": {"rowCount": 40, "columnCount": NCOL,
                               "frozenRowCount": 5, "frozenColumnCount": 2}}}}]}).execute()
    gid = res["replies"][0]["addSheet"]["properties"]["sheetId"]

    # バナー行は固定列(A,B)をまたげないので、テキストはC列(index2)に置きC〜O列をマージ
    BANNER_C = 2
    title_row = ["", "", "オーロラ実加盟店 架電ヒアリング（採用相場軸）｜ 株式会社フライハイト・全国約65店 ｜ 一斉架電 2026-06-01〜"] + [""] * (NCOL - 3)
    note_row = ["", "", NOTE_TOP] + [""] * (NCOL - 3)
    bottom_row = ["", "", NOTE_BOTTOM] + [""] * (NCOL - 3)
    values = [title_row, [""] * NCOL, note_row, [""] * NCOL, HDR] + ROWS + [[""] * NCOL, bottom_row]
    svc.spreadsheets().values().update(
        spreadsheetId=SID, range=f"'{TAB}'!A1",
        valueInputOption="USER_ENTERED", body={"values": values}).execute()

    n_rows = len(values)
    note_top_idx = 2
    hdr_idx = 4
    data_start = hdr_idx + 1
    data_end = data_start + len(ROWS)
    note_bottom_idx = n_rows - 1

    RED = {"red": 0.667, "green": 0.180, "blue": 0.149}
    REDD = {"red": 0.549, "green": 0.141, "blue": 0.114}
    CARD = {"red": 0.945, "green": 0.925, "blue": 0.882}
    REDBG = {"red": 0.957, "green": 0.894, "blue": 0.886}
    WHT = {"red": 1, "green": 1, "blue": 1}
    INK = {"red": 0.1, "green": 0.1, "blue": 0.1}
    # 入力欄（Q列）をうっすら黄色＝ここに手入力、と一目で分かるように
    INBG = {"red": 0.999, "green": 0.984, "blue": 0.882}

    reqs = []
    # 列幅
    widths = [34, 240, 180, 120, 120, 165, 160, 150, 160, 165, 150, 160, 165, 190, 150]
    for i, w in enumerate(widths):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": gid, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    # タイトル行
    reqs.append({"mergeCells": {"range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1,
        "startColumnIndex": BANNER_C, "endColumnIndex": NCOL}, "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1},
        "cell": {"userEnteredFormat": {"backgroundColor": REDD,
            "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 14},
            "verticalAlignment": "MIDDLE", "horizontalAlignment": "LEFT"}},
        "fields": "userEnteredFormat"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS",
        "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 44}, "fields": "pixelSize"}})
    # 注記行
    reqs.append({"mergeCells": {"range": {"sheetId": gid, "startRowIndex": note_top_idx, "endRowIndex": note_top_idx + 1,
        "startColumnIndex": BANNER_C, "endColumnIndex": NCOL}, "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": note_top_idx, "endRowIndex": note_top_idx + 1},
        "cell": {"userEnteredFormat": {"backgroundColor": REDBG,
            "textFormat": {"foregroundColor": REDD, "bold": True, "fontSize": 10},
            "wrapStrategy": "WRAP", "verticalAlignment": "MIDDLE"}},
        "fields": "userEnteredFormat"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS",
        "startIndex": note_top_idx, "endIndex": note_top_idx + 1}, "properties": {"pixelSize": 78}, "fields": "pixelSize"}})
    # ヘッダ行
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": hdr_idx, "endRowIndex": hdr_idx + 1},
        "cell": {"userEnteredFormat": {"backgroundColor": RED,
            "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 10},
            "wrapStrategy": "WRAP", "verticalAlignment": "MIDDLE", "horizontalAlignment": "CENTER"}},
        "fields": "userEnteredFormat"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS",
        "startIndex": hdr_idx, "endIndex": hdr_idx + 1}, "properties": {"pixelSize": 64}, "fields": "pixelSize"}})
    # データ行：折返し＋上寄せ＋交互色
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": data_start, "endRowIndex": data_end},
        "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP",
            "textFormat": {"fontSize": 10, "foregroundColor": INK}}},
        "fields": "userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    for k in range(len(ROWS)):
        r = data_start + k
        bg = CARD if k % 2 == 1 else WHT
        reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": r, "endRowIndex": r + 1},
            "cell": {"userEnteredFormat": {"backgroundColor": bg}}, "fields": "userEnteredFormat.backgroundColor"}})
    # 入力欄(F〜O = col5..14)を薄黄色で塗る＝ここに手入力
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": data_start, "endRowIndex": data_end,
        "startColumnIndex": 5, "endColumnIndex": NCOL},
        "cell": {"userEnteredFormat": {"backgroundColor": INBG}}, "fields": "userEnteredFormat.backgroundColor"}})
    # #列・社名列を強調
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": data_start, "endRowIndex": data_end,
        "startColumnIndex": 0, "endColumnIndex": 1},
        "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER",
            "textFormat": {"bold": True, "fontSize": 11}}},
        "fields": "userEnteredFormat(horizontalAlignment,textFormat)"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": data_start, "endRowIndex": data_end,
        "startColumnIndex": 1, "endColumnIndex": 2},
        "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 10, "foregroundColor": INK}}},
        "fields": "userEnteredFormat.textFormat"}})
    # データ行の高さ
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS",
        "startIndex": data_start, "endIndex": data_end}, "properties": {"pixelSize": 92}, "fields": "pixelSize"}})
    # 相場まとめ行
    reqs.append({"mergeCells": {"range": {"sheetId": gid, "startRowIndex": note_bottom_idx, "endRowIndex": note_bottom_idx + 1,
        "startColumnIndex": BANNER_C, "endColumnIndex": NCOL}, "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": note_bottom_idx, "endRowIndex": note_bottom_idx + 1},
        "cell": {"userEnteredFormat": {"backgroundColor": CARD,
            "textFormat": {"bold": True, "fontSize": 11, "foregroundColor": INK},
            "wrapStrategy": "WRAP", "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS",
        "startIndex": note_bottom_idx, "endIndex": note_bottom_idx + 1}, "properties": {"pixelSize": 54}, "fields": "pixelSize"}})
    # 枠線
    reqs.append({"updateBorders": {"range": {"sheetId": gid, "startRowIndex": hdr_idx, "endRowIndex": data_end,
        "startColumnIndex": 0, "endColumnIndex": NCOL},
        "innerHorizontal": {"style": "SOLID", "color": {"red": 0.85, "green": 0.83, "blue": 0.79}},
        "innerVertical": {"style": "SOLID", "color": {"red": 0.85, "green": 0.83, "blue": 0.79}}}})

    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": reqs}).execute()
    print("DONE tab=", TAB, "gid=", gid)
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__ == "__main__":
    main()
