# -*- coding: utf-8 -*-
"""
日本政策金融公庫「創業計画書」公式様式(kaigyou00_190507a)に忠実なGoogle Sheets版を作成。
オーロラFC船橋・埋まる部分は記入済、本部依存は［本部確認後に修正］。
Tab1=創業計画書(公式様式準拠)、Tab2=本部確認18点。
auth: scripts/sheets_token.pickle (scope=spreadsheets)
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
TITLE = "創業計画書_オーロラFC船橋_公庫様式_260531"
NCOL = 14

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

DOUKI = ("当社（KHD）は不動産物件調達および医療クリニックの承継コンサルティングを営み、医療機関・ケアマネジャーとのネットワークと"
         "事業運営基盤を有している。事業エリアの船橋市は後期高齢者が8万人を超え、在宅・施設での訪問マッサージ（医療保険適用のマッサージ・はり灸）"
         "需要が構造的に拡大している。既存事業で培った医療・介護分野の人的ネットワークを活かせる隣接領域であり、訪問医療マッサージFC「オーロラ」に"
         "加盟することで、集客・保険請求（レセプト）代行・運営ノウハウを本部から取得し、未経験リスクを抑えて確実に立ち上げられると判断し創業する。")
JIYU = ("・船橋市の後期高齢者8万人超という需要基盤と、本部territory承認済みの独占エリアで開業する。\n"
        "・既存のKHD医療・介護ネットワーク（ケアマネ・医療機関）を集客に活用できる。\n"
        "・加盟前にIndeedで採用テストを実施し、即戦力人材の応募を確認済み＝人員確保の蓋然性を実証している。\n"
        "・本部のレセプト代行により保険請求の未収リスクを抑制。慎重に採算を検証したうえで開業判断を行う。")

# 本部確認18点（Tab2）
CONFIRM = [
    ["A 設備資金", "加盟金（税込・支払時期）"],
    ["A 設備資金", "研修費（金額／期間／回数）"],
    ["A 設備資金", "保証金（有無・額・返還条件）"],
    ["A 設備資金", "開業時の初期実費（機材・備品・ユニフォーム・車両・サイネージ等）"],
    ["A 設備資金", "初期システム導入費"],
    ["A 設備資金", "支払いスケジュール（一括／分割・開業何日前）"],
    ["B 収支モデル", "訪問1回あたりの標準保険診療単価（マッサージ／はり灸）"],
    ["B 収支モデル", "施術者1人あたり標準モデル（1日訪問件数／月稼働日数／想定月商）"],
    ["B 収支モデル", "施術者の標準報酬体系（歩合 or 固定＋歩合）"],
    ["B 収支モデル", "ロイヤリティ13.2%の課金ベース・最低保証の有無"],
    ["B 収支モデル", "ロイヤリティ以外の継続費用（広告分担金・システム月額等）"],
    ["B 収支モデル", "レセプト請求→入金までのサイト（立替月数）"],
    ["C 裏付け", "本部の集客支援の範囲（ケアマネ営業の分担）"],
    ["C 裏付け", "レセプト代行の範囲・手数料（ロイヤリティ込みか別か）"],
    ["C 裏付け", "立上げカーブ／BEP到達までの平均月数"],
    ["D 補強材料", "既存店の実績数値（提出可能な範囲）"],
    ["D 補強材料", "territory（船橋市）の独占範囲・契約期間・更新条件"],
    ["D 補強材料", "中途解約時の違約金・競業避止条項"],
]

RED  = {"red": 0.667, "green": 0.180, "blue": 0.149}
REDD = {"red": 0.549, "green": 0.141, "blue": 0.114}
CARD = {"red": 0.945, "green": 0.925, "blue": 0.882}
REDBG= {"red": 0.957, "green": 0.894, "blue": 0.886}
WHT  = {"red": 1, "green": 1, "blue": 1}
INK  = {"red": 0.1, "green": 0.1, "blue": 0.1}

# 行ビルダー
V = []        # 値（14列）
SEC = []      # セクション見出し行
THDR = []     # 表ヘッダ行
MERGE = []    # (r1,r2,c1,c2)
BORDER = []   # (r1,r2,c1,c2)
HEIGHT = {}   # rowidx -> px

def row(*cells):
    r = list(cells) + [""] * (NCOL - len(cells))
    V.append(r[:NCOL]); return len(V) - 1
def sec(txt):
    i = row(txt); SEC.append(i); MERGE.append((i, i+1, 0, NCOL)); HEIGHT[i] = 24; return i
def m(r, c1, c2): MERGE.append((r, r+1, c1, c2))

# ── タイトル ──
t = row("創業計画書"); MERGE.append((t, t+1, 0, 8)); HEIGHT[t] = 38
V[t][8] = "〔令和　　年　　月　　日作成〕　お名前：菊池 研太"; MERGE.append((t, t+1, 8, NCOL))

# ── 1 創業の動機 ──
sec("１　創業の動機（創業されるのは、どのような目的、動機からですか。）")
r = row(DOUKI); m(r, 0, NCOL); HEIGHT[r] = 96

# ── 2 経営者の略歴等 ──
sec("２　経営者の略歴等（勤務先名だけでなく、担当業務・役職・身につけた技能等も記載してください。）")
h = row("年 月", "内 容"); THDR.append(h); m(h, 1, NCOL); b0 = h
r = row("現在", "KHD（kikuchi-hd.net）代表。不動産物件調達（仕入〜加工〜売却）・医療クリニック承継コンサル・EC（韓国輸出）を運営。複数事業の収支・資金繰りを管理。"); m(r, 1, NCOL); HEIGHT[r] = 54
r = row("［年月］", "［要記入：これまでの勤務先・担当業務・役職・身につけた技能］"); m(r, 1, NCOL)
r = row("［年月］", "［要記入］"); m(r, 1, NCOL); BORDER.append((b0, r+1, 0, NCOL))
r = row("過去の事業経験", "☑ 事業を経営していたことがあり、現在もその事業を続けている（事業内容：KHDで不動産・医療コンサル・EC）"); m(r, 1, NCOL)
r = row("取得資格", "☑ 特になし　／　□ 有（　番号等：　　）"); m(r, 1, NCOL)
r = row("許認可（許可・届出等）", "□ 特になし　／　☑ 有（あマ指師の施術所開設届・出張施術業務開始届 等 ※施術者側で要確認）　［要確認］"); m(r, 1, NCOL)
r = row("知的財産権等", "☑ 特になし"); m(r, 1, NCOL)

# ── 3 取扱商品・サービス ──
sec("３　取扱商品・サービス")
r = row("事業内容", "訪問マッサージ／訪問はり灸（医療保険適用、医師の同意書に基づく施術）。訪問医療マッサージFC「オーロラ」加盟。提供エリア：船橋市（本部territory承認済み）。"); m(r, 1, NCOL); HEIGHT[r] = 50
r = row("取扱商品・サービスの内容", "① 訪問マッサージ（売上シェア ［ ］％）"); m(r, 1, NCOL)
r = row("", "② 訪問はり灸（売上シェア ［ ］％）"); m(r, 1, NCOL)
r = row("", "③ ―（売上シェア ［ ］％）"); m(r, 1, NCOL)
r = row("客単価／受注単価", "［本部確認後に修正：訪問1回あたりの保険診療単価（マッサージ／はり灸）］"); m(r, 1, NCOL)
r = row("営業日数（月）", "［ ］日　／　営業時間 ［ ］〜［ ］"); m(r, 1, NCOL)
r = row("海外展開状況", "☑ 無"); m(r, 1, NCOL)
r = row("セールスポイント（自社の強み）", "本部ブランド＋レセプト代行で未収・査定リスクが低い／ケアマネ・施設営業網による安定集客／KHDの医療・介護ネットワークを活用。"); m(r, 1, NCOL); HEIGHT[r] = 50
r = row("販売ターゲット・販売戦略（集客方法）", "船橋市の在宅高齢者・施設入居者。ケアマネ・医療機関への営業、本部の集客支援、紹介ルートで獲得。"); m(r, 1, NCOL); HEIGHT[r] = 50
r = row("競合・市場など自社を取り巻く状況", "船橋市は後期高齢者8万人超で在宅医療マッサージ需要が拡大。本部territory承認により当該エリアで独占的に展開。"); m(r, 1, NCOL); HEIGHT[r] = 50

# ── 4 従業員 ──
sec("４　従業員")
r = row("常勤役員の人数（法人のみ）", "1 人"); m(r, 1, NCOL)
r = row("従業員数（3ヵ月以上継続雇用者）", "［ ］人　（うち家族従業員 ［ ］人 ／ うちパート従業員 ［ ］人）"); m(r, 1, NCOL)
r = row("採用見込み", "Indeed掲載により即戦力1名が応募済（53歳・あマ指師＋鍼灸師、訪問鍼灸マッサージ実務・ケアマネ営業・管理者経験あり）"); m(r, 1, NCOL); HEIGHT[r] = 50

# ── 5 取引先・取引関係等 ──
sec("５　取引先・取引関係等")
h = row("区分", "取引先名（フリガナ）", "", "", "所在地等（市区町村）", "", "シェア%", "掛取引%", "回収・支払の条件"); THDR.append(h)
m(h, 1, 4); m(h, 4, 6); m(h, 8, NCOL); b0 = h
r = row("販売先", "後期高齢者医療広域連合 等（保険者）", "", "", "—", "", "—", "—", "レセプト請求→入金まで約 ［本部確認］ ヶ月（窓口負担分は都度）"); m(r, 1, 4); m(r, 4, 6); m(r, 8, NCOL)
r = row("仕入先", "施術用消耗品ほか ［ ］", "", "", "［ ］", "", "—", "［ ］", "［締日・支払日を記入］"); m(r, 1, 4); m(r, 4, 6); m(r, 8, NCOL)
r = row("外注先", "オーロラ本部", "", "", "［本部所在地］", "", "—", "—", "ロイヤリティ 13.2%（税込）／毎月"); m(r, 1, 4); m(r, 4, 6); m(r, 8, NCOL); BORDER.append((b0, r+1, 0, NCOL))
r = row("人件費の支払", "［ ］日〆　［ ］日支払（ボーナス支給月 ［ ］月、［ ］月）"); m(r, 1, NCOL)

# ── 6 関連企業 ──
sec("６　関連企業（経営する企業がある場合）")
r = row("関連企業①", "企業名：KHD　／　代表者名：菊池 研太　／　所在地：［ ］　／　業種：不動産・医療コンサル・EC"); m(r, 1, NCOL)

# ── 7 お借入の状況 ──
sec("７　お借入の状況（事業用を除く・個人のお借入）")
r = row("", "［事業用以外の個人のお借入（住宅・車・教育・カード等）があれば 借入先／使いみち／残高／年間返済額 を記入。無ければ「該当なし」］"); m(r, 0, NCOL); HEIGHT[r] = 40

# ── 8 必要な資金と調達方法 ──
sec("８　必要な資金と調達方法")
h = row("必要な資金", "", "", "", "", "見積先", "金額(万円)", "", "調達の方法", "", "", "", "金額(万円)"); THDR.append(h)
m(h, 0, 5); m(h, 8, 12); m(h, 12, NCOL); b0 = h
r = row("設備資金", "加盟金・研修費・保証金・備品・機材・初期システム", "", "", "", "オーロラ本部", "［本部］", "", "自己資金", "", "", "", "600"); m(r, 1, 5); m(r, 8, 12); m(r, 12, NCOL)
r = row("運転資金", "立上げ人件費・採用費・広告分担金・家賃・車両・予備費", "", "", "", "—", "［算定］", "", "親・兄弟・知人等からの借入", "", "", "", "—"); m(r, 1, 5); m(r, 8, 12); m(r, 12, NCOL)
r = row("合計", "", "", "", "", "", "［ ］", "", "日本政策金融公庫 国民生活事業 からの借入", "", "", "", "［所要−600万］"); m(r, 1, 5); m(r, 8, 12); m(r, 12, NCOL)
r = row("", "", "", "", "", "", "", "", "他の金融機関等からの借入", "", "", "", "—"); m(r, 0, 8); m(r, 8, 12); m(r, 12, NCOL)
r = row("", "", "", "", "", "", "", "", "合計", "", "", "", "［＝必要資金合計］"); m(r, 0, 8); m(r, 8, 12); m(r, 12, NCOL); BORDER.append((b0, r+1, 0, NCOL))

# ── 9 事業の見通し ──
sec("９　事業の見通し（月平均）")
h = row("項目", "", "", "", "創業当初", "1年後又は軌道に乗った後（  年  月頃）", "", "売上高・売上原価（仕入高）・経費を計算した根拠"); THDR.append(h)
m(h, 0, 4); m(h, 5, 7); m(h, 7, NCOL); b0 = h
def vrow(item, konkyo, ca="", kido=""):
    r = row(item, "", "", "", ca, kido, "", konkyo); m(r, 0, 4); m(r, 5, 7); m(r, 7, NCOL); return r
vrow("売上高①", "施術者数 × 1日訪問件数 × 訪問単価 × 稼働日数　［本部標準モデル受領後に確定］")
vrow("売上原価②（仕入高）", "訪問は仕入原価は僅少（施術用消耗品のみ）")
vrow("経費　人件費（注）", "施術者報酬（歩合／固定）　［本部確認］")
vrow("経費　家賃", "事務所・拠点家賃")
vrow("経費　支払利息", "公庫借入の返済利息")
vrow("経費　その他", "ロイヤリティ13.2%・広告分担金・車両・通信・消耗品")
vrow("合計③", "")
r = vrow("利益　①－②－③", "（＝公庫返済原資）"); BORDER.append((b0, r+1, 0, NCOL))
r = row("（注）個人営業の場合、事業主分は含めません。"); m(r, 0, NCOL)

# ── 10 自由記述欄 ──
sec("10　自由記述欄（アピールポイント、事業を行ううえでの悩み、希望するアドバイス等）")
r = row(JIYU); m(r, 0, NCOL); HEIGHT[r] = 96

def main():
    svc = build("sheets", "v4", credentials=creds(), cache_discovery=False)
    nrow = len(V) + 4
    ss = svc.spreadsheets().create(body={
        "properties": {"title": TITLE},
        "sheets": [
            {"properties": {"title": "創業計画書", "gridProperties": {"rowCount": nrow, "columnCount": NCOL, "frozenRowCount": 1}}},
            {"properties": {"title": "本部確認_18点", "gridProperties": {"rowCount": 30, "columnCount": 4, "frozenRowCount": 2}}},
        ]}, fields="spreadsheetId,sheets.properties").execute()
    sid = ss["spreadsheetId"]
    gid1 = ss["sheets"][0]["properties"]["sheetId"]
    gid2 = ss["sheets"][1]["properties"]["sheetId"]

    svc.spreadsheets().values().update(spreadsheetId=sid, range="'創業計画書'!A1",
        valueInputOption="USER_ENTERED", body={"values": V}).execute()

    reqs = []
    widths = [110, 120, 70, 70, 110, 95, 95, 70, 150, 60, 60, 60, 110, 95]
    for i, w in enumerate(widths):
        reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid1, "dimension": "COLUMNS", "startIndex": i, "endIndex": i+1}, "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    # 全体: 折返し・上寄せ・10pt
    reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": 0, "endRowIndex": len(V)},
        "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP", "textFormat": {"fontSize": 10, "foregroundColor": INK}}},
        "fields": "userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    # A列ラベル太字
    reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": 1, "endRowIndex": len(V), "startColumnIndex": 0, "endColumnIndex": 1},
        "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 10, "foregroundColor": REDD}}}, "fields": "userEnteredFormat.textFormat"}})
    # マージ
    for (r1, r2, c1, c2) in MERGE:
        reqs.append({"mergeCells": {"range": {"sheetId": gid1, "startRowIndex": r1, "endRowIndex": r2, "startColumnIndex": c1, "endColumnIndex": c2}, "mergeType": "MERGE_ALL"}})
    # タイトル装飾
    reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 8},
        "cell": {"userEnteredFormat": {"backgroundColor": REDD, "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 16}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 8, "endColumnIndex": NCOL},
        "cell": {"userEnteredFormat": {"backgroundColor": CARD, "textFormat": {"foregroundColor": INK, "fontSize": 10}, "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    # セクション見出し
    for i in SEC:
        reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": i, "endRowIndex": i+1},
            "cell": {"userEnteredFormat": {"backgroundColor": RED, "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 11}, "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    # 表ヘッダ
    for i in THDR:
        reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": i, "endRowIndex": i+1},
            "cell": {"userEnteredFormat": {"backgroundColor": REDBG, "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": REDD}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    # 行高
    for ridx, px in HEIGHT.items():
        reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid1, "dimension": "ROWS", "startIndex": ridx, "endIndex": ridx+1}, "properties": {"pixelSize": px}, "fields": "pixelSize"}})
    # 罫線
    gl = {"style": "SOLID", "color": {"red": 0.8, "green": 0.78, "blue": 0.74}}
    for (r1, r2, c1, c2) in BORDER:
        reqs.append({"updateBorders": {"range": {"sheetId": gid1, "startRowIndex": r1, "endRowIndex": r2, "startColumnIndex": c1, "endColumnIndex": c2},
            "top": gl, "bottom": gl, "left": gl, "right": gl, "innerHorizontal": gl, "innerVertical": gl}})

    # ── Tab2 本部確認18点 ──
    t2 = [["公庫 創業計画書 ｜ 本部に確認して後で修正する箇所（18点）", "", "", ""], ["#", "区分", "確認項目", "本部回答（記入欄）"]]
    for i, (kbn, item) in enumerate(CONFIRM, 1):
        t2.append([str(i), kbn, item, ""])
    svc.spreadsheets().values().update(spreadsheetId=sid, range="'本部確認_18点'!A1",
        valueInputOption="USER_ENTERED", body={"values": t2}).execute()
    for i, w in enumerate([40, 110, 470, 300]):
        reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid2, "dimension": "COLUMNS", "startIndex": i, "endIndex": i+1}, "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    reqs.append({"mergeCells": {"range": {"sheetId": gid2, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 4}, "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid2, "startRowIndex": 0, "endRowIndex": 1}, "cell": {"userEnteredFormat": {"backgroundColor": REDD, "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 12}, "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid2, "startRowIndex": 1, "endRowIndex": 2}, "cell": {"userEnteredFormat": {"backgroundColor": RED, "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 11}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    c_end = 2 + len(CONFIRM)
    reqs.append({"repeatCell": {"range": {"sheetId": gid2, "startRowIndex": 2, "endRowIndex": c_end}, "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP", "textFormat": {"fontSize": 10, "foregroundColor": INK}}}, "fields": "userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    for k in range(len(CONFIRM)):
        bg = CARD if k % 2 == 1 else WHT
        reqs.append({"repeatCell": {"range": {"sheetId": gid2, "startRowIndex": 2+k, "endRowIndex": 3+k}, "cell": {"userEnteredFormat": {"backgroundColor": bg}}, "fields": "userEnteredFormat.backgroundColor"}})
    reqs.append({"updateBorders": {"range": {"sheetId": gid2, "startRowIndex": 1, "endRowIndex": c_end, "startColumnIndex": 0, "endColumnIndex": 4}, "innerHorizontal": gl, "innerVertical": gl, "top": gl, "bottom": gl, "left": gl, "right": gl}})

    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": reqs}).execute()
    print("DONE sid=", sid)
    print(f"URL: https://docs.google.com/spreadsheets/d/{sid}/edit")

if __name__ == "__main__":
    main()
