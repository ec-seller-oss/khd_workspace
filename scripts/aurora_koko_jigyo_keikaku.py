# -*- coding: utf-8 -*-
"""
日本政策金融公庫「創業計画書」をGoogle Sheetsで新規作成（オーロラFC船橋）。
Tab1=創業計画書(10項目・埋まる部分清書/本部依存は[本部確認後に修正])、Tab2=本部確認18点。
auth: scripts/sheets_token.pickle (scope=spreadsheets)
作成後URL（=ドライブのリンク）を出力。
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
TITLE = "創業計画書_オーロラFC船橋_公庫_260531"

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

# ── Tab1: 創業計画書 本体 ──
SECTIONS = [
    ["① 創業の動機",
     "当社（KHD）は不動産物件調達および医療クリニックの承継コンサルティングを営み、医療機関・ケアマネジャーとのネットワークと事業運営基盤を有している。事業エリアの船橋市は後期高齢者が8万人を超え、在宅・施設での訪問マッサージ（医療保険適用のマッサージ・はり灸）需要が構造的に拡大している。既存事業で培った医療・介護分野の人的ネットワークを活かせる隣接領域であり、訪問医療マッサージFC「オーロラ」に加盟することで、集客・保険請求（レセプト）代行・運営ノウハウを本部から取得し、未経験リスクを抑えて確実に立ち上げられると判断し創業する。"],
    ["② 経営者の略歴等\n（年月・金額は実績に差替）",
     "・KHD（kikuchi-hd.net）代表。不動産物件調達（仕入〜加工〜売却）、医療クリニック承継コンサル、EC（韓国輸出）を運営。\n・不動産事業で仕入〜売却の出口実行・資金管理の実績あり（直近：栄町案件 ほか）。［正確な売却額・粗利を1〜2行で］\n・複数事業を並行運営し、収支・資金繰りを管理してきた経営管理力を有する。\n・取得資格／許認可：［該当あれば記入］"],
    ["③ 取扱商品・サービス",
     "・訪問マッサージ／訪問はり灸（医療保険適用、医師の同意書に基づく施術）\n・提供エリア：船橋市（本部territory承認済み）\n・セールスポイント：本部ブランド＋レセプト代行で未収・査定リスクが低い／ケアマネ・施設営業網による安定集客\n・単価（1回あたり）：［本部確認後に修正］"],
    ["④ 従業員",
     "・常勤役員（本人）1名\n・施術者（あマ指師・鍼灸師）：開業時 ［ ］名 → 軌道後 ［ ］名\n・採用見込み：Indeed掲載で即戦力1名が応募済（53歳・あマ指師＋鍼灸師、訪問鍼灸マッサージ実務・ケアマネ営業・管理者経験あり）"],
    ["⑤ 取引先・取引関係等",
     "・販売先：後期高齢者医療広域連合等の保険者（レセプト請求）＋利用者の窓口負担。回収サイト：レセプト請求→入金まで約 ［本部確認後に修正］ヶ月\n・外注先：オーロラ本部（ロイヤリティ 13.2%・税込）\n・仕入先：施術用消耗品ほか ［ ］／支払サイト：［実態に合わせ記入］"],
    ["⑥ 関連企業",
     "KHD（代表：菊池研太）／不動産物件調達・医療クリニック承継コンサル・EC（韓国輸出）"],
    ["⑦ 借入の状況（個人）",
     "［要記入：住宅ローン・自動車ローン・カードローン等の有無、残高、年間返済額。無ければ「該当なし」］"],
    ["⑧ 必要な資金と調達方法",
     "【設備資金】 加盟金 ［本部］／研修費 ［本部］／保証金・備品・機材・初期システム ［本部］\n【運転資金】 立上げ人件費（入金まで立替＋数ヶ月分）［算定］／採用費・広告分担金 ［本部/算定］／家賃・車両・予備費 ［算定］\n【調達】 自己資金 600万 ／ 公庫融資 ［所要総額−600万］　※自己資金が所要の1/3以上＝目安クリア圏"],
    ["⑨ 事業の見通し（月平均）",
     "売上＝施術者数 × 1日訪問件数 × 訪問単価 × 稼働日数\n・創業当初（〜6ヶ月）／軌道に乗った後 の2列構成。各数値は本部の標準モデル受領後に確定 ［本部確認後に修正］\n・控除：施術者人件費 → ロイヤリティ13.2%（確定）→ 家賃・車両・消耗品・広告 → 営業利益（＝返済原資）"],
    ["⑩ 自由記入欄",
     "・船橋市の後期高齢者8万人超という需要基盤と、本部territory承認済みの独占エリアで開業する。\n・既存のKHD医療・介護ネットワーク（ケアマネ・医療機関）を集客に活用できる。\n・加盟前にIndeedで採用テストを実施し、即戦力人材の応募を確認済み＝人員確保の蓋然性を実証している。\n・本部のレセプト代行により保険請求の未収リスクを抑制。慎重に採算を検証したうえで開業判断を行う。"],
]

# ── Tab2: 本部確認18点 ──
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
WHT  = {"red": 1, "green": 1, "blue": 1}
INK  = {"red": 0.1, "green": 0.1, "blue": 0.1}

def main():
    svc = build("sheets", "v4", credentials=creds(), cache_discovery=False)
    # 新規作成（2タブ）
    ss = svc.spreadsheets().create(body={
        "properties": {"title": TITLE},
        "sheets": [
            {"properties": {"title": "創業計画書", "gridProperties": {"rowCount": 30, "columnCount": 2, "frozenRowCount": 4}}},
            {"properties": {"title": "本部確認_18点", "gridProperties": {"rowCount": 30, "columnCount": 4, "frozenRowCount": 2}}},
        ]}, fields="spreadsheetId,sheets.properties").execute()
    sid = ss["spreadsheetId"]
    gid1 = ss["sheets"][0]["properties"]["sheetId"]
    gid2 = ss["sheets"][1]["properties"]["sheetId"]

    # ── Tab1 値 ──
    t1 = [["創業計画書 ｜ オーロラFC 船橋（日本政策金融公庫 新規開業資金）", ""],
          ["作成日 2026-05-31 ／ 自己資金上限 600万 ／ ［本部確認後に修正］は本部回答後に上書き", ""],
          ["", ""],
          ["項目", "記入内容"]]
    t1 += SECTIONS
    t1 += [["", ""], ["★ 未確定（本部依存）", "③単価／⑤入金サイト／⑧本部費用／⑨見通し数値 → 「本部確認_18点」タブ参照"]]
    svc.spreadsheets().values().update(spreadsheetId=sid, range="'創業計画書'!A1",
        valueInputOption="USER_ENTERED", body={"values": t1}).execute()

    # ── Tab2 値 ──
    t2 = [["公庫 創業計画書 ｜ 本部に確認して後で修正する箇所（18点）", "", "", ""],
          ["#", "区分", "確認項目", "本部回答（記入欄）"]]
    for i, (kbn, item) in enumerate(CONFIRM, 1):
        t2.append([str(i), kbn, item, ""])
    svc.spreadsheets().values().update(spreadsheetId=sid, range="'本部確認_18点'!A1",
        valueInputOption="USER_ENTERED", body={"values": t2}).execute()

    reqs = []
    # Tab1 列幅
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid1, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 200}, "fields": "pixelSize"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid1, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 760}, "fields": "pixelSize"}})
    # Tab1 タイトル
    reqs.append({"mergeCells": {"range": {"sheetId": gid1, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 2}, "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": 0, "endRowIndex": 1}, "cell": {"userEnteredFormat": {"backgroundColor": REDD, "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 13}, "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid1, "dimension": "ROWS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 42}, "fields": "pixelSize"}})
    reqs.append({"mergeCells": {"range": {"sheetId": gid1, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 2}, "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": 1, "endRowIndex": 2}, "cell": {"userEnteredFormat": {"backgroundColor": CARD, "textFormat": {"foregroundColor": INK, "fontSize": 10}}}, "fields": "userEnteredFormat"}})
    # Tab1 ヘッダ(行4=index3)
    reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": 3, "endRowIndex": 4}, "cell": {"userEnteredFormat": {"backgroundColor": RED, "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 11}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    # Tab1 本文（A=項目太字/上, B=折返し/上）
    body_start, body_end = 4, 4 + len(SECTIONS)
    reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": body_start, "endRowIndex": body_end + 2}, "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP", "textFormat": {"fontSize": 10, "foregroundColor": INK}}}, "fields": "userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid1, "startRowIndex": body_start, "endRowIndex": body_end, "startColumnIndex": 0, "endColumnIndex": 1}, "cell": {"userEnteredFormat": {"backgroundColor": CARD, "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": REDD}}}, "fields": "userEnteredFormat(backgroundColor,textFormat)"}})
    # Tab1 枠線
    reqs.append({"updateBorders": {"range": {"sheetId": gid1, "startRowIndex": 3, "endRowIndex": body_end, "startColumnIndex": 0, "endColumnIndex": 2}, "innerHorizontal": {"style": "SOLID", "color": {"red": 0.85, "green": 0.83, "blue": 0.79}}, "innerVertical": {"style": "SOLID", "color": {"red": 0.85, "green": 0.83, "blue": 0.79}}}})

    # Tab2 列幅
    for i, w in enumerate([40, 110, 470, 300]):
        reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid2, "dimension": "COLUMNS", "startIndex": i, "endIndex": i+1}, "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    reqs.append({"mergeCells": {"range": {"sheetId": gid2, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 4}, "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid2, "startRowIndex": 0, "endRowIndex": 1}, "cell": {"userEnteredFormat": {"backgroundColor": REDD, "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 12}, "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid2, "startRowIndex": 1, "endRowIndex": 2}, "cell": {"userEnteredFormat": {"backgroundColor": RED, "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 11}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    c_end = 2 + len(CONFIRM)
    reqs.append({"repeatCell": {"range": {"sheetId": gid2, "startRowIndex": 2, "endRowIndex": c_end}, "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP", "textFormat": {"fontSize": 10, "foregroundColor": INK}}}, "fields": "userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid2, "startRowIndex": 2, "endRowIndex": c_end, "startColumnIndex": 0, "endColumnIndex": 1}, "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER", "textFormat": {"bold": True, "fontSize": 10}}}, "fields": "userEnteredFormat(horizontalAlignment,textFormat)"}})
    for k in range(len(CONFIRM)):
        r = 2 + k
        bg = CARD if k % 2 == 1 else WHT
        reqs.append({"repeatCell": {"range": {"sheetId": gid2, "startRowIndex": r, "endRowIndex": r+1}, "cell": {"userEnteredFormat": {"backgroundColor": bg}}, "fields": "userEnteredFormat.backgroundColor"}})
    reqs.append({"updateBorders": {"range": {"sheetId": gid2, "startRowIndex": 1, "endRowIndex": c_end, "startColumnIndex": 0, "endColumnIndex": 4}, "innerHorizontal": {"style": "SOLID", "color": {"red": 0.85, "green": 0.83, "blue": 0.79}}, "innerVertical": {"style": "SOLID", "color": {"red": 0.85, "green": 0.83, "blue": 0.79}}}})

    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": reqs}).execute()
    print("DONE sid=", sid)
    print(f"URL: https://docs.google.com/spreadsheets/d/{sid}/edit")

if __name__ == "__main__":
    main()
