# -*- coding: utf-8 -*-
"""
オーロラ大元スプシ「医療不動産×オーロラFC_統合事業計画」に
面談ヒアリング比較表タブを追加する（石原洋・6/2対面面談用）。
auth: scripts/sheets_token.pickle (scope=spreadsheets)
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID = "18C4F_EbMnOR9uNisgXBWxTgK4R8DvOnMqv-1sGW5MPc"
TAB = "10_面談ヒアリング_石原(6_2)"

# ── 応募資料 Driveリンク（HYPERLINKで埋め込み）──
L_FOLDER = "https://drive.google.com/drive/folders/1odM59_44GvscPC9KeohtkL1lK9Tk-xXm"
L_RESUME = "https://drive.google.com/file/d/1QrDxN5oBN6xa6bxD07OP9saFwFjITuL0/view"
L_NOTI1  = "https://drive.google.com/file/d/16wjNcMbyNbL_8RVG3dcQquZXk-x6Mdd0/view"
L_NOTI2  = "https://drive.google.com/file/d/14dksnwnnS7cd6qj8v2bf2N0vCXhrDDNa/view"
def hl(url, label):
    return f'=HYPERLINK("{url}","{label}")'

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

# ── 表データ ──
HDR = ["#", "項目", "質問（具体の聞き方）", "◎ 嬉しい回答（GO寄り）", "△ ひっかかる回答（要警戒）", "対面だから見れるポイント"]

ROWS = [
    ["1", "転職理由",
     "「今の治療院、何がきっかけで次を考え始めたんですか？」",
     "前向き・自走志向（裁量を持ちたい／立上げに関わりたい）。不満が“環境”で姿勢は建設的",
     "待遇・人間関係の愚痴が中心、他責、辞め方が曖昧",
     "表情・声のトーン、語る順番で本音度を見る"],
    ["2", "希望条件（給与/稼働/歩合）",
     "「働き方のイメージはありますか？週何日くらい、報酬の考え方は？」",
     "歩合・成果連動に前向き。相場感が現実的でフル稼働可",
     "高い固定給だけ要求して歩合を嫌う／稼働日数が極端に少ない",
     "こちらが数字感を出した時の反応・粘り"],
    ["3", "訪問件数の実績（=売上直結）★中身を要確認",
     "「履歴書の“1日30人以上”は、店舗に来てもらう来店型ですか？それともご自宅を1軒ずつ回る訪問型ですか？訪問だと1日何軒が現実的でした？」",
     "訪問（1対1宅訪）で1日6〜8軒を安定して語れる＝当事業の売上が読める。移動効率も具体的",
     "30人は店舗来店型だった／訪問の実件数は曖昧・極端に少ない、移動段取りを考えていない",
     "“30人”の内訳（来店型か訪問型か）を必ず分解。訪問の1日を具体的に再現して語れるか"],
    ["4", "集客力（ケアマネ営業）",
     "「ケアマネさんや施設への営業は、ご自身でやられてました？どんな風に？」",
     "自分で新規開拓した実績・関係構築の“型”を持つ＝拠点長候補",
     "「営業は会社任せだった」受け身、紹介待ち",
     "営業エピソードの解像度・固有名詞や数字が出るか"],
    ["5", "いつから動けるか",
     "「もし一緒にやるとなったら、いつ頃から動けそうですか？」",
     "1〜2ヶ月以内に動ける。引継ぎ・退職の算段がある",
     "未定／現職を辞める決心が固まっていない",
     "即答できるか、言葉を濁すか"],
    # ── 石原さんならでは（プロフィール特化）──
    ["6", "★現職在籍中の本気度",
     "「今もお勤めされてますよね。転職は具体的に動かれている段階ですか？」",
     "すでに動いている／退職時期の目処を持っている",
     "情報収集どまり・冷やかし感、複数を天秤にかけるだけ",
     "在職リスク（並行応募）を正直に話せるか＝誠実さの試金石"],
    ["7", "★マネジメント経験／拠点長意欲",
     "「マネジメント経験はありますか？」＋「事業が伸びたら、業務委託の若手施術者をまとめる立場をお願いしたい。そうなると報酬も上がる。そういう将来像に興味は？」",
     "実際に人を管理した具体エピソードあり＝拡大時に若手業務委託を任せられる。昇格＆昇給の未来に前向き＝長期コミット",
     "プレイヤー専念希望・短期志向／管理は人任せだった・指示待ち",
     "①若手を束ねられるか（再現性ある管理経験）②昇格＝昇給の未来を提示した時の食いつき。2つの狙いで見る"],
    ["8", "★体力・長期稼働（53歳）",
     "「訪問は移動も多いですが、稼働ペースは今後どのくらいで考えてます？」",
     "現実的な長期稼働プラン。健康・自己管理ができている",
     "無理が見える／体力面の不安を曖昧にする",
     "受け答えの活力・所作。長く一緒にやれる相手か"],
    ["9", "★競業リスク（現職=津田沼・船橋近接）",
     "「エリアが近いですが、現職との関係で気をつける点はありますか？」",
     "競業の線引きを理解し、トラブルにしない配慮がある",
     "現職の顧客・契約を持ち出す気満々（コンプラ赤信号）",
     "倫理観・リスク感度。ここは必ず確認（持出し前提はNG）"],
]

NOTE_TOP = "面談スタンス＝この募集テストの目的は“石橋を叩く”。遠慮せず踏み込んで聞く（正味、尋問でもいい）。引き出さなければ始まらない／こちらは雇う＝組織的立ち位置は崩さず、強めに確認してよい。ただし舐められも卑屈にもならず堂々と。◎が多く△の地雷（特に#9競業）が無ければ前進。n=1で加盟GOは早計、複数応募の再現性も併せて判断。"
NOTE_BOTTOM = "【総合判定メモ】 ◎個数：__  /  地雷(△)：__  /  即戦力度：__  /  拠点長適性：__  /  次アクション：__"


def col(i): return i  # 0-based

def main():
    svc = build("sheets", "v4", credentials=creds(), cache_discovery=False)
    # 既存タブ確認・あれば削除して作り直し
    info = svc.spreadsheets().get(spreadsheetId=SID, fields="sheets.properties").execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in info["sheets"]}
    if TAB in existing:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={
            "requests": [{"deleteSheet": {"sheetId": existing[TAB]}}]}).execute()
    # 追加
    res = svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={
        "requests": [{"addSheet": {"properties": {
            "title": TAB,
            "gridProperties": {"rowCount": 60, "columnCount": 6, "frozenRowCount": 6}}}}]}).execute()
    gid = res["replies"][0]["addSheet"]["properties"]["sheetId"]

    # 値
    title_row = ["採用面談ヒアリング比較表 ｜ 石原 洋（6/2 対面）"] + [""]*5
    links_row = ["📎 応募資料（クリックで開く）",
                 hl(L_FOLDER, "📁 資料フォルダ"), hl(L_RESUME, "📄 履歴書"),
                 hl(L_NOTI1, "🖼 Indeed応募通知①"), hl(L_NOTI2, "🖼 Indeed応募通知②"), ""]
    note_row = [NOTE_TOP] + [""]*5
    values = [title_row, [""]*6, links_row, [""]*6, note_row, HDR] + ROWS + [[""]*6, [NOTE_BOTTOM] + [""]*5]
    svc.spreadsheets().values().update(
        spreadsheetId=SID, range=f"'{TAB}'!A1",
        valueInputOption="USER_ENTERED", body={"values": values}).execute()

    n_rows = len(values)
    links_idx = 2  # 0-based row of 応募資料リンク
    note_top_idx = 4
    hdr_idx = 5  # 0-based row of HDR
    data_start = hdr_idx + 1
    data_end = data_start + len(ROWS)  # exclusive
    note_bottom_idx = n_rows - 1

    RED = {"red": 0.667, "green": 0.180, "blue": 0.149}
    REDD = {"red": 0.549, "green": 0.141, "blue": 0.114}
    CARD = {"red": 0.945, "green": 0.925, "blue": 0.882}
    REDBG = {"red": 0.957, "green": 0.894, "blue": 0.886}
    WHT = {"red": 1, "green": 1, "blue": 1}
    INK = {"red": 0.1, "green": 0.1, "blue": 0.1}

    reqs = []
    # 列幅
    widths = [40, 200, 330, 360, 330, 300]
    for i, w in enumerate(widths):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": gid, "dimension": "COLUMNS", "startIndex": i, "endIndex": i+1},
            "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    # タイトル行 merge＋装飾
    reqs.append({"mergeCells": {"range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1,
        "startColumnIndex": 0, "endColumnIndex": 6}, "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1},
        "cell": {"userEnteredFormat": {"backgroundColor": REDD,
            "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 14},
            "verticalAlignment": "MIDDLE", "horizontalAlignment": "LEFT"}},
        "fields": "userEnteredFormat"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS",
        "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 44}, "fields": "pixelSize"}})
    # 応募資料リンク行
    BLUE = {"red": 0.066, "green": 0.333, "blue": 0.8}
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": links_idx, "endRowIndex": links_idx+1},
        "cell": {"userEnteredFormat": {"backgroundColor": CARD,
            "textFormat": {"bold": True, "fontSize": 11, "foregroundColor": BLUE},
            "verticalAlignment": "MIDDLE", "horizontalAlignment": "LEFT"}},
        "fields": "userEnteredFormat"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": links_idx, "endRowIndex": links_idx+1,
        "startColumnIndex": 0, "endColumnIndex": 1},
        "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 11, "foregroundColor": INK}}},
        "fields": "userEnteredFormat.textFormat"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS",
        "startIndex": links_idx, "endIndex": links_idx+1}, "properties": {"pixelSize": 34}, "fields": "pixelSize"}})
    # 注記行 merge＋装飾
    reqs.append({"mergeCells": {"range": {"sheetId": gid, "startRowIndex": note_top_idx, "endRowIndex": note_top_idx+1,
        "startColumnIndex": 0, "endColumnIndex": 6}, "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": note_top_idx, "endRowIndex": note_top_idx+1},
        "cell": {"userEnteredFormat": {"backgroundColor": REDBG,
            "textFormat": {"foregroundColor": REDD, "bold": True, "fontSize": 10},
            "wrapStrategy": "WRAP", "verticalAlignment": "MIDDLE"}},
        "fields": "userEnteredFormat"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS",
        "startIndex": note_top_idx, "endIndex": note_top_idx+1}, "properties": {"pixelSize": 60}, "fields": "pixelSize"}})
    # ヘッダ行
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": hdr_idx, "endRowIndex": hdr_idx+1},
        "cell": {"userEnteredFormat": {"backgroundColor": RED,
            "textFormat": {"foregroundColor": WHT, "bold": True, "fontSize": 11},
            "wrapStrategy": "WRAP", "verticalAlignment": "MIDDLE", "horizontalAlignment": "CENTER"}},
        "fields": "userEnteredFormat"}})
    # データ行：折返し＋上寄せ＋交互色
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": data_start, "endRowIndex": data_end},
        "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP",
            "textFormat": {"fontSize": 10, "foregroundColor": INK}}},
        "fields": "userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)"}})
    for k in range(len(ROWS)):
        r = data_start + k
        bg = CARD if k % 2 == 1 else WHT
        reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": r, "endRowIndex": r+1},
            "cell": {"userEnteredFormat": {"backgroundColor": bg}}, "fields": "userEnteredFormat.backgroundColor"}})
    # #列・項目列を太字中央/左
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": data_start, "endRowIndex": data_end,
        "startColumnIndex": 0, "endColumnIndex": 1},
        "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER",
            "textFormat": {"bold": True, "fontSize": 11}}},
        "fields": "userEnteredFormat(horizontalAlignment,textFormat)"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": data_start, "endRowIndex": data_end,
        "startColumnIndex": 1, "endColumnIndex": 2},
        "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 10, "foregroundColor": INK}}},
        "fields": "userEnteredFormat.textFormat"}})
    # 総合判定メモ行 merge
    reqs.append({"mergeCells": {"range": {"sheetId": gid, "startRowIndex": note_bottom_idx, "endRowIndex": note_bottom_idx+1,
        "startColumnIndex": 0, "endColumnIndex": 6}, "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": note_bottom_idx, "endRowIndex": note_bottom_idx+1},
        "cell": {"userEnteredFormat": {"backgroundColor": CARD,
            "textFormat": {"bold": True, "fontSize": 11, "foregroundColor": INK},
            "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS",
        "startIndex": note_bottom_idx, "endIndex": note_bottom_idx+1}, "properties": {"pixelSize": 40}, "fields": "pixelSize"}})
    # データ行の高さ確保
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": gid, "dimension": "ROWS",
        "startIndex": data_start, "endIndex": data_end}, "properties": {"pixelSize": 58}, "fields": "pixelSize"}})
    # 枠線
    reqs.append({"updateBorders": {"range": {"sheetId": gid, "startRowIndex": hdr_idx, "endRowIndex": data_end,
        "startColumnIndex": 0, "endColumnIndex": 6},
        "innerHorizontal": {"style": "SOLID", "color": {"red": 0.85, "green": 0.83, "blue": 0.79}},
        "innerVertical": {"style": "SOLID", "color": {"red": 0.85, "green": 0.83, "blue": 0.79}}}})

    svc.spreadsheets().batchUpdate(spreadsheetId=SID, body={"requests": reqs}).execute()
    print("DONE tab=", TAB, "gid=", gid)
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__ == "__main__":
    main()
