# -*- coding: utf-8 -*-
"""
📋 EC運用マニュアル・パイプライン仕様 → Sheetsタブ追加/更新
対象: 1QjyPPOto7J1HiqA_Zb9-UIOe_FQZyqAGSn321R37Tzo
  → 「📋 運用マニュアル」「⚙️ パイプライン仕様」タブを追加
auth: sheets_token.pickle (既存使い回し)
"""
import pickle, datetime
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN  = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SS_ID  = "1QjyPPOto7J1HiqA_Zb9-UIOe_FQZyqAGSn321R37Tzo"

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

def get_service():
    return build("sheets", "v4", credentials=creds(), cache_discovery=False)

def get_sheet_id(svc, title):
    """タブ名 → sheetId。存在しなければ None"""
    meta = svc.spreadsheets().get(spreadsheetId=SS_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == title:
            return s["properties"]["sheetId"]
    return None

def ensure_tab(svc, title, index):
    """タブがなければ追加、あれば既存のsheetIdを返す"""
    sid = get_sheet_id(svc, title)
    if sid is None:
        body = {"requests": [{"addSheet": {"properties": {"title": title, "index": index}}}]}
        resp = svc.spreadsheets().batchUpdate(spreadsheetId=SS_ID, body=body).execute()
        sid = resp["replies"][0]["addSheet"]["properties"]["sheetId"]
        print(f"  ✅ タブ追加: {title} (sheetId={sid})")
    else:
        print(f"  ♻️  タブ既存: {title} (sheetId={sid})")
    return sid

def clear_and_write(svc, tab_name, rows):
    """タブをクリアしてデータ書き込み"""
    svc.spreadsheets().values().clear(
        spreadsheetId=SS_ID, range=f"'{tab_name}'!A1:Z500"
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=SS_ID,
        range=f"'{tab_name}'!A1",
        valueInputOption="USER_ENTERED",
        body={"values": rows}
    ).execute()
    print(f"  📝 {len(rows)}行 書き込み完了: {tab_name}")

def fmt_header(text):
    return [text]

def fmt_section(title):
    return [f"■ {title}"]

def build_manual():
    today = datetime.date.today().strftime("%Y-%m-%d")
    rows = [
        ["📋 EC韓国輸出 運用マニュアル（引き継ぎ用）", "", "", "", ""],
        [f"更新日: {today}  /  対象: クーパン1・クーパン2  /  担当: 麻梨奈（or ゆーし）", "", "", "", ""],
        [],
        ["■ このビジネスの全体像"],
        ["", "クーパン（韓国EC）で受注 → Amazon JP で仕入れ → HANIRO（代行業者）→ 韓国顧客へ配送"],
        ["", "FBA・倉庫は不使用。Amazon JPは『仕入れ先』のみ"],
        [],
        ["■ あなたがやること（1日15分・月2〜3時間）"],
        ["", "作業", "頻度", "時間", "トリガー"],
        ["", "新規注文を販売管理表に記入 + Amazon発注", "注文のたびに", "10〜15分/件", "LINEで「⚠️新規注文」通知が来たら"],
        ["", "HANIROにCSVをアップロード", "週1〜2回", "5分/回", "LINEで「📦HANIRO CSV生成」通知が来たら"],
        ["", "月次粗利の確認（数字を眺めるだけ）", "月1回（25日以降）", "3〜5分", "LINEで「✅粗利パイプライン完了」通知が来たら"],
        ["", "CFダッシュボードのB3セル更新", "月1回（毎月1日）", "1分", "月初に通帳残高を確認して入力"],
        [],
        ["■ 販売管理表への注文入力（最重要）"],
        ["", "①クーパンWingにログイン → 注文管理 → 出荷待ち → 新規注文を確認"],
        ["", "②Drive → 韓国輸出売上 → 当月ファイル（例: 2601_販売管理表1.xlsx）を開く"],
        ["", "③最下行の下に1行追加"],
        ["", "   A列（進捗）: 注文済"],
        ["", "   B列（注文番号）: クーパン注文番号（数字のみ）"],
        ["", "   C列（プラットフォーム）: クーパン1 or クーパン2"],
        ["", "   F列（注文者名）: 注文者名"],
        ["", "   G/H列（商品名）: 韓国語・日本語商品名"],
        ["", "   I列（ASIN）: Amazon ASIN コード"],
        ["", "   M列（仕入）: Amazonで買った金額（円）"],
        ["", "   W列（Quantity）: 数量"],
        ["", "   X列（Sale Price）: クーパンでの売価（ウォン）"],
        ["", "④AmazonでASINを検索して購入"],
        ["", "   → 注文確定後、O列にAmazon注文番号を入力"],
        ["", "   → A列を『注文済』のまま（HANIROアップロード後に変わる）"],
        [],
        ["■ HANIROへのCSVアップロード"],
        ["", "①LINEで『📦 HANIRO CSV生成完了』通知が届く（毎朝9時自動）"],
        ["", "②通知に記載のCSVファイルをDriveの _HANIRO登録CSV フォルダで確認"],
        ["", "③HANIROサイト（haniro.co.kr）にログイン"],
        ["", "④「一括登録」メニュー → CSVアップロード"],
        ["", "⑤完了後、販売管理表のA列を『代行登録済み』に変更"],
        [],
        ["■ 月次粗利確認（毎月25日以降）"],
        ["", "①LINEで『✅月次EC粗利パイプライン完了』通知を受信"],
        ["", "②このスプシ「月次推移」タブを開いて最新月の粗利を確認"],
        ["", "③粗利率が5%以下 or 前月比で大幅マイナスなら研太に連絡"],
        ["", "④問題なければ何もしなくてOK"],
        [],
        ["■ エラー通知が来たときの対応"],
        ["", "エラー内容", "対処"],
        ["", "⚠️ 受注監視 失敗", "研太にLINE → ec_order_monitor.py の再起動を依頼"],
        ["", "⚠️ HANIRO CSV生成 失敗", "研太にLINE → 販売管理表のファイル名を確認"],
        ["", "⚠️ Wing DL失敗（月25日）", "研太にLINE → Wingのセッション切れの可能性"],
        ["", "LINEが来ない（25日過ぎても通知なし）", "研太にLINE → デーモン停止の可能性"],
        [],
        ["■ ファイルの置き場（重要）"],
        ["", "販売管理表", "Google Drive → 共有ドライブ → 01_個人 → 2025_帳票、明細 → 韓国輸出売上"],
        ["", "HANIRO CSV", "同フォルダ → _HANIRO登録CSV"],
        ["", "粗利ダッシュボード", "このスプシ（月次推移タブ）"],
        ["", "CFダッシュボード", "KHD全社WBSスプシ → 資金繰りv4タブ"],
    ]
    return rows

def build_pipeline():
    today = datetime.date.today().strftime("%Y-%m-%d")
    rows = [
        ["⚙️ EC自動化パイプライン仕様書", "", "", "", ""],
        [f"更新日: {today}  /  目的: ゆーし解雇→妻が月2〜3hで運用できる体制", "", "", "", ""],
        [],
        ["■ 自動化レイヤー一覧"],
        ["Layer", "スクリプト", "実行タイミング", "何をするか", "状態"],
        ["Layer1", "ec_price_monitor.py", "2時間毎（常駐）", "Amazon価格監視・赤字商品の価格自動調整", "✅ 稼働中"],
        ["Layer2", "ec_order_monitor.py", "30分毎（常駐）", "新規注文検知 → 販売管理表追記 → Amazon発注補助", "✅ 稼働中（coupang.json要設定）"],
        ["Layer3", "haniro_csv_generator.py", "毎朝9:00（常駐）", "HANIRO未登録の発注済み注文をCSV化 → LINE通知", "✅ 稼働中"],
        ["Layer4", "run_monthly_ec_pipeline.sh", "毎月25日9:00（cron）", "MSF自動DL → 粗利集計 → ダッシュボード更新 → LINE通知", "✅ 稼働中"],
        ["Layer5", "ec_tracking_updater.py", "将来実装", "追跡番号自動取得 → Coupangへ反映", "⏳ 未実装"],
        [],
        ["■ 月次パイプライン詳細（毎月25日9:00自動）"],
        ["実行順", "スクリプト", "内容", "失敗時"],
        ["①", "wing_auto_download.py", "Coupang WingにPlaywrightでログインしてMSF精算ファイルをDL", "スキップして②以降続行 + LINE緊急通知"],
        ["②", "ec_profit_pipeline.py", "販売管理表24本(25*_販売管理表*.xlsx)を読んで粗利集計CSV生成", "②以降失敗 → LINE緊急通知"],
        ["③", "ec_product_strategy.py", "ASIN別4区分（スポット/安定/撤退候補/様子見）分類", "警告のみ（致命的でない）"],
        ["④", "ec_settlement_recon_v2.py", "MSF精算実額とのズレを照合（照合OK率目標95%以上）", "警告のみ"],
        ["⑤", "ec_dashboard_build.py", "このスプシの月次推移・商品戦略タブを更新", "LINE緊急通知"],
        ["⑥", "ec_dashboard_recon_patch.py", "精算実額照合タブを更新", "警告のみ（⑤成功時のみ実行）"],
        ["⑦", "ec_handoff_03.py", "03_事業運営への申し送りファイル生成", "警告のみ"],
        [],
        ["■ 粗利計算ロジック"],
        ["", "純売上（円）", "= 販売管理表AB列（手数料控除後入金額）"],
        ["", "原価（円）", "= 販売管理表M列（Amazon仕入れ額、手入力）"],
        ["", "粗利", "= 純売上 − 原価"],
        ["", "手数料", "= 総売上（X列×為替）− 純売上"],
        ["", "キャンセル/返品", "売上から除外。仕入済みなら原価リスクとして別計上"],
        ["", "税務申告", "橋本税理士のMoneyForward帳簿が正本。このシートはKPI管理専用"],
        ["", "還付申告タイミング", "年4回（5・8・11・2月） / Amazon JP仕入れ消費税10%が還付対象"],
        [],
        ["■ 2026年データへの対応"],
        ["", "現在のglobパターン: 25*_販売管理表*.xlsx（2025年のみ）"],
        ["", "2026年（2601〜）ファイルが来たら ec_profit_pipeline.py のglob行を変更:"],
        ["", "  変更前: glob.glob(f\"{DIR}/25*_販売管理表*.xlsx\")"],
        ["", "  変更後: glob.glob(f\"{DIR}/{25,26}*_販売管理表*.xlsx\")"],
        ["", "  → その後 run_monthly_ec_pipeline.sh を手動1回実行で完了"],
        [],
        ["■ ファイル・スクリプト一覧"],
        ["ファイル", "場所", "役割"],
        ["ec_full_auto_daemon.py", "scripts/", "全常駐ジョブのAPSchedulerデーモン（LaunchAgent登録済）"],
        ["ec_order_monitor.py", "scripts/", "30分毎Wingポーリング・注文DB書き込み・販売管理表更新"],
        ["ec_price_monitor.py", "scripts/", "Amazon価格スクレイピング・赤字商品検知・価格更新"],
        ["haniro_csv_generator.py", "scripts/", "HANIRO一括登録CSV自動生成"],
        ["ec_profit_pipeline.py", "scripts/", "販売管理表24本集計→粗利CSV生成"],
        ["ec_dashboard_build.py", "scripts/", "粗利CSVからSheetsダッシュボード更新"],
        ["ec_dashboard_recon_patch.py", "scripts/", "精算実額照合タブ更新"],
        ["wing_auto_download.py", "scripts/", "Playwright → Wing MSFファイル自動DL"],
        ["ec_automation_db.py", "scripts/", "SQLite DB（orders/products/price_log）"],
        ["ec_notify.py", "scripts/", "LINE Messaging API 通知ヘルパー"],
        ["run_monthly_ec_pipeline.sh", "scripts/", "月25日cron: 全スクリプト直列実行"],
        ["~/.config/khd/coupang.json", "~/.config/khd/", "Coupang認証情報（要設定・chmod 600必須）"],
        ["ec_automation.db", "~/01_honbu_docs_automation/", "SQLite状態DB"],
        [],
        ["■ トラブルシューティング"],
        ["症状", "原因", "対処"],
        ["LINE通知が来ない", "LINE token切れ / coupang.json未設定", "~/.config/khd/coupang.json を確認"],
        ["Wing DL失敗", "セッション切れ / 2FA", "wing_sessions/*.jsonを削除して再実行"],
        ["ダウンロードボタン未検出", "Coupang Wing UI変更", "/tmp/wing_*.png を確認 → セレクタ修正"],
        ["粗利がゼロになる", "販売管理表のパスまたはglob不一致", "DIR変数とファイル名を確認"],
        ["デーモンが止まっている", "PIDファイル確認: cat ~/01_honbu_docs_automation/ec_daemon.pid", "python3 scripts/ec_full_auto_daemon.py &"],
        [],
        ["■ 連絡先・リンク"],
        ["", "EC粗利ダッシュボード（このスプシ）", "https://docs.google.com/spreadsheets/d/1QjyPPOto7J1HiqA_Zb9-UIOe_FQZyqAGSn321R37Tzo"],
        ["", "CFダッシュボードv4", "https://docs.google.com/spreadsheets/d/1vrMXgI5pNrEGFAo4IR1YBrDLT_kI4Ecsy8PxhsVQm0A"],
        ["", "借入7本SSoT（Notion）", "https://www.notion.so/7-v3-3707d27fe2958164ab41ffaa33b5a473"],
        ["", "EC運用ハブ（Notion）", "https://www.notion.so/EC-3717d27fe295810​69cb4e5be1e659776"],
        ["", "自動化設計書", "~/01_honbu_docs_automation/EC粗利_自動化設計書.md"],
    ]
    return rows

def main():
    svc = get_service()

    # タブ確認・作成
    manual_sid  = ensure_tab(svc, "📋 運用マニュアル", 4)
    pipeline_sid = ensure_tab(svc, "⚙️ パイプライン仕様", 5)

    # データ書き込み
    clear_and_write(svc, "📋 運用マニュアル",  build_manual())
    clear_and_write(svc, "⚙️ パイプライン仕様", build_pipeline())

    # A列幅を少し広げる（視認性向上）
    requests = []
    for sid in [manual_sid, pipeline_sid]:
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
                "properties": {"pixelSize": 30},
                "fields": "pixelSize"
            }
        })
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2},
                "properties": {"pixelSize": 280},
                "fields": "pixelSize"
            }
        })
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
                "properties": {"pixelSize": 220},
                "fields": "pixelSize"
            }
        })
    svc.spreadsheets().batchUpdate(spreadsheetId=SS_ID, body={"requests": requests}).execute()
    print("✅ 列幅調整完了")
    print(f"\n🎉 完了: https://docs.google.com/spreadsheets/d/{SS_ID}")

if __name__ == "__main__":
    main()
