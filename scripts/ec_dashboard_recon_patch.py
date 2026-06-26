# -*- coding: utf-8 -*-
"""
📊EC粗利ダッシュボード — 精算実額照合パッチ (01_経営管理 CFO)
ec_settlement_recon_v2.csv を読み込み:
  1. 「精算実額照合」タブを追加（新規 or 上書き）— 月次実精算 vs 見込み比較
  2. サマリータブ末尾に「■ 実精算照合サマリー」セクションを追記
※ 既存3タブ（サマリー/月次推移/商品戦略）の元データは一切書き換えない
"""
import os, csv, pickle, datetime
from collections import defaultdict
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

BASE = "/Users/kikuchikenta/01_honbu_docs_automation"
SCR  = f"{BASE}/scripts"
RECON_CSV = f"{BASE}/ec_settlement_recon_v2.csv"
TOKEN  = f"{SCR}/sheets_token.pickle"
IDFILE = f"{SCR}/ec_dashboard_id.txt"

RECON_TAB = "精算実額照合"

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

def num(v):
    try: return float(v) if v not in ("", None) else 0.0
    except: return 0.0

def load_recon():
    rows = list(csv.DictReader(open(RECON_CSV, encoding="utf-8-sig")))
    return rows

def build_recon_tab(rows):
    # --------- 月次集計（照合OK のみ） ---------
    monthly = defaultdict(lambda: dict(件数=0, 実精算=0.0, 見込みAB=0.0, acct1実=0.0, acct2実=0.0))
    total_ok_seika  = 0.0
    total_ok_mikomI = 0.0
    ok_cnt = 0

    for r in rows:
        ku = r["区分"]
        if ku == "照合OK":
            ym = r["年月"]
            m = monthly[ym]
            m["件数"] += 1
            m["実精算"] += num(r["実精算_円"])
            m["見込みAB"] += num(r["見込みAB_円"])
            if r["アカウント"] == "クーパン1":
                m["acct1実"] += num(r["実精算_円"])
            else:
                m["acct2実"] += num(r["実精算_円"])
            total_ok_seika  += num(r["実精算_円"])
            total_ok_mikomI += num(r["見込みAB_円"])
            ok_cnt += 1

    # --------- フラグ集計 ---------
    flag_cnt  = defaultdict(int)
    flag_seika = defaultdict(float)
    flag_mikom = defaultdict(float)
    for r in rows:
        ku = r["区分"]
        flag_cnt[ku]  += 1
        flag_seika[ku] += num(r["実精算_円"])
        flag_mikom[ku] += num(r["見込みAB_円"])

    total_rows = len(rows)  # 全区分合計件数（精算待ち含む）
    # 管理表 valid 件数 = 照合OK + 精算待ち
    # ※ A/B は settle側にあるが管理表では除外ステータス → mgmt_valid に含まない
    mgmt_valid = (flag_cnt.get("照合OK",0)
                + flag_cnt.get("精算待ち(実精算なし)",0))
    coverage_pct = ok_cnt / mgmt_valid * 100 if mgmt_valid else 0
    diff = total_ok_seika - total_ok_mikomI
    diff_pct = diff / total_ok_mikomI * 100 if total_ok_mikomI else 0

    today = datetime.date.today()

    # --------- ヘッダー ---------
    sheet = []
    sheet.append(["📊 精算実額照合レポート（Coupang Wing 実精算 vs 販売管理表 見込み）"])
    sheet.append([f"更新日: {today}  /  ソース: ec_settlement_recon_v2.csv"])
    sheet.append([])

    # --------- カバレッジサマリー ---------
    sheet.append(["■ 照合カバレッジサマリー（2025通年）"])
    sheet.append(["指標","値","備考"])
    sheet.append(["管理表 売上計上件数（valid）", mgmt_valid, "発送済/代行登録済み/注文済"])
    sheet.append(["照合OK件数", ok_cnt, "MSF+売掛金スナップで精算確認済"])
    sheet.append(["カバレッジ", f"{coverage_pct:.1f}%", "目標90%以上"])
    sheet.append(["実精算合計（円）", round(total_ok_seika), "照合OK分のみ"])
    sheet.append(["見込みAB合計（円）", round(total_ok_mikomI), "手数料11%固定の見積額"])
    sheet.append(["差（実－見）（円）", round(diff), "実手数料が11%未満のため実精算>見込み"])
    sheet.append(["差（%）", f"{diff_pct:+.2f}%", "実精算が見込みより多い理由=実手数料<11%"])
    sheet.append([])

    # --------- 月次実精算 vs 見込み ---------
    sheet.append(["■ 月次 実精算 vs 見込みAB 比較（照合OK分）"])
    sheet.append(["年月","照合OK件数","実精算_円","見込みAB_円","差（円）","差（%）","クーパン1 実精算","クーパン2 実精算"])
    for ym in sorted(monthly.keys()):
        m = monthly[ym]
        d = m["実精算"] - m["見込みAB"]
        d_pct = d / m["見込みAB"] * 100 if m["見込みAB"] else 0
        sheet.append([
            ym,
            m["件数"],
            round(m["実精算"]),
            round(m["見込みAB"]),
            round(d),
            f"{d_pct:+.2f}%",
            round(m["acct1実"]),
            round(m["acct2実"]),
        ])
    # 合計行
    sheet.append([
        "合計",
        ok_cnt,
        round(total_ok_seika),
        round(total_ok_mikomI),
        round(diff),
        f"{diff_pct:+.2f}%",
        "",
        "",
    ])
    sheet.append([])

    # --------- フラグ別内訳 ---------
    sheet.append(["■ 照合結果フラグ別内訳"])
    sheet.append(["区分","件数","実精算_円（合計）","見込みAB_円（合計）","対応"])
    flag_action = {
        "照合OK":                          "問題なし",
        "精算待ち(実精算なし)":             "月次MSF追加DL→再実行で自動消化",
        "A_売上漏れ候補(status要修正)":     "販売管理表のstatusを「発送済」に修正して売上計上",
        "B_キャンセル補償(雑収入)":         "何もしない。雑収入で正しい",
    }
    for ku in ["照合OK","精算待ち(実精算なし)","A_売上漏れ候補(status要修正)","B_キャンセル補償(雑収入)"]:
        cnt = flag_cnt.get(ku, 0)
        seika = round(flag_seika.get(ku, 0))
        mikom = round(flag_mikom.get(ku, 0))
        sheet.append([ku, cnt, seika, mikom, flag_action.get(ku, "")])
    sheet.append([])

    # --------- 月次運用ランブック ---------
    sheet.append(["■ 月次運用ランブック（毎月の作業手順）"])
    sheet.append(["Step","担当","作業"])
    sheet.append(["1", "ゆーし", "毎月25日頃、Coupang Wing → 決済 → 精算状況 → 決済確定で2ヶ月窓DL（両口座）→ _精算実額MSFフォルダへ追加"])
    sheet.append(["2", "菊池(Claude)", "ec_settlement_recon_v2.py を再実行 → カバレッジ❌の月を確認"])
    sheet.append(["3", "菊池(Claude)", "A_売上漏れ候補があれば販売管理表のstatusを「発送済」に修正"])
    sheet.append(["4", "菊池(Claude)", "ec_profit_pipeline.py → ec_product_strategy.py → ec_dashboard_build.py → ec_dashboard_recon_patch.py の順に実行してダッシュボード更新"])
    sheet.append(["5", "菊池(Claude)", "ec_handoff_03.py で03申し送りmarkdown更新"])
    sheet.append([])

    # --------- 注意事項 ---------
    sheet.append(["■ データ解釈メモ"])
    sheet.append(["・実精算 > 見込みAB の理由: Coupangの実際の手数料が11%より低い（実績約10.2〜10.5%）"])
    sheet.append(["・10月・12月の差がマイナスの理由: 為替差（管理表の行別為替レートとMSF入金日のレートの乖離）"])
    sheet.append(["・A_売上漏れ: 入金前に「キャンセル」と誤分類した注文。管理表修正で+¥75,525の売上計上が可能"])
    sheet.append(["・B_キャンセル補償: 出荷後キャンセル時にCoupangが支払う固定補償（約29,100W）。雑収入扱いで正しい"])
    sheet.append(["・精算待ち38件: Coupangの精算は購入確定から2〜4ヶ月後。月次MSF追加DLで自動消化する"])

    return sheet, coverage_pct, ok_cnt, mgmt_valid, total_ok_seika, total_ok_mikomI, diff, diff_pct

def build_summary_append(coverage_pct, ok_cnt, mgmt_valid,
                          total_ok_seika, total_ok_mikomI, diff, diff_pct):
    """サマリータブに追記する実精算照合セクション"""
    rows = []
    rows.append([])
    rows.append(["■ 実精算照合（Coupang Wing 実績 vs 見込み） ★2026-05-30 v2エンジン完成"])
    rows.append(["指標","値"])
    rows.append(["照合カバレッジ", f"{coverage_pct:.0f}% ({ok_cnt}/{mgmt_valid}件)"])
    rows.append(["実精算合計（円）", round(total_ok_seika)])
    rows.append(["見込みAB合計（円）", round(total_ok_mikomI)])
    rows.append(["差（実－見）", f"¥{round(diff):+,} ({diff_pct:+.2f}%)"])
    rows.append(["推定粗利率（実額ベース）", "約9.6%（見込み8.1%より+1.5pt）"])
    rows.append(["", "実手数料が11%を下回るため実精算は見込みより¥179,080多い"])
    rows.append([])
    rows.append(["A_売上漏れ候補", "4件 / +¥75,525の売上計上が可能（管理表status修正要）"])
    rows.append(["B_キャンセル補償", "5件 / ¥13,875は雑収入（対応不要）"])
    rows.append(["精算待ち", "38件 / ¥559,373相当（月次MSF追加DLで自動消化）"])
    return rows

def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def write_tab(svc, sid, tab, rows):
    end = col_letter(max(len(r) for r in rows))
    rng = f"{tab}!A1:{end}{len(rows)}"
    svc.spreadsheets().values().update(
        spreadsheetId=sid, range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": rows}
    ).execute()

def main():
    if not os.path.exists(RECON_CSV):
        print(f"❌ {RECON_CSV} が見つかりません。ec_settlement_recon_v2.py を先に実行してください。")
        return
    if not os.path.exists(IDFILE):
        print(f"❌ {IDFILE} が見つかりません。ec_dashboard_build.py を先に実行してください。")
        return

    sid = open(IDFILE).read().strip()
    rows_recon = load_recon()

    (recon_sheet, coverage_pct, ok_cnt, mgmt_valid,
     total_ok_seika, total_ok_mikomI, diff, diff_pct) = build_recon_tab(rows_recon)

    summary_append = build_summary_append(
        coverage_pct, ok_cnt, mgmt_valid,
        total_ok_seika, total_ok_mikomI, diff, diff_pct
    )

    svc = build("sheets", "v4", credentials=creds())

    # --- 精算実額照合 タブ: 追加 or クリア ---
    info = svc.spreadsheets().get(spreadsheetId=sid).execute()
    have = {s["properties"]["title"] for s in info["sheets"]}
    if RECON_TAB not in have:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [{"addSheet": {"properties": {"title": RECON_TAB}}}]}
        ).execute()
        print(f"タブ追加: {RECON_TAB}")
    else:
        svc.spreadsheets().values().clear(
            spreadsheetId=sid, range=f"{RECON_TAB}!A1:Z500"
        ).execute()
        print(f"タブクリア: {RECON_TAB}")

    write_tab(svc, sid, RECON_TAB, recon_sheet)
    print(f"✅ 精算実額照合 タブ書き込み完了 ({len(recon_sheet)}行)")

    # --- サマリー タブに末尾追記 ---
    # 既存データの最終行を取得してその後ろにAPPEND
    svc.spreadsheets().values().append(
        spreadsheetId=sid,
        range="サマリー!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": summary_append}
    ).execute()
    print(f"✅ サマリー タブ末尾追記完了 ({len(summary_append)}行)")

    url = f"https://docs.google.com/spreadsheets/d/{sid}/edit"
    print(f"\n📊 ダッシュボード URL: {url}")
    print(f"   実精算照合タブ: {url}#gid= → 「精算実額照合」タブを確認")
    print(f"\n=== 実精算照合 サマリー ===")
    print(f"  カバレッジ : {coverage_pct:.0f}% ({ok_cnt}/{mgmt_valid}件)")
    print(f"  実精算合計 : ¥{total_ok_seika:,.0f}")
    print(f"  見込みAB   : ¥{total_ok_mikomI:,.0f}")
    print(f"  差（実－見）: ¥{diff:+,.0f} ({diff_pct:+.2f}%)")

if __name__ == "__main__":
    main()
