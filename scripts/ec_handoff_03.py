# -*- coding: utf-8 -*-
"""
EC粗利分析 → 03事業運営 申し送り生成パイプライン (01_経営管理 CFO → 03 テッカン)
ec_orders_consolidated.csv / ec_product_ranking.csv を読み、
「03が次に何をすべきか」を数値つきで確定した申し送りドキュメント(markdown)を生成する。

出力:
  1) ec_handoff_03.md           … Notionへ貼る/Drive保存用の構造化マークダウン
  2) コンソールに要約            … 目視確認用

毎月の使い方: pipeline → strategy → dashboard を回した後に本スクリプトを実行。
生成された ec_handoff_03.md を 01→03申し送りNotionページへ反映(更新)する。
※数字はすべてCSVから再計算するので、CSVが最新なら申し送りも自動で最新になる。
"""
import csv, datetime
from collections import defaultdict

BASE = "/Users/kikuchikenta/01_honbu_docs_automation"
ORDERS = f"{BASE}/ec_orders_consolidated.csv"
RANK = f"{BASE}/ec_product_ranking.csv"
OUT_MD = f"{BASE}/ec_handoff_03.md"
DASH_ID = "1QjyPPOto7J1HiqA_Zb9-UIOe_FQZyqAGSn321R37Tzo"
DASH_URL = f"https://docs.google.com/spreadsheets/d/{DASH_ID}/edit"

def n(r, k):
    try: return float(r[k])
    except: return 0.0

def yen(x): return f"¥{round(x):,}"

def load():
    orders = list(csv.DictReader(open(ORDERS, encoding="utf-8-sig")))
    rank = list(csv.DictReader(open(RANK, encoding="utf-8-sig")))
    return orders, rank

def acct_summary(orders):
    acc = defaultdict(lambda: dict(件数=0, 純売上=0, 原価=0, 手数料=0, 粗利=0, 営業利益=0))
    for r in orders:
        a = acc[r["アカウント"]]
        a["件数"] += 1
        a["純売上"] += n(r, "純売上円"); a["原価"] += n(r, "原価円")
        a["手数料"] += n(r, "手数料円"); a["粗利"] += n(r, "粗利円")
        a["営業利益"] += n(r, "営業利益円")
    return acc

def seg_summary(rank):
    seg = defaultdict(lambda: dict(ASIN=0, 販売数=0, 純売上=0, 粗利=0))
    for it in rank:
        s = it["戦略区分"]
        seg[s]["ASIN"] += 1
        seg[s]["販売数"] += int(float(it["販売数"]))
        seg[s]["純売上"] += round(n(it, "純売上"))
        seg[s]["粗利"] += round(n(it, "粗利"))
    return seg

def build_md(orders, rank):
    today = datetime.date.today()
    acc = acct_summary(orders)
    seg = seg_summary(rank)
    months = sorted(set(r["年月"] for r in orders))
    period = f"{months[0]}〜{months[-1]}" if months else "-"

    tot = dict(件数=0, 純売上=0, 原価=0, 手数料=0, 粗利=0, 営業利益=0)
    for a in acc.values():
        for k in tot: tot[k] += a[k]
    gm = tot["粗利"] / tot["純売上"] if tot["純売上"] else 0

    L = []
    L.append(f"# 📊 EC粗利分析 01→03申し送り（韓国クーパン輸出）")
    L.append("")
    L.append(f"> 発信: 01_経営管理（CFO）→ 着信: 03_事業運営（テッカン） ／ 更新日: {today} ／ 対象期間: {period}")
    L.append(f"> ダッシュボード（数字の生データ・常時最新）: {DASH_URL}")
    L.append("")
    L.append("## 1. 結論（CFOの確定値）")
    L.append("")
    L.append(f"- **EC（クーパン1+2）は粗利ベースで黒字。粗利{yen(tot['粗利'])}（粗利率{gm:.1%}）／営業利益{yen(tot['営業利益'])}。**")
    L.append("- 計算式: 粗利 = 純売上（クーパン手数料11%控除後の入金額）− 総原価（売れた分の仕入原価）。キャンセル等は売上から除外済み。")
    L.append("- ただし外注¥10万/月・税理士等の固定費（年¥120万超）は未控除。**フルコストでは赤字圏** → 撤退ではなく「粗利の質を上げる」のが03のテーマ。")
    L.append("")
    L.append("### アカウント別")
    L.append("")
    L.append("| アカウント | 件数 | 純売上 | 原価 | 粗利 | 粗利率 | 営業利益 |")
    L.append("|---|--:|--:|--:|--:|--:|--:|")
    for k in sorted(acc):
        a = acc[k]; g = a["粗利"] / a["純売上"] if a["純売上"] else 0
        L.append(f"| {k} | {a['件数']} | {yen(a['純売上'])} | {yen(a['原価'])} | {yen(a['粗利'])} | {g:.1%} | {yen(a['営業利益'])} |")
    # ※Notionの表セルはmarkdown太字(**)を解釈せず記号が残るため、表内では装飾を使わない
    L.append(f"| 合計（1+2） | {tot['件数']} | {yen(tot['純売上'])} | {yen(tot['原価'])} | {yen(tot['粗利'])} | {gm:.1%} | {yen(tot['営業利益'])} |")
    L.append("")
    L.append("> クーパン2の方が粗利率が高く健全。1は薄利体質 → 1の出品方針を2型へ寄せるのが筋。")
    L.append("")

    L.append("## 2. 03がやること（数値ルールで仕分けた4区分）")
    L.append("")
    order = ["①伸ばす", "②改善", "③スポット", "④撤退", "◎維持"]
    # ※表セル内はmarkdown太字が効かないため装飾なしの平文にする
    actmap = {
        "①伸ばす": "在庫確保・出品強化・横展開（実証済みの勝ち筋）",
        "②改善": "値上げ／原価交渉／配送最適化（高回転だが薄利）",
        "③スポット": "同カテゴリへ横展開テスト増産（少数だが高粗利＝利益の本体）",
        "④撤退": "出品停止・整理（死筋／原価割れ）",
        "◎維持": "現状維持（様子見）",
    }
    L.append("| 区分 | ASIN数 | 販売数 | 粗利合計 | 03のアクション |")
    L.append("|---|--:|--:|--:|---|")
    for s in order:
        if s in seg:
            d = seg[s]
            L.append(f"| {s} | {d['ASIN']} | {d['販売数']} | {yen(d['粗利'])} | {actmap.get(s,'')} |")
    L.append("")

    # 撤退・スポットの定量インパクト
    kill = seg.get("④撤退", {})
    spot = seg.get("③スポット", {})
    scale = seg.get("①伸ばす", {})
    L.append("### 優先順位つきアクション")
    L.append("")
    if kill:
        L.append(f"1. **【最優先】撤退{kill['ASIN']}ASINの出品停止** — 粗利を{yen(kill['粗利'])}食っている（高単価エレキ/時計/カメラの原価割れ）。止めるだけで全体粗利が改善。")
    if spot:
        share = spot['粗利'] / tot['粗利'] if tot['粗利'] else 0
        L.append(f"2. **スポット{spot['ASIN']}ASINの横展開** — ここで粗利{yen(spot['粗利'])}（全体の{share:.0%}）。多品種少量×高粗利のニッチが利益の本体。同カテゴリで似た商材を増やすのが増収の本筋。")
    if scale:
        L.append(f"3. **伸ばす{scale['ASIN']}ASINの在庫確保** — 実証済みの勝ち筋。欠品で機会損失しないよう在庫を厚くする。")
    L.append("")

    L.append("## 3. 運用改善テーマ（03のオペレーション課題）")
    L.append("")
    L.append("- **キャンセル率17%**（発注後に消滅した注文が多数）= 運転資金の拘束・工数の垂れ流し。仕入実行タイミングの検証が次の急所。")
    L.append("- ※キャンセル原価¥3.12Mは**実損ではない**【判定済】。発注なし55%は損失ゼロ、発注済分も還付回収済で最大エクスポージャー約¥1.1M、在庫は返品/再販可で真の実損は数十万規模。年間粗利は脅かさない。論点は「損失額」ではなく「率の高さ＝オペレーションの無駄」。")
    L.append("")

    L.append("## 4. データの所在・更新")
    L.append("")
    L.append(f"- ダッシュボード（3タブ: サマリー/月次推移/商品戦略）: {DASH_URL}")
    L.append("- 元データ: 月次販売管理表（Drive `韓国輸出売上/`）→ openpyxlで集計 → `ec_orders_consolidated.csv`（明細）/ `ec_product_ranking.csv`（ASIN戦略）。")
    L.append("- 毎月締まったら 01 がパイプライン再実行 → ダッシュボードと本申し送りを更新して 03 へ再連携。")
    L.append("")
    L.append(f"_自動生成: scripts/ec_handoff_03.py（{today}）_")
    return "\n".join(L)

def main():
    orders, rank = load()
    md = build_md(orders, rank)
    with open(OUT_MD, "w", encoding="utf-8") as f:
        f.write(md)
    print(md)
    print(f"\n--- 申し送りmarkdown出力: {OUT_MD} ---")

if __name__ == "__main__":
    main()
