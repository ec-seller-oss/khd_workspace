# -*- coding: utf-8 -*-
"""
EC売れ筋商品分析＆出品戦略 数値指針 (01_経営管理 → 03へ申し送り)
ec_orders_consolidated.csv を読み、商品(ASIN)別に粗利・回転を集計し
「伸ばす/改善/撤退」を数値ルールで自動仕分け。03事業運営への指針を出す。
"""
import csv
from collections import defaultdict

SRC = "/Users/kikuchikenta/01_honbu_docs_automation/ec_orders_consolidated.csv"

rows = list(csv.DictReader(open(SRC, encoding="utf-8-sig")))
def f(r,k):
    try: return float(r[k])
    except: return 0.0

# ===== ASIN別集計 =====
prod = defaultdict(lambda: dict(商品名="", アカ=set(), 月=set(), 販売数=0, 件数=0,
                                純売上=0, 原価=0, 粗利=0, 営業利益=0))
for r in rows:
    asin = r["ASIN"].strip() or "(ASIN無)"
    p = prod[asin]
    if not p["商品名"]: p["商品名"] = r["商品名"]
    p["アカ"].add(r["アカウント"]); p["月"].add(r["年月"])
    p["販売数"] += int(f(r,"数量")); p["件数"] += 1
    p["純売上"] += f(r,"純売上円"); p["原価"] += f(r,"原価円")
    p["粗利"] += f(r,"粗利円"); p["営業利益"] += f(r,"営業利益円")

items = []
for asin,p in prod.items():
    gm = p["粗利"]/p["純売上"] if p["純売上"] else 0
    items.append(dict(ASIN=asin, 商品名=p["商品名"], 販売数=p["販売数"], 件数=p["件数"],
        月数=len(p["月"]), アカ="/".join(sorted(p["アカ"])),
        純売上=round(p["純売上"]), 粗利=round(p["粗利"]), 粗利率=round(gm,3),
        営業利益=round(p["営業利益"]), 粗利単価=round(p["粗利"]/p["件数"]) if p["件数"] else 0))

print(f"取扱ASIN総数: {len(items)}  / 総注文: {len(rows)}件")

# ===== 売れ筋TOP（粗利額順）=====
print("\n"+"="*112)
print("【売れ筋TOP20｜年間粗利額順】★=伸ばす候補(粗利率10%↑×複数月)  ▲=要改善(粗利率5%未満)")
print("="*112)
print(f"{'商品名':32s}{'数':>4s}{'月':>3s}{'アカ':>10s}{'純売上':>9s}{'粗利':>8s}{'率':>6s}{'判定':>6s}")
for it in sorted(items, key=lambda x:-x["粗利"])[:20]:
    if it["粗利率"]>=0.10 and it["月数"]>=2: tag="★伸ばす"
    elif it["粗利率"]<0.05: tag="▲改善"
    else: tag="◎維持"
    print(f"{it['商品名'][:30]:32s}{it['販売数']:>4d}{it['月数']:>3d}{it['アカ']:>10s}{it['純売上']:>9,d}{it['粗利']:>8,d}{it['粗利率']:>6.0%}{tag:>7s}")

# ===== 数値ルールで戦略仕分け =====
scale, improve, kill, spot = [], [], [], []
for it in items:
    v, gm, gp = it["販売数"], it["粗利率"], it["粗利"]
    if v>=3 and gm>=0.10 and gp>0:          scale.append(it)     # 伸ばす(実証済み勝ち筋)
    elif v>=3 and 0<gm<0.05:                improve.append(it)   # 高回転・低利益→値上/原価交渉
    elif (v<=1 and gm<0.05) or gp<0:        kill.append(it)      # 死筋/赤字→出品停止
    elif gm>=0.20 and v<=2 and gp>0:        spot.append(it)      # 高利益スポット→横展開テスト

def block(title, lst, key, n=12, desc=""):
    print("\n"+"="*112); print(f"【{title}】{desc}  該当{len(lst)}ASIN")
    print("="*112)
    s_v=sum(x["販売数"] for x in lst); s_g=sum(x["粗利"] for x in lst)
    print(f"  小計: 販売{s_v}個 / 粗利{s_g:,}円")
    for it in sorted(lst, key=key)[:n]:
        print(f"  {it['商品名'][:34]:36s} 数{it['販売数']:>3d} 月{it['月数']:>2d} 純売上{it['純売上']:>8,d} 粗利{it['粗利']:>7,d}({it['粗利率']:>5.0%}) {it['アカ']}")

block("① 伸ばす｜実証済み勝ち筋(数量3↑×粗利率10%↑)", scale, lambda x:-x["粗利"],
      desc="→ 03へ:在庫確保・出品強化・横展開の最優先")
block("② 改善｜高回転だが薄利(数量3↑×粗利率5%未満)", improve, lambda x:-x["販売数"],
      desc="→ 03へ:値上げ/原価交渉/配送最適化で粗利率改善")
block("③ 高利益スポット｜少数だが高粗利(粗利率20%↑×数量2以下)", spot, lambda x:-x["粗利率"],
      desc="→ 03へ:同カテゴリ横展開でテスト増産")
block("④ 撤退候補｜死筋/赤字(単発×薄利 or 粗利マイナス)", kill, lambda x:x["粗利"],
      desc="→ 03へ:出品停止・整理")

# CSV出力(03申し送り用)
OUT = "/Users/kikuchikenta/01_honbu_docs_automation/ec_product_ranking.csv"
with open(OUT,"w",newline="",encoding="utf-8-sig") as fo:
    w = csv.DictWriter(fo, fieldnames=["ASIN","商品名","販売数","件数","月数","アカ","純売上","粗利","粗利率","営業利益","粗利単価","戦略区分"])
    w.writeheader()
    for it in sorted(items, key=lambda x:-x["粗利"]):
        v,gm,gp = it["販売数"],it["粗利率"],it["粗利"]
        if v>=3 and gm>=0.10 and gp>0: seg="①伸ばす"
        elif v>=3 and 0<gm<0.05: seg="②改善"
        elif (v<=1 and gm<0.05) or gp<0: seg="④撤退"
        elif gm>=0.20 and v<=2 and gp>0: seg="③スポット"
        else: seg="◎維持"
        it2=dict(it); it2["戦略区分"]=seg; w.writerow(it2)
print(f"\n商品ランキングCSV出力: {OUT}")
