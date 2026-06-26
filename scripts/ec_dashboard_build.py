# -*- coding: utf-8 -*-
"""
📊EC粗利ダッシュボード ビルダー (01_経営管理)
ec_orders_consolidated.csv / ec_product_ranking.csv を読み、
専用Google Sheetsに「サマリー/月次推移/商品戦略」タブを構築・更新する。
初回=新規作成しID保存。2回目以降=同じスプシを更新（毎月締めるたびに再実行）。
auth: sheets_token.pickle (scope=spreadsheets)
"""
import os, csv, pickle, datetime
from collections import defaultdict
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

BASE = "/Users/kikuchikenta/01_honbu_docs_automation"
SCR = f"{BASE}/scripts"
ORDERS = f"{BASE}/ec_orders_consolidated.csv"
RANK = f"{BASE}/ec_product_ranking.csv"
TOKEN = f"{SCR}/sheets_token.pickle"
IDFILE = f"{SCR}/ec_dashboard_id.txt"
TITLE = "📊EC粗利ダッシュボード_韓国輸出(01本部)"

def creds():
    with open(TOKEN, "rb") as f:
        c = pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN, "wb") as f:
            pickle.dump(c, f)
    return c

def load():
    orders = list(csv.DictReader(open(ORDERS, encoding="utf-8-sig")))
    rank = list(csv.DictReader(open(RANK, encoding="utf-8-sig")))
    return orders, rank

def n(r,k):
    try: return float(r[k])
    except: return 0.0

def build_summary(orders, rank):
    # アカウント別＋合計
    acc = defaultdict(lambda: dict(件数=0,純売上=0,原価=0,手数料=0,粗利=0,営業利益=0))
    for r in orders:
        a=acc[r["アカウント"]]
        a["件数"]+=1; a["純売上"]+=n(r,"純売上円"); a["原価"]+=n(r,"原価円")
        a["手数料"]+=n(r,"手数料円"); a["粗利"]+=n(r,"粗利円"); a["営業利益"]+=n(r,"営業利益円")
    rows=[["📊 EC粗利ダッシュボード（韓国クーパン輸出 1+2）"],
          [f"更新日: {datetime.date.today()}  / 計算式: 粗利=純売上(手数料11%控除後)−総原価  / キャンセル等は売上除外"],
          [],
          ["■ 通年サマリー（2025）"],
          ["アカウント","件数","純売上","原価","手数料","粗利","粗利率","営業利益"]]
    tot=dict(件数=0,純売上=0,原価=0,手数料=0,粗利=0,営業利益=0)
    for k in sorted(acc):
        a=acc[k]; gm=a["粗利"]/a["純売上"] if a["純売上"] else 0
        rows.append([k,a["件数"],round(a["純売上"]),round(a["原価"]),round(a["手数料"]),round(a["粗利"]),f"{gm:.1%}",round(a["営業利益"])])
        for x in tot: tot[x]+=a[x]
    gm=tot["粗利"]/tot["純売上"] if tot["純売上"] else 0
    rows.append(["合計(1+2)",tot["件数"],round(tot["純売上"]),round(tot["原価"]),round(tot["手数料"]),round(tot["粗利"]),f"{gm:.1%}",round(tot["営業利益"])])
    rows += [[],["※営業利益=粗利−配送料。外注¥10万/月・税理士等の固定費(年¥120万超)は未控除→フルコストでは赤字圏"],[]]
    # 戦略区分サマリー
    seg=defaultdict(lambda: dict(ASIN=0,販売数=0,粗利=0))
    for it in rank:
        s=it["戦略区分"]; seg[s]["ASIN"]+=1; seg[s]["販売数"]+=int(float(it["販売数"])); seg[s]["粗利"]+=round(n(it,"粗利"))
    rows += [["■ 商品戦略サマリー（591ASIN）→ 03事業運営へ申し送り"],
             ["戦略区分","ASIN数","販売数","粗利合計","アクション"]]
    actmap={"①伸ばす":"在庫確保・出品強化・横展開","②改善":"値上げ/原価交渉/配送最適化",
            "③スポット":"同カテゴリ横展開でテスト増産","④撤退":"出品停止・整理","◎維持":"現状維持"}
    for s in ["①伸ばす","②改善","③スポット","④撤退","◎維持"]:
        if s in seg:
            d=seg[s]; rows.append([s,d["ASIN"],d["販売数"],d["粗利"],actmap.get(s,"")])
    rows += [[],["■ CFO重大発見"],
             ["① 撤退候補123ASINが粗利を▲187,089円食う → 出品停止で粗利+20%（最優先）"],
             ["② 利益の本体=多品種少量×高粗利のニッチ商品。スポット237ASINで粗利の53%"],
             ["③ キャンセル160件に原価¥3,123,696が紐づく→仕入実行タイミングの検証が次の急所"]]
    return rows

def build_monthly(orders):
    ma=defaultdict(lambda: dict(件数=0,純売上=0,原価=0,手数料=0,粗利=0,営業利益=0))
    for r in orders:
        k=(r["年月"],r["アカウント"]); a=ma[k]
        a["件数"]+=1; a["純売上"]+=n(r,"純売上円"); a["原価"]+=n(r,"原価円")
        a["手数料"]+=n(r,"手数料円"); a["粗利"]+=n(r,"粗利円"); a["営業利益"]+=n(r,"営業利益円")
    rows=[["■ 月次推移（アカウント別＋月合計）"],
          ["年月","アカウント","件数","純売上","原価","手数料","粗利","粗利率","営業利益"]]
    months=sorted(set(k[0] for k in ma))
    for m in months:
        msum=dict(件数=0,純売上=0,原価=0,手数料=0,粗利=0,営業利益=0)
        for acct in ["クーパン1","クーパン2"]:
            if (m,acct) in ma:
                a=ma[(m,acct)]; gm=a["粗利"]/a["純売上"] if a["純売上"] else 0
                rows.append([m,acct,a["件数"],round(a["純売上"]),round(a["原価"]),round(a["手数料"]),round(a["粗利"]),f"{gm:.1%}",round(a["営業利益"])])
                for x in msum: msum[x]+=a[x]
        gm=msum["粗利"]/msum["純売上"] if msum["純売上"] else 0
        rows.append([m,"合計",msum["件数"],round(msum["純売上"]),round(msum["原価"]),round(msum["手数料"]),round(msum["粗利"]),f"{gm:.1%}",round(msum["営業利益"])])
    return rows

def build_products(rank):
    rows=[["■ 商品別ランキング（粗利額順）｜戦略区分つき"],
          ["戦略区分","商品名","ASIN","アカ","販売数","月数","純売上","粗利","粗利率","粗利単価"]]
    for it in sorted(rank, key=lambda x:-n(x,"粗利")):
        rows.append([it["戦略区分"],it["商品名"],it["ASIN"],it["アカ"],
                     int(float(it["販売数"])),int(float(it["月数"])),round(n(it,"純売上")),
                     round(n(it,"粗利")),f"{n(it,'粗利率'):.0%}",round(n(it,"粗利単価"))])
    return rows

def col_letter(n):
    s=""
    while n>0:
        n,r=divmod(n-1,26); s=chr(65+r)+s
    return s

def write_tab(svc, sid, tab, rows):
    end=col_letter(max(len(r) for r in rows)); rng=f"{tab}!A1:{end}{len(rows)}"
    svc.spreadsheets().values().update(spreadsheetId=sid, range=rng,
        valueInputOption="USER_ENTERED", body={"values":rows}).execute()

def main():
    orders,rank=load()
    svc=build("sheets","v4",credentials=creds())
    tabs={"サマリー":build_summary(orders,rank),"月次推移":build_monthly(orders),"商品戦略":build_products(rank)}
    sid=None
    if os.path.exists(IDFILE):
        sid=open(IDFILE).read().strip() or None
    if not sid:
        meta={"properties":{"title":TITLE},
              "sheets":[{"properties":{"title":t}} for t in tabs]}
        ss=svc.spreadsheets().create(body=meta, fields="spreadsheetId").execute()
        sid=ss["spreadsheetId"]; open(IDFILE,"w").write(sid)
        print(f"新規作成: {sid}")
    else:
        # 既存タブ確認、無ければ追加、あればクリア
        info=svc.spreadsheets().get(spreadsheetId=sid).execute()
        have={s["properties"]["title"] for s in info["sheets"]}
        reqs=[{"addSheet":{"properties":{"title":t}}} for t in tabs if t not in have]
        if reqs: svc.spreadsheets().batchUpdate(spreadsheetId=sid,body={"requests":reqs}).execute()
        for t in tabs:
            svc.spreadsheets().values().clear(spreadsheetId=sid,range=f"{t}!A1:Z2000").execute()
    for t,rows in tabs.items():
        write_tab(svc,sid,t,rows)
    print(f"URL: https://docs.google.com/spreadsheets/d/{sid}/edit")
    return sid

if __name__=="__main__":
    main()
