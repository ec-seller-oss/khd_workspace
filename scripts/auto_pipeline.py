#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
無人化オーケストレータ：URL群 → 本査定 → 🟢だけ通知
================================================================
ポータル物件URL（健美家等）を一括で本査定し、🟢買い候補だけを通知（LINE/notes）。
入力＝URLリスト(引数 or queue_urls.txt)。スケジュール=cron/launchdで定期起動。
完全自動化の最終段：GASが収集→listing URLをスプシ/ファイルへ→本スクリプトが査定→通知。
"""
import os
import sys
import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import kenbiya_fetch
import rakumachi_fetch
import property_screener as screener
import reinfolib_client as rc

QUEUE = Path(__file__).parent / "queue_urls.txt"


def fetch_fields(url):
    """ポータルURL → fields。ドメインで自動振り分け。"""
    if "kenbiya.com" in url:
        return kenbiya_fetch.fetch(url)
    if "rakumachi.jp" in url:
        return rakumachi_fetch.fetch(url)
    raise ValueError(f"未対応ポータル: {url}")


def assess(url):
    f = fetch_fields(url)
    ref = f.get("_参考情報", {})
    # 路線価 自動（住所→reinfolib）
    pref, city, district = rc.parse_address(ref.get("所在地") or "")
    if pref and city:
        r = rc.lookup(pref, city, district)
        if not r.get("error"):
            f["路線価_円per_m2"] = r["相続税路線価相当"]
    res = screener.calc(f)
    _, verdict, mark, _ = screener.render_console(f, res)
    return {"url": url, "name": ref.get("所在地") or f.get("物件名"),
            "price": f.get("価格_円"), "yield": res.get("表面利回り"),
            "tochi": res.get("土地値割合"), "cf": res.get("CF率"),
            "mark": mark, "verdict": verdict, "fields": f, "res": res}


def line_notify(msg):
    token = os.environ.get("KHD_LINE_TOKEN")
    if not token:
        return False
    import urllib.request
    import urllib.parse
    req = urllib.request.Request("https://notify-api.line.me/api/notify",
                                 data=urllib.parse.urlencode({"message": msg}).encode(),
                                 headers={"Authorization": f"Bearer {token}"})
    try:
        urllib.request.urlopen(req, timeout=15)
        return True
    except Exception:
        return False


def run(urls):
    results = []
    for u in urls:
        try:
            results.append(assess(u))
        except Exception as e:
            results.append({"url": u, "name": "(取得失敗)", "mark": "⚠️", "verdict": str(e)[:40]})
    greens = [r for r in results if r["mark"] == "🟢"]

    def pc(x): return f"{x*100:.1f}%" if x is not None else "—"
    lines = [f"📥 無人査定 {datetime.date.today()}  対象{len(results)}件 → 🟢{len(greens)}"]
    for r in sorted(results, key=lambda x: (x.get("yield") or 0), reverse=True):
        lines.append(f"{r['mark']} {r['name']} 利{pc(r.get('yield'))}/土地値{pc(r.get('tochi'))} {r['verdict']}")
    report = "\n".join(lines)
    print(report)

    notes = Path.home() / "01_honbu_docs_automation" / ".company" / "secretary" / "notes"
    (notes / f"{datetime.date.today()}-auto-pipeline.md").write_text(
        "# 無人査定パイプライン\n\n```\n" + report + "\n```\n", encoding="utf-8")
    if greens:
        gmsg = "🟢買い候補:\n" + "\n".join(
            f"・{r['name']} 利{pc(r.get('yield'))} {r['url']}" for r in greens)
        line_notify(gmsg)
    return results


# 物件一覧スプシ（GASがlisting URLをO列に書き出す）
MATCH_SS = "1a0w6K-fi_BpTGGAVmB1lHqAJYPjnM4M8fw8Rs25ghnc"
MATCH_TAB = "物件マッチング一覧（医療テナント）"


def run_from_sheet():
    """スプシO列(物件URL)から未査定の不動産行を読んで本査定→ステータス/備考へ書戻し。"""
    import pickle
    import warnings
    warnings.filterwarnings("ignore")
    from googleapiclient.discovery import build
    creds = pickle.load(open(Path(__file__).parent / "sheets_token.pickle", "rb"))
    svc = build("sheets", "v4", credentials=creds)
    rows = svc.spreadsheets().values().get(
        spreadsheetId=MATCH_SS, range=f"{MATCH_TAB}!A1:O80").execute().get("values", [])
    targets = []
    for i, row in enumerate(rows, 1):
        if len(row) < 2 or row[1].strip() != "不動産":
            continue
        url = row[14].strip() if len(row) > 14 and row[14] else ""
        done = "本査定" in (row[12] if len(row) > 12 else "")
        if url.startswith("http") and not done:
            targets.append((i, url))
    if not targets:
        print("本査定対象URLなし（GASがO列にlisting URLを書くと処理対象になります）")
        return []
    results, data = [], []
    for (ri, url) in targets:
        try:
            r = assess(url)
        except Exception as e:
            r = {"url": url, "mark": "⚠️", "verdict": str(e)[:40], "name": url}
        results.append(r)
        def pc(x): return f"{x*100:.1f}%" if x is not None else "—"
        biko = f"本査定 {r['mark']}{r['verdict']}／利回り{pc(r.get('yield'))}／土地値{pc(r.get('tochi'))}"
        data.append({"range": f"{MATCH_TAB}!M{ri}", "values": [[f"本査定{r['mark']}"]]})
        data.append({"range": f"{MATCH_TAB}!N{ri}", "values": [[biko]]})
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=MATCH_SS, body={"valueInputOption": "USER_ENTERED", "data": data}).execute()
    # 通知共通
    _notify(results)
    return results


DRIVE_INBOX = Path.home() / "Library/CloudStorage/GoogleDrive-ec-seller@kikuchi-hd.net/マイドライブ/_査定受け箱"


def assess_pdf(pdf_path):
    """マイソクPDF → 抽出→路線価自動→本査定。"""
    import property_intake
    f, meta, _, _ = property_intake.parse(str(pdf_path))
    f["_参考情報"] = meta.get("_参考情報", {})
    addr = f["_参考情報"].get("所在地") or ""
    pref, city, district = rc.parse_address(addr)
    if pref and city:
        r = rc.lookup(pref, city, district)
        if not r.get("error"):
            f["路線価_円per_m2"] = r["相続税路線価相当"]
    res = screener.calc(f)
    _, verdict, mark, _ = screener.render_console(f, res)
    return {"url": str(pdf_path.name), "name": addr or f.get("物件名"),
            "price": f.get("価格_円"), "yield": res.get("表面利回り"),
            "tochi": res.get("土地値割合"), "cf": res.get("CF率"),
            "mark": mark, "verdict": verdict, "fields": f, "res": res}


def run_from_drive():
    """Drive受け箱の新規PDFを自動査定→processedへ移動→🟢通知。"""
    if not DRIVE_INBOX.exists():
        print(f"受け箱なし: {DRIVE_INBOX}"); return []
    proc = DRIVE_INBOX / "processed"; proc.mkdir(exist_ok=True)
    pdfs = sorted(p for p in DRIVE_INBOX.glob("*.pdf"))
    if not pdfs:
        print("受け箱に新規PDFなし"); return []
    results = []
    for p in pdfs:
        try:
            results.append(assess_pdf(p))
        except Exception as e:
            results.append({"url": p.name, "name": p.name, "mark": "⚠️", "verdict": str(e)[:40]})
        p.rename(proc / p.name)   # 処理済みへ移動（再査定防止）
    _notify(results)
    return results


def _notify(results):
    greens = [r for r in results if r["mark"] == "🟢"]
    def pc(x): return f"{x*100:.1f}%" if x is not None else "—"
    rep = [f"📥 無人本査定 {datetime.date.today()}  {len(results)}件 → 🟢{len(greens)}"]
    for r in sorted(results, key=lambda x: (x.get("yield") or 0), reverse=True):
        rep.append(f"{r['mark']} {r.get('name')} 利{pc(r.get('yield'))}/土地値{pc(r.get('tochi'))}")
    report = "\n".join(rep); print(report)
    notes = Path.home() / "01_honbu_docs_automation" / ".company" / "secretary" / "notes"
    (notes / f"{datetime.date.today()}-auto-pipeline.md").write_text(
        "# 無人本査定\n\n```\n" + report + "\n```\n", encoding="utf-8")
    # 専用DB『🤖自動査定結果』へ1件ずつ自動打ち出し（スマホで一覧）
    for r in results:
        notion_post_result(r)
    # 🟢は「最後まで」自動完走：融資資料デッキ生成＋物DB(SSoT)蓄積
    for r in greens:
        finish_green(r)
    if greens:
        body = "🟢買い候補（融資資料デッキ自動生成済）:\n" + "\n".join(
            f"・{r.get('name')} 利{pc(r.get('yield'))}"
            + ("　📊デッキ:_査定結果デッキ/" if r.get("deck") else "")
            + f"　{r['url']}" for r in greens)
        line_notify(body)
        # Googleタスク化（スマホにプッシュ）
        try:
            import gtasks
            gtasks.create_task(
                f"🤖🟢買い候補 {len(greens)}件（{datetime.date.today()}）自動査定",
                body + "\n\n一覧: https://www.notion.so/f3df9251e34c49bb9e493fb990325bf0")
        except Exception:
            pass


ASSESS_DB = "f3df9251e34c49bb9e493fb990325bf0"  # 🤖自動査定結果DB
DECK_OUT = Path.home() / "Library/CloudStorage/GoogleDrive-ec-seller@kikuchi-hd.net/マイドライブ/_査定結果デッキ"


def _safe(s):
    import re
    return re.sub(r"[/\\:*?\"<>|\s]+", "_", str(s or "物件"))[:40]


def finish_green(r):
    """🟢買い判定→最後まで自動：融資資料デッキ生成＋物DB(SSoT)蓄積。"""
    f, res = r.get("fields"), r.get("res")
    if not f:
        return
    # ① 融資資料デッキ生成（Driveの_査定結果デッキへ→スマホで開ける）
    try:
        import screen_property as sp
        import loan_deck
        DECK_OUT.mkdir(parents=True, exist_ok=True)
        deck = sp.assemble_deck(f, res, None)
        out = DECK_OUT / f"融資資料_{_safe(r.get('name'))}.pptx"
        loan_deck.build(deck, str(out))
        r["deck"] = str(out)
    except Exception:
        pass
    # ② 物DB(SSoT)へ蓄積
    try:
        import sheets_db
        sheets_db.upsert_property(f, res, r.get("verdict", "買い"))
    except Exception:
        pass


def _notion_token():
    import json
    try:
        cfg = json.load(open(Path.home() / ".claude.json"))
        return cfg.get("mcpServers", {}).get("notion", {}).get("env", {}).get("NOTION_TOKEN")
    except Exception:
        return None


def notion_post_result(r):
    """査定結果1件を専用DBへ追記。"""
    import json
    import urllib.request
    token = _notion_token()
    if not token:
        return
    def pc(x): return f"{x*100:.1f}%" if x is not None else "—"
    mark2name = {"🟢": "🟢買い", "🟡": "🟡要検討", "🔴": "🔴見送り", "⚠️": "⚠️要確認"}
    u = str(r.get("url") or "")
    entry = "Drive" if u.lower().endswith(".pdf") else ("URL" if u.startswith("http") else "住所")
    props = {
        "物件": {"title": [{"text": {"content": str(r.get("name") or "(無名)")[:90]}}]},
        "判定": {"select": {"name": mark2name.get(r.get("mark"), "⚠️要確認")}},
        "入口": {"select": {"name": entry}},
        "表面利回り": {"rich_text": [{"text": {"content": pc(r.get("yield"))}}]},
        "土地値割合": {"rich_text": [{"text": {"content": pc(r.get("tochi"))}}]},
        "CF率": {"rich_text": [{"text": {"content": pc(r.get("cf"))}}]},
        "所在地": {"rich_text": [{"text": {"content": str(r.get("name") or "")[:200]}}]},
        "査定日": {"date": {"start": datetime.date.today().isoformat()}},
        "メモ": {"rich_text": [{"text": {"content": str(r.get("verdict") or "")[:200]}}]},
    }
    if r.get("price"):
        props["価格"] = {"number": int(r["price"])}
    if u.startswith("http"):
        props["出典URL"] = {"url": u}
    payload = {"parent": {"database_id": ASSESS_DB}, "properties": props}
    req = urllib.request.Request(
        "https://api.notion.com/v1/pages", data=json.dumps(payload).encode(),
        headers={"Authorization": f"Bearer {token}", "Notion-Version": "2022-06-28",
                 "Content-Type": "application/json"}, method="POST")
    try:
        urllib.request.urlopen(req, timeout=20)
    except Exception:
        pass


if __name__ == "__main__":
    if "--all" in sys.argv:        # メール(スプシURL) + Drive受け箱 を両方
        run_from_sheet(); run_from_drive(); sys.exit(0)
    if "--drive" in sys.argv:
        run_from_drive(); sys.exit(0)
    if "--sheet" in sys.argv:
        run_from_sheet()
        sys.exit(0)
    urls = [a for a in sys.argv[1:] if a.startswith("http")]
    if not urls and QUEUE.exists():
        urls = [u.strip() for u in QUEUE.read_text().splitlines() if u.strip()]
    if not urls:
        print("使い方: auto_pipeline.py <URL...> / --sheet(スプシO列) / queue_urls.txt")
        sys.exit(1)
    run(urls)
