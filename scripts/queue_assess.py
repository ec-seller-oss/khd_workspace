#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自動収集 → 査定キュー（業者メール物件 → 一次査定）
================================================================
物件マッチング一覧スプシの「不動産」行（GASが業者メールから収集）を、
件名から 所在/価格/利回り をパース → reinfolibで相場付与 → 一次判定 →
スプシのステータス/備考へ書き戻し、🟢候補をキュー出力。
マイソク以外（業者メール）からの情報集約＝North Starの収集レイヤ。
"""
import re
import pickle
import warnings
import datetime
from pathlib import Path

warnings.filterwarnings("ignore")
import reinfolib_client as rc

SS = "1a0w6K-fi_BpTGGAVmB1lHqAJYPjnM4M8fw8Rs25ghnc"
TAB = "物件マッチング一覧（医療テナント）"
TOKEN = Path(__file__).parent / "sheets_token.pickle"
YIELD_GREEN, YIELD_YELLOW = 0.07, 0.06


def _svc():
    from googleapiclient.discovery import build
    creds = pickle.load(open(TOKEN, "rb"))
    return build("sheets", "v4", credentials=creds)


def parse_price(t):
    m = re.search(r"(\d+)\s*億\s*([\d,]+)?\s*万?円?", t)
    if m:
        oku = int(m.group(1)) * 100_000_000
        man = int((m.group(2) or "0").replace(",", "")) * 10000
        return oku + man
    m = re.search(r"([\d,]+(?:\.\d+)?)\s*万円", t)
    if m:
        return int(float(m.group(1).replace(",", "")) * 10000)
    return None


def parse_yield(t):
    # 「利回り」ラベル付きを最優先
    m = re.search(r"(?:満室想定)?利回り[^\d]{0,4}([\d.]+)\s*[%％]?", t)
    if m:
        return float(m.group(1)) / 100
    # ラベル無し：土地値/積算/ローン等を除外し、妥当レンジ(3〜20%)の%を採用
    for mm in re.finditer(r"([\d.]+)\s*[%％]", t):
        pre = t[max(0, mm.start() - 8):mm.start()]
        if any(w in pre for w in ["土地値", "積算", "土地", "ローン", "建ぺい", "容積", "返済"]):
            continue
        v = float(mm.group(1)) / 100
        if 0.03 <= v <= 0.20:
            return v
    return None


NOISE = ["ライブ", "セミナー", "レポート", "まとめ", "ご案内", "預かり", "REIT", "マーケット", "注目記事"]


def assess_row(subject):
    """件名 → 査定dict or None(物件でない)。"""
    if any(w in subject for w in NOISE):
        return None
    pref, city, district = rc.parse_address(subject)
    price = parse_price(subject)
    yld = parse_yield(subject)
    if not pref and price is None and yld is None:
        return None
    soba = None
    if pref and city:
        r = rc.lookup(pref, city)
        if not r.get("error"):
            soba = r.get("相続税路線価相当")
    # 一次判定（一棟収益＝利回り主軸）
    if yld is None:
        mark, verdict = "❔", "要詳細(利回り不明)"
    elif yld >= YIELD_GREEN:
        mark, verdict = "🟢", "買い候補"
    elif yld >= YIELD_YELLOW:
        mark, verdict = "🟡", "要検討"
    else:
        mark, verdict = "🔴", "利回り低"
    return {"pref": pref, "city": city, "price": price, "yield": yld,
            "soba": soba, "mark": mark, "verdict": verdict}


def run(write=True):
    svc = _svc()
    rows = svc.spreadsheets().values().get(
        spreadsheetId=SS, range=f"{TAB}!A1:N60").execute().get("values", [])
    queue, updates = [], []
    for i, row in enumerate(rows, 1):
        if len(row) < 2 or row[1].strip() != "不動産":
            continue
        if (len(row) > 12 and "査定" in (row[12] or "")):
            continue  # 既査定はスキップ
        subject = row[3] if len(row) > 3 else ""
        a = assess_row(subject)
        if not a:
            continue
        loc = (a["pref"] or "") + (a["city"] or "")
        s_price = f"{a['price']:,}円" if a["price"] else "—"
        s_yield = f"{a['yield']*100:.2f}%" if a["yield"] else "—"
        s_soba = f"{a['soba']:,}円/㎡" if a["soba"] else "—"
        biko = f"一次査定 {a['mark']}{a['verdict']}／価格{s_price}／利回り{s_yield}／相場路線価相当{s_soba}"
        status = f"査定済{a['mark']}"
        updates.append((i, status, biko))
        queue.append({"row": i, "loc": loc or "(地名不明)", "subj": subject[:30], **a})

    if write and updates:
        data = []
        for (ri, st, bk) in updates:
            data.append({"range": f"{TAB}!M{ri}", "values": [[st]]})
            data.append({"range": f"{TAB}!N{ri}", "values": [[bk]]})
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SS, body={"valueInputOption": "USER_ENTERED", "data": data}).execute()
    return queue


def render(queue):
    g = [q for q in queue if q["mark"] == "🟢"]
    y = [q for q in queue if q["mark"] == "🟡"]
    queue.sort(key=lambda q: (q["yield"] or 0), reverse=True)
    L = ["=" * 62, f"  📥 自動収集→査定キュー（業者メール / {datetime.date.today()}）", "=" * 62,
         f"  査定対象 {len(queue)}件 → 🟢{len(g)} / 🟡{len(y)} / その他{len(queue)-len(g)-len(y)}",
         "-" * 62]
    for q in queue:
        py = f"{q['yield']*100:.2f}%" if q['yield'] else "—"
        pr = f"{q['price']:,}円" if q['price'] else "—"
        L.append(f"  {q['mark']} {q['loc']}  利{py} / 価{pr}")
        L.append(f"      {q['subj']}")
    L.append("-" * 62)
    L.append("  ※ 件名パース＋reinfolib相場。詳細(面積/築年)は資料リンクから要取得。")
    L.append("=" * 62)
    return "\n".join(L)


if __name__ == "__main__":
    import sys
    q = run(write=("--dry" not in sys.argv))
    out = render(q)
    print(out)
    notes = Path.home() / "01_honbu_docs_automation" / ".company" / "secretary" / "notes"
    (notes / f"{datetime.date.today()}-assess-queue.md").write_text(
        "# 自動収集→査定キュー\n\n```\n" + out + "\n```\n", encoding="utf-8")
