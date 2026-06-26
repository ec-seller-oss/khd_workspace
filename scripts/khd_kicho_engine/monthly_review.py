#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
月次クローズ前レビュー — MFインポート(UP)前に過去と比較して妥当性をチェックする。
=================================================================
目的: アップロード前に「定例の抜け／初出ベンダー／科目別の増減」を自動で炙り出し、
      記帳の間違い確認を楽にする。

検査:
  ① 当月の事業仕訳（MFインポートCSVに出る分）
  ② 定例ベンダーの抜け  … 過去に毎月あるのに当月に無い = 計上漏れ/カード変更の疑い
  ③ 初出ベンダー        … 過去に無い新顔 = 架空請求/新規契約の確認対象
  ④ 科目別 当月 vs 過去平均 … 例月から外れた増減
  ⑤ 要確認(未HIT)・高額

使い方:
  python3 monthly_review.py --account MB \
      "fixtures/収入・支出詳細_2026-02-01_2026-02-28.csv" \
      "fixtures/収入・支出詳細_2026-03-01_2026-03-31.csv" \
      "fixtures/収入・支出詳細_2026-04-01_2026-04-30.csv" \
      fixtures/mf_bulk_2026-05.csv
  （最後のファイル=当月。それ以前=過去比較データ）
"""
import csv
import os
import re
import sys
import argparse
import collections

from kicho_engine import _open_csv, nfkc, _amt, _date, norm_vendor, HIGH_AMOUNT
from journal_rules import MF_HEADER, CARD_PROFILES, classify

HERE = os.path.dirname(os.path.abspath(__file__))


def ym_of(path):
    m = re.search(r"(20\d{2})[-_/]?(\d{2})", os.path.basename(path))
    return f"{m.group(1)}-{m.group(2)}" if m else os.path.basename(path)


def load_account(path, account):
    out = []
    for r in _open_csv(path):
        if (r.get("保有金融機関") or "").strip() != account:
            continue
        if (r.get("振替") or "") in ("1", "TRUE", "true"):
            continue
        store = nfkc(r.get("内容") or "")
        amt = abs(_amt(r.get("金額（円）")))
        if amt == 0:
            continue
        c = classify(store, {})
        out.append({"date": _date(r.get("日付")), "store": store, "amount": amt,
                    "vendor": norm_vendor(store), **c})
    return out


def yen(n):
    return f"{n:,}円"


def review_account(cur, past):
    """1口座の当月分析を返す。past=[(ym,rows),...]。run_month/CLI共用。"""
    gyo = [r for r in cur if r["区分"] in ("法人", "個人事業")]
    you = [r for r in cur if r["区分"] == "不明"]
    kakei = [r for r in cur if r["区分"] == "家計"]
    past_vendor_months = collections.Counter()
    vendor_label = {}
    for _, rows in past:
        for v in set(r["vendor"] for r in rows):
            past_vendor_months[v] += 1
        for r in rows:
            vendor_label.setdefault(r["vendor"], r["store"][:24])
    n_past = max(1, len(past))
    teirei = {v for v, c in past_vendor_months.items() if c >= max(2, n_past - 1)}
    cur_vendors = set(r["vendor"] for r in cur)
    missing = [vendor_label.get(v, v) for v in sorted(teirei - cur_vendors)]
    shoshutsu = [r for r in cur if r["vendor"] not in past_vendor_months]
    return {"gyo": gyo, "you": you, "kakei": kakei,
            "missing": missing, "shoshutsu": shoshutsu}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--account", default="MB")
    ap.add_argument("files", nargs="+", help="過去...当月 の順。最後が当月")
    a = ap.parse_args()

    months = [(ym_of(f), load_account(f, a.account)) for f in a.files]
    *past, (cur_ym, cur) = months
    prof = CARD_PROFILES.get(a.account, {"kashikata": "未払金", "kashikata_hojo": a.account})

    print("=" * 70)
    print(f" 月次クローズ前レビュー — {a.account} / 当月 {cur_ym}（過去{len(past)}ヶ月と比較）")
    print("=" * 70)

    # ① 当月 事業仕訳
    gyo = [r for r in cur if r["区分"] in ("法人", "個人事業")]
    kakei = [r for r in cur if r["区分"] == "家計"]
    you = [r for r in cur if r["区分"] == "不明"]
    print(f"\n① 当月の事業仕訳（MFインポート対象）: {len(gyo)}件 {yen(sum(r['amount'] for r in gyo))}")
    for r in sorted(gyo, key=lambda x: -x["amount"]):
        print(f"   {r['date']} {r['amount']:>7,}  {r['勘定科目']:6} ← {r['store'][:30]}")

    # ②③ 定例の抜け / 初出
    past_vendor_months = collections.Counter()
    vendor_label = {}
    for _, rows in past:
        for v in set(r["vendor"] for r in rows):
            past_vendor_months[v] += 1
        for r in rows:
            vendor_label.setdefault(r["vendor"], r["store"][:24])
    teirei = {v for v, c in past_vendor_months.items() if c >= max(2, len(past) - 1)}  # 大半の月で出現
    cur_vendors = set(r["vendor"] for r in cur)

    missing = sorted(teirei - cur_vendors)
    print(f"\n② 定例の抜け（過去ほぼ毎月あるのに当月に無い）: {len(missing)}件")
    if not missing:
        print("   ✅ 抜けなし")
    for v in missing:
        print(f"   🔴 {vendor_label.get(v, v)}  ← 計上漏れ/カード変更/解約 を確認")

    shoshutsu = [r for r in cur if r["vendor"] not in past_vendor_months]
    print(f"\n③ 初出ベンダー（過去に無い新顔）: {len(shoshutsu)}件")
    if not shoshutsu:
        print("   ✅ 新顔なし")
    for r in sorted(shoshutsu, key=lambda x: -x["amount"]):
        print(f"   ⚠️ {r['date']} {r['amount']:>7,}  {r['store'][:30]}  ← 架空請求/新規契約を確認")

    # ④ 科目別 当月 vs 過去平均
    def by_kamoku(rows):
        d = collections.Counter()
        for r in rows:
            if r["区分"] in ("法人", "個人事業"):
                d[r["勘定科目"]] += r["amount"]
        return d
    cur_k = by_kamoku(cur)
    past_k_avg = collections.Counter()
    for _, rows in past:
        for k, v in by_kamoku(rows).items():
            past_k_avg[k] += v
    for k in past_k_avg:
        past_k_avg[k] = round(past_k_avg[k] / len(past))
    print(f"\n④ 科目別 当月 vs 過去平均")
    for k in sorted(set(cur_k) | set(past_k_avg), key=lambda x: -cur_k.get(x, 0)):
        c, p = cur_k.get(k, 0), past_k_avg.get(k, 0)
        d = c - p
        mark = "→" if abs(d) < max(3000, p * 0.3) else ("🔺" if d > 0 else "🔻")
        print(f"   {k:8} 当月{c:>8,} / 平均{p:>8,}  {mark}{d:+,}")

    # ⑤ 要確認・高額
    print(f"\n⑤ 要確認(未HIT): {len(you)}件 {yen(sum(r['amount'] for r in you))}")
    for r in sorted(you, key=lambda x: -x["amount"]):
        hi = " 🔴高額" if r["amount"] >= HIGH_AMOUNT else ""
        print(f"   {r['date']} {r['amount']:>7,}  {r['store'][:34]}{hi}")

    # 当月 MF仕訳インポートCSV 出力
    outdir = os.path.join(HERE, "out")
    os.makedirs(outdir, exist_ok=True)
    out = os.path.join(outdir, f"{a.account}_{cur_ym}_MF仕訳インポート.csv")
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(MF_HEADER)
        for r in gyo:
            w.writerow(["", r["date"], r["勘定科目"], "", "", r["vendor"], r["税区分"], "無",
                        r["amount"], "未払金", prof.get("kashikata_hojo", a.account), "", "",
                        "対象外", "", r["amount"], f"{r['メモ']} {r['store']}"[:40], "",
                        f"{a.account}/自動仕訳"])
    print("\n" + "=" * 70)
    print(f"UP前チェック完了。問題なければ取込: {out}")
    print(f"  → 🔴定例抜け {len(missing)}件 / ⚠️初出 {len(shoshutsu)}件 / 要確認 {len(you)}件 を確認してからUP")
    print("=" * 70)


if __name__ == "__main__":
    main()
