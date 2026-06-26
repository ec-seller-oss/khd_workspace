#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
月次ランナー — 全カード/口座を一括処理し、月次にまとめる（記帳自動化の器）。
=================================================================
入力: MF一括CSV（当月＋過去数ヶ月）。最後のファイル＝当月。
出力:
  A. out/月次集約_YYYY-MM_MF仕訳インポート.csv … 全事業口座の事業仕訳（補助科目=口座名）
  B. コンソール＝月次UP前レビュー（口座別サマリ＋定例抜け＋初出＋未分類）
  C. out/月次サマリ_YYYY-MM.csv … 口座×区分×件数×金額（→04PL/05資金繰り連動の素）

使い方:
  python3 run_month.py 過去CSV... 当月CSV
  例) python3 run_month.py "fixtures/収入・支出詳細_2026-02-*.csv" ... fixtures/mf_bulk_2026-05.csv
"""
import csv
import os
import sys
import collections

from kicho_engine import _open_csv, nfkc, _amt, HIGH_AMOUNT
from journal_rules import MF_HEADER, CARD_PROFILES
from monthly_review import load_account, ym_of, review_account, yen

HERE = os.path.dirname(os.path.abspath(__file__))


def resolve_profile(acct):
    """口座名→profile。dictキー一致 or mf_account一致で解決。"""
    p = CARD_PROFILES.get(acct)
    if p:
        return p
    for v in CARD_PROFILES.values():
        if v.get("mf_account") == acct:
            return v
    return None


def accounts_in(path):
    s = set()
    for r in _open_csv(path):
        a = (r.get("保有金融機関") or "").strip()
        if a:
            s.add(a)
    return sorted(s)


def main():
    files = sys.argv[1:]
    if len(files) < 2:
        print("使い方: run_month.py 過去CSV... 当月CSV"); return
    *past_files, cur_file = files
    cur_ym = ym_of(cur_file)
    accts = accounts_in(cur_file)

    journal_rows = []                      # 全事業口座の事業仕訳（集約）
    summary = []                           # 口座×区分集計
    biz_reviews = []                       # 事業口座のレビュー
    kakei_total = 0
    busho_skip = []                        # 物販(別pipeline)
    mibunrui = []                          # 台帳未登録

    for acct in accts:
        prof = resolve_profile(acct)
        cur = load_account(cur_file, acct)
        n = len(cur); amt = sum(r["amount"] for r in cur)
        if not cur:
            continue
        if prof and prof.get("skip"):
            busho_skip.append((acct, n, amt)); summary.append((acct, "物販(別pipeline)", n, amt)); continue
        kubun = prof["default_kubun"] if prof else "未分類"

        if kubun in ("法人", "個人事業"):
            past = [(ym_of(f), load_account(f, acct)) for f in past_files]
            rv = review_account(cur, past)
            for r in rv["gyo"]:
                journal_rows.append(["", r["date"], r["勘定科目"], "", "", r["vendor"],
                                     r["税区分"], "無", r["amount"], prof.get("kashikata", "未払金"),
                                     prof.get("kashikata_hojo", acct), "", "", "対象外", "", r["amount"],
                                     f"{r['メモ']} {r['store']}"[:40], "", f"{acct}/自動仕訳"])
            biz_reviews.append((acct, rv))
            summary.append((acct, "事業", len(rv["gyo"]), sum(r["amount"] for r in rv["gyo"])))
            if rv["you"]:
                summary.append((acct, "要確認", len(rv["you"]), sum(r["amount"] for r in rv["you"])))
        elif kubun == "家計":
            kakei_total += amt; summary.append((acct, "家計", n, amt))
        else:
            mibunrui.append((acct, n, amt)); summary.append((acct, "未分類", n, amt))

    # ===== 出力A: 集約MFインポートCSV =====
    outdir = os.path.join(HERE, "out"); os.makedirs(outdir, exist_ok=True)
    mf_csv = os.path.join(outdir, f"月次集約_{cur_ym}_MF仕訳インポート.csv")
    with open(mf_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f); w.writerow(MF_HEADER); w.writerows(journal_rows)
    # 出力C: 月次サマリ
    sm_csv = os.path.join(outdir, f"月次サマリ_{cur_ym}.csv")
    with open(sm_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f); w.writerow(["口座", "区分", "件数", "金額"]); w.writerows(summary)

    # ===== 出力B: コンソール 月次UP前レビュー =====
    print("=" * 72)
    print(f" 月次クローズ前レビュー（全カード）— 当月 {cur_ym} / 過去{len(past_files)}ヶ月比較")
    print("=" * 72)
    print(f"\n■ 集約 事業仕訳: {len(journal_rows)}件 {yen(sum(r[8] for r in journal_rows))} → MF取込CSV")
    print(f"■ 家計合計: {yen(kakei_total)}（→05資金繰りの生活費）")
    if busho_skip:
        b = sum(x[2] for x in busho_skip)
        print(f"■ 物販仕入(別pipeline): {sum(x[1] for x in busho_skip)}件 {yen(b)}（amazon_to_mf_journal.py）")

    print("\n◆ 事業口座 UP前チェック")
    for acct, rv in biz_reviews:
        print(f"  ▸ {acct}: 事業{len(rv['gyo'])}件 / 要確認{len(rv['you'])}件")
        for m in rv["missing"]:
            print(f"      🔴 定例抜け: {m}（計上漏れ/タイミング確認）")
        for r in sorted(rv["shoshutsu"], key=lambda x: -x["amount"])[:5]:
            hi = " 🔴高額" if r["amount"] >= HIGH_AMOUNT else ""
            print(f"      ⚠️ 初出: {r['date']} {r['amount']:,} {r['store'][:24]}{hi}")
        for r in sorted(rv["you"], key=lambda x: -x["amount"])[:5]:
            print(f"      🟡 要確認: {r['date']} {r['amount']:,} {r['store'][:24]}")

    if mibunrui:
        print("\n◆ 台帳未登録の口座（仕分け方針が必要）")
        for acct, n, amt in mibunrui:
            print(f"  ? {acct}: {n}件 {yen(amt)} ← 事業/家計を台帳(CARD_PROFILES)に登録")

    print("\n" + "=" * 72)
    print(f"A 集約MFインポCSV : {mf_csv}")
    print(f"C 月次サマリ      : {sm_csv}")
    print(f"→ 🔴定例抜け・⚠️初出・🟡要確認 を確認してからMFへUP")
    print("=" * 72)


if __name__ == "__main__":
    main()
