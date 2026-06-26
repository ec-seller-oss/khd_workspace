#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MF一括CSV → 事業経費のMF仕訳インポートCSV を生成（事業仕訳が実際に埋まる絵）。
家計/振替/未HITは出さない（事業=法人/個人事業のみ）。Amazon物販仕入は対象外
（=amazon_to_mf_journal.py の注文履歴パイプラインが担当。二重計上防止）。
"""
import csv
import os
import sys
from kicho_engine import _open_csv, nfkc, _date, _amt, norm_vendor
from journal_rules import MF_HEADER, classify

path = sys.argv[1]
out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(__file__), "out/事業仕訳_MF一括.csv")
prof = {"default_kubun": "?", "kashikata": "未払金"}

rows, total = [], 0
for r in _open_csv(path):
    if (r.get("振替") or "") in ("1", "TRUE", "true"):
        continue
    store = nfkc(r.get("内容") or "")
    amt = abs(_amt(r.get("金額（円）")))
    if amt == 0:
        continue
    acct = (r.get("保有金融機関") or "").strip()
    c = classify(store, prof)
    if c["区分"] not in ("法人", "個人事業"):
        continue
    rows.append(["", _date(r.get("日付")), c["勘定科目"], "", "", norm_vendor(store),
                 c["税区分"], "無", amt,
                 "未払金", acct, "", "", "対象外", "", amt,
                 f"{c['メモ']} {store}"[:40], "", f"{acct}/自動仕訳(MF抽出)"])
    total += amt

with open(out, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.writer(f)
    w.writerow(MF_HEADER)
    w.writerows(rows)

print(f"✅ 事業仕訳 {len(rows)}件 / 借方計 {total:,}円  → {out}\n")
print(f"{'取引日':10} {'借方科目':8} {'金額':>8}  摘要")
for r in rows:
    print(f"{r[1]:10} {r[2]:8} {r[8]:>8,}  {r[16][:36]}")
