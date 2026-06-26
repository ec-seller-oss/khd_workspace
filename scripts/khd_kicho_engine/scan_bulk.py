#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MF一括CSV 全口座を分類器にかけ、事業経費の所在を発見する（RULES育成用）。"""
import sys
import collections
from kicho_engine import _open_csv, nfkc
from journal_rules import classify

path = sys.argv[1]
prof = {"default_kubun": "?", "kashikata": "未払金", "kashikata_hojo": ""}

by_kubun = collections.Counter()
gyomu, mihit = [], []
for r in _open_csv(path):
    store = nfkc(r.get("内容") or "")
    amt = abs(int(float((r.get("金額（円）") or "0").replace(",", "") or 0)))
    acct = (r.get("保有金融機関") or "").strip()
    furikae = (r.get("振替") or "") in ("1", "TRUE", "true")
    if amt == 0 or furikae:
        continue
    c = classify(store, prof)
    by_kubun[c["区分"]] += 1
    if c["区分"] in ("法人", "個人事業"):
        gyomu.append((amt, acct, store, c["勘定科目"]))
    elif c["区分"] == "不明":
        mihit.append((amt, acct, store))

print("=== 区分別 件数 ===")
for k, v in by_kubun.most_common():
    print(f"  {k}: {v}")

print("\n=== 事業HIT（=MF仕訳に出る取引）===")
for amt, acct, store, kamoku in sorted(gyomu, reverse=True):
    print(f"  {amt:>8,}  [{acct}] {store[:30]} → {kamoku}")

print("\n=== 未HIT（要確認）高額TOP20 ＝ 事業経費が埋もれてる候補 ===")
for amt, acct, store in sorted(mihit, reverse=True)[:20]:
    print(f"  {amt:>8,}  [{acct}] {store[:34]}")
