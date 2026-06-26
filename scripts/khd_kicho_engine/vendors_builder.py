#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
既知ベンダー辞書ビルダー — 過去明細CSVから「過去に存在した店」の集合を作る。
出力 vendors.txt を kicho_engine が読み、初出ベンダー（架空請求の一次フィルタ）を検知する。

使い方:
  python3 vendors_builder.py 過去CSV1 過去CSV2 ... --out fixtures/vendors.txt
  ※ MF一括CSV(内容列) / 正本CSV(利用店名・商品名列) どちらも可。norm_vendorで正規化して集約。
"""
import csv
import sys
import os
import argparse

from kicho_engine import norm_vendor, _open_csv  # 同一の正規化を使う

STORE_COLS = ["内容", "利用店名・商品名", "利用店名", "ご利用店名", "content", "store"]


def collect(paths):
    vendors = set()
    per_file = {}
    for p in paths:
        n = 0
        for r in _open_csv(p):
            store = next((r[c] for c in STORE_COLS if r.get(c)), "")
            v = norm_vendor(store)
            if v:
                vendors.add(v)
                n += 1
        per_file[os.path.basename(p)] = n
    return vendors, per_file


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="+", help="過去明細CSV（複数可）")
    ap.add_argument("--out", default=os.path.join(os.path.dirname(__file__), "fixtures/vendors.txt"))
    a = ap.parse_args()
    vendors, per_file = collect(a.inputs)
    with open(a.out, "w", encoding="utf-8") as f:
        f.write("\n".join(sorted(vendors)) + "\n")
    print(f"✅ 既知ベンダー辞書: {len(vendors)}社  → {a.out}")
    for k, v in per_file.items():
        print(f"   {k}: {v}行")
