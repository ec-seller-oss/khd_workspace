#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Drive(スプレッドシート)のテキスト表現(Markdown表) → Amazon注文履歴CSV へ前処理。
read_file_content がスプシをMarkdown表で返すため、それをCSV化して
amazon_to_mf_journal.py に食わせる橋渡し。
使い方: python3 prep_amazon_md.py <drive_text.json or .txt> <out.csv>
"""
import json
import csv
import sys


def clean(x):
    # Markdownエスケープ(\-500 等)を戻す
    return x.replace("\\-", "-").replace("\\", "").strip()


def main(src, out):
    raw = open(src, encoding="utf-8").read()
    try:
        md = json.loads(raw)["fileContent"]
    except (ValueError, KeyError):
        md = raw  # 既に素のMarkdownならそのまま

    table = [l for l in md.splitlines() if l.strip().startswith("|")]
    rows = [[clean(c) for c in l.strip().strip("|").split("|")] for l in table]
    if len(rows) < 3:
        print("⚠️ 表が見つかりません")
        return
    header = rows[0]
    data_rows = [r for r in rows[2:] if len(r) == len(header)]  # 区切り行(:-:)除外・列ずれ除外
    dropped = len(rows[2:]) - len(data_rows)

    with open(out, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(data_rows)
    print(f"✅ CSV化: {len(data_rows)}行 (列ずれ除外{dropped}行) → {out}")
    print(f"   列数={len(header)} 先頭列={header[:5]}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
