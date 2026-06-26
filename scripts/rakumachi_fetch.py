#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
楽待 物件詳細ページ 自動取得パーサ（資料取得→本査定）
================================================================
楽待listing URL → 価格/土地面積/用途地域/建ぺい容積/所在地 を取得。
※楽待は利回り/建物面積/築年が非ログインでblur(非表示)。取れる項目のみ返す。
screen_property が食えるfields dictを返す。pure python。
"""
import re
import sys
import json
import html as htmllib
import urllib.request


def _fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15)"})
    return urllib.request.urlopen(req, timeout=25).read().decode("utf-8", "ignore")


def _thtd(html, label):
    """<th>label</th><td ...>VALUE</td> の VALUE(タグ除去)。"""
    m = re.search(r"<th[^>]*>\s*" + re.escape(label) + r"\s*</th>\s*<td[^>]*>(.*?)</td>", html, re.S)
    if not m:
        return None
    return htmllib.unescape(re.sub(r"<[^>]+>", "", m.group(1))).strip()


def _num(s):
    if not s:
        return None
    m = re.search(r"([\d,]+(?:\.\d+)?)", s)
    return float(m.group(1).replace(",", "")) if m else None


def fetch(url):
    h = _fetch(url)
    m = re.search(r'<span class="price[^"]*">\s*([\d,]+)\s*万円', h)
    price_yen = int(m.group(1).replace(",", "")) * 10000 if m else None
    m = re.search(r'id="js-landArea">\s*([\d,]+\.?\d*)', h)
    tochi = _num(m.group(1)) if m else _num(_thtd(h, "土地面積"))
    youto = _thtd(h, "用途地域")
    yoseki = _num(_thtd(h, "容積率"))
    kenpei = _num(_thtd(h, "建ぺい率") or _thtd(h, "建蔽率"))
    juusho = _thtd(h, "所在地")
    if juusho:
        juusho = re.sub(r"\s+", "", juusho)
        juusho = re.sub(r"(地図|周辺地図|MAP).*$", "", juusho)
    kouzou = _thtd(h, "構造") or _thtd(h, "建物構造")

    name = juusho or url.rstrip("/").split("/")[-2]
    return {
        "物件名": f"楽待_{name}"[:40],
        "価格_円": price_yen,
        "土地面積_m2": tochi,
        "容積率_pct": yoseki,
        "年賃料_円": None,   # 楽待は利回り非表示→収益は別途
        "_参考情報": {"所在地": juusho, "用途地域": youto, "建蔽率_pct": kenpei,
                   "構造": kouzou, "URL": url,
                   "_注記": "楽待は利回り/建物/築年が非ログインで非表示。価格・土地・容積のみ取得"},
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("使い方: python3 rakumachi_fetch.py <楽待listing URL>")
        sys.exit(1)
    print(json.dumps(fetch(sys.argv[1]), ensure_ascii=False, indent=2))
