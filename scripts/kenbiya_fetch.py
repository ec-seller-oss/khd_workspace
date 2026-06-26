#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
健美家 物件詳細ページ 自動取得パーサ（資料取得→本査定）
================================================================
健美家listing URL → 土地面積/建物面積/用途地域/建ぺい容積/築年/構造/価格/利回り を取得。
screen_property が食えるfields dictを返す。WebFetch不要・pure python。
"""
import re
import sys
import json
import html as htmllib
import urllib.request


def _fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15)"})
    return urllib.request.urlopen(req, timeout=25).read().decode("utf-8", "ignore")


def _dd(html, label):
    """<dt>label</dt><dd ...>VALUE</dd> の VALUE(タグ除去)を返す。"""
    m = re.search(r"<dt>\s*" + re.escape(label) + r"\s*</dt>\s*<dd[^>]*>(.*?)</dd>", html, re.S)
    if not m:
        return None
    txt = re.sub(r"<[^>]+>", "", m.group(1))
    return htmllib.unescape(txt).strip()


def _num(s):
    if not s:
        return None
    m = re.search(r"([\d,]+(?:\.\d+)?)", s)
    return float(m.group(1).replace(",", "")) if m else None


def fetch(url):
    h = _fetch(url)
    tochi = _num(_dd(h, "土地面積"))
    tatemono = _num(_dd(h, "建物面積"))
    youto = _dd(h, "用途地域")
    kenpei = _num(_dd(h, "建ぺい率"))
    yoseki = _num(_dd(h, "容積率"))
    # 「建ぺい/容積率」結合ラベル（60 ％ / 200 ％）対応
    combo = _dd(h, "建ぺい/容積率") or _dd(h, "建ぺい率/容積率")
    if combo:
        nums = re.findall(r"([\d.]+)\s*[%％]", combo)
        if len(nums) >= 2:
            kenpei = kenpei or float(nums[0])
            yoseki = yoseki or float(nums[1])
    setsudo = _dd(h, "接道状況") or _dd(h, "接道")
    chiku = _dd(h, "築年月")
    kouzou = _dd(h, "構造") or _dd(h, "建物構造")
    juusho = _dd(h, "住所") or _dd(h, "所在地")
    price = _num(_dd(h, "価格"))
    # 価格は「7,480万円」→円
    price_yen = int(price * 10000) if price and price < 1_000_000 else (int(price) if price else None)
    # 満室時想定年収 or 利回り
    nenshu = _num(_dd(h, "満室時想定年収") or _dd(h, "想定年収") or _dd(h, "年間収入"))
    nenchin_yen = int(nenshu * 10000) if nenshu and nenshu < 1_000_000 else (int(nenshu) if nenshu else None)
    rimawari = _num(_dd(h, "満室時利回り") or _dd(h, "想定利回り") or _dd(h, "表面利回り"))
    if nenchin_yen is None and rimawari and price_yen:
        nenchin_yen = int(price_yen * rimawari / 100)

    name = (juusho or url.rstrip("/").split("/")[-1])
    fields = {
        "物件名": f"健美家_{name}"[:40],
        "価格_円": price_yen,
        "土地面積_m2": tochi,
        "容積率_pct": yoseki,
        "年賃料_円": nenchin_yen,
        "_参考情報": {"所在地": juusho, "用途地域": youto, "建蔽率_pct": kenpei,
                   "構造": kouzou, "築年月": chiku, "建物面積_m2": tatemono,
                   "接道メモ": setsudo, "URL": url},
    }
    return fields


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("使い方: python3 kenbiya_fetch.py <健美家listing URL>")
        sys.exit(1)
    f = fetch(sys.argv[1])
    print(json.dumps(f, ensure_ascii=False, indent=2))
