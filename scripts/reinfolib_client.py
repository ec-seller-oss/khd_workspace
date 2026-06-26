#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
reinfolib（国交省 不動産情報ライブラリ）APIクライアント
================================================================
仕入スクリーナーの「路線価（土地値）自動取得」を担う。
取引価格情報API(XIT001)から 宅地(土地) の実勢㎡単価 中央値を取り、
相続税路線価相当に換算して返す。geocode/タイル変換不要で堅牢。

【APIキー】scripts/.env の REINFOLIB_API_KEY を読む（git管理外・平文コミット禁止）。
【換算】相続税路線価 ≈ 実勢取引㎡単価 × 0.72
        （公示≈実勢×0.9 × 相続税路線価≈公示×0.8）。あくまで近似。
"""
import os
import re
import json
import gzip
import statistics
import urllib.request
import urllib.parse
from pathlib import Path

BASE = "https://www.reinfolib.mlit.go.jp/ex-api/external"
ROSENKA_RATIO = 0.72   # 実勢取引㎡単価 → 相続税路線価相当
ENV_PATH = Path(__file__).parent / ".env"

PREF_CODES = {
    "北海道": "01", "青森県": "02", "岩手県": "03", "宮城県": "04", "秋田県": "05",
    "山形県": "06", "福島県": "07", "茨城県": "08", "栃木県": "09", "群馬県": "10",
    "埼玉県": "11", "千葉県": "12", "東京都": "13", "神奈川県": "14", "新潟県": "15",
    "富山県": "16", "石川県": "17", "福井県": "18", "山梨県": "19", "長野県": "20",
    "岐阜県": "21", "静岡県": "22", "愛知県": "23", "三重県": "24", "滋賀県": "25",
    "京都府": "26", "大阪府": "27", "兵庫県": "28", "奈良県": "29", "和歌山県": "30",
    "鳥取県": "31", "島根県": "32", "岡山県": "33", "広島県": "34", "山口県": "35",
    "徳島県": "36", "香川県": "37", "愛媛県": "38", "高知県": "39", "福岡県": "40",
    "佐賀県": "41", "長崎県": "42", "熊本県": "43", "大分県": "44", "宮崎県": "45",
    "鹿児島県": "46", "沖縄県": "47",
}


def _load_key():
    key = os.environ.get("REINFOLIB_API_KEY")
    if key:
        return key
    if ENV_PATH.exists():
        for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
            if line.startswith("REINFOLIB_API_KEY="):
                return line.split("=", 1)[1].strip()
    raise RuntimeError("REINFOLIB_API_KEY が未設定です（環境変数 or scripts/.env）")


def _call(path, params):
    key = _load_key()
    url = f"{BASE}/{path}?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={
        "Ocp-Apim-Subscription-Key": key, "Accept-Encoding": "gzip"})
    with urllib.request.urlopen(req, timeout=30) as r:
        raw = r.read()
        if r.headers.get("Content-Encoding") == "gzip":
            raw = gzip.decompress(raw)
    return json.loads(raw)


def pref_code(pref_name):
    for k, v in PREF_CODES.items():
        if pref_name and (pref_name in k or k.startswith(pref_name.rstrip("都道府県"))):
            return v
    return None


def city_code(area_code, city_name):
    """XIT002: 市区町村名（部分一致）→ 市区町村コード。政令市は親名でヒット。"""
    d = _call("XIT002", {"area": area_code})
    data = d.get("data", [])
    # 政令市「○○市□□区」表記 → 区一覧は「□□区」のみで載るため、区名も検索キーに加える
    keys = [city_name]
    if city_name and "市" in city_name and "区" in city_name:
        keys.append(city_name.split("市", 1)[1])   # 例: 横浜市鶴見区 → 鶴見区
    for key in keys:                                # 完全一致 優先
        hit = [c for c in data if c["name"] == key]
        if hit:
            return hit[0]
    for key in keys:                                # 次に部分一致
        hit = [c for c in data if key and key in c["name"]]
        if hit:
            return hit[0]
    return None


def land_unit_price(area_code, ccode, district=None, years=(2024, 2023)):
    """XIT001: 宅地(土地)の実勢㎡単価 中央値 + 相続税路線価相当を返す。
    district 指定時はその町名(部分一致)に絞る。少なすぎれば市区町村全域へ自動拡大。"""
    rows = []
    for y in years:
        d = _call("XIT001", {"year": y, "area": area_code, "city": ccode})
        rows += d.get("data", [])

    def units(filter_district):
        out = []
        for r in rows:
            if r.get("Type") != "宅地(土地)":
                continue
            up = r.get("UnitPrice")
            if not up:
                continue
            if filter_district and district not in (r.get("DistrictName") or ""):
                continue
            out.append(int(up))
        return out

    scope = "町名"
    prices = units(True) if district else units(False)
    if district and len(prices) < 3:   # 町名サンプル過少→市区町村全域に拡大
        prices = units(False)
        scope = "市区町村全域(町名サンプル過少のため拡大)"
    elif not district:
        scope = "市区町村全域"

    if not prices:
        return None
    med = int(statistics.median(prices))
    return {
        "実勢㎡単価_中央値": med,
        "件数": len(prices),
        "範囲": (min(prices), max(prices)),
        "集計範囲": scope,
        "相続税路線価相当": int(med * ROSENKA_RATIO),
        "換算係数": ROSENKA_RATIO,
        "対象年": list(years),
    }


def parse_address(addr):
    """住所文字列 → (都道府県, 市区町村, 町名)。政令市の区も拾う。
    例: 岩手県盛岡市高松2丁目34-5 → ('岩手県','盛岡市','高松')
        神奈川県横浜市鶴見区上町1 → ('神奈川県','横浜市鶴見区','上町')"""
    if not addr:
        return (None, None, None)
    pref = None; rest = addr
    # 都道府県名を文字列中から検索（先頭以外/接頭辞ゴミがあってもOK）
    for k in PREF_CODES:
        idx = addr.find(k)
        if idx >= 0:
            pref = k; rest = addr[idx + len(k):]; break
    # 市区町村（政令市は「○○市△△区」／東京特別区は「□□区」）
    m = re.match(r"(.+?[市町村])", rest)
    city = m.group(1) if m else None
    if city:
        after = rest[len(city):]
        m2 = re.match(r"([^\d０-９一二三四五六七八九十]+?区)", after)
        if m2:  # 区名に数字が無ければ政令市の区 → 市に連結
            city = city + m2.group(1)
    else:  # 市町村が無い＝東京特別区など（北区・港区等）
        m3 = re.match(r"([^\d０-９]+?区)", rest)
        city = m3.group(1) if m3 else None
    rest2 = rest[len(city):] if city else rest
    # 町名（先頭の非数字部分・丁目/番手前まで）
    m3 = re.match(r"([^\d０-９]+?)(?:[\d０-９]|[一二三四五六七八九十]+丁目)", rest2)
    district = m3.group(1) if m3 else (rest2[:6] if rest2 else None)
    if district:
        district = district.strip("　 ")
    return (pref, city, district)


def lookup_by_address(addr):
    """住所だけで路線価相当をlookup。"""
    pref, city, district = parse_address(addr)
    if not (pref and city):
        return {"error": f"住所を分解できません: {addr}（pref={pref}/city={city}）"}
    return lookup(pref, city, district)


def lookup(pref_name, city_name, district=None):
    """住所パーツ → 路線価相当。失敗時は理由つきで None系を返す。"""
    pc = pref_code(pref_name)
    if not pc:
        return {"error": f"都道府県を特定できません: {pref_name}"}
    try:
        cc = city_code(pc, city_name)
    except Exception as e:
        return {"error": f"市区町村API失敗: {e}"}
    if not cc:
        return {"error": f"市区町村を特定できません: {city_name}（{pref_name}）"}
    try:
        res = land_unit_price(pc, cc["id"], district)
    except Exception as e:
        hint = "（政令市は区まで指定してください 例: 横浜市鶴見区）" if cc["id"].endswith("100") else ""
        return {"error": f"取引価格API失敗: {e}{hint}", "市区町村": cc}
    if not res:
        return {"error": f"取引データなし: {city_name}", "市区町村": cc}
    res["都道府県"] = pref_name
    res["市区町村"] = cc["name"]
    res["市区町村コード"] = cc["id"]
    res["町名"] = district
    return res


if __name__ == "__main__":
    import sys
    args = sys.argv[1:]
    if len(args) < 2:
        print("使い方: python3 reinfolib_client.py <都道府県> <市区町村> [町名]")
        sys.exit(1)
    res = lookup(args[0], args[1], args[2] if len(args) > 2 else None)
    print(json.dumps(res, ensure_ascii=False, indent=2))
