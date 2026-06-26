#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
物件管理マスターDB(SSoT)への査定結果 自動upsert
================================================================
screen_property の判定データ → 物DB「物」タブに1行 upsert（情報集約DBの本体）。
キー＝物件名（A列/No）。既存は更新、無ければ末尾に追記。既存ロジック(関数)を踏襲。
安全：呼び出しは screen_property の --db 明示時のみ。
"""
import pickle
import datetime
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
SS = "1XTPXFxvJtaoEKVlEaigP3U1VdYfG-IHa_9pqOiZ1-hA"
TOKEN = Path(__file__).parent / "sheets_token.pickle"


def _svc():
    from googleapiclient.discovery import build
    creds = pickle.load(open(TOKEN, "rb"))
    return build("sheets", "v4", credentials=creds)


def _col(n):
    """1-indexed列番号→A1表記。"""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _pct(x):
    return f"{x*100:.2f}%" if x is not None else ""


def upsert_property(fields, res, verdict):
    """物DBへ1行upsert。戻り値=(row, 'updated'|'appended', url)。"""
    svc = _svc()
    # A列読み込み→キー行 or 追記先を決定
    colA = svc.spreadsheets().values().get(
        spreadsheetId=SS, range="物!A1:A400").execute().get("values", [])
    key = str(fields.get("物件名", "")).strip()
    row = None
    last = 4  # ヘッダ等を避け5行目以降
    for i, v in enumerate(colA, 1):
        cell = (v[0].strip() if v else "")
        if cell:
            last = i
            if cell == key:
                row = i
    mode = "updated" if row else "appended"
    if not row:
        row = max(last + 1, 8)

    price = fields.get("価格_円")
    nenchin = fields.get("年賃料_円")
    is_income = res.get("is_income")
    ref = fields.get("_参考情報") or {}
    today = datetime.date.today().strftime("%-m/%-d")

    # 1-indexed列 → 値（物タブ 57列）
    vals = {
        1: key,                                   # No(キー)
        3: today,                                 # 開始
        7: ref.get("所在地") or key,              # 所在地
        8: fields.get("土地面積_m2") or "",       # 敷地面積/㎡
        10: price or "",                          # 金額
        12: "情報",                               # ステータス
        13: f"AI査定:{verdict}",                  # 状況
        20: price or "",                          # 仕入金額
        29: int(nenchin / 12) if nenchin else "",  # 家賃(円/月)
        33: "保有（ローン" if is_income else "転売",  # スキーム
        44: _pct(res.get("表面利回り")) if is_income else "",  # 仕入利回
        46: ref.get("用途地域") or "",            # 用途地域
        47: fields.get("路線価_円per_m2") or "",  # 相続路線価
    }
    # 土地値割合(AV=48) は玉川式A（路線価×面積÷価格・0.4↑が安全）に統一
    if res.get("土地値割合") is not None:
        vals[48] = round(res["土地値割合"], 3)

    # A..AV(1..48)を配列化
    arr = [vals.get(c, "") for c in range(1, 49)]
    svc.spreadsheets().values().update(
        spreadsheetId=SS, range=f"物!A{row}:AV{row}",
        valueInputOption="USER_ENTERED", body={"values": [arr]}).execute()

    # メモ(BE=57)
    kpi = []
    if res.get("土地値割合") is not None:
        kpi.append(f"土地値割合{_pct(res['土地値割合'])}")
    if is_income:
        kpi.append(f"実質利回り{_pct(res.get('実質利回り'))}")
        if res.get("CF率") is not None:
            kpi.append(f"CF率{_pct(res['CF率'])}")
    elif res.get("粗利率") is not None:
        kpi.append(f"粗利率{_pct(res['粗利率'])}")
    memo = f"AI自動査定({today}) 判定:{verdict} / " + " / ".join(kpi) + " ※screen_property自動入力"
    svc.spreadsheets().values().update(
        spreadsheetId=SS, range=f"物!BE{row}",
        valueInputOption="USER_ENTERED", body={"values": [[memo]]}).execute()

    # BF=積算割合 / BG=融資カバー率B（計算式で投入＝価格/路線価/面積が変われば自動再計算）
    bf, bg = "", ""
    if fields.get("路線価_円per_m2") and fields.get("土地面積_m2") and price:
        tatemono = int(res.get("建物積算_円") or 0)  # 建物積算(築古=0)はpython算出の定数
        bf = f"=(AU{row}*H{row}+{tatemono})/J{row}"   # 積算割合=(土地積算+建物積算)÷価格
        bg = f"=T{row}*0.8/(H{row}*AU{row})"          # 融資カバー率B=仕入×0.8÷(面積×路線価)
    svc.spreadsheets().values().update(
        spreadsheetId=SS, range=f"物!BF{row}:BG{row}",
        valueInputOption="USER_ENTERED", body={"values": [[bf, bg]]}).execute()

    url = f"https://docs.google.com/spreadsheets/d/{SS}/edit#gid=1649873874&range=A{row}"
    return row, mode, url
