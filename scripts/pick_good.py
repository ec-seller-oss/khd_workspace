#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
良物件オートピック（物DB SSoT → 🟢買い候補だけ抽出・通知）
================================================================
物DB「物」タブを走査し、判定可能なデータだけを対象に玉川式の閾値で🟢を拾う。
データ欠損/ノイズ(金額に人名・利回り>100%等)は正直に除外し件数を報告する。
通知: コンソール＋notes保存。Notion追記は khd-log 連携で別途。
"""
import re
import sys
import pickle
import warnings
import datetime
from pathlib import Path

warnings.filterwarnings("ignore")
SS = "1XTPXFxvJtaoEKVlEaigP3U1VdYfG-IHa_9pqOiZ1-hA"
TOKEN = Path(__file__).parent / "sheets_token.pickle"

# 閾値（玉川式・収益寄り）
YIELD_MIN = 0.08     # 仕入利回 8%以上
YIELD_MAX = 0.40     # 40%超はデータノイズ扱いで除外
EXCLUDE_STATUS = {"クローズ", "敗退", "売却決済完了"}


def _svc():
    from googleapiclient.discovery import build
    creds = pickle.load(open(TOKEN, "rb"))
    return build("sheets", "v4", credentials=creds)


def _yen(s):
    if not s:
        return None
    m = re.search(r"¥?\s*([\d,]{4,})", str(s))
    return int(m.group(1).replace(",", "")) if m else None


def _pct(s):
    if not s or "#" in str(s):
        return None
    m = re.search(r"([\d.]+)\s*%", str(s))
    return float(m.group(1)) / 100 if m else None


def _f(s):
    try:
        return float(str(s))
    except (ValueError, TypeError):
        return None


def run():
    svc = _svc()
    cols = svc.spreadsheets().values().batchGet(spreadsheetId=SS, ranges=[
        "物!A5:A400", "物!G5:G400", "物!J5:J400", "物!L5:L400",
        "物!AR5:AR400", "物!AV5:AV400", "物!BE5:BE400"]).execute()["valueRanges"]
    def col(i): return [(r[0] if r else "") for r in cols[i].get("values", [])]
    A, G, J, L, AR, AV, BE = col(0), col(1), col(2), col(3), col(4), col(5), col(6)
    n = max(len(A), len(J), len(L), len(AR))
    def g(x, i): return x[i] if i < len(x) else ""

    total = sum(1 for i in range(n) if g(A, i).strip())
    judgeable, picks = 0, []
    for i in range(n):
        no = g(A, i).strip()
        if not no or no == "No":
            continue
        status = g(L, i).strip()
        kin = _yen(g(J, i))
        yld = _pct(g(AR, i))
        tochi = _f(g(AV, i))
        # 判定可能＝金額が数値 かつ 利回りが数値
        if kin is None or yld is None:
            continue
        judgeable += 1
        if status in EXCLUDE_STATUS:
            continue
        if not (YIELD_MIN <= yld <= YIELD_MAX):
            continue
        # 🟢候補
        name = no.replace("メモ - 「", "").replace("」", "")
        picks.append({"name": name[:36], "kin": kin, "yld": yld,
                      "tochi": tochi, "status": status or "—",
                      "memo": (g(BE, i) or "")[:30]})

    picks.sort(key=lambda p: p["yld"], reverse=True)
    return total, judgeable, picks


def render(total, judgeable, picks):
    L = []
    L.append("=" * 60)
    L.append(f"  🟢 良物件オートピック（物DB / {datetime.date.today()}）")
    L.append("=" * 60)
    L.append(f"  全案件 {total}件 → 判定可能データ {judgeable}件 → 🟢候補 {len(picks)}件")
    L.append(f"  基準: 利回り{int(YIELD_MIN*100)}〜{int(YIELD_MAX*100)}% / クローズ・敗退除外")
    L.append("-" * 60)
    if not picks:
        L.append("  該当なし（データ品質が低く、自動査定の蓄積で精度向上）")
    for p in picks:
        t = f"土地値{p['tochi']}" if p['tochi'] is not None else ""
        L.append(f"  🟢 {p['name']}")
        L.append(f"     金額¥{p['kin']:,} / 利回り{p['yld']*100:.1f}% {t} / {p['status']}")
    L.append("-" * 60)
    L.append("  ※ レガシー台帳は金額/利回り欄にノイズ多。screen_property --db の自動査定で")
    L.append("     クリーンなデータが増えるほどオートピック精度は上がる。")
    L.append("=" * 60)
    return "\n".join(L)


if __name__ == "__main__":
    total, judgeable, picks = run()
    out = render(total, judgeable, picks)
    print(out)
    notes = Path.home() / "01_honbu_docs_automation" / ".company" / "secretary" / "notes"
    f = notes / f"{datetime.date.today()}-autopick.md"
    f.write_text("# 良物件オートピック結果\n\n```\n" + out + "\n```\n", encoding="utf-8")
    print(f"\n📄 保存: {f}")
