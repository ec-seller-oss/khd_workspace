#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KHD自動記帳エンジン  Phase 0
=================================================
02_資金調達 / バックオフィス自動化（営業時間創出＝利益直結の凍結例外）/ 2026-06-04

パイプライン:
  [1 取得]   正本CSV（各社サイト＝証憑）＋ MF一括CSV（照合チェックサム）
  [2 照合]   正本×MF を突合 → 抜け/二重/金額ズレ を検出（¥6,469の欠落を自動再現）
  [3 異常検知] 初出ベンダー・重複請求・高額外れ値（架空請求の一次フィルタ）
  [4 自動仕訳] 摘要→区分/勘定科目（journal_rules.py 準拠）。振替除外・家計/事業の切分け。
              ルール未HITは「要確認」キューへ（7割自動＋3割人間）。
  [5 出力]   MFクラウド会計 仕訳インポートCSV（19列）＋ 分類表 ＋ 照合/異常レポート

使い方:
  python3 kicho_engine.py --seihon fixtures/rakuten_2026-05_enavi.csv \
                          --mf fixtures/rakuten_2026-05_mf.csv \
                          --card "楽天カード(8991)" --ym 2026-05
  （引数省略時は楽天5月のfixtureで動く＝今日の実データ突合デモ）
"""
import csv
import os
import re
import argparse
import datetime as dt
import unicodedata

from journal_rules import MF_HEADER, CARD_PROFILES, classify


def nfkc(s):
    return unicodedata.normalize("NFKC", s or "").strip()

HERE = os.path.dirname(os.path.abspath(__file__))
HIGH_AMOUNT = 50000      # 高額フラグ（架空請求一次フィルタ）
DUP_WINDOW_DAYS = 2      # 重複請求とみなす日数窓


# ---------- 取得 ----------
def _amt(s):
    return int(round(float(re.sub(r"[^\d.\-]", "", str(s) or "0") or 0)))


def _date(s):
    s = (s or "").strip().replace("-", "/")
    m = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}/{int(m.group(2)):02d}/{int(m.group(3)):02d}"
    return s


def load_seihon(path):
    """正本（楽天e-NAVI等）CSV → [{date, store, amount(正)}]。半角カナはNFKCで正規化。"""
    out = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            date = _date(r.get("利用日") or r.get("ご利用日") or r.get("date"))
            store = nfkc(r.get("利用店名・商品名") or r.get("利用店名")
                         or r.get("ご利用店名") or r.get("store"))
            amt = abs(_amt(r.get("利用金額") or r.get("金額") or r.get("amount")))
            if not date or amt == 0:
                continue
            out.append({"date": date, "store": store, "amount": amt})
    return out


def _open_csv(path):
    """MF一括CSVはcp932、手作りfixtureはutf-8。両対応。"""
    for enc in ("cp932", "utf-8-sig", "utf-8"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return list(csv.DictReader(f))
        except (UnicodeDecodeError, LookupError):
            continue
    raise IOError(f"encoding判定失敗: {path}")


def load_mf(path, account=None):
    """MF家計簿CSV → [{date, content, amount(正), furikae}]。
    一括CSV(全口座)の場合は account(保有金融機関名)で対象カードに絞る。"""
    out = []
    for r in _open_csv(path):
        if account and (r.get("保有金融機関") or "").strip() != account:
            continue
        date = _date(r.get("日付") or r.get("date"))
        content = nfkc(r.get("内容") or r.get("content"))
        amt = abs(_amt(r.get("金額（円）") or r.get("金額") or r.get("amount")))
        furikae = (r.get("振替") or r.get("furikae") or "").strip() in ("1", "TRUE", "true", "振替", "○")
        if not date or amt == 0:
            continue
        out.append({"date": date, "content": content, "amount": amt, "furikae": furikae})
    return out


def _d(s):
    return dt.datetime.strptime(s, "%Y/%m/%d")


# ---------- 照合 ----------
def reconcile(seihon, mf):
    """正本×MF。amount一致＋日付±3日でマッチング（greedy）。"""
    mf_pool = list(mf)
    matched, missing_in_mf = [], []
    for s in seihon:
        hit = None
        for m in mf_pool:
            if m["amount"] == s["amount"] and abs((_d(m["date"]) - _d(s["date"])).days) <= 3:
                hit = m
                break
        if hit:
            mf_pool.remove(hit)
            matched.append((s, hit))
        else:
            missing_in_mf.append(s)        # 正本にあるがMFに無い＝MFの取りこぼし
    extra_in_mf = mf_pool                   # MFにあるが正本に無い＝二重/不明
    return matched, missing_in_mf, extra_in_mf


# ---------- 異常検知 ----------
def norm_vendor(store):
    t = re.sub(r"(VISA|ＶＩＳＡ)国内利用\s*VS\s*", "", nfkc(store))
    t = re.sub(r"\s+|　", "", t)
    t = re.sub(r"\d+$", "", t)
    return t[:12]


def detect_anomalies(seihon, known_vendors):
    anomalies = []
    # 1) 重複請求（同ベンダー・同額・近接日）
    for i in range(len(seihon)):
        for j in range(i + 1, len(seihon)):
            a, b = seihon[i], seihon[j]
            if a["amount"] == b["amount"] and norm_vendor(a["store"]) == norm_vendor(b["store"]) \
               and abs((_d(a["date"]) - _d(b["date"])).days) <= DUP_WINDOW_DAYS:
                anomalies.append(("重複請求の疑い", b, f"{a['date']}と同額同店"))
    # 2) 高額（架空請求一次フィルタ）
    for s in seihon:
        if s["amount"] >= HIGH_AMOUNT:
            anomalies.append(("高額（要目視）", s, f"{s['amount']:,}円 ≥ {HIGH_AMOUNT:,}"))
    # 3) 初出ベンダー（過去に無い店＝架空請求の一次フィルタ）
    if known_vendors is not None:
        for s in seihon:
            if norm_vendor(s["store"]) not in known_vendors:
                anomalies.append(("初出ベンダー", s, "過去明細に無い店名"))
    return anomalies


# ---------- 自動仕訳 ----------
def auto_journal(seihon, card_name):
    prof = CARD_PROFILES.get(card_name, {"default_kubun": "不明", "kashikata": "未払金",
                                         "kashikata_hojo": card_name})
    classified, mf_rows, kakei, toushi, youkakunin = [], [], [], [], []
    for s in seihon:
        c = classify(s["store"], prof)
        row = {**s, **c}
        classified.append(row)
        if c["区分"] in ("法人", "個人事業") and c["is_keihi"]:
            # 事業経費 → MF仕訳（借方=費用 / 貸方=未払金）
            invoice = "無"  # カード明細単体ではインボイス番号不明→無（経過措置）。証憑PDFで後付け。
            zei = c["税区分"]
            mf_rows.append([
                "", s["date"], c["勘定科目"], "", "", norm_vendor(s["store"]),
                zei, invoice, s["amount"],
                prof["kashikata"], prof.get("kashikata_hojo", ""), "", "", "対象外", "", s["amount"],
                f"{c['メモ']} {s['store']}"[:40], "", f"{card_name}/自動仕訳",
            ])
        elif c["区分"] == "家計":
            kakei.append(row)
        elif c["区分"] in ("投資", "振替"):
            toushi.append(row)
        else:
            youkakunin.append(row)
    return prof, classified, mf_rows, kakei, toushi, youkakunin


# ---------- 出力 ----------
def yen(n):
    return f"{n:,}円"


def run(seihon_path, mf_path, card_name, ym, vendors_path, mf_account=None):
    prof0 = CARD_PROFILES.get(card_name, {})
    mf_account = mf_account or prof0.get("mf_account")
    seihon = load_seihon(seihon_path)
    mf = load_mf(mf_path, account=mf_account)
    known = None
    if vendors_path and os.path.exists(vendors_path):
        known = set(l.strip() for l in open(vendors_path, encoding="utf-8") if l.strip())

    matched, missing, extra = reconcile(seihon, mf)
    anomalies = detect_anomalies(seihon, known)
    prof, classified, mf_rows, kakei, toushi, youkakunin = auto_journal(seihon, card_name)

    outdir = os.path.join(HERE, "out")
    os.makedirs(outdir, exist_ok=True)
    base = f"{card_name}_{ym}".replace("(", "").replace(")", "")

    # MF仕訳インポートCSV（19列・UTF-8 BOM）
    mf_csv = os.path.join(outdir, f"{base}_MF仕訳インポート.csv")
    with open(mf_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(MF_HEADER)
        w.writerows(mf_rows)

    # 分類表（全明細の区分/科目）
    cls_csv = os.path.join(outdir, f"{base}_分類表.csv")
    with open(cls_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["利用日", "店名", "金額", "区分", "勘定科目", "税区分", "メモ"])
        for r in classified:
            w.writerow([r["date"], r["store"], r["amount"], r["区分"],
                        r["勘定科目"], r["税区分"], r["メモ"]])

    # ===== コンソール・レポート =====
    s_sum, m_sum = sum(x["amount"] for x in seihon), sum(x["amount"] for x in mf)
    print("=" * 64)
    print(f" KHD自動記帳エンジン Phase0 — {card_name} / {ym}")
    print("=" * 64)
    print(f"\n[2] 照合（正本 vs MF）")
    print(f"  正本(証憑) : {len(seihon):>3}件  {yen(s_sum)}")
    print(f"  MF(照合)   : {len(mf):>3}件  {yen(m_sum)}")
    print(f"  一致       : {len(matched):>3}件")
    if missing:
        print(f"  🔴 MF取りこぼし（正本にあるがMFに無い）: {len(missing)}件")
        for s in missing:
            print(f"      └ {s['date']}  {yen(s['amount'])}  {s['store']}")
    if extra:
        print(f"  🟡 MF余剰（MFにあるが正本に無い＝要確認）: {len(extra)}件")
        for m in extra:
            print(f"      └ {m['date']}  {yen(m['amount'])}  {m['content']}")
    if not missing and not extra:
        print("  ✅ 抜け・二重なし（完全一致）")

    print(f"\n[3] 異常検知")
    if not anomalies:
        print("  ✅ 異常なし")
    seen = set()
    for kind, s, why in anomalies:
        key = (kind, s["date"], s["amount"])
        if key in seen:
            continue
        seen.add(key)
        print(f"  ⚠️ {kind}: {s['date']} {yen(s['amount'])} {s['store'][:24]} … {why}")
    if known is None:
        print("  ℹ️ 初出ベンダー検知は過去ベンダー辞書(vendors.txt)未投入のためスキップ")

    print(f"\n[4] 自動仕訳（区分: {prof['default_kubun']}カード）")
    print(f"  事業仕訳(MF出力) : {len(mf_rows):>3}件  {yen(sum(r[8] for r in mf_rows))}")
    print(f"  家計(対象外)     : {len(kakei):>3}件  {yen(sum(r['amount'] for r in kakei))}")
    print(f"  投資/振替(除外)  : {len(toushi):>3}件  {yen(sum(r['amount'] for r in toushi))}")
    print(f"  🟡 要確認(未HIT) : {len(youkakunin):>3}件  {yen(sum(r['amount'] for r in youkakunin))}")
    for r in youkakunin:
        print(f"      └ {r['date']} {yen(r['amount'])} {r['store'][:28]}")

    print(f"\n[5] 出力")
    print(f"  MF仕訳インポートCSV : {mf_csv}")
    print(f"  分類表CSV           : {cls_csv}")
    print("=" * 64)
    return {"missing": missing, "anomalies": anomalies, "mf_rows": mf_rows,
            "kakei": kakei, "youkakunin": youkakunin}


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="KHD自動記帳エンジン Phase0")
    ap.add_argument("--seihon", default=os.path.join(HERE, "fixtures/rakuten_2026-05_enavi.csv"))
    ap.add_argument("--mf", default=os.path.join(HERE, "fixtures/rakuten_2026-05_mf.csv"))
    ap.add_argument("--card", default="楽天カード(8991)")
    ap.add_argument("--ym", default="2026-05")
    ap.add_argument("--vendors", default=os.path.join(HERE, "fixtures/vendors.txt"))
    ap.add_argument("--mf-account", default=None,
                    help="MF一括CSVから抽出する保有金融機関名（省略時はカードprofileのmf_account）")
    a = ap.parse_args()
    run(a.seihon, a.mf, a.card, a.ym, a.vendors, a.mf_account)
