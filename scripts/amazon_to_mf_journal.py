#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Amazon注文履歴 → MF仕訳帳インポートCSV 変換（個人EC消費税記帳の自動化）
================================================================
バックオフィス自動化 Tier2「本丸」/ 02_資金調達 / 2026-05-29 設計

目的: Q1で162件あった「注文明細→仕入高」の手入力地獄をゼロ化。
      Amazon注文履歴CSVを、MF会計の仕訳帳インポート形式CSVへ自動変換する。

入力:
  - Amazonビジネス注文履歴CSV（クーパン2 / アカウントグループ=KIKUCHI HOLDINGS K.K. / 日付はJST）
  - 通常Amazon Order History.csv（クーパン1 / 日付はUTC=末尾Z → JST変換が必要）

出力:
  - MF仕訳帳インポートCSV（19列・UTF-8 BOM）
    ※ reference_tax_audit_methodology §9 のヘッダと完全一致

★ 橋本先生の確認待ち（CONFIGで切替）= 確定するまではデフォルト(2025準拠の仮置き)で動く:
  1. 経理方式（税込/税抜）        -> CONFIG["keiri"]
  2. ギフト払いの貸方科目          -> CONFIG["kashikata"]（2025準拠=獲得ポイント残高 資産取崩）
  3. インボイス番号 無し の税区分  -> CONFIG["invoice_nashi_kubun"]（適格 or 経過措置80%）
  4. 送料・割引を仕入高に含めるか  -> CONFIG["soryo_included"]
  5. 輸出免税対応の控除            -> 申告側の処理（本スクリプト対象外）

使い方:
  python3 amazon_to_mf_journal.py 入力.csv --account 2 --out 出力.csv
  python3 amazon_to_mf_journal.py 入力.csv --account 1 --utc   # クーパン1(UTC日付)
"""
import csv
import sys
import argparse
import datetime as dt

# ===== CONFIG（税務ルール）★2026-05-29 菊池OKで確定（税込／獲得ポイント残高／経過措置80%） =====
CONFIG = {
    "keiri": "税込",                 # 確定: 税込経理（2025準拠）
    "kashikata": "獲得ポイント残高",  # 確定: ギフト払いの貸方=獲得ポイント残高（資産取崩）
    "kashikata_hojo": "",
    "invoice_ari_kubun": "課仕 10%",          # 確定: インボイスT番号有=適格
    "invoice_nashi_kubun": "課仕 10%（経過措置80%）",  # 確定: 2026年は経過措置80%が制度上正
    "soryo_included": True,          # 確定: 送料・割引は仕入高に含む
}

# MF仕訳帳インポート 19列ヘッダ（§9 と完全一致）
MF_HEADER = [
    "取引No", "取引日", "借方勘定科目", "借方補助科目", "借方部門", "借方取引先",
    "借方税区分", "借方インボイス", "借方金額(円)", "貸方勘定科目", "貸方補助科目",
    "貸方部門", "貸方取引先", "貸方税区分", "貸方インボイス", "貸方金額(円)",
    "摘要", "タグ", "メモ",
]

# Amazon注文履歴CSV の使用列名（80列中）
COL = {
    "date": "注文日",
    "order_no": "注文番号",
    "group": "アカウントグループ",
    "subtotal_excl": "注文の小計（税抜）",
    "tax": "注文の消費税額",
    "total_incl": "注文の合計（税込）",
    "shipping_excl": "注文の配送料および手数料（税抜）",
    "discount_incl": "注文の割引（税込）",
    "invoice_no": "適格請求書発行事業者登録番号",
    "invoice_issuer": "適格請求書（または支払い明細書）発行者名",
    "item_name": "商品名",
    "asin": "ASIN",
    "seller": "出品者名",
    "category": "商品カテゴリー",
}

# ギフトカードチャージのカテゴリ。これは「仕入」ではなく獲得ポイント残高への
# チャージ（カード明細側で計上）なので、注文履歴側では除外する（二重計上防止）。
# ※「商品名にギフト」で判定すると VICTORINOX 等の実商品を誤除外するため、必ずカテゴリで判定。
GIFT_CHARGE_CATEGORY = "Reload Gift Card Balance"


def to_jst(s, is_utc):
    """注文日 → 取引日(YYYY/MM/DD)。クーパン1はUTC→JST(+9h)。クーパン2はJSTそのまま。"""
    s = (s or "").strip()
    if not s:
        return ""
    if is_utc and "T" in s:
        utc = dt.datetime.strptime(s[:19], "%Y-%m-%dT%H:%M:%S")
        return (utc + dt.timedelta(hours=9)).strftime("%Y/%m/%d")
    return s.replace("-", "/")


def to_int(s):
    try:
        return int(round(float(str(s).replace(",", "").strip() or 0)))
    except ValueError:
        return 0


def convert(in_path, account, is_utc, out_path, from_ym=None, to_ym=None):
    hojo_karikata = "クーパン2" if account == 2 else "クーパン1"
    rows_out = []
    skipped = 0
    gift_skipped = 0
    period_skipped = 0

    with open(in_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for r in reader:
            order_no = (r.get(COL["order_no"]) or "").strip()
            if not order_no:
                continue  # 商品行のみ・空行スキップ

            # ギフトカードチャージは仕入ではない → 除外（二重計上防止）
            if (r.get(COL["category"]) or "").strip() == GIFT_CHARGE_CATEGORY:
                gift_skipped += 1
                continue

            date = to_jst(r.get(COL["date"]), is_utc)

            # 期間フィルタ（YYYYMM整数で比較）
            if from_ym or to_ym:
                p = date.replace("-", "/").split("/")
                ymv = int(p[0]) * 100 + int(p[1]) if len(p) >= 2 and p[0].isdigit() else 0
                if (from_ym and ymv < from_ym) or (to_ym and ymv > to_ym):
                    period_skipped += 1
                    continue

            # 仕入高に計上する金額（経理方式で列を選ぶ）
            if CONFIG["keiri"] == "税込":
                amount = to_int(r.get(COL["total_incl"]))
                if not CONFIG["soryo_included"]:
                    amount -= (to_int(r.get(COL["shipping_excl"])) + to_int(r.get(COL["tax"])))
            else:  # 税抜
                amount = to_int(r.get(COL["subtotal_excl"]))

            if amount <= 0:
                skipped += 1
                continue

            invoice_no = (r.get(COL["invoice_no"]) or "").strip()
            zeikubun = CONFIG["invoice_ari_kubun"] if invoice_no.startswith("T") \
                else CONFIG["invoice_nashi_kubun"]

            tekiyo = (r.get(COL["item_name"]) or r.get(COL["seller"]) or "")[:40]
            memo = f"注文{order_no}/ASIN{r.get(COL['asin'],'')}/{invoice_no}"

            rows_out.append([
                "", date, "仕入高", hojo_karikata, "", "",
                zeikubun, ("有" if invoice_no.startswith("T") else "無"), amount,
                CONFIG["kashikata"], CONFIG["kashikata_hojo"], "", "", "対象外", "", amount,
                tekiyo, "", memo,
            ])

    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(MF_HEADER)
        w.writerows(rows_out)

    total = sum(row[8] for row in rows_out)
    print(f"✅ 変換完了: {len(rows_out)}件 / 仕入高合計 {total:,}円")
    print(f"   除外: 金額0={skipped}件 / ギフトチャージ={gift_skipped}件 / 期間外={period_skipped}件")
    print(f"   出力: {out_path}")
    print(f"   ⚠️ 税ルール(CONFIG)は橋本先生の確認後に確定すること:")
    print(f"      経理={CONFIG['keiri']} / 貸方={CONFIG['kashikata']} / "
          f"インボイス無={CONFIG['invoice_nashi_kubun']}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Amazon注文履歴 → MF仕訳CSV 変換")
    ap.add_argument("input", help="Amazon注文履歴CSV")
    ap.add_argument("--account", type=int, choices=[1, 2], default=2,
                    help="1=通常Amazon(クーパン1) / 2=ビジネス(クーパン2)")
    ap.add_argument("--utc", action="store_true",
                    help="注文日がUTC(末尾Z)の場合に指定（通常Order History用）")
    ap.add_argument("--out", default="mf_import.csv", help="出力CSVパス")
    ap.add_argument("--from-ym", type=int, default=None, help="開始年月(YYYYMM) 例:202601")
    ap.add_argument("--to-ym", type=int, default=None, help="終了年月(YYYYMM) 例:202603")
    a = ap.parse_args()
    convert(a.input, a.account, a.utc, a.out, a.from_ym, a.to_ym)
