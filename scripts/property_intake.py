#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
物件PDF 抽出インテーク（融資資料自動化TTP・自社版 スコープA / PDF抽出ピース）
================================================================================
マイソク/物件概要PDF → 価格・土地面積・容積率・用途地域・所在地・接道を抽出し、
property_screener.py が食えるJSONを生成する。

【2段構え（B&S設計書 原則③を踏襲）】
  1) pdfplumber でテキスト抽出 → 正規表現で定型フィールドを拾う（高速・再現性）
  2) テキストが薄い / 拾えない項目 → "要Claude確認" として残し、画像PDFはClaude推論
     （Readツールでのマルチモーダル読取）にフォールバックする。
  取れない値は勝手に補完しない。missing は正直に残す。

使い方:
  python3 property_intake.py 物件.pdf
  → out_screener/intake_<物件名>.json を生成（screenerにそのまま渡せる）
  → 続けて: python3 property_screener.py out_screener/intake_<物件名>.json
"""
import sys
import re
import json
import unicodedata
import datetime
from pathlib import Path

OUT_DIR = Path.home() / "01_honbu_docs_automation" / "out_screener"
TSUBO_TO_M2 = 3.305785

# property_screener.py が要求するキー
SCREENER_KEYS = ["物件名", "価格_円", "土地面積_m2", "路線価_円per_m2",
                 "容積率_pct", "想定売却額_円", "建築費_円"]


def extract_text(pdf_path):
    """pdfplumber で全ページのテキストを連結。空なら画像PDFの可能性。"""
    import pdfplumber
    chunks = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            chunks.append(t)
    # NFKC正規化：athome等のCJK互換文字(⼟⾯)→通常字(土面)、全角→半角に統一
    return unicodedata.normalize("NFKC", "\n".join(chunks))


def yen_from(text):
    """「3,000万円」「30,000,000円」「3000万」→ 円(int)。最初の妥当値を返す。"""
    # 万円表記
    m = re.search(r"(?:価格|売値|販売価格|物件価格)[^\d]{0,8}([\d,]+(?:\.\d+)?)\s*万", text)
    if m:
        return int(float(m.group(1).replace(",", "")) * 10000)
    m = re.search(r"([\d,]+(?:\.\d+)?)\s*万円", text)
    if m:
        return int(float(m.group(1).replace(",", "")) * 10000)
    # 円表記（8桁前後）
    m = re.search(r"(?:価格|売値)[^\d]{0,8}([\d,]{7,12})\s*円", text)
    if m:
        return int(m.group(1).replace(",", ""))
    return None


def area_m2_from(text):
    """土地/敷地面積 ㎡ または 坪 → ㎡(float)。"""
    m = re.search(r"(?:土地面積|敷地面積|地積)[^\d]{0,6}([\d,]+\.?\d*)\s*(?:㎡|m2|平米|平方メートル)", text)
    if m:
        return float(m.group(1).replace(",", ""))
    # athome系：「公簿 335.66㎡」（土地面積ラベルが縦割れ）
    m = re.search(r"公簿\s*([\d,]+\.?\d*)\s*(?:㎡|m2|平米)", text)
    if m:
        return float(m.group(1).replace(",", ""))
    m = re.search(r"(?:土地面積|敷地面積|地積)[^\d]{0,6}([\d,]+\.?\d*)\s*坪", text)
    if m:
        return round(float(m.group(1).replace(",", "")) * TSUBO_TO_M2, 2)
    return None


def yoseki_from(text):
    """容積率 % を抽出。「容積率 200%」「容 積 率 200%」「200/100」など。"""
    m = re.search(r"容\s*積\s*率[^\d]{0,4}(\d{2,4})\s*[%％]", text)
    if m:
        return float(m.group(1))
    m = re.search(r"容\s*積\s*率[^\d]{0,4}(\d{2,4})\s*/\s*100", text)
    if m:
        return float(m.group(1))
    return None


def nenchin_from(text):
    """年額賃料（円）。「年額賃料等：3,528,000円」等。月額しか無ければ×12。"""
    m = re.search(r"年額賃料[等]?\s*[：:]\s*([\d,]+)\s*円", text)
    if m:
        return int(m.group(1).replace(",", ""))
    m = re.search(r"月額賃料[等]?\s*[：:]\s*([\d,]+)\s*円", text)
    if m:
        return int(m.group(1).replace(",", "")) * 12
    return None


def kenpei_from(text):
    m = re.search(r"建[蔽ぺ][率][^\d]{0,4}(\d{2,3})\s*[%％]", text)
    return float(m.group(1)) if m else None


def youto_from(text):
    """用途地域。代表的13種を順にマッチ。"""
    for y in ["第一種低層住居専用地域", "第二種低層住居専用地域",
              "第一種中高層住居専用地域", "第二種中高層住居専用地域",
              "第一種住居地域", "第二種住居地域", "準住居地域",
              "近隣商業地域", "商業地域", "準工業地域", "工業地域",
              "工業専用地域", "田園住居地域"]:
        if y in text:
            return y
    return None


def address_from(text):
    """所在地。都道府県から始まる行を拾う。"""
    m = re.search(r"((?:北海道|青森県|岩手県|宮城県|秋田県|山形県|福島県|茨城県|栃木県|群馬県|"
                  r"埼玉県|千葉県|東京都|神奈川県|新潟県|富山県|石川県|福井県|山梨県|長野県|"
                  r"岐阜県|静岡県|愛知県|三重県|滋賀県|京都府|大阪府|兵庫県|奈良県|和歌山県|"
                  r"鳥取県|島根県|岡山県|広島県|山口県|徳島県|香川県|愛媛県|高知県|福岡県|"
                  r"佐賀県|長崎県|熊本県|大分県|宮崎県|鹿児島県|沖縄県)[^\n　 ]{2,30})", text)
    return m.group(1).strip() if m else None


def setsudo_from(text):
    """接道。「南 6m 公道」のような記述を拾う（参考情報）。"""
    m = re.search(r"(接道|前面道路)[^\n]{0,30}", text)
    return m.group(0).strip() if m else None


def parse(pdf_path):
    name = Path(pdf_path).stem
    text = extract_text(pdf_path)
    text_len = len(text.strip())

    fields = {
        "物件名": name,
        "価格_円": yen_from(text),
        "土地面積_m2": area_m2_from(text),
        "容積率_pct": yoseki_from(text),
        "年賃料_円": nenchin_from(text),   # 収益マイソク→インカムモード自動起動
        # screener が使うが PDF からは通常取れない → None（後フェーズで路線価スクレイパー / 手入力）
        "路線価_円per_m2": None,
        "想定売却額_円": None,
        "建築費_円": None,
    }
    ref = {  # 参考情報（screener対象外だが判断に効く）
        "用途地域": youto_from(text),
        "建蔽率_pct": kenpei_from(text),
        "所在地": address_from(text),
        "接道メモ": setsudo_from(text),
    }

    missing = [k for k in ("価格_円", "土地面積_m2", "容積率_pct") if fields[k] is None]
    intake_meta = {
        "_抽出元PDF": str(pdf_path),
        "_抽出日時": datetime.datetime.now().isoformat(timespec="seconds"),
        "_抽出方式": "pdfplumber+regex",
        "_テキスト文字数": text_len,
        "_要Claude確認": missing,
        "_画像PDF疑い": text_len < 50,
        "_参考情報": ref,
    }
    return fields, intake_meta, text_len, missing


def main():
    if len(sys.argv) < 2:
        print("使い方: python3 property_intake.py <物件PDFパス>")
        sys.exit(1)
    pdf_path = sys.argv[1]
    if not Path(pdf_path).exists():
        print(f"❌ ファイルが見つかりません: {pdf_path}")
        sys.exit(1)

    fields, meta, text_len, missing = parse(pdf_path)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / f"intake_{fields['物件名']}.json"
    payload = {**fields, **meta}
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    print("=" * 56)
    print(f"  PDF抽出インテーク — {fields['物件名']}")
    print("=" * 56)
    print(f"  抽出方式: pdfplumber+regex  (テキスト{text_len}文字)")
    if meta["_画像PDF疑い"]:
        print("  ⚠️ 画像PDFの可能性大（テキストほぼ無し）")
        print("     → Claude推論フォールバック: ClaudeにPDFをReadさせて各項目を埋める")
    print("-" * 56)
    print(f"  価格    : {fields['価格_円']:,} 円" if fields['価格_円'] else "  価格    : ❌ 未抽出")
    print(f"  土地面積: {fields['土地面積_m2']} ㎡" if fields['土地面積_m2'] else "  土地面積: ❌ 未抽出")
    print(f"  容積率  : {fields['容積率_pct']}%" if fields['容積率_pct'] else "  容積率  : ❌ 未抽出")
    r = meta["_参考情報"]
    print(f"  用途地域: {r['用途地域'] or '—'}  建蔽率: {r['建蔽率_pct'] or '—'}")
    print(f"  所在地  : {r['所在地'] or '—'}")
    print("-" * 56)
    if missing:
        print(f"  要Claude確認/補完: {', '.join(missing)}")
        print(f"  （路線価・想定売却額・建築費は別途入力／後フェーズで自動化）")
    print("=" * 56)
    print(f"\n📄 抽出JSON: {out}")
    print(f"▶ 次: python3 scripts/property_screener.py {out}")


if __name__ == "__main__":
    main()
