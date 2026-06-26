#!/usr/bin/env python3
"""
ウエルシア物件 × 先生ニーズ 月次照合パイプライン（半自動）
使い方:  python3 welcia_match_pipeline.py
  Drive同期の最新ウエルシアzipを展開→全PDFテキスト化→先生エリア×科目で自動照合
  → マッチ候補レポートを標準出力。これを見て福井下書きを作る（最終判断は人/Claude）。

先生が増えたら下の SENSEI リストに1行追記するだけ（先生ニーズリストスプシと同期）。
"""
import zipfile, os, re, glob
import pdfplumber

# === 先生ニーズ（先生ニーズリストスプシと同期。増えたら追記）===
SENSEI = [
    {"name": "曾我(眼科)",     "pref": ["東京都", "千葉県", "埼玉県", "神奈川県"], "area": ["台東", "足立", "千葉", "埼玉", "神奈川", "横浜", "川崎"], "sci": ["眼科"], "tsubo_min": 60, "yen_max": 15000},
    {"name": "内山(内科)",     "pref": ["東京都"], "area": ["渋谷", "中野", "新宿", "西新宿"],         "sci": ["内科"], "tsubo_min": 35, "yen_max": None},
    {"name": "長西(婦人科)",   "pref": ["東京都"], "area": ["練馬", "板橋", "吉祥寺", "中野", "杉並"], "sci": ["婦人科", "産婦"], "tsubo_min": 30, "yen_max": None},
    {"name": "星山(内科)",     "pref": ["神奈川県"], "area": ["長津田", "十日市場", "鴨居", "小机", "緑区", "青葉"], "sci": ["内科"], "tsubo_min": None, "yen_max": None},
    {"name": "桑原(産婦人科)", "pref": ["長野県"], "area": ["御代田", "小諸", "佐久", "軽井沢"],       "sci": ["産婦", "婦人科"], "tsubo_min": None, "yen_max": None},
    {"name": "奥村(耳鼻科)",   "pref": ["東京都"], "area": ["中野"],                                 "sci": ["耳鼻"], "tsubo_min": 55, "yen_max": None},
    {"name": "眞木(泌尿器科)", "pref": ["東京都"], "area": ["四谷", "新宿", "四ツ谷"],               "sci": ["泌尿"], "tsubo_min": 40, "yen_max": None},
    {"name": "秋元(内科)",     "pref": ["東京都"], "area": ["西池袋", "池袋", "豊島"],               "sci": ["内科"], "tsubo_min": None, "yen_max": None},
]

ZIP_GLOB = os.path.expanduser(
    '~/Library/CloudStorage/GoogleDrive-ec-seller@kikuchi-hd.net/マイドライブ/281_先生一覧/ウエルシア物件情報*.zip')
WORK = '/tmp/welcia_pipeline'


def extract_zip():
    src = sorted(glob.glob(ZIP_GLOB))[-1]
    os.makedirs(WORK, exist_ok=True)
    z = zipfile.ZipFile(src)
    pdfs = []
    for info in z.infolist():
        if info.is_dir():
            continue
        try:
            name = info.filename.encode('cp437').decode('cp932')
        except Exception:
            name = info.filename
        if not name.lower().endswith('.pdf') or '公開削除' in name:
            continue
        flat = name.split('/', 1)[1].replace('/', '__') if '/' in name else name
        out = os.path.join(WORK, flat)
        with open(out, 'wb') as f:
            f.write(z.read(info))
        pdfs.append(out)
    return src, pdfs


def pdf_text(path):
    try:
        with pdfplumber.open(path) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        return ""


def parse_property(text, fname):
    m = re.search(r'所\s*在\s*地\s*[:：．\.]*\s*([^\n]+)', text)
    addr = m.group(1).strip() if m else ""
    m = re.search(r'募集科目\s*[:：]?\s*([^\n]+)', text)
    sci = m.group(1).strip() if m else ""
    tsubo = [float(x) for x in re.findall(r'([0-9]+\.?[0-9]*)\s*坪', text)]
    m = re.search(r'坪\s*([0-9,]+)\s*円', text)
    yen = int(m.group(1).replace(',', '')) if m else None
    full = addr + " " + sci  # エリア判定は所在+科目欄テキスト両方を対象
    return {"file": os.path.basename(fname), "addr": addr, "sci": sci,
            "tsubo": tsubo, "yen": yen, "blob": text[:400]}


def get_pref(addr):
    m = re.search(r'(東京都|北海道|京都府|大阪府|[^\s　]{2,3}県)', addr)
    return m.group(1) if m else ""


def match(prop):
    hits = []
    target = prop["addr"] or prop["blob"]
    pref = get_pref(prop["addr"])
    for s in SENSEI:
        # 都道府県チェック（他県の同名地名を弾く）
        if pref and pref not in s["pref"]:
            continue
        if not any(a in target for a in s["area"]):
            continue
        tokens = re.split(r'[・、,／/\s]+', prop["sci"])
        sci_ok = ("全" in prop["sci"]) or any(k in tokens for k in s["sci"])
        # 「以外」表記（◯◯科以外を募集）は科目不一致扱い
        if prop["sci"] and "以外" in prop["sci"]:
            sci_ok = False
        if not sci_ok:
            continue
        notes = []
        if s["tsubo_min"] and prop["tsubo"]:
            maxt = max(prop["tsubo"])
            if maxt < s["tsubo_min"]:
                notes.append(f"面積△(最大{maxt}坪<{s['tsubo_min']})")
        if s["yen_max"] and prop["yen"] and prop["yen"] > s["yen_max"]:
            notes.append(f"賃料△(坪{prop['yen']}>{s['yen_max']})")
        hits.append((s["name"], " ".join(notes) or "◎条件OK"))
    return hits


def main():
    src, pdfs = extract_zip()
    print(f"# ウエルシア × 先生ニーズ 照合レポート")
    print(f"元zip: {os.path.basename(src)} / 現役PDF {len(pdfs)}件\n")
    matched = 0
    for p in pdfs:
        text = pdf_text(p)
        if not text.strip():
            continue
        prop = parse_property(text, p)
        hits = match(prop)
        if hits:
            matched += 1
            print(f"## {prop['file']}")
            print(f"- 所在: {prop['addr']}")
            print(f"- 科目: {prop['sci']}")
            print(f"- 坪: {prop['tsubo']} / 坪単価: {prop['yen']}")
            for name, note in hits:
                print(f"  → {name} : {note}")
            print()
    print(f"# マッチ物件: {matched}件 / 現役{len(pdfs)}件")


if __name__ == "__main__":
    main()
