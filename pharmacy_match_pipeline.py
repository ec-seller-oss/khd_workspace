#!/usr/bin/env python3
"""
薬局全社 物件 × 先生ニーズ 月次照合パイプライン（半自動）
対象: 281_先生一覧 内の全薬局zip（ウエルシア / ご案内物件=日本調剤 など）
使い方: python3 pharmacy_match_pipeline.py
  各zip展開(Shift-JIS)→ファイル名で先生エリア一次フィルタ→該当のみPDFで科目/坪確認→照合レポート
先生が増えたら SENSEI に1行追記。依存: pip install pdfplumber
"""
import zipfile, os, re, glob
import pdfplumber

SENSEI = [
    {"name": "曾我(眼科)",     "pref": ["東京都", "千葉県", "埼玉県", "神奈川県"], "area": ["台東", "足立", "千葉", "埼玉", "神奈川", "横浜", "川崎"], "sci": ["眼科"], "tsubo_min": 60},
    {"name": "内山(内科)",     "pref": ["東京都"], "area": ["渋谷", "中野", "新宿", "西新宿"],         "sci": ["内科"], "tsubo_min": 35},
    {"name": "長西(婦人科)",   "pref": ["東京都"], "area": ["練馬", "板橋", "吉祥寺", "中野", "杉並"], "sci": ["婦人科", "産婦"], "tsubo_min": 30},
    {"name": "星山(内科)",     "pref": ["神奈川県"], "area": ["長津田", "十日市場", "鴨居", "小机", "緑区", "青葉"], "sci": ["内科"], "tsubo_min": None},
    {"name": "桑原(産婦人科)", "pref": ["長野県"], "area": ["御代田", "小諸", "佐久", "軽井沢"],       "sci": ["産婦", "婦人科"], "tsubo_min": None},
    {"name": "奥村(耳鼻科)",   "pref": ["東京都"], "area": ["中野"],                                 "sci": ["耳鼻"], "tsubo_min": 55},
    {"name": "眞木(泌尿器科)", "pref": ["東京都"], "area": ["四谷", "新宿", "四ツ谷"],               "sci": ["泌尿"], "tsubo_min": 40},
    {"name": "秋元(内科)",     "pref": ["東京都"], "area": ["西池袋", "池袋", "豊島"],               "sci": ["内科"], "tsubo_min": None},
]

BASE = os.path.expanduser('~/Library/CloudStorage/GoogleDrive-ec-seller@kikuchi-hd.net/マイドライブ/281_先生一覧')
WORK = '/tmp/pharmacy_pipeline'


def pharmacy_name(src):
    b = os.path.basename(src)
    if 'ウエルシア' in b:
        return 'ウエルシア'
    if 'ご案内' in b:
        return '日本調剤'
    return b


def iter_pdfs():
    for src in glob.glob(os.path.join(BASE, '*.zip')):
        ph = pharmacy_name(src)
        z = zipfile.ZipFile(src)
        for info in z.infolist():
            if info.is_dir():
                continue
            try:
                name = info.filename.encode('cp437').decode('cp932')
            except Exception:
                name = info.filename
            if not name.lower().endswith('.pdf') or '公開削除' in name:
                continue
            yield ph, name, info, z


def get_pref(s):
    m = re.search(r'(東京都|北海道|京都府|大阪府|[^\s　]{2,3}県)', s)
    return m.group(1) if m else ""


def area_hit(text, sensei):
    pref = get_pref(text)
    if pref and pref not in sensei["pref"]:
        return False
    return any(a in text for a in sensei["area"])


def main():
    os.makedirs(WORK, exist_ok=True)
    results = []
    scanned = 0
    for ph, name, info, z in iter_pdfs():
        scanned += 1
        fname = name.split('/')[-1]
        # 1次: ファイル名(所在含む)で先生エリア該当を判定
        cands = [s for s in SENSEI if area_hit(name, s)]
        if not cands:
            continue
        # 2次: PDFで科目・坪を確認
        out = os.path.join(WORK, fname.replace('/', '_'))
        with open(out, 'wb') as f:
            f.write(z.read(info))
        try:
            with pdfplumber.open(out) as pdf:
                text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        except Exception:
            text = ""
        m = re.search(r'募集科目\s*[:：]?\s*([^\n]+)', text)
        sci = m.group(1).strip() if m else ""
        tsubo = [float(x) for x in re.findall(r'([0-9]+\.?[0-9]*)\s*坪', text)]
        sci_tokens = re.split(r'[・、,／/\s]+', sci)
        broad = ('クリニック募集' in fname) or ('医療モール' in fname) or ('全' in sci)
        am = re.search(r'所\s*在\s*地\s*[:：．\.]*\s*([^\n]+)', text)
        addr_pref = get_pref(am.group(1)) if am else ""
        for s in cands:
            if addr_pref and addr_pref not in s["pref"]:
                continue
            sci_ok = broad or any(k in sci_tokens for k in s["sci"])
            if sci and '以外' in sci:
                sci_ok = False
            if not sci_ok:
                continue
            note = []
            if s["tsubo_min"] and tsubo and max(tsubo) < s["tsubo_min"]:
                note.append(f"面積△(最大{max(tsubo)}坪<{s['tsubo_min']})")
            if not sci:
                note.append("科目はPDF未記載→要確認")
            results.append((ph, fname, s["name"], " ".join(note) or "◎条件OK"))

    print("# 薬局全社 × 先生ニーズ 照合レポート")
    print(f"走査PDF: {scanned}件 / マッチ: {len(results)}件\n")
    cur = None
    for ph, fname, sname, note in results:
        if ph != cur:
            print(f"\n=== {ph} ===")
            cur = ph
        print(f"- {fname[:55]}")
        print(f"    → {sname} : {note}")


if __name__ == "__main__":
    main()
