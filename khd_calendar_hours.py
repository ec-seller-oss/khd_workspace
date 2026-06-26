# -*- coding: utf-8 -*-
# KHD カレンダー実績h 自動集計エンジン v1（2026-06-02）
# 統合ダッシュボード②本部マトリクスの「実績h」を、Googleカレンダーから自動集計して書き込む配線。
# 運用：① Claudeがカレンダーをlist_events→JSON保存 → ② このスクリプトで本部別に集計 → ③ マトリクスへ書込。
# 仕分けキー：予定タイトル先頭の本部番号(00_/01_/03_/04_/05_)を最優先＋キーワードで事業へサブ分類。
import json, re, sys
from datetime import datetime

# ---- 本部マトリクスの事業バケツ（②の行ラベルに対応） ----
# matrix行: 不動産 / 医療 / EC / テレアポ / 調査士 / メディア / 家族  ＋ 非事業: 内務 / その他
def classify(title):
    t = title.strip()
    m = re.match(r'^\s*(\d{2})[_\.]', t)
    honbu = m.group(1) if m else None
    has_addr = any(k in t for k in ['丁目','番地','県','市','区','町','字'])

    # 00_ = 家族 or 物件住所
    if honbu == '00':
        if any(k in t for k in ['親子','家族','葵斗','モーニング','会議']): return '家族'
        if has_addr: return '不動産'
        return '家族'
    # 01/02 = 経営・資金（内務）
    if honbu in ('01','02'): return '内務'
    # 03 = 事業運営（EC or 不動産 or 運営内務）
    if honbu == '03':
        if any(k in t for k in ['韓国輸出','クーパン','EC','せどり']): return 'EC'
        if any(k in t for k in ['バイセル','物上げ','仕入']): return '不動産'
        return '内務'   # 朝礼/終礼(組織化)等の運営
    # 04 = コンサル・調査士
    if honbu == '04':
        if any(k in t for k in ['調査士','土地家屋','マン菅','賃管','診断士']): return '調査士'
        if any(k in t for k in ['オーロラ','テレアポ','インディード','鍼灸']): return 'テレアポ'
        if any(k in t for k in ['YouTube','メディア','HP']): return 'メディア'
        return '医療'   # TAW/歯科/セミナー/クリニックDX/リベ大家の会 等
    # 05 = 物件調達
    if honbu == '05': return '不動産'

    # --- プレフィックス無し：キーワードのみ ---
    if any(k in t for k in ['親子','家族','葵斗']): return '家族'
    if any(k in t for k in ['調査士','土地家屋','自己投資']): return '調査士'
    if any(k in t for k in ['韓国輸出','クーパン']): return 'EC'
    if any(k in t for k in ['オーロラ','テレアポ','インディード','石原','鍼灸','採用']): return 'テレアポ'
    if any(k in t for k in ['そうけん','YouTube','My AI','MyAI','メディア','HP','デッキ']): return 'メディア'
    if any(k in t for k in ['TAW','歯科','医療','セミナー','福井','クリニック','診療','野口','ソニー生命']): return '医療'
    if any(k in t for k in ['バイセル','物上げ','仕入','決済']) or has_addr: return '不動産'
    if any(k in t for k in ['台帳','BS','DB','ダッシュ','パイプライン','週次','KPI','報告']): return '内務'
    return 'その他'

def hours(ev):
    s = ev['start'].get('dateTime'); e = ev['end'].get('dateTime')
    if not s or not e: return 0.0  # 終日予定は時間集計しない
    fmt = lambda x: datetime.fromisoformat(x)
    return (fmt(e) - fmt(s)).total_seconds() / 3600.0

def aggregate(events, start, end):
    """[start,end) JST の時間指定イベントを本部別に集計"""
    lo = datetime.fromisoformat(start); hi = datetime.fromisoformat(end)
    agg = {}
    detail = {}
    for ev in events:
        s = ev['start'].get('dateTime')
        if not s: continue
        st = datetime.fromisoformat(s)
        if not (lo <= st < hi): continue
        b = classify(ev.get('summary',''))
        h = hours(ev)
        agg[b] = agg.get(b, 0.0) + h
        detail.setdefault(b, []).append((ev.get('summary',''), round(h,1)))
    return agg, detail

if __name__ == '__main__':
    jsonpath = sys.argv[1]
    start = sys.argv[2]   # 例 2026-06-01T00:00:00+09:00
    end   = sys.argv[3]
    data = json.load(open(jsonpath))
    events = data.get('events', data if isinstance(data, list) else [])
    agg, detail = aggregate(events, start, end)

    order = ['不動産','医療','EC','テレアポ','調査士','メディア','家族','内務','その他']
    biz = ['不動産','医療','EC','テレアポ','調査士','メディア','家族']
    print(f"=== 本部別 実績h（{start[:10]} 〜 {end[:10]}）===")
    tot=0.0; biz_tot=0.0; eigyo=0.0
    for k in order:
        if k in agg:
            print(f"  {k:6s}: {agg[k]:5.1f}h")
            tot += agg[k]
            if k in biz: biz_tot += agg[k]
            if k in ('不動産','医療','テレアポ'): eigyo += agg[k]
    print(f"  {'─'*18}")
    print(f"  合計      : {tot:5.1f}h（うち事業{biz_tot:.1f}h／営業直結(不動産+医療+テレアポ){eigyo:.1f}h＝{eigyo/tot*100 if tot else 0:.0f}%）")
    # JSON出力（マトリクス書込用）
    print("\n__HOURS__" + json.dumps({k: round(agg.get(k,0),1) for k in biz}, ensure_ascii=False))

    # ---- 第4引数にxlsxパスがあれば②本部マトリクスの実績h(D列)へ書込 ----
    # 全活動を計上（内務・その他含む）。営業直結%・構成比・ROIはシート側の数式が自動計算。
    if len(sys.argv) >= 5:
        import openpyxl
        xlsx = sys.argv[4]
        wb = openpyxl.load_workbook(xlsx)
        mx = wb['②本部マトリクス']
        # マトリクス行(5-13)とバケツの対応（新レイアウト）
        row_of = {5:'不動産',6:'医療',7:'EC',8:'調査士',9:'メディア',10:'テレアポ',11:'内務',12:'その他',13:'家族'}
        for r,b in row_of.items():
            mx.cell(r,3, round(agg.get(b,0.0),1))   # C列=実績h
        mx['A2'] = (f"時間は総時間=ゼロサム予算。実績h=カレンダー自動集計（窓 {start[:10]}〜{end[:10]}・"
                    f"全{tot:.1f}h計上）。ROI=粗利実績÷実績h。営業直結%/構成比/ROIは下の数式で自動算出。"
                    f"日次/週次=時間規律・営業直結%／月次=ROIで続ける・減らす・やめる判断。")
        wb.save(xlsx)
        print(f"\n✅ ②本部マトリクスのD列(実績h・全活動)へ書込完了: {xlsx}")
