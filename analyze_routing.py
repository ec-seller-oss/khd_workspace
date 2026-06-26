#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""クローズ案件を「物件タイプ×エリア」で分類し、最適な"投げ先"をタグ付け。
   広く=自分で再アプローチ/ツクビト再販/福井(医療用地)/自社(新築用地)/投資家/現地つて。"""
import json, csv, re
from pathlib import Path
from collections import Counter

F = "/Users/kikuchikenta/.claude/projects/-Users-kikuchikenta-01-honbu-docs-automation/12fdcfcd-e76e-45f9-b967-67bacdb8b466/tool-results/mcp-794fe495-8a1a-4fce-993e-122ed2414029-read_file_content-1781065189809.txt"
t = json.load(open(F))['fileContent']
rows = t.split(',e ')[1:]

# 医療(福井)が効く＝人口集積の市街地。新築AP/自社が効く＝土地値の出る郊外〜近郊
IRYO_AREA = ['さいたま','川口','越谷','草加','春日部','所沢','上尾','川越','八潮','三郷','蕨','戸田']

def classify(jusho, joukyo):
    s = jusho + ' ' + joukyo
    if re.search(r'マンション|ダイアパレス|ライオンズ|壁芯|管理費|区分|パレス|コーポ|レーベン|サーパス', s):
        return '区分'
    if re.search(r'更地|土地|分筆|セットバック|地目|畑|農地|建築条件|坪単価|㎡単価', s):
        return '土地'
    if re.search(r'戸建|築\d|木造|平屋|階建|リフォーム|空き家|再建築|ボロ|戸建て', s):
        return '戸建'
    return '不明'

def route(typ, jusho, joukyo):
    tags = []
    area_hit = any(a in jusho for a in IRYO_AREA)
    # 自分で再アプローチ(価格下落/意思継続)
    if re.search(r'意思は変わらず|売りたい|手放す|下げ|値下げ|反響なし|売れな|待ち|また連絡|再度連絡', joukyo):
        tags.append('①自分で再アプローチ')
    # ツクビト再販(戸建/区分の再販)
    if typ in ('戸建', '区分'):
        tags.append('②ツクビト再販')
    # 福井(医療/クリニック用地)=土地×市街地
    if typ == '土地' and area_hit:
        tags.append('③福井(医療用地)')
    # 自社/叔父(新築AP用地)=土地
    if typ == '土地':
        tags.append('④自社/TAW(新築用地)')
    # 投資家(賃貸可能性)
    if re.search(r'賃貸|投資|利回り|オーナーチェンジ|入居', joukyo):
        tags.append('⑤投資家')
    # 松戸=込山の地元
    if '松戸' in jusho:
        tags.append('⑥込山(松戸地元)')
    if not tags:
        tags.append('保留')
    return tags, area_hit

records = []
KILL = ['決済','引き渡し','引渡','契約完了','契約実施','売買契約','成約','他社','他業者','取られた','白紙','逝去']
type_c = Counter(); route_c = Counter()
for r in rows:
    f = r.split(',', 8)
    if len(f) < 9: continue
    date, gyosha, tanto, renraku, jusho, status, eigyo, kaden, joukyo = [x.strip().strip('"') for x in f[:9]]
    if 'クローズ' not in status: continue
    dead = any(k in joukyo for k in KILL)
    typ = classify(jusho, joukyo)
    tags, area_hit = route(typ, jusho, joukyo)
    type_c[typ]+=1
    if not dead:
        for tg in tags: route_c[tg]+=1
    records.append({
        'date':date,'gyosha':gyosha,'jusho':jusho,'type':typ,'eigyo':eigyo,
        'renraku':renraku,'dead':dead,'routes':' / '.join(tags),
        'summary':joukyo[-140:]
    })

print('=== 物件タイプ分布 ===')
for k,v in type_c.most_common(): print(f'  {k}: {v}')
print('=== 投げ先タグ別(終了案件除く) ===')
for k,v in route_c.most_common(): print(f'  {k}: {v}件')

# 福井(医療用地)候補を具体表示
print('=== ③福井(医療用地)候補 ===')
for r in records:
    if '③福井' in r['routes'] and not r['dead']:
        print(f"  {r['date']} {r['gyosha']} {r['jusho']} | {r['summary'][:70]}")

# CSV(終了案件除く・全ルーティング)
out = Path.home()/"01_honbu_docs_automation"/"out_screener"/"05_クローズ案件_投げ先ルーティング表.csv"
live=[r for r in records if not r['dead']]
live.sort(key=lambda x:(x['type'],x['jusho']))
with open(out,'w',encoding='utf-8') as fp:
    w=csv.writer(fp)
    w.writerow(['物件タイプ','投げ先候補','紹介日','業者名','物件住所','元担当','連絡先','状況(最新)'])
    for r in live:
        w.writerow([r['type'],r['routes'],r['date'],r['gyosha'],r['jusho'],r['eigyo'],r['renraku'],r['summary']])
print('CSV:',out,' 生存案件:',len(live))
