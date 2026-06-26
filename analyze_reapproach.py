#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""進捗管理表(クローズ案件)を解析し、再アプローチ(荷電)で復活しそうな順に優先リスト化。
   軸: ①売却意思の継続 ②価格目線が下がる/下がった示唆 ③他社成約・決済済は除外。"""
import json, csv, re
from pathlib import Path

F = "/Users/kikuchikenta/.claude/projects/-Users-kikuchikenta-01-honbu-docs-automation/12fdcfcd-e76e-45f9-b967-67bacdb8b466/tool-results/mcp-794fe495-8a1a-4fce-993e-122ed2414029-read_file_content-1781065189809.txt"
t = json.load(open(F))['fileContent']
recs = t.split(',e ')
rows = recs[1:]

# 復活見込み(加点): 売却意思継続 / 価格が下がる・下がった / 待ち・再連絡 / 反響なし高値
HIGH = {
  '意思は変わらず':3,'売却の意思':3,'売りたい':3,'売る意思':3,
  '待ち':2,'まで待':2,'再度連絡':2,'また連絡':2,'後で連絡':2,'連絡もらう':2,'数ヶ月':2,'近くなったら':2,
  '反響なし':2,'反響はない':2,'反響無し':2,'売れない':2,'売れず':2,
  '高値':2,'高く出':1,'下げ':2,'値下げ':2,'下がっ':2,'下がる':2,'価格を下げ':3,
  'ウォッチ':2,'謄本':1,'レインズ':1,'金額合わず':2,'目線合わ':2,'目線が合':2,
  '専任':1,'募集':1,'希望価格':1,'時間を掛けれ':1,'安くても':2,'手放す':2,
}
# 復活不可(除外): 決済・契約完了・他社成約・白紙・逝去
KILL = ['決済','引き渡し','引渡','契約完了','契約実施','売買契約','成約','他社','他業者','他で買','取られた','白紙','逝去','死亡','解体済','更地済']

cand = []
killed = 0
for r in rows:
    f = r.split(',', 8)
    if len(f) < 9:
        continue
    date, gyosha, tanto, renraku, jusho, status, eigyo, kaden, joukyo = f[:9]
    joukyo = joukyo.strip().strip('"').strip()
    if 'クローズ' not in status:
        continue
    kills = [k for k in KILL if k in joukyo]
    score = 0
    hits = []
    for k, w in HIGH.items():
        if k in joukyo:
            score += w
            hits.append(k)
    if kills:
        killed += 1
        continue  # 成約済・他社・物理的に不可は除外
    if score <= 0:
        continue
    # 状況要約(末尾の最新動向を優先して120字)
    summ = joukyo[-160:] if len(joukyo) > 160 else joukyo
    cand.append({
        'score': score, 'date': date.strip(), 'gyosha': gyosha.strip(),
        'tanto': tanto.strip(), 'renraku': renraku.strip(), 'jusho': jusho.strip(),
        'eigyo': eigyo.strip(), 'hits': '/'.join(hits), 'summary': summ,
    })

cand.sort(key=lambda x: -x['score'])
print(f"総レコード:{len(rows)}  クローズ解析  除外(成約/他社/決済等):{killed}  再アプローチ候補:{len(cand)}")
print("=== 上位15件プレビュー ===")
for c in cand[:15]:
    print(f"[{c['score']}] {c['date']} {c['gyosha']}({c['eigyo']}) {c['jusho']} | {c['hits']}")

# CSV出力
out = Path.home()/"01_honbu_docs_automation"/"out_screener"/"05_クローズ案件_再アプローチ優先リスト.csv"
with open(out, 'w', encoding='utf-8') as fp:
    w = csv.writer(fp)
    w.writerow(['優先度','スコア','紹介日','業者名','担当者','連絡先','物件住所','元営業担当','復活シグナル','状況(最新)'])
    for i, c in enumerate(cand, 1):
        rank = 'A' if c['score'] >= 5 else ('B' if c['score'] >= 3 else 'C')
        w.writerow([rank, c['score'], c['date'], c['gyosha'], c['tanto'], c['renraku'], c['jusho'], c['eigyo'], c['hits'], c['summary']])
print("CSV:", out)
