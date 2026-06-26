#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KHD 土地から新築 一次スクリーニング・パイプライン  v1
--------------------------------------------------------
マイソク1枚の数値を入れると、5ステップで「何階・何造・何戸・粗利の枠」の
"当たり" を出す、菊池の自前ツール。

※これは菊池研太が自分の言葉・自分のワークフローとして組んだ一次スクリーニング機。
  根拠は建築基準法・各自治体条例という公知の法令の一般的な考え方であり、
  特定有料コンテンツの本文・図・具体例の複製ではない。
  正式なボリューム判断は必ず建築士・自治体窓口で確認すること（本ツールはあくまで仕入の足切り用）。

使い方:
    python3 land_screening_pipeline.py            # サンプル物件で実行
    （実物件は SITE のdictを書き換えて再実行）
"""

from dataclasses import dataclass, field


# ============================================================
# 入力：マイソクから読み取る敷地条件
# ============================================================
@dataclass
class Site:
    name: str                 # 物件名
    area_m2: float            # 敷地面積(㎡)
    yoto: str                 # 用途地域（'住居系' or '商業系'）
    kenpei: float             # 建蔽率(%)  ※角地緩和・防火緩和は反映後の値を入れる
    yoseki: float             # 指定容積率(%)
    road_w_m: float           # 前面道路幅員(m)
    bouka: str                # 防火指定（'防火' / '準防火' / '法22条' / 'なし'）
    hokugawa_shasen: bool     # 北側斜線（高度地区/低層住専）あり=True
    nichiei: bool             # 日影規制 あり=True
    juko_men_m2: float = 25.0 # 想定1住戸の専有面積(㎡)
    rosenka_man_per_m2: float = 0.0  # 路線価(万円/㎡) ※土地値割合の判定用(任意)
    shiire_man: float = 0.0          # 想定仕入価格(万円)  ※KPI判定用(任意)


# ============================================================
# Step1: 敷地条件を正確に読む（法定上限を出す）
# ============================================================
def step1_site_caps(s: Site):
    # 道路幅員による容積低減（住居系×0.4 / 商業系×0.6 が一般的係数）
    coef = 0.4 if s.yoto == '住居系' else 0.6
    road_yoseki = s.road_w_m * coef * 100         # 道路幅員容積(%)
    eff_yoseki = min(s.yoseki, road_yoseki)        # 実効容積率(%)

    max_kenchiku = s.area_m2 * s.kenpei / 100      # 法定建築面積上限(㎡)
    max_enshou   = s.area_m2 * eff_yoseki / 100    # 法定延床上限(㎡)
    return {
        'eff_yoseki': eff_yoseki,
        'road_yoseki': road_yoseki,
        'yoseki_binding': '道路幅員' if road_yoseki < s.yoseki else '指定容積',
        'max_kenchiku': max_kenchiku,
        'max_enshou': max_enshou,
    }


# ============================================================
# Step2: 駐車場を先に計画（必要台数の当たり）
# ============================================================
def step2_parking(s: Site, units_rough: float):
    # 単身向け(専有25㎡前後)は附置義務ゆるめ・ファミリーは厚め、の粗い当たり
    ratio = 0.3 if s.juko_men_m2 < 30 else 0.7
    need = round(units_rough * ratio, 1)
    return {'parking_ratio': ratio, 'parking_need': need,
            'note': '建物より先に駐車場配置を検討。台数が出口に直結する地域は要注意。'}


# ============================================================
# Step3: 建築可能範囲（斜線・離隔のチェック）
# ============================================================
def step3_envelope(s: Site):
    # 道路斜線：住居系1.25 / 商業系1.5 を勾配係数とした高さの目安
    slope = 1.25 if s.yoto == '住居系' else 1.5
    road_height = round(s.road_w_m * slope, 2)  # 道路境界での高さ目安(m)
    flags = []
    flags.append(f'道路斜線の高さ目安 ≒ {road_height}m（幅員{s.road_w_m}m×{slope}）')
    if s.hokugawa_shasen:
        flags.append('北側斜線/高度地区あり → 北側を削る前提。3階の北側住戸が痩せやすい')
    if s.nichiei:
        flags.append('日影規制あり → 4階以上(高さ10m超)は要日影チェック。東西に長い敷地ほど不利')
    flags.append('隣地離隔 約0.5〜0.6m、敷地内通路/窓先空地（自治体条例）を別途確認')
    return {'road_height': road_height, 'flags': flags}


# ============================================================
# Step4: 構造3択を法規から逆算
# ============================================================
def step4_structure(floors_guess: int, max_enshou: float):
    """木三共 / 耐火木造 / RC の当たり"""
    cands = []
    if floors_guess <= 3 and max_enshou < 200:
        cands.append(('木造(耐火不要域)', '延床200㎡未満・3階以下なら避難/採光規制が一気に緩む。最有利ゾーン'))
    if floors_guess <= 3:
        cands.append(('木三共(木造3階建共同住宅)', 'コスト最安級。条件A/Bの充足が要件。延床は伸ばしにくい'))
        cands.append(('耐火木造3階', '木三共で延床が伸びない時の解。バルコニー/共用部を屋内化し戸数を稼げる'))
    if floors_guess >= 4:
        cands.append(('RC造', '4階以上・容積を使い切る時。コスト高。日影/斜線をRCで詰める'))
        cands.append(('耐火木造', '4階でもコスト圧縮狙い。要構造検討'))
    return cands


# ============================================================
# Step5: 戸数を現実に絞る（単純割り算を信じない）
# ============================================================
def step5_units(max_enshou: float, juko_men: float):
    naive = max_enshou / juko_men                    # 容積÷住戸面積(上限の幻)
    real = naive * 0.75                              # 共用部(廊下/階段/EV/バルコニー)で約75%
    return {
        'naive_units': round(naive, 1),
        'real_units': int(real),
        'warn': f'単純割り算{round(naive,1)}戸は上限の幻。廊下・階段・避難通路で現実は約{int(real)}戸。',
    }


# ============================================================
# KPI：土地値割合・粗利の枠（任意入力があれば）
# ============================================================
def kpi_check(s: Site):
    out = []
    if s.rosenka_man_per_m2 > 0 and s.shiire_man > 0:
        landval = s.rosenka_man_per_m2 * s.area_m2
        ratio = landval / s.shiire_man
        verdict = 'OK(担保価値あり)' if ratio >= 0.4 else '注意(土地値薄い)'
        out.append(f'土地値割合 = {round(ratio,2)} （路線価ベース土地値{round(landval)}万 ÷ 仕入{s.shiire_man}万） → {verdict}（基準0.4）')
    else:
        out.append('土地値割合: 路線価・仕入価格を入れると自動判定（玉川式基準0.4以上が安全圏）')
    return out


# ============================================================
# 実行：5ステップを順に回して当たりを出す
# ============================================================
def screen(s: Site, floors_guess: int = 3):
    print('=' * 60)
    print(f'■ 物件: {s.name}')
    print('=' * 60)

    st1 = step1_site_caps(s)
    print('\n【Step1 敷地条件・法定上限】')
    print(f'  実効容積率 = {st1["eff_yoseki"]:.0f}%（拘束＝{st1["yoseki_binding"]} / 道路容積{st1["road_yoseki"]:.0f}%）')
    print(f'  建築面積上限 = {st1["max_kenchiku"]:.1f}㎡ / 延床上限 = {st1["max_enshou"]:.1f}㎡')
    if s.bouka == 'なし':
        print('  ⚠ 防火指定「なし」: 建蔽の緩和を勝手に+10%しない。法22条/準防火を必ず確認')

    st5 = step5_units(st1['max_enshou'], s.juko_men_m2)
    st2 = step2_parking(s, st5['real_units'])
    print('\n【Step2 駐車場（建物より先に）】')
    print(f'  必要台数の当たり ≒ {st2["parking_need"]}台（戸数{st5["real_units"]}×係数{st2["parking_ratio"]}）')
    print(f'  {st2["note"]}')

    st3 = step3_envelope(s)
    print('\n【Step3 建築可能範囲（斜線・離隔）】')
    for f in st3['flags']:
        print(f'  - {f}')

    print('\n【Step4 構造3択（法規から逆算）】')
    for name, why in step4_structure(floors_guess, st1['max_enshou']):
        print(f'  ◎ {name}: {why}')

    print('\n【Step5 戸数（単純割り算を信じない）】')
    print(f'  ⚠ {st5["warn"]}')
    print(f'  → 当たり戸数: 約{st5["real_units"]}戸（{floors_guess}階建て想定）')

    print('\n【KPI】')
    for k in kpi_check(s):
        print(f'  - {k}')

    print('\n【一次判定サマリ】')
    print(f'  当たり: {floors_guess}階建て / 延床上限{st1["max_enshou"]:.0f}㎡ / 約{st5["real_units"]}戸')
    print('  ※構造は木造200㎡未満→木三共→耐火木造→RCの順で法規/コストから選ぶ')
    print('  ※正式判断は建築士・自治体確認。本ツールは仕入の足切り用')
    print()


# ------------------------------------------------------------
# サンプル物件（実物件は値を書き換えて再実行）
# ------------------------------------------------------------
if __name__ == '__main__':
    sample = Site(
        name='サンプル：首都圏 角地 第一種住居',
        area_m2=180.0,
        yoto='住居系',
        kenpei=60.0,        # 角地緩和込み
        yoseki=200.0,
        road_w_m=4.0,
        bouka='準防火',
        hokugawa_shasen=False,
        nichiei=True,
        juko_men_m2=25.0,
        rosenka_man_per_m2=18.0,
        shiire_man=4500.0,
    )
    screen(sample, floors_guess=3)
