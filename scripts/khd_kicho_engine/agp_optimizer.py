#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGP 解約最適化 — Amexプリファードゴールドの「いつ移行→解約がベストか」分析。
=================================================================
背景: AGP1(研太)/AGP2(麻梨奈)はマイル移行律速で解約できず年会費を垂れ流し中。
      Amex MR→ANA移行ルールに照合し、最適な解約タイミングと損益を提示する。

★Amex MR→ANA移行ルール（2026-06時点・要再確認）:
  - 年間移行上限 = 80,000ポイント（=40,000マイル）/年（1/1移行〜12/31処理）
  - レート: 通常 2,000P→1,000マイル / MRプラス(年3,300円)加入で 1,000P→1,000マイル
  - ANAコース 年5,500円別途（プラチナ無料・プリファードゴールドは有料）
  - ★解約するとMRポイントは失効 → 移行しきる前に解約できない＝律速
"""
import math

# ===== ルール定数（改定時はここだけ直す）=====
ANA_ANNUAL_CAP_MILES = 40000      # 年間移行上限（マイル）
RATE_NORMAL = 0.5                 # 2000P→1000M
RATE_PLUS = 1.0                   # 1000P→1000M（MRプラス加入時）
ANA_COURSE_FEE = 5500             # ANAコース年会費
MR_PLUS_FEE = 3300                # MRプラス年会費
DEFAULT_MILE_VALUE = 2.0          # マイル単価(円)。ANA特典航空券想定の控えめ値


def analyze(name, annual_fee, points, mr_plus=True, mile_value=DEFAULT_MILE_VALUE,
            renewal_month=None, ana_course=True):
    rate = RATE_PLUS if mr_plus else RATE_NORMAL
    total_miles = int(points * rate)
    years = max(1, math.ceil(total_miles / ANA_ANNUAL_CAP_MILES))  # 移行に要する年数

    # 移行期間中に発生するコスト
    course_cost = years * (ANA_COURSE_FEE + (MR_PLUS_FEE if mr_plus else 0))
    # 年会費は「移行を待つ間」に更新ごと発生。最終移行直後に解約 → 概ね years回ぶん負担
    fee_drag = annual_fee * years

    mile_value_total = total_miles * mile_value
    # 移行する価値があるか（増分）: 得るマイル価値 vs 移行に要する追加コスト
    net_if_transfer = mile_value_total - (course_cost + fee_drag)
    # 即解約: ポイント放棄(価値0)・将来年会費0
    transfer_worth = mile_value_total > (course_cost + fee_drag)

    return {
        "name": name, "annual_fee": annual_fee, "points": points, "rate": rate,
        "total_miles": total_miles, "years": years,
        "course_cost": course_cost, "fee_drag": fee_drag,
        "mile_value_total": int(mile_value_total), "net_if_transfer": int(net_if_transfer),
        "transfer_worth": transfer_worth, "renewal_month": renewal_month,
    }


def report(cards):
    print("=" * 74)
    print(" AGP 解約最適化分析（Amex MR→ANA移行ルール照合）")
    print("=" * 74)
    grand_fee = 0
    for c in cards:
        r = analyze(**c)
        grand_fee += r["annual_fee"]
        print(f"\n■ {r['name']}  年会費 {r['annual_fee']:,}円" +
              (f" / 更新月 {r['renewal_month']}月" if r['renewal_month'] else " / 更新月: 要確認"))
        print(f"   保有 {r['points']:,}P × レート{r['rate']} = {r['total_miles']:,}マイル")
        print(f"   移行所要 {r['years']}年（年上限{ANA_ANNUAL_CAP_MILES:,}マイル）"
              f" ＝ その間 年会費を {r['years']}回負担 = {r['fee_drag']:,}円 垂れ流し")
        print(f"   移行コスト(ANAコース+MRプラス×{r['years']}年): {r['course_cost']:,}円")
        print(f"   獲得マイル価値: {r['mile_value_total']:,}円（@{DEFAULT_MILE_VALUE}円/マイル）")
        if r["transfer_worth"]:
            print(f"   ✅ 推奨: 移行する価値あり（純益 +{r['net_if_transfer']:,}円）")
            print(f"      → 毎年1/1に上限まで移行し、{r['years']}年目の移行完了直後"
                  f"（次回更新月の前）に解約 = 年会費を最小化")
        else:
            print(f"   🔴 推奨: 即解約が合理的（移行コスト+年会費 {r['course_cost']+r['fee_drag']:,}円"
                  f" > マイル価値 {r['mile_value_total']:,}円）。ポイント放棄でも年会費を止める方が得")
    print("\n" + "-" * 74)
    print(f"現状の年会費 垂れ流し: 合計 {grand_fee:,}円/年（{'＋'.join(c['name'] for c in cards)}）")
    print("=" * 74)
    print("※ 保有ポイント・更新月・MRプラス/ANAコース加入状況の実値で精度確定。")


if __name__ == "__main__":
    # ▼ 例示シナリオ（実値が入るまでの仮置き。年会費のみ実値）
    DEMO = [
        {"name": "AGP1(研太)", "annual_fee": 45273, "points": 120000, "mr_plus": True, "renewal_month": None},
        {"name": "AGP2(麻梨奈)", "annual_fee": 68025, "points": 200000, "mr_plus": True, "renewal_month": None},
    ]
    report(DEMO)
