# -*- coding: utf-8 -*-
"""v9デッキ用チャート3点（KHD配色・Hiragino Sans）"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

plt.rcParams["font.family"] = ["Hiragino Sans"]
plt.rcParams["axes.unicode_minus"] = False
CREAM="#F9F6EF"; RED="#AA2E26"; INK="#1A1A1A"; RULE="#DAD6CF"
GREY="#6B6B6B"; LIGHTRED="#C96A62"; STAYGREEN="#7BA77C"

OUT = "/Users/kikuchikenta/01_honbu_docs_automation/"

# ── ① Slide3用：削減効果の幅（レンジバー・横） ──
fig, ax = plt.subplots(figsize=(11.6, 1.85), dpi=200)
fig.patch.set_facecolor(CREAM); ax.set_facecolor(CREAM)
items = [("問診・予約", 40, 70), ("電話対応", 40, 80), ("レセプト点検", 80, 95)]
ypos = np.arange(len(items))[::-1]
for y,(lab,lo,hi) in zip(ypos, items):
    ax.barh(y, hi-lo, left=lo, height=0.52, color=LIGHTRED, zorder=3)
    ax.barh(y, 6, left=hi-6, height=0.52, color=RED, zorder=4)
    ax.text(lo-1.5, y, f"{lo}%", ha="right", va="center", fontsize=10, color=GREY)
    ax.text(hi+1.5, y, f"最大 {hi}%削減", ha="left", va="center", fontsize=11, color=RED, fontweight="bold")
ax.set_yticks(ypos); ax.set_yticklabels([i[0] for i in items], fontsize=11, color=INK)
ax.set_xlim(0, 115); ax.set_xticks([0,25,50,75,100]); ax.set_xticklabels(["0%","25%","50%","75%","100%"], fontsize=8.5, color=GREY)
for sp in ["top","right","left"]: ax.spines[sp].set_visible(False)
ax.spines["bottom"].set_color(RULE)
ax.grid(axis="x", color=RULE, lw=0.6, zorder=0)
ax.text(108, ypos[-1]-0.62, "※書類作成は 月30時間以上の削減（時間ベース）", fontsize=8.5, color=GREY, ha="right")
plt.tight_layout()
plt.savefig(OUT+"_chart_range.png", facecolor=CREAM, bbox_inches="tight")
print("saved _chart_range.png")

# ── ② Slide6用：削減時間の内訳（縦棒） ──
fig, ax = plt.subplots(figsize=(5.4, 3.6), dpi=200)
fig.patch.set_facecolor(CREAM); ax.set_facecolor(CREAM)
labs = ["問診", "レセプト", "電話", "書類"]
vals = [42, 45, 40, 12]
cols = [LIGHTRED, RED, LIGHTRED, LIGHTRED]
bars = ax.bar(labs, vals, width=0.58, color=cols, zorder=3)
for b,v in zip(bars, vals):
    ax.text(b.get_x()+b.get_width()/2, v+1.2, f"−{v}h", ha="center", fontsize=12, color=RED, fontweight="bold")
ax.set_ylim(0, 56)
ax.set_ylabel("削減時間（h/月）", fontsize=9.5, color=INK)
ax.tick_params(axis="x", labelsize=11); ax.tick_params(axis="y", labelsize=8.5)
for sp in ["top","right"]: ax.spines[sp].set_visible(False)
for sp in ["left","bottom"]: ax.spines[sp].set_color(RULE)
ax.grid(axis="y", color=RULE, lw=0.6, zorder=0)
ax.set_title("月139時間の内訳（標準モデル）", fontsize=12, color=INK, fontweight="bold", pad=10)
plt.tight_layout()
plt.savefig(OUT+"_chart_breakdown.png", facecolor=CREAM, bbox_inches="tight")
print("saved _chart_breakdown.png")

# ── ③ 新スライド用：使うほど院に残る（既存 _savings_chart.png を再生成・同一ロジック） ──
SAVE=240000; RATE=0.40; MAINT=50000
months=np.arange(1,13)
pay=np.where(months<=6, SAVE*RATE, MAINT)
stay=SAVE-pay; cum=np.cumsum(stay); man=10000.0
fig, ax = plt.subplots(figsize=(8.6,4.7), dpi=200)
fig.patch.set_facecolor(CREAM); ax.set_facecolor(CREAM)
ax.bar(months, stay/man, width=0.62, color=STAYGREEN, label="院に残る効果", zorder=3)
ax.bar(months, pay/man, bottom=stay/man, width=0.62, color=RED, label="お支払い", zorder=3)
ax.axvline(6.5, color=GREY, lw=1.0, ls=(0,(4,3)), zorder=2)
ax.text(6.5, 26.6, "成果報酬6ヶ月  →  7ヶ月目〜 保守 月5万", color=GREY, fontsize=9, ha="center", va="bottom")
for m in [1,6,7,12]:
    ax.text(m,(stay[m-1]/man)/2, f"{stay[m-1]/man:.1f}万", color="white", fontsize=8.5, ha="center", va="center", fontweight="bold", zorder=4)
ax.set_ylim(0,28); ax.set_xticks(months)
ax.set_xlabel("導入からの月数", fontsize=9, color=INK); ax.set_ylabel("月額（万円）", fontsize=9, color=INK)
ax.tick_params(labelsize=8.5)
for sp in ["top","right"]: ax.spines[sp].set_visible(False)
for sp in ["left","bottom"]: ax.spines[sp].set_color(RULE)
ax.grid(axis="y", color=RULE, lw=0.6, zorder=0)
ax2=ax.twinx()
ax2.plot(months, cum/man, color=INK, lw=2.2, marker="o", ms=4.5, label="院に残る（累計）", zorder=5)
ax2.scatter([12],[cum[-1]/man], s=60, color=RED, zorder=6)
ax2.annotate(f"12ヶ月で\n院に約{cum[-1]/man:.0f}万円", xy=(12,cum[-1]/man), xytext=(9.0,120),
             fontsize=10.5, color=RED, fontweight="bold", ha="center",
             arrowprops=dict(arrowstyle="->", color=RED, lw=1.4))
ax2.set_ylim(0,230); ax2.set_ylabel("院に残る効果（累計・万円）", fontsize=9, color=INK)
ax2.tick_params(labelsize=8.5); ax2.spines["top"].set_visible(False); ax2.spines["right"].set_color(RULE)
h1,l1=ax.get_legend_handles_labels(); h2,l2=ax2.get_legend_handles_labels()
ax.legend(h1+h2, l1+l2, loc="upper left", fontsize=8.5, frameon=False)
plt.tight_layout()
plt.savefig(OUT+"_savings_chart.png", facecolor=CREAM, bbox_inches="tight")
print("saved _savings_chart.png")
