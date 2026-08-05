#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KPIエンジン：対話 → DB転記 → 前日翌日の連動 → KPI変動の即時報告
================================================================
菊池さんが「やるべきこと」だけに集中するための計測基盤。

設計の骨：
  1) 追記のみ（.company/CLAUDE.md の原則）。イベントは append-only、
     状態変更も「同じidの新レコードを追記」で表現し、最後の1件が勝つ。
  2) 9指標のうち手入力が必要なのは ⑧オーロラ だけ。他は物証層／対話層から入る。
  3) 抜け漏れは「昨日開いて今日閉じていない次アクション」として自動で浮く。

サブコマンド：
  brief        セッション開始時に出す3行KPI＋未消化の次アクション＋要判断の候補
  log          イベントを1件追記（Claudeが対話中に呼ぶ）
  next         次アクションを1件追記
  done         次アクションを消化済みにする
  report       今日／今週／先週／今月の全指標レポート
  carry        未消化の次アクションだけを出す（経過日数つき）
  scan         対話ログ(jsonl)から接触候補を抽出して候補ファイルへ（判断はしない）
  import-csv   作業DBのCSVエクスポートを取り込む（学習h・営業直結・提案数等）

使い方の例：
  python3 scripts/kpi_engine.py log daseki --who 天野先生 --what "診療圏調査を1件無料で送った" --kind 提案
  python3 scripts/kpi_engine.py log study --h 5.0
  python3 scripts/kpi_engine.py next --who 銀行 --what "住宅ローン借換の3数字を持って30分" --due 2026-08-08
  python3 scripts/kpi_engine.py brief
"""
import argparse
import csv
import datetime as dt
import json
import os
import re
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
KPI_DIR = ROOT / ".company" / "secretary" / "kpi"
EVENTS = KPI_DIR / "events.jsonl"
NEXTS = KPI_DIR / "next_actions.jsonl"
PENDING = KPI_DIR / "pending_candidates.jsonl"

# ---- 目標値（genddouryoku_board_260804.html の9指標と一致させる） ----
TARGET = {
    "study_week_h": 35.0,      # ① 週35時間
    "study_day_h": 5.0,        # ① 1日5時間最低
    "oyako_week_days": 5,      # ③ 週5日以上
    "daseki_week": 2,          # ④ 週2打席（10/19から週5）
    "daseki_week_after": 5,
    "contact_week": 1,         # ⑤ 週1人（10/19から週3人）
    "contact_week_after": 3,
    "mutual_total": 10,        # ⑨ 2027/1/31までに10人
    "bs_net_assets_man": 10000,  # ⑦ ざっくり1億円
}
SWITCH_DATE = dt.date(2026, 10, 19)   # 宅建試験の翌日＝発信を始める日
MUTUAL_DEADLINE = dt.date(2027, 1, 31)

EVENT_TYPES = {
    "study": "① 試験勉強",
    "pyramid": "② 1日1回ピラミッド",
    "oyako": "③ 18-20時の親子タイム",
    "daseki": "④ 提案・GIVEの打席",
    "contact": "⑤ 30分以上1対1で話した人",
    "mutual": "⑨ 相互認知",
    "bs": "⑦ BSの純資産",
    "aurora": "⑧ オーロラの3指標",
}


# ================================================================
# 基盤：読み書き
# ================================================================
def today():
    return dt.date.fromisoformat(os.environ.get("KHD_TODAY") or dt.date.today().isoformat())


def now_ts():
    return dt.datetime.now().strftime("%Y-%m-%dT%H:%M")


def read_jsonl(path):
    if not path.exists():
        return []
    out = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        try:
            out.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return out


def append_jsonl(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")


def latest_by_id(records):
    """追記のみで状態を表すための畳み込み。同じidは後勝ちでマージ。"""
    merged = {}
    for r in records:
        rid = r.get("id")
        if not rid:
            continue
        if rid in merged:
            merged[rid].update(r)
        else:
            merged[rid] = dict(r)
    return merged


# ================================================================
# 期間の計算（週は日曜始まり＝これまでの報告と揃える）
# ================================================================
def week_start(d):
    return d - dt.timedelta(days=(d.weekday() + 1) % 7)


def period(name, base=None):
    d = base or today()
    ws = week_start(d)
    if name == "today":
        return d, d
    if name == "week":
        return ws, d
    if name == "lastweek":
        return ws - dt.timedelta(days=7), ws - dt.timedelta(days=1)
    if name == "month":
        return d.replace(day=1), d
    if name == "lastmonth":
        first = d.replace(day=1)
        prev_end = first - dt.timedelta(days=1)
        return prev_end.replace(day=1), prev_end
    raise ValueError(name)


def in_range(datestr, a, b):
    try:
        x = dt.date.fromisoformat(datestr)
    except (ValueError, TypeError):
        return False
    return a <= x <= b


# ================================================================
# 集計
# ================================================================
def aggregate(events, a, b):
    """期間 [a,b] の9指標を集計。件数だけでなく中身（誰と何を）も返す。

    type="period" は作業DB（スプレッドシート）側の確定集計を持ち込むための箱。
    [from,to] が [a,b] に完全に含まれるときだけ足す。部分重なりは足さない
    （日別の内訳を持っていない集計を切り分けると必ず嘘になるため）。
    """
    ev = [e for e in events if in_range(e.get("date", ""), a, b)]
    days = (b - a).days + 1
    agg = {
        "days": days,
        "study_h": 0.0,
        "study_by_day": defaultdict(float),
        "pyramid_days": set(),
        "oyako_days": set(),
        "daseki": [],
        "contact": [],
        "mutual": [],
        "bs": [],
        "aurora": [],
        "eigyo_chokketsu": 0,
        "teian": 0, "give": 0, "sodan": 0,
        "jitsu_min": 0.0,
        "oyako_extra": 0,     # 作業DB集計から持ち込んだ親子タイム日数
        "rec_days": 0,        # 記録が存在した日数（1日あたりを出すのに使う）
        "from_period": [],    # 持ち込んだ確定集計の出典
    }
    for e in events:
        if e.get("type") != "period":
            continue
        try:
            pa = dt.date.fromisoformat(e["from"])
            pb = dt.date.fromisoformat(e["to"])
        except (KeyError, ValueError):
            continue
        if not (a <= pa and pb <= b):
            continue
        agg["study_h"] += float(e.get("study_h") or 0)
        agg["jitsu_min"] += float(e.get("jitsu_min") or 0)
        agg["eigyo_chokketsu"] += int(e.get("eigyo") or 0)
        agg["teian"] += int(e.get("teian") or 0)
        agg["give"] += int(e.get("give") or 0)
        agg["sodan"] += int(e.get("sodan") or 0)
        agg["oyako_extra"] += int(e.get("oyako_days") or 0)
        agg["rec_days"] += int(e.get("rec_days") or ((pb - pa).days + 1))
        agg["from_period"].append(f"{e['from']}〜{e['to']}({e.get('src','')})")
    for e in ev:
        t = e.get("type")
        d = e.get("date")
        if t == "study":
            agg["study_h"] += float(e.get("h") or 0)
            agg["study_by_day"][d] += float(e.get("h") or 0)
        elif t == "pyramid":
            agg["pyramid_days"].add(d)
        elif t == "oyako":
            agg["oyako_days"].add(d)
        elif t == "daseki":
            agg["daseki"].append(e)
        elif t == "contact":
            if float(e.get("min") or 0) >= 30:
                agg["contact"].append(e)
        elif t == "mutual":
            agg["mutual"].append(e)
        elif t == "bs":
            agg["bs"].append(e)
        elif t == "aurora":
            agg["aurora"].append(e)
        elif t == "daily":   # 作業DBからの日次ロールアップ
            agg["study_h"] += float(e.get("study_h") or 0)
            agg["study_by_day"][d] += float(e.get("study_h") or 0)
            agg["jitsu_min"] += float(e.get("jitsu_min") or 0)
            agg["eigyo_chokketsu"] += int(e.get("eigyo") or 0)
            agg["teian"] += int(e.get("teian") or 0)
            agg["give"] += int(e.get("give") or 0)
            agg["sodan"] += int(e.get("sodan") or 0)
            if e.get("oyako_min"):
                agg["oyako_days"].add(d)
    agg["contact_people"] = sorted({c.get("who", "") for c in agg["contact"] if c.get("who")})
    agg["oyako_count"] = len(agg["oyako_days"]) + agg["oyako_extra"]
    agg["rec_days"] += len({e["date"] for e in ev if e.get("type") in ("study", "daily")})
    agg["study_per_day"] = agg["study_h"] / agg["rec_days"] if agg["rec_days"] else 0.0
    return agg


def daseki_target(d=None):
    d = d or today()
    return TARGET["daseki_week"] if d < SWITCH_DATE else TARGET["daseki_week_after"]


def contact_target(d=None):
    d = d or today()
    return TARGET["contact_week"] if d < SWITCH_DATE else TARGET["contact_week_after"]


def arrow(cur, prev):
    if prev == 0 and cur == 0:
        return "→ ±0"
    if prev == 0:
        return f"↑ +{cur:g}（先週0）"
    diff = cur - prev
    ratio = cur / prev
    sign = "↑" if diff > 0 else ("↓" if diff < 0 else "→")
    return f"{sign} {diff:+g}（{ratio:.1f}倍）"


# ================================================================
# レポート
# ================================================================
def three_lines(events, d=None):
    """毎回の対話で出す3行。モチベは『変動』にしか宿らないので必ず対比で書く。"""
    d = d or today()
    wa, wb = period("week", d)
    la, lb = period("lastweek", d)
    tw = aggregate(events, wa, wb)
    lw = aggregate(events, la, lb)

    # ① 週35時間への残り
    need = TARGET["study_week_h"] - tw["study_h"]
    rest_days = (week_start(d) + dt.timedelta(days=6) - d).days + 1
    pace = (need / rest_days) if rest_days > 0 and need > 0 else 0.0
    l1 = (f"① 学習 {tw['study_h']:.1f}h / 35h　{arrow(tw['study_h'], lw['study_h'])}"
          f"　→ 残り{rest_days}日で1日 {pace:.2f}h 必要"
          if need > 0 else
          f"① 学習 {tw['study_h']:.1f}h / 35h　✅週次クリア（⑥達成）")

    # ④⑤ 営業の打席と人
    dt_ = daseki_target(d)
    ct_ = contact_target(d)
    l2 = (f"④ 打席 {len(tw['daseki'])}/{dt_}　{arrow(len(tw['daseki']), len(lw['daseki']))}"
          f"　／ ⑤ 30分1対1 {len(tw['contact_people'])}人/{ct_}人"
          f"　{arrow(len(tw['contact_people']), len(lw['contact_people']))}")

    # ③ 親子タイム
    l3 = (f"③ 親子タイム {tw['oyako_count']}/{min(tw['days'], 7)}日（目標 週5日）"
          f"　{arrow(tw['oyako_count'], lw['oyako_count'])}"
          f"　／ ② ピラミッド 今日 {'✓' if d.isoformat() in tw['pyramid_days'] else '✗'}")
    return [l1, l2, l3]


def cmd_brief(args):
    """セッション開始フックから呼ばれる。ここだけで『今日やること』が決まる形にする。"""
    d = today()
    events = read_jsonl(EVENTS)
    print(f"━━ KPI {d.isoformat()}（{'日月火水木金土'[(d.weekday()+1)%7]}） ━━")
    for line in three_lines(events, d):
        print("  " + line)

    # 未消化の次アクション＝抜け漏れの正体
    opens = open_actions(d)
    if opens:
        print(f"\n━━ 未消化の次アクション {len(opens)}件 ━━")
        for a in opens[:8]:
            age = (d - dt.date.fromisoformat(a["opened"])).days
            od = a.get("due") or ""
            flag = ""
            if od:
                dd = (dt.date.fromisoformat(od) - d).days
                flag = "🔴期限切れ" if dd < 0 else ("🟠今日" if dd == 0 else f"あと{dd}日")
            print(f"  [{a['id']}] {a.get('who','')}／{a.get('what','')}"
                  f"　（{age}日前に発生 {flag}）")
    else:
        print("\n未消化の次アクション：なし")

    # 対話ログから拾った要判断の候補（接触か言及かはClaudeが判断する）
    pend = [p for p in read_jsonl(PENDING) if not p.get("resolved")]
    if pend:
        print(f"\n━━ 対話ログからの接触候補 {len(pend)}件（言及か接触かは要判断） ━━")
        for p in pend[:8]:
            print(f"  ? {p.get('date','')} {p.get('who','')}：{(p.get('snippet') or '')[:60]}")

    print("\n※ 手入力が必要なのは ⑧オーロラの3指標 だけ。他は物証層／対話層から入る。")
    return 0


def open_actions(d=None):
    d = d or today()
    merged = latest_by_id(read_jsonl(NEXTS))
    opens = [a for a in merged.values() if a.get("status", "open") == "open"]

    def key(a):
        due = a.get("due") or "9999-12-31"
        return (due, a.get("opened", ""))
    return sorted(opens, key=key)


def cmd_report(args):
    events = read_jsonl(EVENTS)
    d = today()
    spans = [("今日", *period("today", d)), ("今週", *period("week", d)),
             ("先週", *period("lastweek", d)), ("今月", *period("month", d)),
             ("先月", *period("lastmonth", d))]
    print(f"══ KPIレポート {d.isoformat()} ══\n")
    rows = []
    for name, a, b in spans:
        g = aggregate(events, a, b)
        rows.append((name, f"{a.month}/{a.day}–{b.month}/{b.day}", g, g["rec_days"]))
    hdr = f"{'期間':<6}{'日付':<12}{'学習h':>7}{'/日':>7}{'打席':>5}{'1対1':>5}{'親子':>5}{'営業直結':>7}"
    print(hdr)
    print("-" * len(hdr))
    for name, span, g, rec in rows:
        print(f"{name:<6}{span:<12}{g['study_h']:>7.1f}{g['study_per_day']:>7.2f}"
              f"{len(g['daseki']) + g['teian'] + g['give']:>5}{len(g['contact_people']):>5}"
              f"{g['oyako_count']:>5}{g['eigyo_chokketsu']:>7}")

    for name, span, g, rec in rows:
        if rec == 0 and g["study_h"] == 0 and name not in ("今日", "今週"):
            print(f"※ {name}（{span}）は記録が0日。作業DBの運用開始が2026-07-26なので、"
                  f"こことの対比は成立しない。0と読まず『比較不能』と読む。")

    # ⑨ 相互認知は累計
    all_mutual = sorted({e.get("who") for e in events if e.get("type") == "mutual" and e.get("who")})
    left = (MUTUAL_DEADLINE - d).days
    print(f"\n⑨ 相互認知 {len(all_mutual)}/10人（2027/1/31まで あと{left}日）"
          f"{'：' + '・'.join(all_mutual) if all_mutual else '：まだ0人'}")

    # ⑦ BS
    bs = [e for e in events if e.get("type") == "bs"]
    if bs:
        bs.sort(key=lambda e: e.get("date", ""))
        cur = float(bs[-1].get("net_assets_man") or 0)
        prev = float(bs[-2].get("net_assets_man") or 0) if len(bs) > 1 else 0
        print(f"⑦ BS純資産 {cur:,.0f}万円 {arrow(cur, prev)}　目標 10,000万円")
    else:
        print("⑦ BS純資産：未記録（財務司令塔から月1回 log bs）")

    au = [e for e in events if e.get("type") == "aurora"]
    print(f"⑧ オーロラ：{'記録あり ' + au[-1].get('date','') if au else '11/4開業後に手入力（唯一の手入力）'}")

    print("\n── 今週の打席の中身（④は件数ではなく『誰に何を渡したか』が本体）──")
    g = aggregate(events, *period("week", d))
    if g["daseki"]:
        for e in g["daseki"]:
            print(f"  {e['date']} {e.get('kind','')}／{e.get('who','')}：{e.get('what','')}")
    else:
        print("  なし。打席0＝この週は営業していない、という事実がそのまま残る。")
    return 0


def cmd_carry(args):
    d = today()
    opens = open_actions(d)
    if not opens:
        print("未消化の次アクション：なし")
        return 0
    print(f"未消化 {len(opens)}件")
    for a in opens:
        age = (d - dt.date.fromisoformat(a["opened"])).days
        print(f"[{a['id']}] {a.get('who','')}／{a.get('what','')}"
              f"　開始{a['opened']}（{age}日）期限{a.get('due','-')} KPI{a.get('kpi','-')}")
    return 0


# ================================================================
# 書き込み系
# ================================================================
def cmd_log(args):
    d = args.date or today().isoformat()
    e = {"ts": now_ts(), "date": d, "type": args.type, "src": args.src}
    for k in ("who", "what", "kind", "topic", "note"):
        v = getattr(args, k, None)
        if v:
            e[k] = v
    for k in ("h", "min", "net_assets_man", "shinkan", "revisit", "sales_man"):
        v = getattr(args, k, None)
        if v is not None:
            e[k] = v
    append_jsonl(EVENTS, e)
    print(f"✅ 記録: {EVENT_TYPES.get(args.type, args.type)}"
          f"{' / ' + e['who'] if e.get('who') else ''}"
          f"{' / ' + e['what'] if e.get('what') else ''}")
    # 記録したら必ずKPIの変動を返す（都度のモチベーション）
    for line in three_lines(read_jsonl(EVENTS)):
        print("  " + line)
    return 0


def new_action_id(d):
    same = [a for a in read_jsonl(NEXTS) if a.get("id", "").startswith(f"NA-{d:%Y%m%d}")]
    return f"NA-{d:%Y%m%d}-{len(same)+1:02d}"


def cmd_next(args):
    d = today()
    obj = {"id": args.id or new_action_id(d), "opened": d.isoformat(),
           "who": args.who or "", "what": args.what, "status": "open"}
    if args.due:
        obj["due"] = args.due
    if args.kpi:
        obj["kpi"] = args.kpi
    if args.src:
        obj["src"] = args.src
    append_jsonl(NEXTS, obj)
    print(f"➕ 次アクション [{obj['id']}] {obj['who']}／{obj['what']}"
          f"{' 期限' + obj['due'] if obj.get('due') else ''}")
    return 0


def cmd_done(args):
    merged = latest_by_id(read_jsonl(NEXTS))
    if args.id not in merged:
        print(f"⚠️ 見つからない: {args.id}", file=sys.stderr)
        return 1
    append_jsonl(NEXTS, {"id": args.id, "status": "done",
                         "closed": today().isoformat(),
                         "result": args.result or ""})
    a = merged[args.id]
    print(f"✔ 消化: [{args.id}] {a.get('who','')}／{a.get('what','')}")
    return 0


def cmd_drop(args):
    append_jsonl(NEXTS, {"id": args.id, "status": "dropped",
                         "closed": today().isoformat(),
                         "reason": args.reason or ""})
    print(f"🗑 取り下げ: [{args.id}] {args.reason or ''}")
    return 0


# ================================================================
# 対話ログのスキャン（判断はしない。候補を出すだけ）
# ================================================================
# 「渡した／送った／会った」等の動作語。言及と接触を分けるのは正規表現では不可能で、
# 2026-08-05のテストでは13ヒット全部が引用・分析文の誤検知だった。
# よってここは「候補の絞り込み」までで止め、接触かどうかはClaudeが判断する。
VERB = re.compile(r"(渡した|送った|話した|会った|訊いた|連絡した|投稿した|公開した|"
                  r"書いて出した|提案した|決済|契約|面談|打ち合わせ|セミナーした|電話した)")
NAME = re.compile(r"([一-龥ぁ-んァ-ヶA-Za-z]{2,6}(?:先生|さん|社長|様|支店長|部長))")
# 引用・分析・書き起こしの巻き込みを弾く
NOISE = re.compile(r"(https?://|^\s*[>｜|]|【|■|◆|字幕|書き起こし|とは\?|とは？"
                   r"|①|②|③|④|⑤|⑥|⑦|⑧|⑨|KPI|指標)")
SELF = re.compile(r"(菊池|自分|相手|お客様|先方)")   # 本人・総称は接触相手ではない


def scan_transcript(path, since=None):
    """Claude Codeの対話ログ(jsonl)からユーザー発言だけ抜き、接触候補を返す。"""
    cands = []
    p = Path(path)
    if not p.exists():
        return cands
    with p.open(encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                o = json.loads(line)
            except json.JSONDecodeError:
                continue
            m = o.get("message") or {}
            if m.get("role") != "user":
                continue
            c = m.get("content")
            if isinstance(c, str):
                txt = c
            elif isinstance(c, list):
                txt = " ".join(b.get("text", "") for b in c
                               if isinstance(b, dict) and b.get("type") == "text")
            else:
                continue
            if not txt or "<system-reminder>" in txt[:80] or "tool_result" in txt[:60]:
                continue
            ts = (o.get("timestamp") or "")[:10]
            if since and ts and ts < since:
                continue
            for sent in re.split(r"[。\n]", txt):
                if not VERB.search(sent) or NOISE.search(sent):
                    continue
                for who in set(NAME.findall(sent)):
                    if SELF.search(who):
                        continue
                    cands.append({"ts": now_ts(), "date": ts or today().isoformat(),
                                  "who": who, "snippet": sent.strip()[:160],
                                  "src": "dialogue", "resolved": False})
    return cands


def cmd_scan(args):
    path = args.transcript
    if not path:
        # 既定＝このプロジェクトの最新セッション
        base = Path.home() / ".claude" / "projects"
        pats = list(base.glob("*khd*/*.jsonl")) + list(base.glob("*honbu*/*.jsonl"))
        if not pats:
            print("対話ログが見つからない", file=sys.stderr)
            return 1
        path = max(pats, key=lambda p: p.stat().st_mtime)
    cands = scan_transcript(path, since=args.since)
    known = {(p.get("date"), p.get("who"), p.get("snippet")) for p in read_jsonl(PENDING)}
    added = 0
    for c in cands:
        if (c["date"], c["who"], c["snippet"]) in known:
            continue
        append_jsonl(PENDING, c)
        added += 1
    print(f"scan: {path} → 候補{len(cands)}件（新規{added}件）")
    if args.show:
        for c in cands:
            print(f"  ? {c['date']} {c['who']}: {c['snippet'][:70]}")
    print("※ 接触か言及かは正規表現では判定不能。Claudeが1件ずつ判断して log する。")
    return 0


def cmd_resolve(args):
    """候補を裁く。--as contact|daseki|none"""
    pend = read_jsonl(PENDING)
    hit = [p for p in pend if p.get("who") == args.who and not p.get("resolved")]
    if not hit:
        print(f"該当候補なし: {args.who}", file=sys.stderr)
        return 1
    for p in hit:
        p["resolved"] = True
        p["verdict"] = args.as_
        append_jsonl(PENDING, p)
    print(f"裁定: {args.who} × {len(hit)}件 → {args.as_}")
    return 0


# ================================================================
# 作業DBのCSV取り込み（学習h・営業直結・提案数はここが唯一の真実）
# ================================================================
def cmd_import_csv(args):
    """作業DB（スプレッドシート）をCSVエクスポートしたものを日次ロールアップに変換。"""
    added = 0
    with open(args.path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    by_day = defaultdict(lambda: {"study_h": 0.0, "jitsu_min": 0.0, "eigyo": 0,
                                  "teian": 0, "give": 0, "sodan": 0, "oyako_min": 0.0})

    def num(v):
        try:
            return float(str(v).replace(",", "").strip() or 0)
        except ValueError:
            return 0.0

    for r in rows:
        d = (r.get("日付") or "").strip()[:10]
        if not re.match(r"\d{4}-\d{2}-\d{2}", d):
            continue
        mins = num(r.get("実所要(分)") or r.get("実所要分"))
        honbu = str(r.get("本部") or "")
        content = str(r.get("予定・タスク内容") or r.get("案件・相手") or "")
        b = by_day[d]
        b["jitsu_min"] += mins
        if "04" in honbu or "調査士" in content or "宅建" in content or "勉強" in content:
            b["study_h"] += mins / 60.0
        if "親子" in content:
            b["oyako_min"] += mins
        if str(r.get("営業直結") or "").strip() in ("○", "◯", "o", "O", "1"):
            b["eigyo"] += 1
        b["teian"] += int(num(r.get("提案数")))
        b["give"] += int(num(r.get("GIVE数")))
        b["sodan"] += int(num(r.get("相談数")))

    existing = {(e.get("date"), e.get("type")) for e in read_jsonl(EVENTS)}
    for d, b in sorted(by_day.items()):
        if (d, "daily") in existing and not args.force:
            continue
        append_jsonl(EVENTS, {"ts": now_ts(), "date": d, "type": "daily",
                              "src": "作業DB", **{k: round(v, 2) for k, v in b.items()}})
        added += 1
    print(f"import-csv: {len(by_day)}日分 → 追記{added}件")
    return 0


# ================================================================
def main():
    ap = argparse.ArgumentParser(description="KHD KPIエンジン")
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("brief").set_defaults(func=cmd_brief)
    sub.add_parser("report").set_defaults(func=cmd_report)
    sub.add_parser("carry").set_defaults(func=cmd_carry)

    p = sub.add_parser("log")
    p.add_argument("type", choices=list(EVENT_TYPES.keys()))
    p.add_argument("--date")
    p.add_argument("--who"); p.add_argument("--what"); p.add_argument("--kind")
    p.add_argument("--topic"); p.add_argument("--note")
    p.add_argument("--h", type=float); p.add_argument("--min", type=float)
    p.add_argument("--net-assets-man", type=float, dest="net_assets_man")
    p.add_argument("--shinkan", type=int); p.add_argument("--revisit", type=int)
    p.add_argument("--sales-man", type=float, dest="sales_man")
    p.add_argument("--src", default="dialogue")
    p.set_defaults(func=cmd_log)

    p = sub.add_parser("next")
    p.add_argument("--what", required=True)
    p.add_argument("--who"); p.add_argument("--due"); p.add_argument("--kpi")
    p.add_argument("--id"); p.add_argument("--src", default="dialogue")
    p.set_defaults(func=cmd_next)

    p = sub.add_parser("done")
    p.add_argument("id"); p.add_argument("--result")
    p.set_defaults(func=cmd_done)

    p = sub.add_parser("drop")
    p.add_argument("id"); p.add_argument("--reason")
    p.set_defaults(func=cmd_drop)

    p = sub.add_parser("scan")
    p.add_argument("--transcript"); p.add_argument("--since")
    p.add_argument("--show", action="store_true")
    p.set_defaults(func=cmd_scan)

    p = sub.add_parser("resolve")
    p.add_argument("who")
    p.add_argument("--as", dest="as_", required=True,
                   choices=["contact", "daseki", "none"])
    p.set_defaults(func=cmd_resolve)

    p = sub.add_parser("import-csv")
    p.add_argument("path"); p.add_argument("--force", action="store_true")
    p.set_defaults(func=cmd_import_csv)

    args = ap.parse_args()
    sys.exit(args.func(args))


if __name__ == "__main__":
    main()
