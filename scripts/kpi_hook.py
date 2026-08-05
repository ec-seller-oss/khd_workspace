#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KPIフック：対話の"たび"に自動で回すための接続部品
================================================================
菊池さんが入力コストを払わずに済むように、Claude Code のフックから
KPIエンジンを自動起動する。3箇所に挿す。

  --mode prompt    UserPromptSubmit → 対話のたびにKPIの変動と未消化を注入
                   （＝Claudeが毎回それを見た状態で答える。都度報告の実装）
  --mode stop      Stop → その回の対話ログを走査して接触候補を溜める
                   （判断はしない。正規表現では言及と接触を分けられない）
  --mode session   SessionStart → 朝の3行＋前日からの繰り越しを表示

フックは失敗しても対話を止めてはいけないので、全経路で必ず exit 0 する。
"""
import argparse
import datetime as dt
import io
import json
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENGINE = Path(__file__).resolve().parent / "kpi_engine.py"
STATE = ROOT / ".company" / "secretary" / "kpi" / ".hook_state.json"
THROTTLE_MIN = 90   # 同じ内容を何度も注入しない間隔（分）


def run_engine(*a):
    try:
        r = subprocess.run([sys.executable, str(ENGINE), *a],
                           capture_output=True, text=True, timeout=20)
        return r.stdout.strip()
    except Exception:
        return ""


def load_state():
    try:
        return json.loads(STATE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_state(s):
    try:
        STATE.parent.mkdir(parents=True, exist_ok=True)
        STATE.write_text(json.dumps(s, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def read_hook_input():
    try:
        raw = sys.stdin.read()
        return json.loads(raw) if raw.strip() else {}
    except Exception:
        return {}


# ================================================================
def mode_prompt():
    """対話のたびに注入する。長すぎると邪魔なので3行＋未消化の見出しだけ。"""
    st = load_state()
    now = dt.datetime.now()
    last = st.get("last_prompt_inject")
    compact = False
    if last:
        try:
            if (now - dt.datetime.fromisoformat(last)).total_seconds() < THROTTLE_MIN * 60:
                compact = True
        except ValueError:
            pass

    sys.path.insert(0, str(ENGINE.parent))
    try:
        import kpi_engine as ke
    except Exception:
        print(json.dumps({}))
        return

    events = ke.read_jsonl(ke.EVENTS)
    lines = ke.three_lines(events)
    opens = ke.open_actions()
    pend = [p for p in ke.read_jsonl(ke.PENDING) if not p.get("resolved")]

    buf = io.StringIO()
    buf.write("<khd-kpi>\n")
    if compact:
        buf.write(lines[0] + "\n" + lines[1] + "\n")
    else:
        buf.write("\n".join(lines) + "\n")
    if opens:
        buf.write(f"未消化の次アクション {len(opens)}件: "
                  + " ／ ".join(f"[{a['id']}]{a.get('who','')}:{a.get('what','')[:24]}"
                                for a in opens[:5]) + "\n")
    if pend:
        buf.write(f"要判断の接触候補 {len(pend)}件（接触か言及かを判断して log / resolve）\n")
    buf.write("ルール: この対話で①〜⑨に当たる事実が出たら kpi_engine.py log で即転記し、"
              "新しい約束・宿題は next で追記する。返答の最後にKPIの変動を1行で返す。\n")
    buf.write("</khd-kpi>")

    if not compact:
        st["last_prompt_inject"] = now.isoformat()
        save_state(st)

    print(json.dumps({
        "hookSpecificOutput": {
            "hookEventName": "UserPromptSubmit",
            "additionalContext": buf.getvalue(),
        }
    }, ensure_ascii=False))


def mode_stop():
    """その回の対話ログを走査。ユーザーには何も出さない（沈黙のルールを守る）。"""
    data = read_hook_input()
    tp = data.get("transcript_path")
    args = ["scan", "--since", dt.date.today().isoformat()]
    if tp and Path(tp).exists():
        args += ["--transcript", tp]
    run_engine(*args)
    # 記録漏れの自己申告用に、その回の最後のログ時刻を残す
    st = load_state()
    st["last_stop"] = dt.datetime.now().isoformat()
    save_state(st)


def mode_session():
    out = run_engine("brief")
    if out:
        print(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", required=True, choices=["prompt", "stop", "session"])
    a = ap.parse_args()
    try:
        {"prompt": mode_prompt, "stop": mode_stop, "session": mode_session}[a.mode]()
    except Exception as e:
        # フックは絶対に対話を止めない
        if a.mode == "session":
            print(f"（KPIブリーフ生成に失敗: {e}）")
    sys.exit(0)


if __name__ == "__main__":
    main()
