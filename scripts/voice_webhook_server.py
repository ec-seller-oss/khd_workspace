#!/usr/bin/env python3
"""
KHD 音声報連相 Webhook サーバー（AI API不要・ルールベース）
POST /report {"text": "バイセル記帳終わった"} → パターン解析 → Apps Script更新 → LINE確認
POST /property {"text": "物件情報テキスト"} → claude解析 → 01_物件検討DB自動追記
"""
import os, re, json, requests, subprocess, threading
from flask import Flask, request, jsonify
from datetime import datetime

app = Flask(__name__)

APPS_SCRIPT_URL = os.environ.get(
    "KHD_APPS_SCRIPT_URL",
    "https://script.google.com/macros/s/AKfycbw1uDeXqnzVvAMJ5Rt2BnCUNKdNqjpgS5rQPTr8Dpujzwqo1JgPlfZNGSQNmzB0S7KrlA/exec"
)
SECRET_TOKEN = "khd_report_2026"
LINE_TOKEN = os.environ.get("LINE_CHANNEL_TOKEN", "")


# ── Apps Script 呼び出し ──────────────────────────────────
def call_apps_script(action_dict):
    payload = {"token": SECRET_TOKEN, **action_dict}
    resp = requests.post(APPS_SCRIPT_URL, json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()


# ── LINE確認通知 ──────────────────────────────────────────
def send_line(text):
    if not LINE_TOKEN:
        return
    requests.post(
        "https://api.line.me/v2/bot/message/broadcast",
        headers={"Authorization": f"Bearer {LINE_TOKEN}", "Content-Type": "application/json"},
        json={"messages": [{"type": "text", "text": text}]},
        timeout=10,
    )


# ── パターン解析 ──────────────────────────────────────────
def parse_text(text):
    """報告テキストからアクションリストを返す"""
    actions = []
    t = text.strip()

    # ① WBS: 完了
    # 例: 「バイセル記帳終わった」「〇〇完了した」「〇〇できた」
    m = re.search(r"(.+?)(?:が|を|は)?(?:終わっ|完了|できた|片付|済み|やり切|リリース|上げ)", t)
    if m:
        task = m.group(1).strip()
        # 学習・KPI系は除外
        if not re.search(r"朝活|勉強|学習|民法|不登|KPI|他力|朝活|EQ|家族", task):
            actions.append({
                "action": "update_wbs",
                "task": task,
                "status": "完了",
                "progress": "100%",
            })

    # ② WBS: 対応中
    # 例: 「〇〇着手した」「〇〇対応中」「〇〇始めた」
    m2 = re.search(r"(.+?)(?:に|を)?(?:着手|対応中|始め|取り掛|やり始|進め)", t)
    if m2 and not actions:
        task = m2.group(1).strip()
        if not re.search(r"朝活|勉強|学習|民法|不登|KPI|他力|EQ|家族", task):
            actions.append({
                "action": "update_wbs",
                "task": task,
                "status": "対応中",
            })

    # ③ 学習ログ
    # 例: 「朝活2時間、民法30問、ミス5個」
    study = {}
    hm = re.search(r"(?:朝活|勉強|学習)[^\d]*(\d+(?:\.\d+)?)\s*(?:時間|h|H)", t)
    if hm:
        study["hours"] = float(hm.group(1))

    rm = re.search(r"(?:民法|不登法|択一|記述|区分|表題)[^\d]*(\d+[~〜\-]?\d*)\s*(?:問|p|P|ページ|頁)?", t)
    if rm:
        study["range_takuitsu"] = rm.group(0).strip()

    mm = re.search(r"ミス[^\d]*(\d+)\s*(?:個|問|つ)?", t)
    if mm:
        study["miss_count"] = int(mm.group(1))

    km = re.search(r"(?:完遂|完璧|全完|全問)[：:]?\s*([○×〇])", t)
    if km:
        study["kansei"] = km.group(1)

    if study:
        actions.append({"action": "update_study", **study})

    # ④ KPI
    # 例: 「他力○、朝活○、EQ×」「KPI全部達成」「KPI①②③」
    kpi = {}
    kpi_map = [
        ("tariki",   r"他力[：:\s]*([○×〇✕])"),
        ("eq",       r"EQ[：:\s]*([○×〇✕])"),
        ("chokatsu", r"朝活[：:\s]*([○×〇✕])"),
        ("sen_sabo", r"戦サボ[：:\s]*([○×〇✕])"),
        ("kazoku",   r"家族[：:\s]*([○×〇✕])"),
    ]
    for key, pat in kpi_map:
        m3 = re.search(pat, t)
        if m3:
            val = m3.group(1).replace("〇", "○").replace("✕", "×")
            kpi[key] = val

    # 「KPI全部達成」「KPI①②③④⑤」など全達成パターン
    if re.search(r"KPI.{0,10}(?:全部|全て|全達成|達成|○|①②③④⑤)", t) and not kpi:
        for key, _ in kpi_map:
            kpi[key] = "○"

    cm = re.search(r"(?:TOP3|今日の3|top3)[：:\s]*(.+?)(?:$|、|,|\n)", t, re.IGNORECASE)
    if cm:
        kpi["top3"] = cm.group(1).strip()

    if kpi:
        actions.append({"action": "update_kpi", **kpi})

    # ⑤ 日次ログ追記（マッチしなかった場合も追記候補）
    # 例: 「〇〇をやった」「〇〇に参加した」
    if re.search(r"(?:やった|参加|mtg|ミーティング|打ち合|会議|報告|確認|連絡|提出|送った|送信)", t, re.IGNORECASE):
        actions.append({
            "action": "append_log",
            "content": t,
            "category": "業務記録",
        })

    return actions


# ── Webhookエンドポイント ─────────────────────────────────
@app.route("/report", methods=["POST"])
def report():
    data = request.get_json(force=True)
    text = data.get("text", "").strip()
    if not text:
        return jsonify({"ok": False, "error": "text is empty"}), 400

    actions = parse_text(text)
    if not actions:
        msg = f"⚠️ パターン未検出: {text[:40]}\n手動で確認してください"
        send_line(msg)
        return jsonify({"ok": False, "error": "パターン未検出", "text": text}), 200

    results = []
    summaries = [f"📋 受信: {text[:40]}"]

    for act in actions:
        r = call_apps_script(act)
        results.append(r)
        if r.get("ok") and r.get("result"):
            res = r["result"]
            upd = res.get("updated", [])
            if upd:
                summaries.append("✅ " + " / ".join(upd))
            else:
                summaries.append(f"⚠️ 更新なし (sheet={res.get('sheet','')} row={res.get('row','')})")

    send_line("\n".join(summaries))
    return jsonify({"ok": True, "results": results, "actions": actions})


PROPERTY_KEYWORDS = ["物件", "利回り", "万円", "㎡", "不動産", "売り", "築", "土地", "アパート", "マンション", "CF", "利率"]
PROPERTY_SHEET_ID = "1-mf4JxVXLyghcDcyfxOnh3kNYp7Wrfgv4neGHjyEOlw"
CLAUDE_BIN = "/Users/kikuchikenta/.npm-global/bin/claude"
LOG_DIR = "/Users/kikuchikenta/01_honbu_docs_automation/scripts/logs"

def _run_property_intake(text):
    """バックグラウンドでclaudeを呼んで物件DBに追記し、LINEに結果を通知する。"""
    today = datetime.now().strftime("%Y-%m-%d")
    prompt = f"""以下の物件情報を玉川式で評価し、Google Sheetsに追記してください。

【物件情報】
{text}

【実行手順】
1. 所在地・物件名・購入価格(万円)・表面利回り(%)・土地面積(㎡)・建物面積(㎡)・構造・築年数を抽出（不明は「不明」）
2. 玉川式で評価:
   - 実質利回り = 表面利回り × 0.85（経費15%控除）
   - CF率 = (実質利回り - 融資金利2.5%) × 0.7（目標1.5%以上）
   - 土地値割合 = 路線価×面積 ÷ 購入価格（目標0.4以上）
   - ステータス判定: CF率1.5%以上かつ土地値0.4以上→「検討中」、どちらか未達→「要調査」、両方未達→「見送り」
3. sheets_token.pickle を使い Google Sheets API で以下IDに追記:
   スプレッドシートID: {PROPERTY_SHEET_ID}
   シート名: 01_物件検討DB
   列順(A〜Q): 検討日,物件名,所在地,用途地域,土地面積,建物面積,構造,築年数,購入価格,表面利回り,実質利回り,年間CF,融資打診先,融資結果,ステータス,見送り理由,備考
   検討日は {today}
   sheets_token.pickleのパス: /Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle

4. 追記後、以下フォーマットで結果を出力:
「[物件名] [所在地] 購入価格:X万円 CF率:X.X% 土地値:X.XX → [判断]」

Google Sheets APIのコードはpython3で書いて実行してください。"""

    log_path = os.path.join(LOG_DIR, "property_line_intake.log")
    try:
        result = subprocess.run(
            [CLAUDE_BIN, "--print", "--dangerously-skip-permissions", "-p", prompt],
            capture_output=True, text=True, timeout=180,
            env={**os.environ, "HOME": "/Users/kikuchikenta",
                 "PATH": "/Users/kikuchikenta/.npm-global/bin:/usr/local/bin:/usr/bin:/bin"}
        )
        output = result.stdout.strip() or "（出力なし）"
        with open(log_path, "a") as f:
            f.write(f"\n[{datetime.now().isoformat()}] LINE物件受付\n入力: {text[:100]}\n結果: {output[:300]}\n")
        send_line(f"📊 物件DB追記完了\n{output[:200]}")
    except Exception as e:
        send_line(f"⚠️ 物件処理エラー: {str(e)[:100]}")


@app.route("/property", methods=["POST"])
def property_intake():
    """LINEや任意の経路から物件情報を受け取り、claudeで評価してDBに自動追記。"""
    data = request.get_json(force=True)
    text = data.get("text", "").strip()
    if not text:
        return jsonify({"ok": False, "error": "text is empty"}), 400

    send_line(f"📥 物件情報受付\n解析中... しばらくお待ちください\n\n{text[:80]}")
    threading.Thread(target=_run_property_intake, args=(text,), daemon=True).start()
    return jsonify({"ok": True, "message": "物件情報を受け付けました。処理後にLINEに通知します。"})


@app.route("/ping", methods=["GET", "POST"])
def ping():
    return jsonify({"pong": True, "time": datetime.now().isoformat()})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5055))
    print(f"KHD Voice Webhook listening on :{port}")
    app.run(host="0.0.0.0", port=port)
