#!/usr/bin/env python3
"""
KHD WBS スプレッドシート更新ライブラリ。
Google Apps Script Web App 経由で書き込む（Sheets API 不要）。
"""
import os
import requests
from datetime import datetime

# Apps Script デプロイURL（デプロイ後に設定）
APPS_SCRIPT_URL = os.environ.get(
    "KHD_APPS_SCRIPT_URL",
    "https://script.google.com/macros/s/AKfycbzj_irD6brENSL6KL2nrqopx3hVScul3LWkhrsJZxyIcs6UWca5ekd8ZvUhtHhq-CNfMQ/exec"
)
SECRET_TOKEN = "khd_report_2026"

# 日報専用 web app（2026-06-05 デプロイ・openByIdで日報スプシに書く）
NIPPO_APPS_SCRIPT_URL = os.environ.get(
    "KHD_NIPPO_SCRIPT_URL",
    "https://script.google.com/macros/s/AKfycbxkmH3gj95Yg1cRYCYLc7fAN26N58euHHE2bx2GracXdL_8OOQ0LN_weTjzj3cRfQ/exec"
)


def _call(action, _url=None, **kwargs):
    url = _url or APPS_SCRIPT_URL
    payload = {"token": SECRET_TOKEN, "action": action, **kwargs}
    resp = requests.post(url, json=payload, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    if not data.get("ok"):
        raise RuntimeError(f"Apps Script エラー: {data.get('error')}")
    return data.get("result", {})


def update_wbs_task(task_partial, status=None, progress=None):
    """WBS タスクのステータス・進捗率を更新。"""
    return _call("update_wbs",
                 task=task_partial,
                 status=status,
                 progress=progress)


def update_study_log(hours=None, range_takuitsu=None, range_kijutsu=None,
                     miss_count=None, comment=None, kansei=None):
    """今日の調査士学習ログを更新。"""
    return _call("update_study",
                 hours=hours,
                 range_takuitsu=range_takuitsu,
                 range_kijutsu=range_kijutsu,
                 miss_count=miss_count,
                 comment=comment,
                 kansei=kansei)


def update_kpi(tariki=None, eq=None, chokatsu=None, sen_sabo=None,
               kazoku=None, comment=None, top3=None):
    """今日の KPI 5指標を更新。"""
    return _call("update_kpi",
                 tariki=tariki, eq=eq, chokatsu=chokatsu,
                 sen_sabo=sen_sabo, kazoku=kazoku,
                 comment=comment, top3=top3)


def append_daily_log(content, category="完了報告", dept="",
                     project="", manager_comment="", status="✅"):
    """14_日次ログに1行追記。"""
    return _call("append_log",
                 content=content,
                 category=category,
                 dept=dept,
                 project=project,
                 manager_comment=manager_comment,
                 status=status,
                 time=datetime.now().strftime("%H:%M"))


import csv as _csv

NIPPO_CSV = os.path.join(os.path.dirname(__file__), "..",
                         ".company", "secretary", "nippo", "KHD_日報ログ.csv")


def _read_nippo_rows():
    with open(NIPPO_CSV, newline="") as f:
        return [row for row in _csv.reader(f)]


def write_nippo(rows=None):
    """日報スプシをローカルCSV全体で全置換ミラー（URL固定）。
    rows未指定なら .company/secretary/nippo/KHD_日報ログ.csv を丸ごと反映。"""
    if rows is None:
        rows = _read_nippo_rows()
    return _call("write_nippo", _url=NIPPO_APPS_SCRIPT_URL, rows=rows)


def append_nippo(rows):
    """日報スプシの末尾にデータ行を追記（ヘッダ除く2D配列）。"""
    return _call("append_nippo", _url=NIPPO_APPS_SCRIPT_URL, rows=rows)


# 02_作業DB 書き込み専用 web app（2026-06-07 v2デプロイ・read_db02/update_db02/delete_db02_rows追加）
DB02_WEBAPP_URL = os.environ.get(
    "KHD_DB02_SCRIPT_URL",
    "https://script.google.com/macros/s/AKfycbz7QJAh9cdlyGjd0Uep4LIifw1_fiHDi0oWZ2B6VLphX1T7G3DffgnWVrhpWq-eG_g_/exec"
)


def append_db02(row, host_id="1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc", _url=None):
    """02_作業DBに1行追記。row={列名:値}の辞書。ヘッダ名で対応づけ＝列順を動かしても壊れない。
    戻り値のunmatchedに、対応する列が見つからなかったキーが入る。"""
    return _call("append_db02", _url=_url or DB02_WEBAPP_URL, row=row, hostId=host_id)


def db02_headers(host_id="1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc", _url=None):
    """02_作業DBのヘッダ名一覧を取得（対応づけ確認・学習用）。"""
    return _call("append_db02", _url=_url or DB02_WEBAPP_URL, headers_only=True, hostId=host_id)


PIPELINE_CSV = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            ".company", "secretary", "nippo", "KHD_収益パイプライン.csv")


# 資金繰り司令塔スプシ（収益パイプライン同居先）
CASHFLOW_ID = "1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc"


def write_pipeline(rows=None, host_id=CASHFLOW_ID):
    """収益パイプラインを資金繰りスプシの「収益パイプライン」タブへ全置換ミラー。
    rows未指定なら .company/secretary/nippo/KHD_収益パイプライン.csv を丸ごと反映。
    host_id を渡せばGAS再デプロイなしで書込先を切替可能。
    "="始まりのセル（期待値・SUM・ギャップ）は数式として展開される。"""
    if rows is None:
        with open(PIPELINE_CSV, newline="") as f:
            rows = [row for row in _csv.reader(f)]
    return _call("write_pipeline", rows=rows, hostId=host_id)


def format_rules():
    """運用ルールの色分けを資金繰りスプシ各タブへ適用（🟡入力/🟧毎日/🟦週次月次/⬜自動）。"""
    return _call("format_rules")


def finance_fixes():
    """財務修正バンドル：税理士正値・経常利益を結論色・営業ドライバー色ズレ修正・未来会計図表のPL参照化。"""
    return _call("finance_fixes")


def tidy_tabs():
    """タブ整理：操縦席11枚を流れ順に前へ／データ・スポット16枚を非表示（名前は変えず参照は維持）。"""
    return _call("tidy_tabs")


def ping():
    """接続テスト。"""
    return _call("ping")


if __name__ == "__main__":
    import sys
    cmd = sys.argv[1] if len(sys.argv) > 1 else "ping"
    if cmd == "ping":
        print(ping())
    elif cmd == "wbs":
        print(update_wbs_task(sys.argv[2], status=sys.argv[3] if len(sys.argv) > 3 else None))
    elif cmd == "kpi":
        print(update_kpi(tariki="○", chokatsu="○"))
    elif cmd == "study":
        print(update_study_log(hours=sys.argv[2] if len(sys.argv) > 2 else None))
    elif cmd == "log":
        print(append_daily_log(sys.argv[2] if len(sys.argv) > 2 else "テスト"))
