#!/bin/bash
# patrol.py の自己診断。REINSには一切アクセスしない（モック画面が相手）。
#
#   bash reins_patrol/tests/run_tests.sh
#
# 何を確かめるか:
#   1. 図面PDFから11項目を正規表現で抜けるか（test_parse_and_diff.py）
#   2. 新着差分（state/seen.json）が効くか
#   3. ログイン〜保存済み検索〜図面一括取得〜PDF保存の一連が、
#      Playwrightのコードとして実際に動くか（test_browser_flow.py）
#
# 3は「REINSの実画面と一致するか」の検証ではない。実画面の確認は
#   patrol.py --headed --pause-after-login
# でしか取れない。ここが通っていても初回は必ず目視すること。
set -uo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"

PY="${PY:-python3}"
VENV="$HERE/.venv-test"
[ -d "$VENV" ] || "$PY" -m venv "$VENV"
"$VENV/bin/pip" install -q -r "$HERE/../requirements.txt" reportlab

RC=0
echo "───────── 1. PDF解析 ＆ 新着差分 ─────────"
"$VENV/bin/python" "$HERE/test_parse_and_diff.py" || RC=1
echo ""
echo "───────── 2. ブラウザ操作 通し（モック） ─────────"
"$VENV/bin/python" -m playwright install chromium >/dev/null 2>&1
"$VENV/bin/python" "$HERE/test_browser_flow.py" || RC=1

echo ""
[ $RC -eq 0 ] && echo "✅ 自己診断すべて通過" || echo "❌ 失敗あり"
exit $RC
