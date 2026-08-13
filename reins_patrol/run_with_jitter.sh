#!/bin/bash
# run_with_jitter.sh — cronから呼ぶラッパー
#
# なぜジッターを入れるか:
#   毎日きっかり同じ秒にアクセスする挙動は、人間の操作としては不自然に見える。
#   cronの起動時刻から 0〜10分 ランダムに遅らせて、時刻の規則性を消す。
#
# crontab例（1日2回まで。増やさない）:
#   10 8  * * 1-5 /Users/kikuchikenta/reins-patrol/run_with_jitter.sh >> /Users/kikuchikenta/reins-patrol/logs/cron.log 2>&1
#   10 18 * * 1-5 /Users/kikuchikenta/reins-patrol/run_with_jitter.sh >> /Users/kikuchikenta/reins-patrol/logs/cron.log 2>&1
set -uo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

MAX_JITTER=${MAX_JITTER:-600}          # 秒。既定10分
SLEEP=$((RANDOM % (MAX_JITTER + 1)))
echo "[$(date '+%F %T')] jitter ${SLEEP}s 待機"
sleep "$SLEEP"

./.venv/bin/python patrol.py
RC=$?

case $RC in
  0) echo "[$(date '+%F %T')] 正常終了" ;;
  5) echo "[$(date '+%F %T')] メンテナンス時間帯のためスキップ（異常ではない）" ;;
  *) echo "[$(date '+%F %T')] ❌ 異常終了 rc=$RC — logs/ の最新スクショを確認" ;;
esac
exit $RC
