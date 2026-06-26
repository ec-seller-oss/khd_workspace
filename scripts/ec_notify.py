# -*- coding: utf-8 -*-
"""
KHD EC通知ヘルパー — LINE Messaging API 経由で菊池に結果通知
呼び出し方:
  from ec_notify import notify
  notify("✅ 月次EC粗利パイプライン完了\n実精算¥XX万 / 粗利率X.X%")
"""
import json, os, requests, traceback

CFG = os.path.expanduser("~/.config/khd/coupang.json")

def _load_cfg():
    try:
        with open(CFG) as f:
            return json.load(f)
    except Exception:
        return {}

def notify(message: str, urgent: bool = False) -> bool:
    """LINE で菊池に送信。送信失敗しても例外を上げない（ベストエフォート）"""
    cfg = _load_cfg()
    token   = cfg.get("line_channel_token", os.environ.get("LINE_CHANNEL_TOKEN", ""))
    user_id = cfg.get("line_user_id",       os.environ.get("LINE_USER_ID", ""))
    if not token or not user_id:
        print(f"[notify] LINE設定なし（~/.config/khd/coupang.json を確認）: {message}")
        return False
    prefix = "🚨 " if urgent else "📊 "
    try:
        r = requests.post(
            "https://api.line.me/v2/bot/message/push",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"to": user_id, "messages": [{"type": "text", "text": prefix + message}]},
            timeout=10,
        )
        return r.status_code == 200
    except Exception:
        print(f"[notify] LINE送信失敗: {traceback.format_exc()}")
        return False

if __name__ == "__main__":
    import sys
    msg = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else "テスト通知"
    ok = notify(msg)
    print("送信OK" if ok else "送信NG（設定確認が必要）")
