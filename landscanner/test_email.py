"""メール送信テスト — python3 test_email.py で実行"""
import sys
sys.path.insert(0, '.')
from dotenv import load_dotenv
load_dotenv()

from notifier import send_email

TO = input("送信先メールアドレス: ").strip()

html = """
<html><body style="font-family:sans-serif;padding:20px;">
<h2 style="color:#1a2332;">Land Scanner — テストメール</h2>
<p>このメールが届いていれば Gmail SMTP 設定は正常です。</p>
<table border="1" cellpadding="8" style="border-collapse:collapse;font-size:14px;">
  <tr><th>物件</th><th>価格</th><th>面積</th><th>利回り</th></tr>
  <tr><td>千葉県 上総湊駅</td><td>130万円</td><td>135㎡</td><td style="color:green;font-weight:bold;">8.2%</td></tr>
  <tr><td>千葉県 三門駅</td><td>200万円</td><td>291㎡</td><td style="color:green;font-weight:bold;">8.2%</td></tr>
</table>
<p style="margin-top:20px;font-size:12px;color:#666;">Land Scanner 自動送信</p>
</body></html>
"""

ok = send_email(TO, "【Land Scanner】テストメール", html)
print("送信成功 ✓" if ok else "送信失敗 — .env の SMTP_USER / SMTP_PASSWORD を確認")
