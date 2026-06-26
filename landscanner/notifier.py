"""
Land Scanner — 通知エンジン

パターンマッチング → NotificationHistory 作成 → Gmail 送信
"""
import os
import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, date
from typing import Optional

from sqlalchemy.orm import Session

from models import User, SearchPattern, Property, NotificationHistory
from calculator import calc_yield_for_property

logger = logging.getLogger(__name__)

# ---- SMTP 設定（.env から取得） -----------------------------------------------
SMTP_HOST     = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT     = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER     = os.getenv("SMTP_USER", "")       # Gmail アドレス
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")   # Gmail App Password（16桁）
FROM_NAME     = os.getenv("FROM_NAME", "Land Scanner")
APP_URL       = os.getenv("APP_URL", "http://127.0.0.1:8765")


# ---- メール送信 ---------------------------------------------------------------

def send_email(to: str, subject: str, html_body: str) -> bool:
    """Gmail SMTP でメール送信。未設定なら WARNING を出して終了。"""
    if not SMTP_USER or not SMTP_PASSWORD:
        logger.warning("[notifier] SMTP 未設定 (SMTP_USER / SMTP_PASSWORD が空). メール送信をスキップ.")
        return False

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = f"{FROM_NAME} <{SMTP_USER}>"
    msg["To"]      = to
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=15) as server:
            server.ehlo()
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, to, msg.as_string())
        logger.info(f"[notifier] メール送信 OK → {to}: {subject}")
        return True
    except Exception as e:
        logger.error(f"[notifier] メール送信失敗 → {to}: {e}")
        return False


# ---- パターンマッチング --------------------------------------------------------

def match_property(prop: Property, pat: SearchPattern) -> Optional[float]:
    """
    物件がパターンにマッチするか判定。
    マッチすれば yield_pct を返す。マッチしなければ None。
    """
    # エリア条件（駅 or 沿線のどちらかが指定されていること）
    station_names = {s["name"] for s in (pat.stations or [])}
    lines = set(pat.lines or [])
    if not station_names and not lines:
        return None  # エリア未設定パターンは通知しない

    station_ok = (prop.nearest_station and prop.nearest_station in station_names)
    line_ok    = (prop.line and prop.line in lines)
    if not station_ok and not line_ok:
        return None

    # 面積・価格が欠損なら計算不能
    if not prop.area or not prop.price:
        return None

    # 面積フィルター
    if pat.min_area and prop.area < pat.min_area:
        return None
    if pat.max_area and prop.area > pat.max_area:
        return None

    # 土地価格上限
    if pat.max_land_price and prop.price > pat.max_land_price:
        return None

    # 利回り計算
    yield_pct = calc_yield_for_property(
        land_price=prop.price,
        land_area=prop.area,
        far=prop.far or 200.0,
        monthly_rent=pat.monthly_rent,
        room_area=pat.room_area,
        build_cost_per_tsubo=pat.build_cost_per_tsubo,
        living_ratio=pat.living_ratio,
        demolition_unit=pat.demolition_unit,
        has_old_house=prop.has_old_house,
    )

    # 総投資額上限
    if pat.max_total_investment:
        from calculator import calc_volume
        from schemas import VolumeCheckInput
        result = calc_volume(VolumeCheckInput(
            land_price=prop.price, land_area=prop.area,
            far=prop.far or 200.0, monthly_rent=pat.monthly_rent,
            room_area=pat.room_area, build_cost_per_tsubo=pat.build_cost_per_tsubo,
            living_ratio=pat.living_ratio, demolition_unit=pat.demolition_unit,
            has_old_house=prop.has_old_house,
        ))
        if result.total_investment > pat.max_total_investment:
            return None

    # 利回り閾値
    if yield_pct < pat.yield_threshold:
        return None

    return yield_pct


# ---- 新着物件 → 通知ディスパッチ（スクレイプ後に呼ぶ）-----------------------

def dispatch_notifications(db: Session, new_properties: list[Property]) -> int:
    """
    新着物件リストをすべての有効パターンと照合し:
      - NotificationHistory を作成
      - notify_timing == 'realtime' かつ notify_email == True のユーザーには即時メール
    戻り値: 作成した NotificationHistory の件数
    """
    if not new_properties:
        return 0

    # 有効なパターンを全ユーザー分取得
    patterns = (
        db.query(SearchPattern)
        .filter(SearchPattern.enabled == True)
        .all()
    )
    if not patterns:
        return 0

    # 既通知チェック用: (property_id, pattern_id) の集合
    already_notified = {
        (nh.property_id, nh.pattern_id)
        for nh in db.query(NotificationHistory.property_id, NotificationHistory.pattern_id).all()
    }

    created = 0
    realtime_queue: dict[int, list[tuple[Property, float, str]]] = {}  # user_id → [(prop, yield, pattern_name)]

    for prop in new_properties:
        for pat in patterns:
            if (prop.id, pat.id) in already_notified:
                continue

            yield_pct = match_property(prop, pat)
            if yield_pct is None:
                continue

            # NotificationHistory 作成
            nh = NotificationHistory(
                user_id=pat.user_id,
                pattern_id=pat.id,
                property_id=prop.id,
                yield_pct=yield_pct,
            )
            db.add(nh)
            already_notified.add((prop.id, pat.id))
            created += 1

            # リアルタイム通知キューに積む
            if pat.notify_email and pat.notify_timing == "realtime":
                realtime_queue.setdefault(pat.user_id, [])
                realtime_queue[pat.user_id].append((prop, yield_pct, pat.name))

    db.commit()

    # リアルタイムメール送信
    for user_id, items in realtime_queue.items():
        user = db.query(User).filter(User.id == user_id).first()
        if not user:
            continue
        subject = f"【Land Scanner】新着物件 {len(items)}件"
        html = _build_html(items, realtime=True)
        send_email(user.email, subject, html)

    logger.info(f"[notifier] dispatch完了: {created}件 通知作成, {len(realtime_queue)}ユーザーにリアルタイム送信")
    return created


# ---- 日次ダイジェスト（毎朝7時に呼ぶ）---------------------------------------

def send_daily_digest(db: Session) -> int:
    """
    notify_timing == 'daily' のユーザーに昨日～今日の通知をまとめてメール。
    戻り値: 送信件数
    """
    today = date.today()

    # 今日分の NotificationHistory を取得（daily ユーザーのみ）
    rows = (
        db.query(NotificationHistory, Property, SearchPattern, User)
        .join(Property,      NotificationHistory.property_id  == Property.id)
        .join(SearchPattern, NotificationHistory.pattern_id   == SearchPattern.id)
        .join(User,          NotificationHistory.user_id      == User.id)
        .filter(SearchPattern.notify_email  == True)
        .filter(SearchPattern.notify_timing == "daily")
        .filter(SearchPattern.enabled       == True)
        .all()
    )

    # ユーザーごとにグループ化
    per_user: dict[int, dict] = {}
    for nh, prop, pat, user in rows:
        if nh.sent_at and nh.sent_at.date() != today:
            continue
        uid = user.id
        if uid not in per_user:
            per_user[uid] = {"email": user.email, "items": []}
        per_user[uid]["items"].append((prop, nh.yield_pct, pat.name))

    sent = 0
    for uid, data in per_user.items():
        if not data["items"]:
            continue
        subject = f"【Land Scanner】本日の新着物件 {len(data['items'])}件"
        html = _build_html(data["items"], realtime=False)
        if send_email(data["email"], subject, html):
            sent += 1

    logger.info(f"[notifier] 日次ダイジェスト: {sent}ユーザーに送信")
    return sent


# ---- HTML メール本文 ----------------------------------------------------------

def _yield_color(y: Optional[float]) -> str:
    if y is None:
        return "#dc2626"
    if y >= 8.0:
        return "#16a34a"
    if y >= 7.0:
        return "#ca8a04"
    return "#dc2626"


def _build_html(items: list[tuple[Property, float, str]], realtime: bool) -> str:
    rows_html = ""
    for prop, yield_pct, pat_name in items:
        color = _yield_color(yield_pct)
        title = prop.title or prop.address or "（物件名なし）"
        url   = prop.url or f"{APP_URL}/search"
        price = f"{prop.price:.0f}万円" if prop.price else "-"
        area  = f"{prop.area:.0f}㎡"   if prop.area  else "-"
        rows_html += f"""
        <tr>
          <td style="padding:10px 8px;border-bottom:1px solid #e5e7eb;">
            <a href="{url}" style="color:#2563eb;text-decoration:none;font-weight:600;">{title}</a>
            <div style="font-size:11px;color:#6b7280;margin-top:2px;">パターン: {pat_name}</div>
          </td>
          <td style="padding:10px 8px;border-bottom:1px solid #e5e7eb;white-space:nowrap;">{price}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #e5e7eb;white-space:nowrap;">{area}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #e5e7eb;font-weight:700;color:{color};white-space:nowrap;">{yield_pct:.1f}%</td>
        </tr>"""

    timing_label = "リアルタイム通知" if realtime else "本日のまとめ"

    return f"""<!DOCTYPE html>
<html lang="ja">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:'Helvetica Neue',Arial,sans-serif;">
  <div style="max-width:640px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.08);">

    <!-- ヘッダー -->
    <div style="background:#1a2332;padding:24px 28px;">
      <div style="color:#60a5fa;font-size:12px;font-weight:700;letter-spacing:2px;margin-bottom:4px;">LAND SCANNER</div>
      <div style="color:#fff;font-size:20px;font-weight:700;">新着物件のお知らせ ({timing_label})</div>
      <div style="color:#94a3b8;font-size:13px;margin-top:4px;">{datetime.now().strftime('%Y年%m月%d日 %H:%M')} 時点</div>
    </div>

    <!-- テーブル -->
    <div style="padding:20px 28px;">
      <p style="color:#374151;font-size:14px;margin:0 0 16px;">
        条件に合致する物件が <strong>{len(items)}件</strong> 見つかりました。
      </p>
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <thead>
          <tr style="background:#f8fafc;">
            <th style="padding:10px 8px;text-align:left;color:#6b7280;font-weight:600;border-bottom:2px solid #e5e7eb;">物件</th>
            <th style="padding:10px 8px;text-align:left;color:#6b7280;font-weight:600;border-bottom:2px solid #e5e7eb;">価格</th>
            <th style="padding:10px 8px;text-align:left;color:#6b7280;font-weight:600;border-bottom:2px solid #e5e7eb;">面積</th>
            <th style="padding:10px 8px;text-align:left;color:#6b7280;font-weight:600;border-bottom:2px solid #e5e7eb;">利回り</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>

    <!-- フッター -->
    <div style="padding:20px 28px;border-top:1px solid #e5e7eb;background:#f9fafb;">
      <a href="{APP_URL}/search"
         style="display:inline-block;background:#2563eb;color:#fff;text-decoration:none;padding:10px 22px;border-radius:8px;font-size:13px;font-weight:600;">
        物件を確認する →
      </a>
      <p style="margin:16px 0 0;font-size:11px;color:#9ca3af;">
        通知設定の変更は <a href="{APP_URL}/settings" style="color:#2563eb;">こちら</a> から。
      </p>
    </div>

  </div>
</body>
</html>"""
