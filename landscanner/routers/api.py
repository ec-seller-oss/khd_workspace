"""API エンドポイント"""
from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse
from sqlalchemy.orm import Session
from sqlalchemy import or_, and_
from typing import Optional, List
from datetime import datetime, timedelta
import json

from database import get_db
from models import (
    User, SearchPattern, Property, NotificationHistory,
    Station, Announcement, FeatureRequest
)
from schemas import (
    SearchPatternCreate, SearchQuery, VolumeCheckInput,
    HistoryQuery, FeatureRequestCreate
)
from calculator import calc_volume, calc_yield_for_property
from auth import SESSION_COOKIE, decode_session_token

router = APIRouter(prefix="/api")


def get_user_from_request(request: Request, db: Session) -> Optional[User]:
    token = request.cookies.get(SESSION_COOKIE)
    if not token:
        return None
    data = decode_session_token(token)
    if not data:
        return None
    return db.query(User).filter(User.id == data["user_id"]).first()


def require_user(request: Request, db: Session = Depends(get_db)) -> User:
    user = get_user_from_request(request, db)
    if not user:
        raise HTTPException(status_code=401, detail="Unauthorized")
    return user


# ---- 検索パターン ----

@router.get("/settings/patterns")
async def get_patterns(request: Request, db: Session = Depends(get_db)):
    user = require_user(request, db)
    patterns = db.query(SearchPattern).filter(SearchPattern.user_id == user.id).all()
    result = []
    for p in patterns:
        result.append({
            "id": p.id,
            "name": p.name,
            "enabled": p.enabled,
            "stations": p.stations or [],
            "lines": p.lines or [],
            "build_cost_per_tsubo": p.build_cost_per_tsubo,
            "room_area": p.room_area,
            "demolition_unit": p.demolition_unit,
            "living_ratio": p.living_ratio,
            "monthly_rent": p.monthly_rent,
            "yield_threshold": p.yield_threshold,
            "max_land_price": p.max_land_price,
            "max_total_investment": p.max_total_investment,
            "min_area": p.min_area,
            "max_area": p.max_area,
            "notify_email": p.notify_email,
            "notify_announcement": p.notify_announcement,
            "notify_timing": p.notify_timing,
        })
    return result


@router.post("/settings/pattern")
async def create_pattern(
    request: Request,
    data: SearchPatternCreate,
    db: Session = Depends(get_db),
):
    user = require_user(request, db)
    count = db.query(SearchPattern).filter(SearchPattern.user_id == user.id).count()
    if count >= 3:
        raise HTTPException(status_code=400, detail="最大3パターンまでです")
    pattern = SearchPattern(
        user_id=user.id,
        **data.model_dump()
    )
    db.add(pattern)
    db.commit()
    db.refresh(pattern)
    return {"id": pattern.id, "message": "作成しました"}


@router.put("/settings/pattern/{pattern_id}")
async def update_pattern(
    pattern_id: int,
    request: Request,
    data: SearchPatternCreate,
    db: Session = Depends(get_db),
):
    user = require_user(request, db)
    pattern = db.query(SearchPattern).filter(
        SearchPattern.id == pattern_id,
        SearchPattern.user_id == user.id
    ).first()
    if not pattern:
        raise HTTPException(status_code=404, detail="パターンが見つかりません")

    for key, value in data.model_dump().items():
        setattr(pattern, key, value)
    db.commit()
    return {"message": "保存しました"}


@router.delete("/settings/pattern/{pattern_id}")
async def delete_pattern(
    pattern_id: int,
    request: Request,
    db: Session = Depends(get_db),
):
    user = require_user(request, db)
    pattern = db.query(SearchPattern).filter(
        SearchPattern.id == pattern_id,
        SearchPattern.user_id == user.id
    ).first()
    if not pattern:
        raise HTTPException(status_code=404, detail="パターンが見つかりません")
    db.delete(pattern)
    db.commit()
    return {"message": "削除しました"}


# ---- 物件検索 ----

@router.post("/search")
async def search_properties(
    request: Request,
    query: SearchQuery,
    db: Session = Depends(get_db),
):
    user = require_user(request, db)

    since = datetime.now() - timedelta(days=query.days)
    q = db.query(Property).filter(Property.scraped_at >= since)

    if query.prefectures:
        q = q.filter(Property.prefecture.in_(query.prefectures))

    if query.lines:
        line_filters = [Property.line.contains(line) for line in query.lines]
        q = q.filter(or_(*line_filters))

    if query.min_area:
        q = q.filter(Property.area >= query.min_area)
    if query.max_area:
        q = q.filter(Property.area <= query.max_area)
    if query.max_land_price:
        q = q.filter(Property.price <= query.max_land_price)

    props = q.order_by(Property.scraped_at.desc()).limit(500).all()

    # 駅別家賃マップ（キャッシュ）
    station_rent_map: dict = {}
    for st in db.query(Station).all():
        station_rent_map[st.name] = st.monthly_rent

    # 利回り計算してフィルタリング
    results = []
    for prop in props:
        if not prop.price or not prop.area:
            continue
        # 最寄り駅の賃料相場を使用（なければクエリのデフォルト値）
        rent = station_rent_map.get(prop.nearest_station, query.monthly_rent)
        yield_pct = calc_yield_for_property(
            land_price=prop.price,
            land_area=prop.area,
            far=prop.far or 200.0,
            monthly_rent=rent,
            room_area=query.room_area,
            build_cost_per_tsubo=query.build_cost_per_tsubo,
            living_ratio=query.living_ratio,
            demolition_unit=query.demolition_unit,
            has_old_house=prop.has_old_house,
        )

        if query.min_yield and yield_pct < query.min_yield:
            continue

        # 総投資額フィルター
        if query.max_total_investment:
            from calculator import calc_volume
            from schemas import VolumeCheckInput
            inp = VolumeCheckInput(
                land_price=prop.price,
                land_area=prop.area,
                far=prop.far or 200.0,
                monthly_rent=query.monthly_rent,
                room_area=query.room_area,
                build_cost_per_tsubo=query.build_cost_per_tsubo,
                living_ratio=query.living_ratio,
                demolition_unit=query.demolition_unit,
                has_old_house=prop.has_old_house,
            )
            vol = calc_volume(inp)
            if vol.total_investment > query.max_total_investment:
                continue

        results.append({
            "id": prop.id,
            "source": prop.source,
            "title": prop.title,
            "address": prop.address,
            "prefecture": prop.prefecture,
            "price": prop.price,
            "area": prop.area,
            "nearest_station": prop.nearest_station,
            "walk_minutes": prop.walk_minutes,
            "line": prop.line,
            "url": prop.url,
            "scraped_at": prop.scraped_at.strftime("%Y-%m-%d") if prop.scraped_at else None,
            "yield_pct": round(yield_pct, 2),
            "far": prop.far,
        })

    # ページネーション
    total = len(results)
    start = (query.page - 1) * query.per_page
    end = start + query.per_page
    page_results = results[start:end]

    return {
        "total": total,
        "page": query.page,
        "per_page": query.per_page,
        "results": page_results,
    }


# ---- ボリュームチェック ----

@router.post("/volume-check")
async def volume_check_api(
    request: Request,
    data: VolumeCheckInput,
    db: Session = Depends(get_db),
):
    user = require_user(request, db)
    result = calc_volume(data)
    return result


# ---- 通知履歴 ----

@router.get("/history")
async def get_history(
    request: Request,
    pattern_id: Optional[int] = None,
    keyword: Optional[str] = None,
    page: int = 1,
    per_page: int = 50,
    db: Session = Depends(get_db),
):
    user = require_user(request, db)

    q = (
        db.query(NotificationHistory, Property, SearchPattern)
        .join(Property, NotificationHistory.property_id == Property.id)
        .join(SearchPattern, NotificationHistory.pattern_id == SearchPattern.id)
        .filter(NotificationHistory.user_id == user.id)
    )

    if pattern_id:
        q = q.filter(NotificationHistory.pattern_id == pattern_id)

    if keyword:
        q = q.filter(
            or_(
                Property.title.contains(keyword),
                Property.address.contains(keyword),
                Property.nearest_station.contains(keyword),
            )
        )

    total = q.count()
    items = q.order_by(NotificationHistory.sent_at.desc()).offset((page - 1) * per_page).limit(per_page).all()

    results = []
    for nh, prop, pattern in items:
        results.append({
            "id": nh.id,
            "property_id": prop.id,
            "source": prop.source,
            "title": prop.title,
            "address": prop.address,
            "price": prop.price,
            "area": prop.area,
            "nearest_station": prop.nearest_station,
            "walk_minutes": prop.walk_minutes,
            "url": prop.url,
            "yield_pct": nh.yield_pct,
            "sent_at": nh.sent_at.strftime("%Y-%m-%d %H:%M") if nh.sent_at else None,
            "pattern_name": pattern.name,
        })

    return {
        "total": total,
        "page": page,
        "per_page": per_page,
        "results": results,
    }


# ---- お知らせ ----

@router.get("/announcements")
async def get_announcements(db: Session = Depends(get_db)):
    items = db.query(Announcement).order_by(Announcement.date.desc()).all()
    return [
        {
            "id": a.id,
            "date": a.date.strftime("%Y-%m-%d") if a.date else None,
            "title": a.title,
            "content": a.content,
        }
        for a in items
    ]


# ---- 機能リクエスト ----

@router.post("/feature-request")
async def post_feature_request(
    request: Request,
    data: FeatureRequestCreate,
    db: Session = Depends(get_db),
):
    user = require_user(request, db)
    if len(data.content) > 1000:
        raise HTTPException(status_code=400, detail="1000文字以内で入力してください")
    fr = FeatureRequest(content=data.content)
    db.add(fr)
    db.commit()
    return {"message": "送信しました。ありがとうございます！"}


# ---- 管理者: 手動スクレイピングトリガー ----

@router.post("/admin/scrape")
async def trigger_scrape(request: Request, db: Session = Depends(get_db)):
    """手動でスクレイピングを実行（管理者のみ）"""
    user = require_user(request, db)
    if not user.is_admin:
        raise HTTPException(status_code=403, detail="Admin only")

    import asyncio
    from main import scheduled_scrape
    asyncio.create_task(scheduled_scrape())
    return {"message": "スクレイピングを開始しました（バックグラウンド実行中）"}


# ---- 駅データ ----

@router.get("/stations")
async def get_stations(db: Session = Depends(get_db)):
    stations = db.query(Station).all()
    return [
        {
            "id": s.id,
            "name": s.name,
            "line": s.line,
            "prefecture": s.prefecture,
            "lat": s.lat,
            "lng": s.lng,
            "monthly_rent": s.monthly_rent,
            "service_count": s.service_count,
        }
        for s in stations
    ]
