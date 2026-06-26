"""ページルーター（HTML レスポンス）"""
from fastapi import APIRouter, Depends, Request, Query
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from typing import Optional
import json
import os

from database import get_db
from models import User, SearchPattern, Property, NotificationHistory, Station, Announcement
from auth import get_current_user_optional, SESSION_COOKIE, decode_session_token

router = APIRouter()
templates = Jinja2Templates(directory="templates")

# 80沿線リスト
LINES = [
    "JR山手線", "JR中央線", "JR総武線", "JR京浜東北線", "JR常磐線",
    "JR東海道線", "JR横須賀線", "JR南武線", "JR横浜線", "JR京葉線",
    "JR武蔵野線", "JR埼京線", "JR湘南新宿ライン", "JR上野東京ライン",
    "JR外房線", "JR内房線", "JR成田線", "JR総武快速線",
    "東急東横線", "東急田園都市線", "東急目黒線", "東急大井町線",
    "東急池上線", "東急多摩川線", "東急世田谷線",
    "小田急小田原線", "小田急江ノ島線", "小田急多摩線",
    "京王線", "京王井の頭線", "京王相模原線", "京王高尾線",
    "西武池袋線", "西武新宿線", "西武多摩湖線", "西武拝島線",
    "東武東上線", "東武スカイツリーライン", "東武アーバンパークライン",
    "東武東武宇都宮線",
    "京成本線", "京成押上線", "京成千葉線",
    "北総線", "新京成線", "東葉高速線",
    "つくばエクスプレス", "流鉄流山線",
    "東京メトロ銀座線", "東京メトロ丸ノ内線", "東京メトロ日比谷線",
    "東京メトロ東西線", "東京メトロ千代田線", "東京メトロ有楽町線",
    "東京メトロ半蔵門線", "東京メトロ南北線", "東京メトロ副都心線",
    "都営浅草線", "都営三田線", "都営新宿線", "都営大江戸線",
    "横浜市営地下鉄ブルーライン", "横浜市営地下鉄グリーンライン",
    "相鉄本線", "相鉄いずみ野線",
    "みなとみらい線", "江ノ島電鉄",
    "埼玉高速鉄道", "東武伊勢崎線", "秩父鉄道",
    "多摩都市モノレール", "東京臨海高速鉄道りんかい線",
    "湘南モノレール", "大船軌道線",
    "ゆりかもめ", "東京モノレール",
    "舞浜リゾートライン", "千葉都市モノレール",
]

PREFECTURES = ["東京都", "千葉県", "埼玉県", "神奈川県", "茨城県", "栃木県", "群馬県"]


def get_user_from_request(request: Request, db: Session) -> Optional[User]:
    token = request.cookies.get(SESSION_COOKIE)
    if not token:
        return None
    data = decode_session_token(token)
    if not data:
        return None
    return db.query(User).filter(User.id == data["user_id"]).first()


def stations_json_path():
    return os.path.join(os.path.dirname(__file__), "..", "data", "stations.json")


def load_stations():
    path = stations_json_path()
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return []


@router.get("/", response_class=HTMLResponse)
async def home(request: Request, db: Session = Depends(get_db)):
    user = get_user_from_request(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    stations = load_stations()
    return templates.TemplateResponse("home.html", {
        "request": request,
        "user": user,
        "stations": json.dumps(stations, ensure_ascii=False),
    })


@router.get("/settings", response_class=HTMLResponse)
async def settings(request: Request, db: Session = Depends(get_db)):
    user = get_user_from_request(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    raw_patterns = db.query(SearchPattern).filter(SearchPattern.user_id == user.id).all()
    patterns_list = [
        {
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
        }
        for p in raw_patterns
    ]
    stations = load_stations()
    return templates.TemplateResponse("settings.html", {
        "request": request,
        "user": user,
        "patterns": raw_patterns,
        "patterns_json": json.dumps(patterns_list, ensure_ascii=False),
        "lines": LINES,
        "stations_json": json.dumps(stations, ensure_ascii=False),
        "max_patterns": 3,
    })


@router.get("/search", response_class=HTMLResponse)
async def search(request: Request, db: Session = Depends(get_db)):
    user = get_user_from_request(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    return templates.TemplateResponse("search.html", {
        "request": request,
        "user": user,
        "lines": LINES,
        "prefectures": PREFECTURES,
    })


@router.get("/volume-check", response_class=HTMLResponse)
async def volume_check(
    request: Request,
    db: Session = Depends(get_db),
    propertyId: Optional[int] = None,
    source: Optional[str] = None,
):
    user = get_user_from_request(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=302)

    property_data = None
    if propertyId:
        prop = db.query(Property).filter(Property.id == propertyId).first()
        if prop:
            property_data = {
                "id": prop.id,
                "title": prop.title,
                "price": prop.price,
                "area": prop.area,
                "far": prop.far,
                "bcr": prop.bcr,
                "has_old_house": prop.has_old_house,
                "nearest_station": prop.nearest_station,
                "address": prop.address,
            }

    return templates.TemplateResponse("volume_check.html", {
        "request": request,
        "user": user,
        "property_data": json.dumps(property_data, ensure_ascii=False) if property_data else "null",
    })


@router.get("/history", response_class=HTMLResponse)
async def history(request: Request, db: Session = Depends(get_db)):
    user = get_user_from_request(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    patterns = db.query(SearchPattern).filter(SearchPattern.user_id == user.id).all()
    return templates.TemplateResponse("history.html", {
        "request": request,
        "user": user,
        "patterns": patterns,
    })


@router.get("/reins", response_class=HTMLResponse)
async def reins(request: Request, db: Session = Depends(get_db)):
    user = get_user_from_request(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    return templates.TemplateResponse("reins.html", {
        "request": request,
        "user": user,
        "lines": LINES,
        "prefectures": PREFECTURES,
    })


@router.get("/announcements", response_class=HTMLResponse)
async def announcements(request: Request, db: Session = Depends(get_db)):
    user = get_user_from_request(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    items = db.query(Announcement).order_by(Announcement.date.desc()).all()
    return templates.TemplateResponse("announcements.html", {
        "request": request,
        "user": user,
        "announcements": items,
    })


@router.get("/feature-request", response_class=HTMLResponse)
async def feature_request(request: Request, db: Session = Depends(get_db)):
    user = get_user_from_request(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    return templates.TemplateResponse("feature_request.html", {
        "request": request,
        "user": user,
    })
