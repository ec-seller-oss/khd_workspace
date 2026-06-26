from pydantic import BaseModel, EmailStr
from typing import Optional, List
from datetime import datetime


class UserCreate(BaseModel):
    email: EmailStr
    password: str


class UserLogin(BaseModel):
    email: EmailStr
    password: str


class UserOut(BaseModel):
    id: int
    email: str
    created_at: datetime

    class Config:
        from_attributes = True


class SearchPatternCreate(BaseModel):
    name: str = "パターン1"
    enabled: bool = True
    stations: List[dict] = []
    lines: List[str] = []
    build_cost_per_tsubo: float = 80.0
    room_area: float = 20.0
    demolition_unit: float = 3.0
    living_ratio: float = 70.0
    monthly_rent: float = 5.0
    yield_threshold: float = 7.0
    max_land_price: Optional[float] = None
    max_total_investment: Optional[float] = None
    min_area: Optional[float] = None
    max_area: Optional[float] = None
    notify_email: bool = True
    notify_announcement: bool = True
    notify_timing: str = "daily"


class SearchPatternOut(SearchPatternCreate):
    id: int
    user_id: int

    class Config:
        from_attributes = True


class PropertyOut(BaseModel):
    id: int
    source: str
    title: Optional[str]
    address: Optional[str]
    prefecture: Optional[str]
    price: Optional[float]
    area: Optional[float]
    far: Optional[float]
    bcr: Optional[float]
    nearest_station: Optional[str]
    walk_minutes: Optional[int]
    line: Optional[str]
    url: Optional[str]
    image_url: Optional[str]
    scraped_at: Optional[datetime]
    has_old_house: bool
    lat: Optional[float]
    lng: Optional[float]

    class Config:
        from_attributes = True


class SearchQuery(BaseModel):
    lines: List[str] = []
    prefectures: List[str] = []
    min_yield: Optional[float] = None
    max_land_price: Optional[float] = None
    max_total_investment: Optional[float] = None
    min_area: Optional[float] = None
    max_area: Optional[float] = None
    room_area: float = 20.0
    build_cost_per_tsubo: float = 80.0
    demolition_unit: float = 3.0
    living_ratio: float = 70.0
    monthly_rent: float = 6.5
    days: int = 7
    page: int = 1
    per_page: int = 50


class VolumeCheckInput(BaseModel):
    land_price: float             # 万円
    land_area: float              # ㎡
    far: float = 200.0            # 容積率 %
    bcr: float = 60.0             # 建蔽率 %
    monthly_rent: float = 5.0     # 万円/部屋
    room_area: float = 20.0       # ㎡/部屋
    rooms: Optional[int] = None   # 指定しない場合は自動計算
    build_cost_per_tsubo: float = 80.0  # 万円/坪
    living_ratio: float = 70.0    # 想定居住面積割合 %
    demolition_unit: float = 3.0  # 万円/㎡
    has_old_house: bool = False
    misc_cost: Optional[float] = None  # 諸費用 万円 (空欄で自動)
    max_floors: Optional[int] = None
    road_width: Optional[float] = None
    use_zone: Optional[str] = None


class VolumeCheckResult(BaseModel):
    rooms: int
    floor_area: float        # 延床面積 ㎡
    annual_rent: float       # 年間賃料 万円
    construction: float      # 建設費用 万円
    demolition: float        # 解体費 万円
    misc: float              # 諸費用 万円
    total_investment: float  # 総投資額 万円
    yield_pct: float         # 利回り %
    tsubo_price: float       # 坪単価 万円
    ichi_tsubo_price: float  # 一種単価 万円/坪 (容積率割戻)


class HistoryQuery(BaseModel):
    pattern_id: Optional[int] = None
    keyword: Optional[str] = None
    page: int = 1
    per_page: int = 50


class FeatureRequestCreate(BaseModel):
    content: str


class AnnouncementOut(BaseModel):
    id: int
    date: datetime
    title: Optional[str]
    content: Optional[str]

    class Config:
        from_attributes = True
