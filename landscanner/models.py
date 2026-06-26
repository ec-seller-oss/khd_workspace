from sqlalchemy import (
    Column, Integer, String, Float, Boolean, DateTime, Text, ForeignKey, JSON
)
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func
from database import Base


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(255), unique=True, nullable=False, index=True)
    password_hash = Column(String(255), nullable=False)
    is_admin = Column(Boolean, default=False)
    created_at = Column(DateTime, server_default=func.now())

    patterns = relationship("SearchPattern", back_populates="user")
    notifications = relationship("NotificationHistory", back_populates="user")


class SearchPattern(Base):
    __tablename__ = "search_patterns"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False)
    name = Column(String(100), default="パターン1")
    enabled = Column(Boolean, default=True)
    # 駅リスト (JSON: [{id, name, line}])
    stations = Column(JSON, default=list)
    # 沿線リスト (JSON: [line_name, ...])
    lines = Column(JSON, default=list)
    # 投資計算設定
    build_cost_per_tsubo = Column(Float, default=80.0)   # 万円/坪
    room_area = Column(Float, default=20.0)              # 部屋面積 ㎡
    demolition_unit = Column(Float, default=3.0)         # 解体費 万円/㎡
    living_ratio = Column(Float, default=70.0)           # 想定居住面積割合 %
    monthly_rent = Column(Float, default=5.0)            # 月額家賃 万円/部屋
    # フィルター設定
    yield_threshold = Column(Float, default=7.0)         # 利回り閾値 %
    max_land_price = Column(Float, nullable=True)        # 土地価格上限 万円
    max_total_investment = Column(Float, nullable=True)  # 総投資額上限 万円
    min_area = Column(Float, nullable=True)              # 最小面積 ㎡
    max_area = Column(Float, nullable=True)              # 最大面積 ㎡
    # 通知設定
    notify_email = Column(Boolean, default=True)
    notify_announcement = Column(Boolean, default=True)
    notify_timing = Column(String(20), default="daily")  # daily/realtime

    user = relationship("User", back_populates="patterns")
    notifications = relationship("NotificationHistory", back_populates="pattern")


class Property(Base):
    __tablename__ = "properties"

    id = Column(Integer, primary_key=True, index=True)
    source = Column(String(50), nullable=False)           # suumo/athome/homes/etc
    title = Column(String(500))
    address = Column(String(500))
    prefecture = Column(String(20))
    city = Column(String(100))
    price = Column(Float)                                 # 万円
    area = Column(Float)                                  # ㎡
    far = Column(Float, default=200.0)                    # 容積率 %
    bcr = Column(Float, default=60.0)                     # 建蔽率 %
    nearest_station = Column(String(100))
    walk_minutes = Column(Integer)
    line = Column(String(100))
    url = Column(String(1000))
    image_url = Column(String(1000))
    scraped_at = Column(DateTime, server_default=func.now())
    has_old_house = Column(Boolean, default=False)
    lat = Column(Float, nullable=True)
    lng = Column(Float, nullable=True)

    notifications = relationship("NotificationHistory", back_populates="property")


class NotificationHistory(Base):
    __tablename__ = "notification_history"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False)
    pattern_id = Column(Integer, ForeignKey("search_patterns.id"), nullable=False)
    property_id = Column(Integer, ForeignKey("properties.id"), nullable=False)
    sent_at = Column(DateTime, server_default=func.now())
    yield_pct = Column(Float)

    user = relationship("User", back_populates="notifications")
    pattern = relationship("SearchPattern", back_populates="notifications")
    property = relationship("Property", back_populates="notifications")


class Station(Base):
    __tablename__ = "stations"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False)
    line = Column(String(100), nullable=False)
    prefecture = Column(String(20))
    lat = Column(Float)
    lng = Column(Float)
    monthly_rent = Column(Float, default=5.0)  # 万円 (1K・20㎡・徒歩10分相場)
    service_count = Column(Integer, default=1)  # 対応サービス数


class Announcement(Base):
    __tablename__ = "announcements"

    id = Column(Integer, primary_key=True, index=True)
    date = Column(DateTime, server_default=func.now())
    content = Column(Text)
    title = Column(String(200))


class FeatureRequest(Base):
    __tablename__ = "feature_requests"

    id = Column(Integer, primary_key=True, index=True)
    content = Column(Text, nullable=False)
    created_at = Column(DateTime, server_default=func.now())
