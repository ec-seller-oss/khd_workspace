"""初期データ投入スクリプト"""
import json
import os
import sys
from datetime import datetime, timedelta
import random

# パスを通す
sys.path.insert(0, os.path.dirname(__file__))

from database import engine, SessionLocal, init_db
from models import User, SearchPattern, Property, NotificationHistory, Station, Announcement
from auth import hash_password

def seed():
    init_db()
    db = SessionLocal()
    try:
        # ---- ユーザー ----
        demo_user = db.query(User).filter(User.email == "demo@landscanner.com").first()
        if not demo_user:
            demo_user = User(
                email="demo@landscanner.com",
                password_hash=hash_password("demo1234"),
            )
            db.add(demo_user)
            db.flush()
            print(f"[seed] デモユーザー作成: demo@landscanner.com / demo1234")
        else:
            print(f"[seed] デモユーザーは既に存在します")

        # ---- 検索パターン ----
        if db.query(SearchPattern).filter(SearchPattern.user_id == demo_user.id).count() == 0:
            patterns = [
                SearchPattern(
                    user_id=demo_user.id,
                    name="千葉・柏エリア",
                    enabled=True,
                    stations=[{"id": 1, "name": "柏", "line": "JR常磐線"}],
                    lines=["JR常磐線"],
                    build_cost_per_tsubo=80.0,
                    room_area=20.0,
                    demolition_unit=3.0,
                    living_ratio=70.0,
                    monthly_rent=5.5,
                    yield_threshold=7.0,
                    max_land_price=5000.0,
                    max_total_investment=15000.0,
                    min_area=100.0,
                    max_area=500.0,
                    notify_email=True,
                ),
                SearchPattern(
                    user_id=demo_user.id,
                    name="埼玉・大宮エリア",
                    enabled=True,
                    stations=[{"id": 12, "name": "大宮", "line": "JR京浜東北線"}],
                    lines=["JR京浜東北線", "東武東上線"],
                    build_cost_per_tsubo=82.0,
                    room_area=22.0,
                    demolition_unit=3.0,
                    living_ratio=70.0,
                    monthly_rent=5.5,
                    yield_threshold=7.5,
                    max_land_price=6000.0,
                    max_total_investment=18000.0,
                    notify_email=True,
                ),
            ]
            for p in patterns:
                db.add(p)
            db.flush()
            print(f"[seed] 検索パターン {len(patterns)}件 作成")

        # ---- 駅データ ----
        if db.query(Station).count() == 0:
            stations_path = os.path.join(os.path.dirname(__file__), "data", "stations.json")
            with open(stations_path, encoding="utf-8") as f:
                stations_data = json.load(f)
            for s in stations_data:
                station = Station(
                    id=s["id"],
                    name=s["name"],
                    line=s["line"],
                    prefecture=s["prefecture"],
                    lat=s["lat"],
                    lng=s["lng"],
                    monthly_rent=s["monthly_rent"],
                    service_count=s["service_count"],
                )
                db.add(station)
            db.flush()
            print(f"[seed] 駅データ {len(stations_data)}件 登録")
        else:
            print(f"[seed] 駅データは既に存在します")

        # ---- サンプル物件 ----
        if db.query(Property).count() == 0:
            sample_props = [
                {
                    "source": "suumo", "prefecture": "千葉県",
                    "title": "柏市柏 土地 全日照 角地",
                    "address": "千葉県柏市柏1丁目",
                    "price": 1980.0, "area": 165.3, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "柏駅", "walk_minutes": 8, "line": "JR常磐線",
                    "url": "https://suumo.jp/sample/001", "has_old_house": False,
                    "lat": 35.8682, "lng": 139.9756,
                    "days_ago": 1,
                },
                {
                    "source": "athome", "prefecture": "千葉県",
                    "title": "松戸市松戸 売地 整形地 閑静な住宅街",
                    "address": "千葉県松戸市松戸1丁目",
                    "price": 2480.0, "area": 198.5, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "松戸駅", "walk_minutes": 5, "line": "JR常磐線",
                    "url": "https://www.athome.co.jp/sample/002", "has_old_house": True,
                    "lat": 35.7879, "lng": 139.9027,
                    "days_ago": 1,
                },
                {
                    "source": "homes", "prefecture": "埼玉県",
                    "title": "川越市脇田町 土地 住宅用地",
                    "address": "埼玉県川越市脇田町",
                    "price": 1650.0, "area": 220.8, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "川越駅", "walk_minutes": 12, "line": "東武東上線",
                    "url": "https://www.homes.co.jp/sample/003", "has_old_house": False,
                    "lat": 35.9250, "lng": 139.4856,
                    "days_ago": 2,
                },
                {
                    "source": "suumo", "prefecture": "千葉県",
                    "title": "船橋市飯山満町 売地 南道路 建築条件なし",
                    "address": "千葉県船橋市飯山満町2丁目",
                    "price": 2850.0, "area": 172.4, "far": 150.0, "bcr": 50.0,
                    "nearest_station": "飯山満駅", "walk_minutes": 6, "line": "東葉高速線",
                    "url": "https://suumo.jp/sample/004", "has_old_house": True,
                    "lat": 35.7083, "lng": 140.0302,
                    "days_ago": 2,
                },
                {
                    "source": "athome", "prefecture": "神奈川県",
                    "title": "横浜市旭区中白根 土地 平坦地 全日照",
                    "address": "神奈川県横浜市旭区中白根",
                    "price": 3200.0, "area": 143.2, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "二俣川駅", "walk_minutes": 9, "line": "相鉄本線",
                    "url": "https://www.athome.co.jp/sample/005", "has_old_house": False,
                    "lat": 35.4706, "lng": 139.5282,
                    "days_ago": 3,
                },
                {
                    "source": "reins", "prefecture": "千葉県",
                    "title": "千葉市若葉区みつわ台 売地 住宅用途",
                    "address": "千葉県千葉市若葉区みつわ台",
                    "price": 1350.0, "area": 263.7, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "みつわ台駅", "walk_minutes": 3, "line": "千葉都市モノレール",
                    "url": "https://www.reins.or.jp/sample/006", "has_old_house": False,
                    "lat": 35.6448, "lng": 140.1411,
                    "days_ago": 3,
                },
                {
                    "source": "suumo", "prefecture": "埼玉県",
                    "title": "さいたま市見沼区東大宮 土地 全日照 角地",
                    "address": "埼玉県さいたま市見沼区東大宮",
                    "price": 2200.0, "area": 185.6, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "東大宮駅", "walk_minutes": 10, "line": "JR宇都宮線",
                    "url": "https://suumo.jp/sample/007", "has_old_house": False,
                    "lat": 35.9524, "lng": 139.6342,
                    "days_ago": 4,
                },
                {
                    "source": "homes", "prefecture": "千葉県",
                    "title": "柏市豊四季 売地 整形地 上水道接続可",
                    "address": "千葉県柏市豊四季",
                    "price": 1520.0, "area": 241.3, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "豊四季駅", "walk_minutes": 7, "line": "東武アーバンパークライン",
                    "url": "https://www.homes.co.jp/sample/008", "has_old_house": True,
                    "lat": 35.8430, "lng": 139.9553,
                    "days_ago": 4,
                },
                {
                    "source": "suumo", "prefecture": "神奈川県",
                    "title": "相模原市中央区相模原 土地 閑静住宅地",
                    "address": "神奈川県相模原市中央区相模原",
                    "price": 1880.0, "area": 195.0, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "相模原駅", "walk_minutes": 8, "line": "JR横浜線",
                    "url": "https://suumo.jp/sample/009", "has_old_house": False,
                    "lat": 35.5703, "lng": 139.3720,
                    "days_ago": 5,
                },
                {
                    "source": "athome", "prefecture": "埼玉県",
                    "title": "越谷市増林 売地 第一種住居地域 建築条件なし",
                    "address": "埼玉県越谷市増林",
                    "price": 1280.0, "area": 303.5, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "越谷駅", "walk_minutes": 15, "line": "東武スカイツリーライン",
                    "url": "https://www.athome.co.jp/sample/010", "has_old_house": True,
                    "lat": 35.8929, "lng": 139.7946,
                    "days_ago": 5,
                },
                {
                    "source": "suumo", "prefecture": "千葉県",
                    "title": "流山市おおたかの森 土地 駅近 生活利便性高",
                    "address": "千葉県流山市おおたかの森南",
                    "price": 3680.0, "area": 128.6, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "流山おおたかの森駅", "walk_minutes": 4, "line": "つくばエクスプレス",
                    "url": "https://suumo.jp/sample/011", "has_old_house": False,
                    "lat": 35.8636, "lng": 139.9320,
                    "days_ago": 6,
                },
                {
                    "source": "homes", "prefecture": "茨城県",
                    "title": "守谷市守谷 売地 TX守谷駅徒歩圏",
                    "address": "茨城県守谷市守谷",
                    "price": 1150.0, "area": 328.4, "far": 200.0, "bcr": 60.0,
                    "nearest_station": "守谷駅", "walk_minutes": 11, "line": "つくばエクスプレス",
                    "url": "https://www.homes.co.jp/sample/012", "has_old_house": True,
                    "lat": 35.9040, "lng": 139.9768,
                    "days_ago": 6,
                },
            ]

            properties = []
            for p in sample_props:
                days_ago = p.pop("days_ago")
                scraped_at = datetime.now() - timedelta(days=days_ago)
                prop = Property(**p, scraped_at=scraped_at)
                db.add(prop)
                properties.append(prop)
            db.flush()

            # 通知履歴を一部作成
            patterns = db.query(SearchPattern).filter(SearchPattern.user_id == demo_user.id).all()
            if patterns:
                from calculator import calc_yield_for_property
                for i, prop in enumerate(properties[:6]):
                    if prop.price and prop.area:
                        y = calc_yield_for_property(
                            land_price=prop.price,
                            land_area=prop.area,
                            far=prop.far or 200.0,
                        )
                        pattern = patterns[i % len(patterns)]
                        nh = NotificationHistory(
                            user_id=demo_user.id,
                            pattern_id=pattern.id,
                            property_id=prop.id,
                            yield_pct=round(y, 2),
                        )
                        db.add(nh)
            print(f"[seed] サンプル物件 {len(sample_props)}件 登録")
        else:
            print(f"[seed] 物件データは既に存在します")

        # ---- お知らせ ----
        if db.query(Announcement).count() == 0:
            announcements = [
                Announcement(
                    title="Land Scanner サービス開始のお知らせ",
                    content="首都圏投資用土地物件の自動収集・利回り計算サービスを開始しました。対応サイトは順次拡大予定です。ご不明点はお気軽にフィードバックからご連絡ください。",
                    date=datetime.now() - timedelta(days=1),
                ),
                Announcement(
                    title="対応路線を80路線に拡大しました",
                    content="つくばエクスプレス・東葉高速線・北総線など、千葉県主要路線を新たに追加しました。設定画面から沿線を選択してご利用ください。",
                    date=datetime.now() - timedelta(days=7),
                ),
                Announcement(
                    title="ボリュームチェック機能をアップデート",
                    content="ボリュームチェックにおいて、前面道路幅員・用途地域・階数制限を入力できるようになりました。より精度の高い試算が可能です。",
                    date=datetime.now() - timedelta(days=14),
                ),
                Announcement(
                    title="レインズ分析機能（β）公開",
                    content="限定機能として、REINSデータを活用した物件分析機能をβ公開しました。フィードバックをお待ちしています。",
                    date=datetime.now() - timedelta(days=21),
                ),
            ]
            for a in announcements:
                db.add(a)
            print(f"[seed] お知らせ {len(announcements)}件 登録")
        else:
            print(f"[seed] お知らせは既に存在します")

        db.commit()
        print("[seed] 完了!")

    except Exception as e:
        db.rollback()
        print(f"[seed] エラー: {e}")
        raise
    finally:
        db.close()


if __name__ == "__main__":
    seed()
