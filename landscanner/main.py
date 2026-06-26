"""Land Scanner - FastAPI アプリケーション エントリーポイント"""
import asyncio
import os
import logging
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

load_dotenv()

# ロガー設定
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

# DB初期化
from database import init_db
init_db()

# FastAPI アプリ作成
app = FastAPI(
    title="Land Scanner",
    description="首都圏投資用土地物件 自動収集・利回り計算サービス",
    version="1.0.0",
)

# CORSミドルウェア（開発用）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 静的ファイル
BASE_DIR = os.path.dirname(__file__)
static_dir = os.path.join(BASE_DIR, "static")
if os.path.exists(static_dir):
    app.mount("/static", StaticFiles(directory=static_dir), name="static")

# ルーター登録
from routers import pages, api
from routers.auth import router as auth_router

app.include_router(auth_router)
app.include_router(pages.router)
app.include_router(api.router)

# スケジューラ（APScheduler）
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger

scheduler = AsyncIOScheduler()


async def scheduled_scrape():
    """毎日AM6時に自動スクレイピング → リアルタイム通知ディスパッチ"""
    logger.info("[scheduler] スクレイピング開始")
    try:
        from database import SessionLocal
        from models import Property
        from scrapers import SuumoScraper, AthomeScraper
        from notifier import dispatch_notifications

        db = SessionLocal()
        scrapers = [AthomeScraper()]  # SuumoScraper は別途対応
        new_props: list[Property] = []
        prefectures = ["千葉", "埼玉", "神奈川", "東京"]

        for scraper in scrapers:
            for pref in prefectures:
                try:
                    props = await scraper.scrape(prefecture=pref, pages=2)
                    await asyncio.sleep(10)  # WAFレート制限回避
                    for sp in props:
                        prop = Property(
                            source=sp.source,
                            title=sp.title,
                            address=sp.address,
                            prefecture=sp.prefecture or pref + "県",
                            price=sp.price,
                            area=sp.area,
                            far=sp.far,
                            bcr=sp.bcr,
                            nearest_station=sp.nearest_station,
                            walk_minutes=sp.walk_minutes,
                            line=sp.line,
                            url=sp.url,
                            image_url=sp.image_url,
                            has_old_house=sp.has_old_house,
                            lat=sp.lat,
                            lng=sp.lng,
                        )
                        db.add(prop)
                        new_props.append(prop)
                except Exception as e:
                    logger.warning(f"[scheduler] scrape error {scraper.source_name}/{pref}: {e}")
            await scraper.close()

        db.commit()
        # IDを確定させてから通知判定
        for p in new_props:
            db.refresh(p)
        logger.info(f"[scheduler] スクレイピング完了: {len(new_props)}件")

        # リアルタイム通知（notify_timing='realtime' のユーザーに即時メール）
        dispatch_notifications(db, new_props)
        db.close()
    except Exception as e:
        logger.error(f"[scheduler] エラー: {e}")


async def scheduled_daily_digest():
    """毎日AM7時に日次ダイジェストメール送信"""
    logger.info("[scheduler] 日次ダイジェスト開始")
    try:
        from database import SessionLocal
        from notifier import send_daily_digest
        db = SessionLocal()
        send_daily_digest(db)
        db.close()
    except Exception as e:
        logger.error(f"[scheduler] ダイジェストエラー: {e}")


@app.on_event("startup")
async def startup_event():
    logger.info("Land Scanner 起動")
    # 毎日AM6時にスクレイピング
    scheduler.add_job(
        scheduled_scrape,
        CronTrigger(hour=6, minute=0),
        id="daily_scrape",
        replace_existing=True,
    )
    scheduler.add_job(
        scheduled_daily_digest,
        CronTrigger(hour=7, minute=0),
        id="daily_digest",
        replace_existing=True,
    )
    scheduler.start()
    logger.info("[scheduler] スケジューラ起動完了 (scrape=6:00 / digest=7:00)")


@app.on_event("shutdown")
async def shutdown_event():
    scheduler.shutdown()
    logger.info("Land Scanner 停止")


@app.get("/health")
async def health():
    return {"status": "ok", "service": "Land Scanner"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
    )
