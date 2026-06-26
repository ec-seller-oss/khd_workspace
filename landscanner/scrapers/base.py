"""スクレイパー基底クラス"""
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional, List
from datetime import datetime
import httpx
import logging

logger = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}


@dataclass
class ScrapedProperty:
    source: str
    title: Optional[str] = None
    address: Optional[str] = None
    prefecture: Optional[str] = None
    city: Optional[str] = None
    price: Optional[float] = None          # 万円
    area: Optional[float] = None           # ㎡
    far: Optional[float] = 200.0           # 容積率 %
    bcr: Optional[float] = 60.0            # 建蔽率 %
    nearest_station: Optional[str] = None
    walk_minutes: Optional[int] = None
    line: Optional[str] = None
    url: Optional[str] = None
    image_url: Optional[str] = None
    has_old_house: bool = False
    lat: Optional[float] = None
    lng: Optional[float] = None


class BaseScraper(ABC):
    source_name: str = "unknown"
    base_url: str = ""

    def __init__(self):
        self.client = httpx.AsyncClient(
            headers=HEADERS,
            timeout=30.0,
            follow_redirects=True,
        )

    async def close(self):
        await self.client.aclose()

    @abstractmethod
    async def scrape(self, **kwargs) -> List[ScrapedProperty]:
        pass

    async def fetch(self, url: str, retries: int = 2, **kwargs) -> Optional[str]:
        import asyncio
        for attempt in range(retries + 1):
            try:
                resp = await self.client.get(url, **kwargs)
                if resp.status_code == 503 and attempt < retries:
                    await asyncio.sleep(3 * (attempt + 1))
                    continue
                resp.raise_for_status()
                return resp.text
            except Exception as e:
                if attempt < retries:
                    await asyncio.sleep(2)
                    continue
                logger.warning(f"[{self.source_name}] fetch failed: {url} - {e}")
                return None
