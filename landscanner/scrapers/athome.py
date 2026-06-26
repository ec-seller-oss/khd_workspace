"""at home 土地物件スクレイパー（Playwright版）"""
from bs4 import BeautifulSoup
from typing import List, Optional
import re
import logging
import asyncio
from .base import ScrapedProperty

logger = logging.getLogger(__name__)


class AthomeScraper:
    source_name = "athome"
    base_url = "https://www.athome.co.jp"

    PREFECTURE_SLUGS = {
        "千葉": "chiba",
        "埼玉": "saitama",
        "神奈川": "kanagawa",
        "東京": "tokyo",
        "茨城": "ibaraki",
        "栃木": "tochigi",
        "群馬": "gunma",
    }

    def __init__(self):
        self._browser = None
        self._playwright = None

    async def _ensure_browser(self):
        if self._browser is None:
            from playwright.async_api import async_playwright
            from playwright_stealth import Stealth
            self._playwright = await async_playwright().start()
            Stealth().hook_playwright_context(self._playwright)
            self._browser = await self._playwright.chromium.launch(headless=True)

    async def close(self):
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()
        self._browser = None
        self._playwright = None

    async def scrape(self, prefecture: str = "千葉", pages: int = 2) -> List[ScrapedProperty]:
        slug = self.PREFECTURE_SLUGS.get(prefecture)
        if not slug:
            return []

        await self._ensure_browser()
        properties = []

        context = await self._browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="ja-JP",
        )
        page = await context.new_page()

        try:
            for page_num in range(1, pages + 1):
                url = f"{self.base_url}/tochi/{slug}/?BKFLG=1&sort=01"
                if page_num > 1:
                    url += f"&page={page_num}"

                try:
                    await page.goto(url, wait_until="domcontentloaded", timeout=20000)
                    await page.wait_for_selector(".bukken-item", timeout=10000)
                    html = await page.content()
                except Exception as e:
                    logger.warning(f"[athome] page {page_num} load failed: {e}")
                    break

                props = self._parse_list(html, prefecture)
                properties.extend(props)
                logger.info(f"[athome] {prefecture} p{page_num}: {len(props)}件")

                if len(props) == 0:
                    break
                await asyncio.sleep(5)
        finally:
            await context.close()

        return properties

    def _parse_list(self, html: str, prefecture: str) -> List[ScrapedProperty]:
        soup = BeautifulSoup(html, "html.parser")
        properties = []
        for card in soup.select(".bukken-item"):
            try:
                prop = self._parse_card(card, prefecture)
                if prop:
                    properties.append(prop)
            except Exception as e:
                logger.debug(f"[athome] parse error: {e}")
        return properties

    def _parse_card(self, card, prefecture: str) -> Optional[ScrapedProperty]:
        link_el = card.select_one("a[href]")
        url = ""
        if link_el and link_el.get("href"):
            href = link_el["href"]
            url = href if href.startswith("http") else self.base_url + href

        desc = card.select_one(".bukken-item__description")
        desc_text = desc.get_text(" ", strip=True) if desc else ""

        station = None
        walk = None
        m = re.search(r"「(.+?)」駅\s*徒歩(\d+)分", desc_text)
        if m:
            station = m.group(1)
            walk = int(m.group(2))

        price = None
        m = re.search(r"([\d,]+)万円", desc_text)
        if m:
            price = float(m.group(1).replace(",", ""))

        area = None
        m = re.search(r"([\d.]+)m²", desc_text)
        if m:
            area = float(m.group(1))

        kind_el = card.select_one(".bukken-item__type")
        title = kind_el.get_text(strip=True) if kind_el else "住宅用地"

        if not price and not area:
            return None

        return ScrapedProperty(
            source="athome",
            title=title,
            address=None,
            prefecture=prefecture,
            price=price,
            area=area,
            nearest_station=station,
            walk_minutes=walk,
            url=url,
        )
