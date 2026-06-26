"""SUUMO土地物件スクレイパー"""
from bs4 import BeautifulSoup
from typing import List, Optional
import re
import logging
from .base import BaseScraper, ScrapedProperty

logger = logging.getLogger(__name__)


class SuumoScraper(BaseScraper):
    source_name = "suumo"
    base_url = "https://suumo.jp"

    # 首都圏土地検索URL (例: 千葉県)
    SEARCH_URLS = {
        "千葉": "https://suumo.jp/joken/chiba/sk_23/",
        "埼玉": "https://suumo.jp/joken/saitama/sk_11/",
        "神奈川": "https://suumo.jp/joken/kanagawa/sk_14/",
        "東京": "https://suumo.jp/joken/tokyo/sk_13/",
    }

    async def scrape(self, prefecture: str = "千葉", pages: int = 2) -> List[ScrapedProperty]:
        properties = []
        url = self.SEARCH_URLS.get(prefecture)
        if not url:
            return properties

        for page in range(1, pages + 1):
            page_url = f"{url}?page={page}" if page > 1 else url
            html = await self.fetch(page_url)
            if not html:
                continue
            props = self._parse_list(html, prefecture)
            properties.extend(props)
            if len(props) == 0:
                break

        return properties

    def _parse_list(self, html: str, prefecture: str) -> List[ScrapedProperty]:
        soup = BeautifulSoup(html, "html.parser")
        properties = []

        # SUUMOの物件カード
        for card in soup.select(".property_unit-content, .cassette-detail"):
            try:
                prop = self._parse_card(card, prefecture)
                if prop:
                    properties.append(prop)
            except Exception as e:
                logger.debug(f"[suumo] parse error: {e}")

        return properties

    def _parse_card(self, card, prefecture: str) -> Optional[ScrapedProperty]:
        # タイトル
        title_el = card.select_one(".property_unit-title a, h2.cassette-detail__title a")
        title = title_el.get_text(strip=True) if title_el else None
        url = ""
        if title_el and title_el.get("href"):
            href = title_el["href"]
            url = href if href.startswith("http") else self.base_url + href

        # 価格
        price = None
        price_el = card.select_one(".dottable-value, .cassette-price--emphasis")
        if price_el:
            price_text = price_el.get_text(strip=True)
            m = re.search(r"([\d,]+)\s*万円", price_text)
            if m:
                price = float(m.group(1).replace(",", ""))

        # 面積
        area = None
        for el in card.select(".dottable-value, .cassette-detail__value"):
            text = el.get_text(strip=True)
            m = re.search(r"([\d.]+)\s*㎡", text)
            if m:
                area = float(m.group(1))
                break

        # 最寄駅
        station = None
        walk = None
        station_el = card.select_one(".property_unit-station, .cassette-detail__station")
        if station_el:
            text = station_el.get_text(strip=True)
            m = re.search(r"(.+?駅)\s*徒歩(\d+)分", text)
            if m:
                station = m.group(1)
                walk = int(m.group(2))

        # 住所
        address = None
        addr_el = card.select_one(".property_unit-address, .cassette-detail__address")
        if addr_el:
            address = addr_el.get_text(strip=True)

        if not price and not area:
            return None

        return ScrapedProperty(
            source="suumo",
            title=title,
            address=address,
            prefecture=prefecture,
            price=price,
            area=area,
            nearest_station=station,
            walk_minutes=walk,
            url=url,
        )
