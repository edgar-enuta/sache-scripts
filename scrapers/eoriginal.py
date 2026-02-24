import os
import re

import requests
from bs4 import BeautifulSoup

from scrapers.base import BaseScraper

BASE_URL = "https://www3.eoriginal.ro"

USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

XHR_HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Accept": "application/json, text/plain, */*",
    "Referer": f"{BASE_URL}/search/",
}


class EOriginalScraper(BaseScraper):
    name = "eoriginal"

    def __init__(self):
        self.session = None
        self.base_url = BASE_URL
        self.columns = {}

    def _xhr_headers(self):
        """Return headers for XHR/API calls (not for HTML page loads)."""
        xsrf = self.session.cookies.get("XSRF-TOKEN", "")
        return {**XHR_HEADERS, "X-XSRF-TOKEN": xsrf}

    def authenticate(self) -> bool:
        username = os.getenv("EORIGINAL_USER")
        password = os.getenv("EORIGINAL_PASSWORD")
        if not username or not password:
            print(f"[{self.name}] Missing EORIGINAL_USER or EORIGINAL_PASSWORD in .env")
            return False

        self.session = requests.Session()
        self.session.headers["User-Agent"] = USER_AGENT

        # Step 1: GET login page, parse XSRF token
        r = self.session.get(self.base_url)
        soup = BeautifulSoup(r.text, "html.parser")
        token_input = soup.find("input", {"name": "X-XSRF-TOKEN"})
        if not token_input:
            raise Exception(f"[{self.name}] Could not find XSRF token on login page")
        token = token_input["value"]

        # Step 2: POST /login
        self.session.post(f"{self.base_url}/login", data={
            "X-XSRF-TOKEN": token,
            "language": "ro",
            "username": username,
            "password": password,
            "default_cookie": "1",
            "analytics_cookie": "1",
        })

        # Step 3: Device check
        r = self.session.get(f"{self.base_url}/device/check-device")
        soup = BeautifulSoup(r.text, "html.parser")
        save_form = soup.find("form", action=re.compile(r"device/save-device"))
        if save_form:
            dev_token_input = save_form.find("input", {"name": "X-XSRF-TOKEN"})
            dev_token = dev_token_input["value"] if dev_token_input else token
            self.session.post(f"{self.base_url}/device/save-device", data={
                "X-XSRF-TOKEN": dev_token,
                "deviceName": "sacke",
            })

        print(f"[{self.name}] Authenticated successfully")
        return True

    def search_articles(self, term: str) -> dict:
        headers = self._xhr_headers()

        # Register search
        self.session.post(f"{self.base_url}/search/doSearch", json={
            "search": term,
            "searchType": 0,
            "section": "auto",
            "ref": "search",
            "autoSelection": 0,
        }, headers=headers)

        # Get articles
        r = self.session.post(f"{self.base_url}/search/getArticles", json={
            "search": term,
            "section": "auto",
            "ref": "search-page",
        }, headers=headers)
        data = r.json()
        return data.get("result", {}).get("articles", {})

    def get_prices(self, articles: dict) -> dict:
        headers = self._xhr_headers()

        ppd = {}
        for rid, art in articles.items():
            ppd[rid] = {
                "ART_SUP_ID": art["sid"],
                "ITEMKEY": art["ikey"],
                "QTY": art.get("p_qty", 1),
                "ATX_SEARCH_SEQ": art["atx_search_seq"],
                "type": art["type"],
            }

        r = self.session.post(f"{self.base_url}/getPrices", json={
            "articles": ppd,
            "qty": 1,
            "price": 1,
            "stoc": 1,
            "promo": 0,
            "source": "auto",
        }, headers=headers)
        return r.json().get("result", {})

    @staticmethod
    def filter_articles(articles: dict, search_term: str) -> list[dict]:
        matched = []
        for art in articles.values():
            # Must match the searched part number
            if art.get("anr") != search_term:
                continue
            brand = art.get("brand", "")
            # Exclude " - R" variants (but keep "Stoc AD")
            if " - R" in brand and "Stoc AD" not in brand:
                continue
            matched.append(art)
        return matched

    @staticmethod
    def extract_price(price_data: dict) -> float | None:
        try:
            pret = price_data.get("pret")
            if not pret or not isinstance(pret, dict):
                return None
            price_base = pret.get("priceBase")
            if not price_base or not isinstance(price_base, dict):
                return None
            return price_base.get("pa")
        except (AttributeError, TypeError):
            return None

    @staticmethod
    def collate_results(matched: list[dict], prices: dict,
                        search_term: str, columns: dict) -> dict | None:
        if not matched:
            return None

        regular_price = None
        stoc_ad_price = None

        for art in matched:
            rid = art["rID"]
            price_data = prices.get(rid, {})
            price = EOriginalScraper.extract_price(price_data)
            if "Stoc AD" in art.get("brand", ""):
                stoc_ad_price = price
            else:
                regular_price = price

        return {
            columns["search_term"]: search_term,
            columns["price_regular"]: regular_price,
            columns["price_stoc_ad"]: stoc_ad_price,
        }

    def scrape(self, search_terms: list[str]) -> list[dict]:
        results = []
        for term in search_terms:
            try:
                articles = self.search_articles(term)
                matched = self.filter_articles(articles, term)

                if not matched:
                    print(f"[{self.name}] No matching articles for '{term}'")
                    continue

                # Only fetch prices for matched articles
                matched_dict = {a["rID"]: a for a in matched}
                prices = self.get_prices(matched_dict)

                row = self.collate_results(matched, prices, term, self.columns)
                if row:
                    results.append(row)

            except Exception as e:
                print(f"[{self.name}] Error processing '{term}': {e}")
                continue

        return results

    def close(self) -> None:
        if self.session:
            self.session.close()
            self.session = None
