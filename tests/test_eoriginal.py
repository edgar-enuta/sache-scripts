import json
import os
import unittest
from unittest.mock import patch, MagicMock

FIXTURES_DIR = os.path.join(os.path.dirname(__file__), "fixtures")


def load_fixture(name):
    with open(os.path.join(FIXTURES_DIR, name), "r") as f:
        return json.load(f)


class TestFilterArticles(unittest.TestCase):
    """Pure logic tests for filter_articles — no mocking needed."""

    def setUp(self):
        from scrapers.eoriginal import EOriginalScraper
        self.filter = EOriginalScraper.filter_articles
        self.fixture = load_fixture("eoriginal_articles.json")
        self.articles = self.fixture["result"]["articles"]

    def test_typical_4_articles_returns_2_matches(self):
        result = self.filter(self.articles, "31372760")
        self.assertEqual(len(result), 2)
        brands = {a["brand"] for a in result}
        self.assertEqual(brands, {"VOLVO", "VOLVO - Stoc AD"})

    def test_filters_out_r_brand(self):
        result = self.filter(self.articles, "31372760")
        r_brands = [a for a in result if " - R" in a["brand"] and "Stoc AD" not in a["brand"]]
        self.assertEqual(len(r_brands), 0)

    def test_filters_out_aftermarket(self):
        result = self.filter(self.articles, "31372760")
        aftermarket = [a for a in result if a["brand"] == "AJUSA"]
        self.assertEqual(len(aftermarket), 0)

    def test_empty_input(self):
        result = self.filter({}, "31372760")
        self.assertEqual(result, [])

    def test_no_matches(self):
        result = self.filter(self.articles, "99999999")
        self.assertEqual(result, [])

    def test_only_regular_match(self):
        # Remove the Stoc AD article
        articles = {k: v for k, v in self.articles.items()
                     if "Stoc AD" not in v.get("brand", "")}
        result = self.filter(articles, "31372760")
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["brand"], "VOLVO")

    def test_ford_brand(self):
        articles = {
            "a_1": {
                "rID": "a_1", "articleName": "FORD 1234567", "anr": "1234567",
                "brand": "FORD", "type": "article", "sid": -10, "ikey": "1234567",
                "p_qty": 1, "atx_search_seq": 100,
            },
            "a_2": {
                "rID": "a_2", "articleName": "FORD - Stoc AD 1234567", "anr": "1234567",
                "brand": "FORD - Stoc AD", "type": "article", "sid": -20, "ikey": "1234567",
                "p_qty": 1, "atx_search_seq": 100,
            },
        }
        result = self.filter(articles, "1234567")
        self.assertEqual(len(result), 2)
        brands = {a["brand"] for a in result}
        self.assertEqual(brands, {"FORD", "FORD - Stoc AD"})


class TestExtractPrice(unittest.TestCase):
    """Pure logic tests for extract_price."""

    def setUp(self):
        from scrapers.eoriginal import EOriginalScraper
        self.extract = EOriginalScraper.extract_price
        self.prices_fixture = load_fixture("eoriginal_prices.json")

    def test_valid_response(self):
        price_data = self.prices_fixture["result"]["a_d611435f"]
        self.assertEqual(self.extract(price_data), 34.2)

    def test_missing_pret_key(self):
        self.assertIsNone(self.extract({}))

    def test_pret_false(self):
        self.assertIsNone(self.extract({"pret": False}))

    def test_missing_price_base(self):
        self.assertIsNone(self.extract({"pret": {"sets": []}}))

    def test_zero_price(self):
        data = {"pret": {"priceBase": {"pa": 0}}}
        self.assertEqual(self.extract(data), 0)

    def test_stoc_ad_price(self):
        price_data = self.prices_fixture["result"]["a_a3ef19e3"]
        self.assertEqual(self.extract(price_data), 38.23)


class TestCollateResults(unittest.TestCase):
    """Pure logic tests for collate_results."""

    def setUp(self):
        from scrapers.eoriginal import EOriginalScraper
        self.collate = EOriginalScraper.collate_results
        self.columns = {
            "search_term": "CodArticol",
            "price_regular": "PretNormal",
            "price_stoc_ad": "PretStocAD",
        }

    def test_two_matches(self):
        matched = [
            {"rID": "a_1", "brand": "VOLVO"},
            {"rID": "a_2", "brand": "VOLVO - Stoc AD"},
        ]
        prices = {
            "a_1": {"pret": {"priceBase": {"pa": 34.2}}},
            "a_2": {"pret": {"priceBase": {"pa": 38.23}}},
        }
        result = self.collate(matched, prices, "31372760", self.columns)
        self.assertEqual(result, {
            "CodArticol": "31372760",
            "PretNormal": 34.2,
            "PretStocAD": 38.23,
        })

    def test_one_match_no_stoc_ad(self):
        matched = [{"rID": "a_1", "brand": "VOLVO"}]
        prices = {"a_1": {"pret": {"priceBase": {"pa": 34.2}}}}
        result = self.collate(matched, prices, "31372760", self.columns)
        self.assertEqual(result, {
            "CodArticol": "31372760",
            "PretNormal": 34.2,
            "PretStocAD": None,
        })

    def test_zero_matches(self):
        result = self.collate([], {}, "31372760", self.columns)
        self.assertIsNone(result)

    def test_price_missing_for_one(self):
        matched = [
            {"rID": "a_1", "brand": "VOLVO"},
            {"rID": "a_2", "brand": "VOLVO - Stoc AD"},
        ]
        prices = {
            "a_1": {"pret": {"priceBase": {"pa": 34.2}}},
            "a_2": {"pret": False},
        }
        result = self.collate(matched, prices, "31372760", self.columns)
        self.assertEqual(result, {
            "CodArticol": "31372760",
            "PretNormal": 34.2,
            "PretStocAD": None,
        })


class TestScrapeIntegration(unittest.TestCase):
    """Integration tests — mock search_articles and get_prices."""

    def setUp(self):
        from scrapers.eoriginal import EOriginalScraper
        self.articles_fixture = load_fixture("eoriginal_articles.json")
        self.prices_fixture = load_fixture("eoriginal_prices.json")
        self.scraper = EOriginalScraper.__new__(EOriginalScraper)
        self.scraper.session = MagicMock()
        self.scraper.base_url = "https://www3.eoriginal.ro"
        self.scraper.columns = {
            "search_term": "CodArticol",
            "price_regular": "PretNormal",
            "price_stoc_ad": "PretStocAD",
        }

    @patch.object(
        __import__("scrapers.eoriginal", fromlist=["EOriginalScraper"]).EOriginalScraper,
        "get_prices",
    )
    @patch.object(
        __import__("scrapers.eoriginal", fromlist=["EOriginalScraper"]).EOriginalScraper,
        "search_articles",
    )
    def test_happy_path_two_terms(self, mock_search, mock_prices):
        mock_search.return_value = self.articles_fixture["result"]["articles"]
        mock_prices.return_value = self.prices_fixture["result"]

        results = self.scraper.scrape(["31372760", "31372760"])
        self.assertEqual(len(results), 2)
        self.assertEqual(results[0]["CodArticol"], "31372760")
        self.assertIsNotNone(results[0]["PretNormal"])
        self.assertIsNotNone(results[0]["PretStocAD"])

    @patch.object(
        __import__("scrapers.eoriginal", fromlist=["EOriginalScraper"]).EOriginalScraper,
        "get_prices",
    )
    @patch.object(
        __import__("scrapers.eoriginal", fromlist=["EOriginalScraper"]).EOriginalScraper,
        "search_articles",
    )
    def test_one_term_fails(self, mock_search, mock_prices):
        def side_effect(term):
            if term == "BAD":
                raise Exception("Network error")
            return self.articles_fixture["result"]["articles"]

        mock_search.side_effect = side_effect
        mock_prices.return_value = self.prices_fixture["result"]

        results = self.scraper.scrape(["BAD", "31372760"])
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["CodArticol"], "31372760")

    @patch.object(
        __import__("scrapers.eoriginal", fromlist=["EOriginalScraper"]).EOriginalScraper,
        "search_articles",
    )
    def test_no_articles(self, mock_search):
        mock_search.return_value = {}
        results = self.scraper.scrape(["31372760"])
        self.assertEqual(results, [])

    @patch.object(
        __import__("scrapers.eoriginal", fromlist=["EOriginalScraper"]).EOriginalScraper,
        "get_prices",
    )
    @patch.object(
        __import__("scrapers.eoriginal", fromlist=["EOriginalScraper"]).EOriginalScraper,
        "search_articles",
    )
    def test_articles_but_no_prices(self, mock_search, mock_prices):
        mock_search.return_value = self.articles_fixture["result"]["articles"]
        mock_prices.return_value = {
            "a_d611435f": {"pret": False},
            "a_a3ef19e3": {"pret": False},
        }
        results = self.scraper.scrape(["31372760"])
        self.assertEqual(len(results), 1)
        self.assertIsNone(results[0]["PretNormal"])
        self.assertIsNone(results[0]["PretStocAD"])


class TestAuthentication(unittest.TestCase):
    """Auth tests — mock HTTP responses."""

    def setUp(self):
        from scrapers.eoriginal import EOriginalScraper
        self.EOriginalScraper = EOriginalScraper

    @patch("scrapers.eoriginal.os.getenv")
    @patch("scrapers.eoriginal.requests.Session")
    def test_xsrf_token_parsed(self, mock_session_cls, mock_getenv):
        mock_getenv.side_effect = lambda k, d=None: {
            "EORIGINAL_USER": "user",
            "EORIGINAL_PASSWORD": "pass",
        }.get(k, d)

        session = MagicMock()
        mock_session_cls.return_value = session

        # GET / returns HTML with XSRF token
        login_page = MagicMock()
        login_page.text = '<html><input name="X-XSRF-TOKEN" value="abc123"></html>'

        # POST /login returns redirect header
        login_resp = MagicMock()
        login_resp.headers = {"Refresh": "0;url=https://www3.eoriginal.ro/device/check-device"}

        # GET /device/check-device returns page without save-device form (returning device)
        device_check = MagicMock()
        device_check.text = '<html><body>Welcome back</body></html>'

        session.get.side_effect = [login_page, device_check]
        session.post.return_value = login_resp
        session.cookies = MagicMock()
        session.cookies.get.return_value = "newtoken456"

        scraper = self.EOriginalScraper()
        result = scraper.authenticate()

        self.assertTrue(result)

    @patch("scrapers.eoriginal.os.getenv")
    @patch("scrapers.eoriginal.requests.Session")
    def test_missing_xsrf_token_raises(self, mock_session_cls, mock_getenv):
        mock_getenv.side_effect = lambda k, d=None: {
            "EORIGINAL_USER": "user",
            "EORIGINAL_PASSWORD": "pass",
        }.get(k, d)

        session = MagicMock()
        mock_session_cls.return_value = session

        login_page = MagicMock()
        login_page.text = "<html><body>No token here</body></html>"
        session.get.return_value = login_page

        scraper = self.EOriginalScraper()
        with self.assertRaises(Exception):
            scraper.authenticate()

    @patch("scrapers.eoriginal.os.getenv")
    def test_missing_credentials(self, mock_getenv):
        mock_getenv.return_value = None
        scraper = self.EOriginalScraper()
        result = scraper.authenticate()
        self.assertFalse(result)

    @patch("scrapers.eoriginal.os.getenv")
    @patch("scrapers.eoriginal.requests.Session")
    def test_device_check_new_device(self, mock_session_cls, mock_getenv):
        mock_getenv.side_effect = lambda k, d=None: {
            "EORIGINAL_USER": "user",
            "EORIGINAL_PASSWORD": "pass",
        }.get(k, d)

        session = MagicMock()
        mock_session_cls.return_value = session

        login_page = MagicMock()
        login_page.text = '<html><input name="X-XSRF-TOKEN" value="abc123"></html>'

        login_resp = MagicMock()
        login_resp.headers = {"Refresh": "0;url=https://www3.eoriginal.ro/device/check-device"}

        # Device check page has save-device form with token
        device_check = MagicMock()
        device_check.text = (
            '<html><form action="https://www3.eoriginal.ro/device/save-device">'
            '<input name="X-XSRF-TOKEN" value="devtoken789">'
            '<input name="deviceName">'
            '</form></html>'
        )

        device_save = MagicMock()
        device_save.headers = {"Refresh": "0;url=https://www3.eoriginal.ro/"}

        session.get.side_effect = [login_page, device_check]
        session.post.side_effect = [login_resp, device_save]
        session.cookies = MagicMock()
        session.cookies.get.return_value = "newtoken456"

        scraper = self.EOriginalScraper()
        result = scraper.authenticate()
        self.assertTrue(result)

        # Verify save-device was called
        self.assertEqual(session.post.call_count, 2)
        save_call = session.post.call_args_list[1]
        self.assertIn("device/save-device", save_call[0][0])


if __name__ == "__main__":
    unittest.main()
