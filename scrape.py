import yaml

import excel
from scrapers import SCRAPER_CLASSES


def load_scraper_config(path="scraper_config.yaml"):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def process_scrape(search_terms=None):
    """
    Web scraping processing function.
    Returns a tuple: (success: bool, file_path: str or None, count: int)
    """
    if search_terms is None:
        print("No search terms provided. Skipping scrape.")
        return True, None, 0

    config = load_scraper_config()
    all_results = []

    for scraper_cls in SCRAPER_CLASSES:
        scraper = scraper_cls()
        try:
            scraper_config = config.get(scraper.name, {})
            scraper.columns = scraper_config.get("columns", {})

            print(f"[{scraper.name}] Authenticating...")
            if not scraper.authenticate():
                print(f"[{scraper.name}] Authentication failed, skipping")
                continue

            print(f"[{scraper.name}] Scraping {len(search_terms)} terms...")
            results = scraper.scrape(search_terms)
            all_results.extend(results)
            print(f"[{scraper.name}] Got {len(results)} results")

        except Exception as e:
            print(f"[{scraper.name}] Error: {e}")
            continue
        finally:
            scraper.close()

    if not all_results:
        print("No scrape results to export.")
        return True, None, 0

    created_path = excel.export_to_excel(all_results)
    return True, created_path, len(all_results)
