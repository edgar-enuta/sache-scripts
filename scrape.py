import os

import yaml
from dotenv import load_dotenv

import excel
from scrapers import SCRAPER_CLASSES

load_dotenv()


def load_scraper_config(path="scraper_config.yaml"):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def resolve_input_path():
    """Resolve the input file path from SCRAPER_PATH dir + SCRAPER_INPUT filename."""
    scraper_path = os.getenv("SCRAPER_PATH", "")
    scraper_input = os.getenv("SCRAPER_INPUT", "")
    if not scraper_input:
        raise ValueError("SCRAPER_INPUT not set in .env file.")
    base_dir = os.path.dirname(scraper_path)
    return os.path.join(base_dir, scraper_input) if base_dir else scraper_input


def run_scrapers(config, search_terms):
    """Run all scraper classes on the given search terms. Returns list of result dicts."""
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

    return all_results


def process_scrape(search_terms=None):
    """
    Web scraping processing function.
    Returns a tuple: (success: bool, file_path: str or None, count: int)

    If search_terms is None, reads from an input Excel file (SCRAPER_INPUT),
    runs scrapers, and merges results back with original columns.
    If search_terms is provided explicitly, uses the old export_to_excel path.
    """
    config = load_scraper_config()

    if search_terms is not None:
        # Explicit search terms: use old path
        all_results = run_scrapers(config, search_terms)
        if not all_results:
            print("No scrape results to export.")
            return True, None, 0
        created_path = excel.export_to_excel(all_results)
        return True, created_path, len(all_results)

    # Read input Excel
    id_column = config.get("input", {}).get("id_column")
    if not id_column:
        raise ValueError("input.id_column not set in scraper_config.yaml")

    input_path = resolve_input_path()
    print(f"Reading input file: {input_path}")
    search_terms, original_rows, original_columns = excel.read_excel_input(input_path, id_column)

    if not search_terms:
        print("No search terms found in input file.")
        return True, None, 0

    print(f"Found {len(search_terms)} search terms from input file")

    # Run scrapers
    all_results = run_scrapers(config, search_terms)

    # Generate output path and export merged Excel
    scraper_path = os.getenv("SCRAPER_PATH", "output/scraper_results.xlsx")
    output_path = excel.generate_timestamped_filename(scraper_path)
    created_path = excel.export_scraper_excel(
        original_rows, original_columns, all_results, id_column, output_path
    )
    return True, created_path, len(all_results)
