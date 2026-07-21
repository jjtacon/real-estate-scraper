import re
from collections import defaultdict
from pathlib import Path
import openpyxl
import requests
from bs4 import BeautifulSoup

# Configuration
BASE_URL = "https://www.infocasas.com.uy"
SEARCH_URL = f"{BASE_URL}/venta/inmuebles/montevideo"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# Dynamically set the output path to the script's directory
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_FILE = SCRIPT_DIR / "infocasas_properties.xlsx"


def clean_number(text: str) -> int | None:
    """Extracts digits from a string and converts to integer."""
    if not text:
        return None
    digits = re.sub(r"\D", "", text)
    return int(digits) if digits else None


def fetch_properties(url: str) -> list[dict]:
    """Fetches and parses real estate property listings from the target URL."""
    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print(f"[Error] Failed to retrieve data from {url}: {err}")
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    property_cards = soup.find_all("div", class_="listingCard")
    parsed_properties = []

    for card in property_cards:
        # 1. Price and Location
        price_tag = card.find("p", class_="main-price")
        location_tag = card.find("strong", class_="lc-location")

        price_str = price_tag.text.strip() if price_tag else "N/A"
        location = location_tag.text.strip() if location_tag else "N/A"

        # Ignore range prices (e.g. "Desde U$S...") for clean numerical analysis
        is_price_range = "desde" in price_str.lower()
        price_num = None if is_price_range else clean_number(price_str)

        # 2. Extract Direct Link
        link_tag = card.find("a", class_="lc-data")
        href = link_tag.get("href") if link_tag else None
        link = f"{BASE_URL}{href}" if href else "N/A"

        # 3. Extract Area (m²) and Bedrooms
        sqm_num = None
        bedrooms_num = None

        tags = card.find_all("span", class_="lc-typologyTag__item")
        for tag in tags:
            tag_text = tag.text.strip()
            if "m²" in tag_text:
                sqm_num = clean_number(tag_text)
            elif "Dorm" in tag_text:
                bedrooms_num = clean_number(tag_text)

        # 4. Calculate USD / m² ratio
        price_per_sqm = None
        if price_num and sqm_num and sqm_num > 0:
            if "USD" in price_str.upper() or "U$S" in price_str.upper():
                price_per_sqm = price_num / sqm_num

        parsed_properties.append({
            "price_str": price_str,
            "location": location,
            "sqm": sqm_num if sqm_num else "N/A",
            "bedrooms": bedrooms_num if bedrooms_num else "N/A",
            "price_per_sqm": price_per_sqm,
            "link": link,
        })

    return parsed_properties


def export_to_excel(properties: list[dict], file_path: Path) -> None:
    """Calculates market averages and exports data to an Excel workbook."""
    if not properties:
        print("[Warning] No properties found to export.")
        return

    # Calculate Neighborhood Averages
    sqm_by_location = defaultdict(list)
    for prop in properties:
        if prop["price_per_sqm"] and prop["location"] != "N/A":
            sqm_by_location[prop["location"]].append(prop["price_per_sqm"])

    neighborhood_averages = {
        loc: sum(vals) / len(vals)
        for loc, vals in sqm_by_location.items()
        if vals
    }

    # Print summary to console
    print("\n--- REAL MARKET AVERAGES (USD / m²) ---")
    for loc, avg in neighborhood_averages.items():
        print(f"📍 {loc}: USD {avg:,.0f} / m²")
    print("-" * 50 + "\n")

    # Create Excel Workbook
    wb = openpyxl.Workbook()

    # Sheet 1: All Properties
    ws_all = wb.active
    ws_all.title = "All Properties"
    ws_all.append([
        "Price", "Location", "Area (m²)", "Bedrooms", "USD / m²", "Link"
    ])

    for p in properties:
        sqm_formatted = (
            f"{p['price_per_sqm']:,.0f}" if p["price_per_sqm"] else "N/A"
        )
        ws_all.append([
            p["price_str"],
            p["location"],
            p["sqm"],
            p["bedrooms"],
            sqm_formatted,
            p["link"],
        ])

    # Sheet 2: Deals Below Market Average (15% discount threshold)
    ws_deals = wb.create_sheet(title="Opportunities")
    ws_deals.append([
        "Price",
        "Location",
        "Area (m²)",
        "Bedrooms",
        "Property USD/m²",
        "Neighborhood Average",
        "Link",
    ])

    for p in properties:
        loc = p["location"]
        p_sqm = p["price_per_sqm"]

        if p_sqm and loc in neighborhood_averages:
            avg_sqm = neighborhood_averages[loc]
            if p_sqm < (avg_sqm * 0.85):  # 15% under market average
                ws_deals.append([
                    p["price_str"],
                    loc,
                    p["sqm"],
                    p["bedrooms"],
                    f"{p_sqm:,.0f}",
                    f"{avg_sqm:,.0f}",
                    p["link"],
                ])
                print(
                    f"Opportunity in {loc}! USD {p_sqm:,.0f}/m² (Average:"
                    f" USD {avg_sqm:,.0f}/m²)"
                )

    wb.save(file_path)
    print(f"\n✅ Excel successfully created at: {file_path}")


if __name__ == "__main__":
    print("Starting Infocasas scraper...")
    data = fetch_properties(SEARCH_URL)
    export_to_excel(data, OUTPUT_FILE)
