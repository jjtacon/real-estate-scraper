# Infocasas Real Estate Scraper

A Python script that scrapes real estate listings from Infocasas (Montevideo), calculates average price per square meter (USD/m^2) by neighborhood, and exports the data to Excel, highlighting properties priced below the neighborhood average.

## Features

- Scrapes price, location, total area (m^2), bedrooms, and direct URLs.
- Computes real-time average USD/m^2 per neighborhood.
- Identifies potential deals (properties listed at least 15% below the neighborhood average).
- Exports data to `infocasas_properties.xlsx` with two sheets: `All Properties` and `🔥 Opportunities`.

## Requirements

Install the required dependencies:

```bash
pip install -r requirements.txt
```
## Usage

Run the scraper directly from the terminal:
```bash
python realestatescraper.py
```
The Excel file will be generated in the same directory as the script.

## Tech Stack

    -Python 3

    -Requests & BeautifulSoup4

    -OpenPyXL
