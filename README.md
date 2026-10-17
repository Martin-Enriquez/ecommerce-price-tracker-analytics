# 🏷️ E-Commerce Price & Stock Tracker

![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![License](https://img.shields.io/badge/license-see%20LICENSE-lightgrey)

Python automation tool for e-commerce price monitoring. It extracts product prices and stock status, converts them to USD, keeps a historical record of every run, and exports a formatted Excel report with conditional formatting.

## Overview

The tracker scrapes a product catalog, normalizes prices to US dollars, compares them against the previous run, and produces an analyst-ready `.xlsx` report. It is designed for sellers and analysts operating in USD-based markets such as the United States and Ecuador.

## Features

- **Price & stock extraction:** Crawls catalog pages with pagination, request timeouts, and error handling.
- **Currency conversion:** Converts source prices (GBP) to USD using a configurable exchange rate.
- **Historical tracking:** Appends a timestamped snapshot to `price_history.csv` on every run and calculates price variation versus the previous run.
- **Excel reporting:** Styled headers, explicit USD formatting (`$`), auto-fitted columns, frozen header row, auto-filters, and conditional formatting (green for high savings, red for price increases).
- **Summary sheet:** Key stats (average, lowest and highest price, savings, price changes) and a price-distribution chart.

## Quick Start

```bash
git clone https://github.com/Martin-Enriquez/ecommerce-price-tracker-analytics.git
cd ecommerce-price-tracker-analytics
pip install -r requirements.txt
python main.py
```

A full run of the whole catalog (50 pages) takes about a minute. To try it quickly, limit the pages: `MAX_PAGES=3 python main.py`.

To override the default exchange rate:

```bash
# macOS / Linux
GBP_USD_RATE=1.34 python main.py
```

```powershell
# Windows (PowerShell)
$env:GBP_USD_RATE="1.34"; python main.py
```

## Configuration

| Setting | Where | Default | Description |
|---|---|---|---|
| `GBP_USD_RATE` | Environment variable | `1.34` | GBP → USD exchange rate. Update it before each real run. |
| `CATALOG_URL` | Environment variable | Whole books.toscrape.com catalog (~1,000 products, 50 pages) | Catalog page to start from. Pagination is followed automatically. |
| `MAX_PAGES` | Environment variable | all pages | Limit the number of pages crawled (e.g. `MAX_PAGES=5`). |
| `REQUEST_DELAY` | `main.py` | `0.5` s | Pause between requests, to be polite to the site. |
| `HIGHLIGHT_DISCOUNT` | `main.py` | `0.13` | Savings at or above this fraction are highlighted in green. |
| `SIM_MARKUP_RANGE` | `main.py` | `(0.10, 0.20)` | Range used to simulate the list price. |

## Output

**`price_monitor_report.xlsx`**

| Sheet | Contents |
|---|---|
| `Price Monitor` | Product, rating, simulated list price ($), offer price ($), savings ($ and %), stock status, variation vs. previous run (%), and a clickable product link. |
| `Summary` | Run date, source, currencies, exchange rate, product count, price and savings statistics, count of price increases/drops, and a price-distribution chart. |

**`price_history.csv`**

One row per product per run: timestamp, product, price in GBP, exchange rate, price in USD, and stock status.

## Project Structure

```
.
├── main.py            # Scraping, analysis, history, and Excel export
├── requirements.txt   # Python dependencies
├── README.md
└── LICENSE
```

## Scope & Limitations

- The data source is [books.toscrape.com](http://books.toscrape.com), a sandbox site built for scraping practice.
- **List prices are simulated** (10%–20% markup, stable per product) because the sandbox does not publish real list prices.
- Prices exclude taxes (VAT/sales tax), import duties, and shipping.
- The sandbox's prices never change, so the variation column will show 0% until a live source is used.

## Roadmap

- [ ] Add a second real source for competitor price comparison
- [ ] Country-specific tax columns (Ecuador VAT, US state sales tax)
- [ ] Scheduled runs (cron / Task Scheduler)
- [ ] Price-drop alerts by email

## Tech Stack

Python 3.10+ · Requests · BeautifulSoup4 · Pandas · OpenPyXL

## License

See the [LICENSE](LICENSE) file for details.

<img width="1193" height="208" alt="Captura de pantalla 2026-10-14 a la(s) 10 40 36 p  m" src="https://github.com/user-attachments/assets/be5f928f-cfb6-468c-811f-9a8e321bb5c5" />
