# 🏷️ E-Commerce Price & Stock Tracker with Analytics Export

A Python automation tool designed for e-commerce price monitoring, competitor analysis, and automated Excel reporting with conditional formatting.

## 🚀 Key Features
- **Price Tracking:** Extracts current pricing, list pricing, and stock status.
- **Analytics & Calculations:** Automatically calculates percentage savings and price gaps using Pandas.
- **Automated Excel Formatting:** Formats currency columns, applies conditional formatting rules (highlighting top discounts in green), freezes header rows, and auto-adjusts column widths with OpenPyXL.

## 🛠️ Tech Stack
Python 3.10+ · Requests · BeautifulSoup4 · Pandas · OpenPyXL

## 📦 Installation & Usage

```bash
# 1. Clone the repository
git clone [https://github.com/Martin-Enriquez/ecommerce-price-tracker-analytics.git](https://github.com/Martin-Enriquez/ecommerce-price-tracker-analytics.git)
cd ecommerce-price-tracker-analytics

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the monitor
python main.py