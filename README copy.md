# Books to Scrape — Data Scraper

A Python scraper for [books.toscrape.com](https://books.toscrape.com). It walks every catalog page, collects structured book data, applies configurable filters, shows live progress in the terminal, and exports a professional two-sheet Excel workbook.

## Features

- Scrapes **title, price, star rating, availability, category, and product URL**
- Paginates through the full catalog automatically
- Configurable filters: minimum rating, maximum price, category, export limit
- Live terminal progress (`Scraped 45 / 1000 books`)
- Excel output with **Results** and **Scrape Summary** sheets
- Retries failed pages up to 3 times; missing fields are left blank instead of crashing

## Requirements

- Python 3.10+
- Dependencies listed in `requirements.txt`

## Installation

```bash
cd "Data Scraper"
python3 -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Configuration

Edit `config.py` to set default filters:

| Setting       | Example   | Description                          |
|---------------|-----------|--------------------------------------|
| `MIN_RATING`  | `4`       | Only books with this rating or higher |
| `MAX_PRICE`   | `30.0`    | Maximum price in GBP                 |
| `CATEGORY`    | `"Mystery"` | Partial category name match        |
| `MAX_RESULTS` | `100`     | Cap on exported rows                 |
| `OUTPUT_FILE` | `"books_output.xlsx"` | Output path              |

Set any filter to `None` (or use CLI flags below) to disable it.

## Usage

Run with `config.py` defaults:

```bash
python scraper.py
```

Override filters from the command line:

```bash
python scraper.py --min-rating 4 --max-price 25 --category Mystery --max-results 50 --output mystery_books.xlsx
```

Disable a filter with `0` (rating/price/results):

```bash
python scraper.py --min-rating 0 --max-price 0 --max-results 0
```

## Output

The scraper writes an `.xlsx` file with two sheets:

1. **Results** — filtered books with columns: Title, Price, Rating, Availability, Category, URL  
2. **Scrape Summary** — timestamp, books scanned, filters applied, export count, and any skipped pages/items

A sample file is included: `sample_books_output.xlsx`.

## Project structure

```
Data Scraper/
├── scraper.py              # Main scraper script
├── config.py               # Default filter settings
├── requirements.txt        # Python dependencies
├── sample_books_output.xlsx
└── README.md
```

## Notes

- books.toscrape.com is a public sandbox built for scraping practice.
- Category is read from each book’s detail page breadcrumb, so a full catalog run performs one extra request per book.
- For faster test runs, set `MAX_RESULTS` to a small number.
