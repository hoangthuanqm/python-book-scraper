# Book Catalogue Scraper

A fast, reliable **Python web scraper** that collects a full product catalogue and exports it to
**CSV, JSON and a formatted Excel report**.

Target: [books.toscrape.com](https://books.toscrape.com), a public sandbox site built for scraping practice.
The same approach works for e-commerce stores, directories, real-estate listings, job boards and similar sites.

## Results

- **1,000 products in about 80 seconds** (12 parallel workers)
- Fields: title, category, price, star rating, stock quantity, UPC, product URL
- Excel report with styled headers, filters, frozen header row and a **summary sheet by category**

See [`sample-output/`](sample-output) for the real files produced by one run.

## Features

- Follows pagination automatically
- Parallel requests with `ThreadPoolExecutor`
- Automatic **retries with back-off** for timeouts and 429/5xx errors
- Polite crawling (small delay between listing pages)
- Errors on a single page are logged and skipped, so the run continues
- Filter by category and limit pages from the command line
- Clean, typed code (dataclasses, type hints) that is easy to adapt to a new site

## Getting started

Requires Python 3.10+.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate    macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python scraper.py                          # whole catalogue
python scraper.py --category Travel        # one category
python scraper.py --max-pages 3            # first 3 listing pages (60 books)
python scraper.py --workers 12 --out data  # more parallel requests, custom folder
```

Output files are written to `output/` by default:

| File | Description |
|---|---|
| `books.csv` | UTF-8 CSV, opens correctly in Excel |
| `books.json` | JSON array for APIs and databases |
| `books.xlsx` | Formatted workbook + "Summary by category" sheet |

## Sample log

```
14:45:41  INFO    Listing page 1: 20 books so far
14:45:42  INFO    Listing page 2: 40 books so far
...
14:59:12  INFO    Parsed 1000/1000 books
14:59:13  INFO    Done: 1000 books in 80.3s -> sample-output
```

---

Need data scraped from a website or a repetitive task automated? Contact me on Fiverr.
Only publicly available data is collected, and the target site's terms are respected.
