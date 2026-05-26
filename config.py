"""
Filter and output settings for the Books to Scrape scraper.
Edit these values, or override them from the command line when running scraper.py.
"""

# Minimum star rating (1–5). Set to None to disable.
MIN_RATING = 4

# Maximum price in GBP. Set to None to disable.
MAX_PRICE = 30.0

# Category name (case-insensitive partial match), e.g. "Mystery", "Science".
# Set to None to include all categories.
CATEGORY = None

# Maximum number of matching books to export. Set to None for no limit.
MAX_RESULTS = 100

# Base URL for the bookstore
BASE_URL = "https://books.toscrape.com/"

# Output Excel file path
OUTPUT_FILE = "books_output.xlsx"

# HTTP request timeout (seconds)
REQUEST_TIMEOUT = 15

# Retries per page before skipping
MAX_RETRIES = 3
