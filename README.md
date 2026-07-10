# Philadelphia Sheriff Sale — Auction Results

Interactive report comparing each property's **opening minimum ("initial ask")** to its
**final bid** at the Philadelphia County Sheriff real-property foreclosure auctions,
with final-bid data pulled live from each [Bid4Assets](https://www.bid4assets.com) listing.

**Live:** https://phila-sheriff-sales.vercel.app

## Repo layout

```
index.html            # the deployed report (generated — do not hand-edit)
scripts/
  scrape_bids.py      # fetch final bids for every auction ID -> data/*_with_final_bids.csv
  build_report.py     # CSV -> formatted .xlsx + printed summary
  build_html.py       # CSV -> index.html (the site)
data/
  <source>.xlsx                    # the county sale list (the input; drop new ones here)
  <source>_with_final_bids.csv     # scraped output
  <source>_with_final_bids.xlsx    # formatted spreadsheet
```

The scripts auto-detect the **newest** `*.xlsx` in `data/` (ignoring `*_with_final_bids`)
as the source, so updating for a new sale date is just: drop the new list in `data/` and re-run.

## Updating for a future sale

Requires Python 3 with `openpyxl` (`pip install openpyxl`).

```bash
# 1. Put the new county sale-list .xlsx in data/
# 2. Scrape final bids (~5 min, polite throttling), then rebuild the site + spreadsheet:
python scripts/scrape_bids.py
python scripts/build_report.py
python scripts/build_html.py

# 3. Commit & push — Vercel auto-deploys on push to the default branch.
git add -A
git commit -m "Update: <sale date>"
git push
```

## How it works

Bid4Assets renders the result **server-side** once a normal browser `User-Agent` is sent
(a plain bot fetch gets a 403, which is why the price looks "JavaScript-only"). Bank
repossessions show as `Final Bid` / `Status: Sold to Plaintiff`; third-party buyer wins
show as `Winning Bid` / `Status: Sold`. Postponed and stayed auctions carry no price.

## Notes

Educational summary of public auction results — not legal, financial, or investment advice.
Sheriff-sale properties sell as-is; verify occupancy, condition, taxes, and title/liens independently.
