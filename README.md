# Dhoot Transmission IPO | ECM case study

**A reconstruction and critique of how Dhoot Transmission's ₹3,067 crore IPO (NSE/BSE, listed 17 August 2026) was priced, sized and received.** It is built only from primary sources: the RHP, the Prospectus, the price band advertisement, the anchor allocation report, the basis-of-allotment notice, and NSE/BSE data and filings.

**Result: priced at the top of the band, at a discount to peers, and still up 36% on day one.** At ₹871 the stock was valued at 44.9x FY26 P/E on post-issue shares and 22.8x EV/EBITDA. That is a 16% and 21% discount to the medians of the four RHP-listed peers (53.5x and 29.0x on 4-Aug-2026 closes). The book closed 74.2x subscribed, with QIBs at 212.9x. The listing-day close of ₹1,187.05 left about ₹1,113 cr on the table.

| Open | What it is |
|---|---|
| [Deck](output/Dhoot_IPO_Case.pptx) | 8 slides: transaction, valuation vs peers, proceeds, demand, aftermarket, my pricing view, limitations |
| [Workbook](output/Dhoot_IPO_Case.xlsx) | Terms, Financials, Peers, Valuation, Proceeds & Dilution, Demand, Aftermarket, Checks. Every output is a live formula |
| [Case write-up](docs/ipo_case.md) | Method, peer selection, my price range, limitations, needs-manual-entry list, five interview questions |
| [Data](data/) | One CSV per topic; every row carries source URL, page and date |

## Headline numbers

| | |
|---|---:|
| Issue size | ₹3,066.89 cr |
| Fresh issue / OFS | ₹1,400.00 cr / ₹1,666.89 cr (46% / 54%) |
| Price band → issue price | ₹829–871 → ₹871 |
| Post-issue market cap | ₹17,816 cr |
| P/E (post-issue) vs peer median | 44.9x vs 53.5x |
| EV/EBITDA vs peer median | 22.8x vs 29.0x |
| Subscription: overall / QIB | 74.2x / 212.9x |
| Listing-day close / gain | ₹1,187.05 / +36.3% |
| Money left on the table | ₹1,113 cr |
| Return at +30 days (Nifty 50 / Nifty Auto) | +82.7% (−5.0% / −8.4%) |
| Return at 12-Nov-2026, +87 days (Nifty 50 / Nifty Auto) | +86.8% (+3.5% / −3.9%) |
| **My recommended band (analyst view)** | **₹860–910, priced at ₹910** |

## Rebuild

```bash
pip install -r requirements.txt
npm install
python src/ipo_case.py
python build_workbook.py
powershell -ExecutionPolicy Bypass -File recalc.ps1
python -m pytest -q
node build_deck.js
```

In order, these steps:

1. Install the Python and Node dependencies.
2. Compute everything from the CSVs and write `output/results.json`.
3. Write the workbook, with every output as a formula.
4. Recalculate the workbook in desktop Excel so the values are cached.
5. Run 22 tests: every row sourced, multiples recompute, share counts reconcile, and Excel values match Python.
6. Build the deck from `output/results.json`.

## Layout

```
data/        ipo_terms, financials, peers, subscription, post_listing, anchor_investors (.csv)
src/         ipo_case.py: valuation, peers, proceeds and dilution, demand, aftermarket, analyst view
tests/       test_ipo_case.py
docs/        ipo_case.md
scripts/     apply_theme.js (writes the deck's theme colours)
output/      Dhoot_IPO_Case.xlsx, Dhoot_IPO_Case.pptx, results.json, cell_map.json
```
