# Dhoot Transmission IPO: how it was priced, sized and received

**Dhoot Transmission Limited** describes itself in its RHP as one of India's leading automotive electrical and electronics (E&E) companies. Its core product is wiring harnesses. It listed on NSE and BSE on 17 August 2026.

The IPO raised **₹3,066.89 crore**: a ₹1,400 crore fresh issue and a ₹1,666.89 crore offer for sale (OFS). It priced at **₹871, the top of a ₹829–871 band**. The book closed 74.2x subscribed and the stock closed its first day 36.3% higher.

INR crore unless stated (1 crore = 10 million). Every reported figure is sourced row by row in `data/*.csv`, with URL, page and date. The views in [How I would have priced it](#how-i-would-have-priced-it-analyst-view) are mine, not reported facts.

## Method

**Sources.** Only primary documents were used:

| Document | Date | Used for |
|---|---|---|
| [Red Herring Prospectus](https://www.sebi.gov.in/filings/public-issues/aug-2026/dhoot-transmission-limited-rhp_103333.html) (SEBI) | 3 Aug 2026 | Restated financials (pp.73–76, 457–459), capitalisation (p.479), indebtedness (p.480), Basis for Offer Price and peers (pp.163–165) |
| [Prospectus](https://nsearchives.nseindia.com/corporate/FP_INE01NH01023_13AUG2026.pdf) (NSE) | 13 Aug 2026 | Final share counts (p.96), shareholding before and after the offer (p.110), objects and net proceeds (pp.117–118), offer expenses (pp.159–160) |
| [Price band advertisement](https://nsearchives.nseindia.com/content/ipo/RATIOS_DHOOTTRANS.zip) (Financial Express p.18) | 4 Aug 2026 | Price band and its announcement date |
| [Anchor allocation report](https://nsearchives.nseindia.com/content/ipo/ANCHOR_DHOOTTRANS.zip) | 7 Aug 2026 | All 72 anchor allocations (`data/anchor_investors.csv`) |
| [Basis of allotment notice](https://nsearchives.nseindia.com/corporate/ADV_INE01NH01023_14AUG2026.pdf) | 14 Aug 2026 | Applications by category before rejections, allotment, anchor mix |
| NSE IPO bid book (`/api/ipo-detail`) | 12 Aug 2026 19:00 | Final consolidated NSE+BSE subscription by category |
| NSE end-of-day prices and index history | various | Dhoot and peer closes; Nifty 50 and Nifty Auto |
| Peers' FY26 results XBRL and 30-Jun-2026 shareholding patterns (NSE) | Apr–Jul 2026 | Peer financials and share counts |

**Definitions.** Dhoot and its peers are measured the same way:

- **EBITDA** is PBT before exceptional items, plus finance costs and D&A, less other income. This is the RHP's own definition (p.457).
- **Net debt** is borrowings less cash and other bank balances. Leases are excluded on both sides because the peers' XBRL does not tag lease liabilities. Dhoot has ₹76.9 crore of leases at FY26.
- **EV** is market cap plus net debt plus non-controlling interest. For Dhoot, net debt is post-issue: FY26 net debt less the ₹1,327 crore of net fresh proceeds.
- **P/E** for Dhoot is post-issue market cap divided by FY26 PAT attributable to owners. For peers it is the 4-Aug-2026 close divided by reported FY26 diluted EPS. Both measure the same thing, because no peer issued shares near the date.
- **P/B** for Dhoot uses post-money book: FY26 owners' equity plus net proceeds.

**Pricing date.** Peer prices are the NSE close on **4 August 2026**, the day the price band advertisement ran. The RHP's own peer table used BSE closes on 31 July 2026. Our EPS inputs reproduce the RHP's peer P/Es exactly at those prices (tested).

**Aftermarket.** Returns are measured from the ₹871 issue price, and from index closes on 12 August 2026, the bid close date. Checkpoints are calendar days after listing; if a checkpoint falls on a holiday, the next trading day's close is used.

**Reproducibility.** The project runs in this order:

1. `src/ipo_case.py` computes everything from the CSVs.
2. `build_workbook.py` writes the workbook, in which every output is an Excel formula.
3. `recalc.ps1` recalculates it in Excel.
4. `tests/` checks sourcing, recomputation, share reconciliation and Excel-vs-Python parity.
5. `build_deck.js` builds the deck from `output/results.json`.

## Key results

| | Reported or computed |
|---|---|
| Issue size | ₹3,066.89 cr: fresh ₹1,400 cr (45.6%), OFS ₹1,666.89 cr (54.4%) |
| Post-issue market cap at ₹871 | ₹17,816 cr on 204,548,057 shares |
| P/E, FY26 (post-issue shares) | 44.9x vs peer median 53.5x, a **16.1% discount** |
| EV/EBITDA, FY26 | 22.8x vs peer median 29.0x, a **21.4% discount** |
| EV/Sales, FY26 | 3.6x vs peer median 3.4x, a 7.0% premium |
| P/B (post-money) | 4.7x vs peer median 9.3x, a 49.0% discount |
| P/E, RHP basis (price / FY26 diluted EPS) | 35.7x, the headline in the RHP |
| Subscription, final consolidated book | Total 74.2x · QIB (ex-anchor) 212.9x · NII 51.9x · Retail 8.1x · Employee 8.1x |
| Anchor book | ₹918.3 cr to 36 investors; 60.0% of the QIB portion; 61% domestic mutual funds |
| Listing-day close | ₹1,187.05 (+36.3%) |
| Money left on the table | ₹1,113 cr (₹508 cr fresh issue, ₹605 cr OFS) |
| Return to +30 days (16 Sep) | +82.7% vs Nifty 50 −5.0% and Nifty Auto −8.4% |
| Return to 12 Nov (+87 days, latest) | +86.8% vs Nifty 50 +3.5% and Nifty Auto −3.9% |

## Peer selection

The RHP's Basis for Offer Price (p.165) names four listed peers. **All four are used and none is excluded:**

| Peer | Why it is comparable | Caveat |
|---|---|---|
| Motherson Sumi Wiring India (MSUMI) | Closest pure-play: wiring harnesses for Indian OEMs | Files standalone results only, so standalone is used |
| Uno Minda | Diversified auto electricals, lighting and switches; overlaps Dhoot's E&E products | Its EBITDA excludes JV profits, about ₹249 cr in FY26. That raises its EV/EBITDA |
| Minda Corporation | Wiring harnesses, mechatronics and plastics; the closest in size (₹17,266 cr market cap) | none |
| Sona BLW Precision Forgings | EV driveline and motors; growth-premium stock | Margin of 24.7% and EV/Sales of 10.9x make it an outlier on every multiple |

Sona is kept because the RHP chose it and the medians are robust to one outlier. Without Sona, the medians fall to 25.4x EV/EBITDA and 47.9x P/E. At ₹871, Dhoot would then sit at a 10% discount on EV/EBITDA and 6% on P/E: still cheaper, but by less.

## How I would have priced it (analyst view)

> This section is my judgement. It is not part of the reported record.

**Fundamental range: ₹806–909 per share.** I took the peer-median FY26 multiples, applied a 15–25% new-issue discount, and converted each to a price on the post-issue share count. I then averaged the two methods:

| Discount to peer median | EV/EBITDA method | P/E method | Average |
|---|---|---|---|
| 25% (low end) | 21.8x → ₹834 | 40.1x → ₹778 | ₹806 |
| 15% (high end) | 24.7x → ₹935 | 45.5x → ₹882 | ₹909 |

The 15–25% discount reflects three things:

- It is an IPO.
- The private-equity seller is exiting: BC Asia sells 16.0m shares.
- Margins have compressed. EBITDA margin fell from 18.3% (FY24) to 17.2% (FY25) to 15.7% (FY26), while revenue grew 31% in FY26.

**Recommended band: ₹860–910, priced at ₹910.** At ₹910, Dhoot would trade at 46.8x FY26 P/E and 23.9x EV/EBITDA. That is still a 13% and 18% discount to the peer medians, and inside the fundamental range.

What supports the top of my band:

- Revenue growth of 31% and an EBITDA margin above three of the four peers.
- Net cash of about ₹1,581 cr after the issue.
- A thin 17.2% free float.
- A book with QIB demand at 213x. Anything this oversubscribed would have cleared at a higher price.

What argues against going higher:

- The margin trend.
- An anchor book that was only 1.01x covered at ₹871.
- An acquisitions and GCP bucket of ₹410 cr (31% of net proceeds, 29% of gross fresh proceeds) that buyers cannot underwrite.

**What it would have changed.** At ₹910 the selling shareholders would have received about ₹75 cr more. The fresh issue is fixed in rupees, so the company would have issued fewer shares: about 15.4m instead of 16.1m. The listing pop would have shrunk from 36% to about 30%. I would not have pushed the price toward the ₹1,187 listing close. Trading at that level on a 17% float, in the first week, is not evidence that long-only investors would have bought the full ₹3,067 cr there.

## Limitations

- **Trailing multiples only.** No forecasts were used, and none are in the primary sources. Growth differences between Dhoot and its peers are judged qualitatively.
- **Mixed reporting bases.** Motherson Sumi Wiring is standalone. Sona BLW's XBRL has no "profit attributable to owners" tag, so peer P/E uses reported diluted EPS throughout.
- **Leases are excluded from net debt** on both sides. Including Dhoot's ₹76.9 cr of leases would lift its EV by 0.5%.
- **Dhoot's FY26 balance sheet holds ₹1,022.6 cr of undeployed Bain money** (BC Asia tranche 2, 20-Mar-2026; RHP p.459). That makes FY26 net debt negative and flatters P/B.
- **Two subscription measures exist.** The exchange measure (74.2x) excludes anchors from the denominator. The registrar's before-rejection figure (52.94x) includes them. NSE's own-platform figure (50.0x) excludes BSE bids. The deck uses the consolidated exchange figure and shows the registrar's for reconciliation.
- **Day 1 and Day 2 subscription are missing.** NSE now publishes only the final consolidated book, so these are marked "needs manual entry".
- **+90-day and +180-day returns have not elapsed.** The +90-day checkpoint is 15 Nov 2026 (a Sunday, so use 16 Nov). The +180-day checkpoint is 13 Feb 2027 (a Saturday, so use 15 Feb).

## Needs manual entry

| Item | Where to get it |
|---|---|
| Day 1 (10 Aug) and Day 2 (11 Aug) subscription by category | Exchange day-end bid notices (BSE "Cumulative Demand" or NSE bid details, dated) |
| +90-day close: Dhoot, Nifty 50, Nifty Auto (16 Nov 2026) | NSE historical data, after the date |
| +180-day close: Dhoot, Nifty 50, Nifty Auto (15 Feb 2027) | NSE historical data, after the date |
| Sona BLW FY26 profit attributable to owners | Sona BLW Q4 FY26 results PDF (NSE, 30 Apr 2026). Not used in any calculation |

## Five questions an interviewer would ask

**1. The RHP says 35.7x P/E. You say 44.9x. Which is right?**

Both, measured differently. The RHP divides ₹871 by FY26 diluted EPS of ₹24.40, which uses 162.6m weighted-average shares (p.458). That count barely includes the 22.2m shares issued to BC Asia on 20 March 2026, and none of the 16.1m IPO shares.

An investor buys into 204.5m post-issue shares. On that count, P/E is ₹17,816 cr ÷ ₹396.7 cr = 44.9x. The headline understates the multiple by about 9 turns. That is a routine but important adjustment when a company has issued shares close to the IPO.

**2. Was ₹1,113 crore really "left on the table"?**

Mechanically, yes: ₹316.05 listing gain × 35.2m offer shares. Of that, ₹508 cr is the company's (fresh shares) and ₹605 cr is the sellers' (OFS).

Economically, less. First-day trading on a 17% float, with 213x QIB demand that went largely unfilled, sets a marginal price, not a clearing price for ₹3,067 cr of stock.

The anchor book is the better test. Anchors were only 1.01x covered at ₹871, so the institutions who committed first did not signal much room above the band. A ₹910 price (+4.5%) was defensible. ₹1,187 was not.

**3. Why raise ₹1,400 crore of primary capital when the company already had net cash?**

At FY26, cash and bank balances were ₹1,096 cr against ₹841 cr of borrowings, but ₹1,022.6 cr of that cash was Bain's March 2026 tranche, not yet deployed.

The IPO repays ₹767 cr of debt, 91% of FY26 borrowings, and funds two plants (₹150 cr). That leaves ₹410 cr for unidentified acquisitions and GCP, 29% of gross fresh proceeds and under SEBI's 35% cap. Dhoot lists with about ₹1,581 cr of net cash.

The honest answer is that the primary raise buys balance-sheet capacity for M&A and makes the stock more palatable. The debt is cheap (7.75–8.90% term loans, RHP p.480), so the earnings uplift from repaying it is modest.

**4. Is Sona BLW a fair peer, and does it inflate the comparison?**

It is the RHP's choice, and the least comparable. Sona earns a 24.7% EBITDA margin and trades at 44.3x EV/EBITDA against 25.0–32.6x for the other three.

With Sona, the peer medians are 29.0x EV/EBITDA and 53.5x P/E. Without it, they are 25.4x and 47.9x. Dhoot's discount at ₹871 narrows from 21% to 10% on EV/EBITDA, and from 16% to 6% on P/E.

The closest pure-play is Motherson Sumi Wiring, at 25.0x and 42.3x. On that comparison ₹871 was a full price, not a cheap one. That is why my band stops at ₹910.

**5. The stock is up 87% while Nifty Auto is down 4%. Does that prove the IPO was underpriced?**

It proves demand exceeded supply at ₹871. It does not prove the fundamental value was wrong.

Most of the gain came in the first month: +82.7% by 16 September, against a falling market (Nifty 50 −5.0%, Nifty Auto −8.4%). Since then the stock has moved only 2.3%. That pattern looks like a one-off repricing of a small float to meet post-listing demand, not a continuous re-rating.

The real tests come later:

- The +90 and +180-day checkpoints.
- The end of the anchor lock-ins.
- The first post-IPO results, which will show whether FY26's margin decline (15.7%) has stopped.

None of these has happened yet.
