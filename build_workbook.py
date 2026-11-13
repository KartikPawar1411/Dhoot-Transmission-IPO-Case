"""Build output/Dhoot_IPO_Case.xlsx.

Input sheets (Terms, Financials, Peers inputs, Demand inputs, Aftermarket inputs) hold
the CSV values with their sources. Every analytical number is an Excel formula.
Run recalc.ps1 afterwards so Excel calculates and caches the results.
"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT / "src"))
import ipo_case as m  # noqa: E402

OUT = ROOT / "output" / "Dhoot_IPO_Case.xlsx"
NAVY = "1F3864"
BLUE = Font(name="Calibri", size=10, color="0000FF")
BLACK = Font(name="Calibri", size=10)
BOLD = Font(name="Calibri", size=10, bold=True)
WHITE_B = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
TITLE = Font(name="Calibri", size=14, bold=True, color=NAVY)
NOTE = Font(name="Calibri", size=9, italic=True, color="595959")
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
VIEW_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(bottom=Side(style="thin", color="BFBFBF"))
F_CR = '#,##0.0;(#,##0.0);"-"'
F_MN = '#,##0.00;(#,##0.00);"-"'
F_SH = '#,##0'
F_X = '0.0"x"'
F_PCT = '0.0%;(0.0%)'
F_PX = '#,##0.00'

wb = Workbook()
cells = {}  # python-key -> "Sheet!A1" for tests


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def sheet(title, heading, sub=None):
    ws = wb.create_sheet(title)
    ws["A1"] = heading
    ws["A1"].font = TITLE
    if sub:
        ws["A2"] = sub
        ws["A2"].font = NOTE
    ws.sheet_view.showGridLines = False
    return ws


def header(ws, row, labels, col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font, c.fill = WHITE_B, HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def put(ws, r, c, v, fmt=None, font=None, key=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font or (BLUE if not (isinstance(v, str) and v.startswith("=")) else BLACK)
    if fmt:
        cell.number_format = fmt
    if key:
        cells[key] = f"'{ws.title}'!{cell.coordinate}"
    return cell


def widths(ws, w):
    for i, x in enumerate(w, start=1):
        ws.column_dimensions[get_column_letter(i)].width = x


def absref(ws, cell):
    col = "".join(ch for ch in cell if ch.isalpha())
    row = "".join(ch for ch in cell if ch.isdigit())
    return f"'{ws.title}'!${col}${row}"


wb.remove(wb.active)

# ------------------------------------------------------------------ Terms
ws = sheet("Terms", "Dhoot Transmission IPO: offer terms", "Blue = input from primary source (see source columns). Amounts INR million unless stated.")
header(ws, 4, ["Field", "Value", "Unit", "Source URL", "Page", "Source date", "Notes"])
r = 5
for row in m.read("ipo_terms.csv"):
    v = m._num(row["value"])
    fmt = F_SH if row["unit"] in ("shares", "count") else (F_MN if row["unit"] == "INR mn" else (F_PX if row["unit"].startswith("INR/") else None))
    put(ws, r, 1, row["field"], font=BLACK)
    put(ws, r, 2, v, fmt)
    for j, k in enumerate(["unit", "source_url", "source_page", "source_date", "notes"], start=3):
        put(ws, r, j, row[k], font=NOTE if k == "notes" else BLACK)
    if isinstance(v, float):
        name("t_" + row["field"], absref(ws, f"B{r}"))
    r += 1
widths(ws, [32, 18, 10, 60, 18, 12, 60])
ws.freeze_panes = "B5"

# ------------------------------------------------------------------ Financials
ws = sheet("Financials", "Restated consolidated financials (RHP)", "INR million. Source: RHP dated 2026-08-03; page per row. No stub period (restated to 2026-03-31).")
header(ws, 4, ["Metric", "FY2024", "FY2025", "FY2026", "Unit", "Source URL", "Page", "Source date", "Notes"])
fins = m.financials()
meta = {}
for row in m.read("financials.csv"):
    meta.setdefault(row["metric"], row)
r = 5
fin_row = {}
for metric, vals in fins.items():
    if metric == "stub_period":
        continue
    put(ws, r, 1, metric, font=BLACK)
    for j, fy in enumerate(["FY2024", "FY2025", "FY2026"], start=2):
        put(ws, r, j, vals[fy], F_PX if metric == "diluted_eps" else F_MN)
        name(f"f_{metric}_{fy}", absref(ws, f"{get_column_letter(j)}{r}"))
    md = meta[metric]
    for j, k in enumerate(["unit", "source_url", "source_page", "source_date", "notes"], start=5):
        put(ws, r, j, md[k], font=NOTE if k == "notes" else BLACK)
    fin_row[metric] = r
    r += 1
r += 1
put(ws, r, 1, "Derived (formulas)", font=BOLD)
r += 1
derived = [
    ("Revenue growth", lambda c, p: f"={c}{fin_row['revenue_from_operations']}/{p}{fin_row['revenue_from_operations']}-1", F_PCT),
    ("EBITDA growth", lambda c, p: f"={c}{fin_row['ebitda']}/{p}{fin_row['ebitda']}-1", F_PCT),
    ("PAT growth", lambda c, p: f"={c}{fin_row['pat']}/{p}{fin_row['pat']}-1", F_PCT),
]
for lab, fn, fmt in derived:
    put(ws, r, 1, lab, font=BLACK)
    for c, p in (("C", "B"), ("D", "C")):
        put(ws, r, ord(c) - 64, fn(c, p), fmt)
    r += 1
for lab, num, fmt in [("EBITDA margin", "ebitda", F_PCT), ("PAT margin", "pat", F_PCT)]:
    put(ws, r, 1, lab, font=BLACK)
    for c in "BCD":
        put(ws, r, ord(c) - 64, f"={c}{fin_row[num]}/{c}{fin_row['revenue_from_operations']}", fmt, key=f"fin_{num}_margin_{c}")
    r += 1
put(ws, r, 1, "Net debt, borrowings basis (ex-leases)", font=BLACK)
for c in "BCD":
    put(ws, r, ord(c) - 64, f"={c}{fin_row['total_borrowings']}-{c}{fin_row['cash_and_equivalents']}-{c}{fin_row['other_bank_balances']}", F_MN)
r += 1
put(ws, r, 1, "Check: RHP net debt recomputed", font=BLACK)
for c in "BCD":
    put(ws, r, ord(c) - 64, f"={c}{fin_row['total_borrowings']}+{c}{fin_row['lease_liabilities']}-{c}{fin_row['cash_and_equivalents']}-{c}{fin_row['other_bank_balances']}-{c}{fin_row['net_debt_rhp']}", F_MN, key=f"fin_netdebt_check_{c}")
r += 1
put(ws, r, 1, "Check: EBITDA recomputed from components", font=BLACK)
for c in "BCD":
    put(ws, r, ord(c) - 64, f"={c}{fin_row['pbt_before_exceptional']}+{c}{fin_row['finance_costs']}+{c}{fin_row['depreciation_amortisation']}-{c}{fin_row['other_income']}-{c}{fin_row['ebitda']}", F_MN, key=f"fin_ebitda_check_{c}")
widths(ws, [38, 13, 13, 13, 10, 60, 10, 12, 60])
ws.freeze_panes = "B5"

# ------------------------------------------------------------------ Peers
ws = sheet("Peers", "RHP-listed peers at price band announcement (NSE close, 4 Aug 2026)",
           "Inputs: NSE close 2026-08-04; shares from 30-Jun-2026 shareholding pattern; FY2026 results XBRL (INR million). All four RHP peers included.")
P = m.peers()
syms = m.PEER_ORDER
header(ws, 4, ["Input"] + syms + ["Unit"])
put(ws, 5, 1, "Company", font=BLACK)
for j, s in enumerate(syms, start=2):
    put(ws, 5, j, P[s]["name"], font=BLACK).alignment = Alignment(wrap_text=True)
ws.row_dimensions[5].height = 30
inputs = [("close_price_2026_08_04", "INR/share", F_PX), ("shares_outstanding", "shares", F_SH)] + [
    ("fy26_" + k, "INR/share" if k == "diluted_eps" else "INR mn", F_PX if k == "diluted_eps" else F_MN)
    for k in ["revenue", "other_income", "finance_costs", "da", "pbt_pre_exc", "pat", "pat_owners", "diluted_eps",
              "borrow_nc", "borrow_c", "cash", "other_bank", "current_investments", "equity_owners", "nci"]]
prow = {}
r = 6
for key, unit, fmt in inputs:
    put(ws, r, 1, key, font=BLACK)
    for j, s in enumerate(syms, start=2):
        v = P[s].get(key)
        put(ws, r, j, v if v is not None else None, fmt)
    put(ws, r, 6, unit, font=BLACK)
    prow[key] = r
    r += 1
put(ws, r, 1, "Source URLs, pages and dates for each figure: data/peers.csv (long format, one row per figure).", font=NOTE)
r += 2
header(ws, r, ["Calculated (INR crore)"] + syms + ["Median", "Low", "High"])
r += 1
calc = [
    ("Market cap", lambda c: f"={c}{prow['close_price_2026_08_04']}*{c}{prow['shares_outstanding']}/1E7", F_CR, "market_cap"),
    ("Net debt (ex-leases)", lambda c: f"=({c}{prow['fy26_borrow_nc']}+{c}{prow['fy26_borrow_c']}-{c}{prow['fy26_cash']}-{c}{prow['fy26_other_bank']})/10", F_CR, "net_debt"),
    ("Non-controlling interest", lambda c: f"=N({c}{prow['fy26_nci']})/10", F_CR, "nci"),
    ("Enterprise value", None, F_CR, "ev"),
    ("Revenue FY26", lambda c: f"={c}{prow['fy26_revenue']}/10", F_CR, "revenue"),
    ("EBITDA FY26", lambda c: f"=({c}{prow['fy26_pbt_pre_exc']}+{c}{prow['fy26_finance_costs']}+{c}{prow['fy26_da']}-{c}{prow['fy26_other_income']})/10", F_CR, "ebitda"),
    ("EBITDA margin", None, F_PCT, "ebitda_margin"),
    ("P/E (price / diluted EPS)", lambda c: f"={c}{prow['close_price_2026_08_04']}/{c}{prow['fy26_diluted_eps']}", F_X, "pe"),
    ("EV / EBITDA", None, F_X, "ev_ebitda"),
    ("EV / Sales", None, F_X, "ev_sales"),
    ("Price / Book", lambda c: f"={c}{r + 0}/({c}{prow['fy26_equity_owners']}/10)", F_X, "pb"),
]
crow = {}
for lab, fn, fmt, key in calc:
    crow[key] = r
    put(ws, r, 1, lab, font=BLACK)
    r += 1
for lab, fn, fmt, key in calc:
    rr = crow[key]
    for j, s in enumerate(syms, start=2):
        c = get_column_letter(j)
        if key == "ev":
            f = f"={c}{crow['market_cap']}+{c}{crow['net_debt']}+{c}{crow['nci']}"
        elif key == "ebitda_margin":
            f = f"={c}{crow['ebitda']}/{c}{crow['revenue']}"
        elif key == "ev_ebitda":
            f = f"={c}{crow['ev']}/{c}{crow['ebitda']}"
        elif key == "ev_sales":
            f = f"={c}{crow['ev']}/{c}{crow['revenue']}"
        elif key == "pb":
            f = f"={c}{crow['market_cap']}/({c}{prow['fy26_equity_owners']}/10)"
        else:
            f = fn(c)
        put(ws, rr, j, f, fmt, key=f"peer_{s}_{key}")
    if key in m.MULTIPLES or key == "ebitda_margin":
        for j, fn_ in zip((6, 7, 8), ("MEDIAN", "MIN", "MAX")):
            put(ws, rr, j, f"={fn_}(B{rr}:E{rr})", fmt, key=f"peer_{fn_.lower()}_{key}")
        name(f"peer_median_{key}", absref(ws, f"F{rr}"))
widths(ws, [30, 16, 16, 16, 16, 12, 10, 10])
ws.freeze_panes = "B6"

# ------------------------------------------------------------------ Valuation
ws = sheet("Valuation", "Valuation at the price band and final price (INR crore)",
           "IPO: P/E = post-issue market cap / FY26 PAT to owners; EV = market cap + post-issue net debt (FY26 net debt less net fresh proceeds) + NCI.")
header(ws, 4, ["", "Lower band", "Upper band", "Final price"])
vrow = {}
labels = [("price", "Price (INR/share)", F_PX), ("post_shares", "Post-issue shares", F_SH), ("market_cap", "Market cap", F_CR),
          ("net_debt_pre", "Net debt, FY26 (ex-leases)", F_CR), ("net_debt_post", "Net debt, post-issue", F_CR), ("nci", "Non-controlling interest", F_CR),
          ("ev", "Enterprise value", F_CR), ("pe", "P/E, FY26 (post-issue shares)", F_X), ("ev_ebitda", "EV / EBITDA, FY26", F_X),
          ("ev_sales", "EV / Sales, FY26", F_X), ("pb", "Price / Book (post-money book)", F_X), ("pe_rhp_basis", "P/E, RHP basis (price / FY26 diluted EPS)", F_X)]
r = 5
for key, lab, fmt in labels:
    vrow[key] = r
    put(ws, r, 1, lab, font=BLACK)
    r += 1
FY = "FY2026"
for col, case, price_ref in (("B", "lower_band", "t_price_band_low"), ("C", "upper_band", "t_price_band_high"), ("D", "final_price", "t_final_issue_price")):
    R = lambda k: f"{col}{vrow[k]}"  # noqa: E731
    f = {
        "price": f"={price_ref}",
        "post_shares": f"=t_pre_issue_shares+ROUNDDOWN(t_fresh_issue_amount*1E6/{R('price')},0)" if case == "lower_band" else "=t_post_issue_shares",
        "market_cap": f"={R('post_shares')}*{R('price')}/1E7",
        "net_debt_pre": f"=(f_total_borrowings_{FY}-f_cash_and_equivalents_{FY}-f_other_bank_balances_{FY})/10",
        "net_debt_post": f"={R('net_debt_pre')}-t_net_proceeds/10",
        "nci": f"=f_non_controlling_interest_{FY}/10",
        "ev": f"={R('market_cap')}+{R('net_debt_post')}+{R('nci')}",
        "pe": f"={R('market_cap')}/(f_pat_attributable_to_owners_{FY}/10)",
        "ev_ebitda": f"={R('ev')}/(f_ebitda_{FY}/10)",
        "ev_sales": f"={R('ev')}/(f_revenue_from_operations_{FY}/10)",
        "pb": f"={R('market_cap')}/((f_equity_attributable_to_owners_{FY}+t_net_proceeds)/10)",
        "pe_rhp_basis": f"={R('price')}/f_diluted_eps_{FY}",
    }
    for key, _, fmt in labels:
        put(ws, vrow[key], ord(col) - 64, f[key], fmt, key=f"val_{case}_{key}")
r += 1
header(ws, r, ["vs peers (final price)", "IPO", "Peer median", "Premium / (discount)", "Peer low", "Peer high"])
r += 1
ps = wb["Peers"]
for key, lab in (("pe", "P/E"), ("ev_ebitda", "EV / EBITDA"), ("ev_sales", "EV / Sales"), ("pb", "Price / Book")):
    put(ws, r, 1, lab, font=BLACK)
    put(ws, r, 2, f"=D{vrow[key]}", F_X)
    put(ws, r, 3, f"=peer_median_{key}", F_X)
    put(ws, r, 4, f"=B{r}/C{r}-1", F_PCT, key=f"prem_{key}")
    put(ws, r, 5, f"={cells[f'peer_min_{key}']}", F_X)
    put(ws, r, 6, f"={cells[f'peer_max_{key}']}", F_X)
    r += 1
r += 1
put(ws, r, 1, "ANALYST VIEW: how I would have priced it (my judgement, not reported fact)", font=BOLD)
ws.cell(row=r, column=1).fill = VIEW_FILL
r += 1
put(ws, r, 1, "New-issue discount to peer median", font=BLACK)
put(ws, r, 2, m.NEW_ISSUE_DISCOUNT["low"], F_PCT)
put(ws, r, 3, m.NEW_ISSUE_DISCOUNT["high"], F_PCT)
disc_r = r
ws.cell(row=r - 1, column=2, value="Low end").font = BOLD
ws.cell(row=r - 1, column=3, value="High end").font = BOLD
r += 1
view = [
    ("Target EV / EBITDA", lambda c: f"=peer_median_ev_ebitda*(1-{c}{disc_r})", F_X, "ev_mult"),
    ("Target P/E", lambda c: f"=peer_median_pe*(1-{c}{disc_r})", F_X, "pe_mult"),
    ("Price from EV / EBITDA", None, F_PX, "price_ev_ebitda"),
    ("Price from P/E", None, F_PX, "price_pe"),
    ("Fundamental value per share (average)", None, F_PX, "price"),
]
vr = {}
for lab, fn, fmt, key in view:
    vr[key] = r
    put(ws, r, 1, lab, font=BLACK)
    r += 1
for col, side in (("B", "low"), ("C", "high")):
    fm = {
        "ev_mult": f"=peer_median_ev_ebitda*(1-{col}{disc_r})",
        "pe_mult": f"=peer_median_pe*(1-{col}{disc_r})",
        "price_ev_ebitda": f"=({col}{vr['ev_mult']}*f_ebitda_{FY}/10-($D${vrow['net_debt_post']}+$D${vrow['nci']}))*1E7/t_post_issue_shares",
        "price_pe": f"={col}{vr['pe_mult']}*f_pat_attributable_to_owners_{FY}/10*1E7/t_post_issue_shares",
        "price": f"=AVERAGE({col}{vr['price_ev_ebitda']},{col}{vr['price_pe']})",
    }
    for lab, fn, fmt, key in view:
        put(ws, vr[key], ord(col) - 64, fm[key], fmt, key=f"view_{side}_{key}")
put(ws, r, 1, "Recommended price band (INR/share)", font=BOLD)
put(ws, r, 2, m.RECOMMENDED_BAND[0], F_PX)
put(ws, r, 3, m.RECOMMENDED_BAND[1], F_PX)
band_r = r
r += 1
for lab, key, fmt in (("Implied post-issue shares", "sh", F_SH), ("Implied market cap", "mc", F_CR), ("Implied P/E", "pe", F_X), ("Implied EV / EBITDA", "ev", F_X)):
    put(ws, r, 1, lab, font=BLACK)
    for col, side in (("B", "low"), ("C", "high")):
        if key == "sh":
            f = f"=t_pre_issue_shares+ROUNDDOWN(t_fresh_issue_amount*1E6/{col}{band_r},0)"
        elif key == "mc":
            f = f"={col}{r - 1}*{col}{band_r}/1E7"
        elif key == "pe":
            f = f"={col}{r - 1}/(f_pat_attributable_to_owners_{FY}/10)"
        else:
            f = f"=({col}{r - 2}+$D${vrow['net_debt_post']}+$D${vrow['nci']})/(f_ebitda_{FY}/10)"
        put(ws, r, ord(col) - 64, f, fmt, key=f"band_{side}_{key}")
    r += 1
for rr in range(disc_r - 1, r):
    for cc in range(1, 4):
        ws.cell(row=rr, column=cc).fill = VIEW_FILL
widths(ws, [44, 16, 16, 18, 12, 12])

# ------------------------------------------------------------------ Proceeds & Dilution
ws = sheet("Proceeds & Dilution", "Use of proceeds, net debt and dilution (INR crore)")
rows = [
    ("Gross offer size", "=t_total_offer_amount/10", F_CR, "gross_offer"),
    ("Fresh issue (gross)", "=t_gross_fresh_proceeds/10", F_CR, "fresh_gross"),
    ("Offer for sale", "=t_ofs_amount/10", F_CR, "ofs"),
    ("Fresh issue share of offer", "=t_fresh_issue_amount/t_total_offer_amount", F_PCT, "fresh_share_of_offer"),
    ("OFS share of offer", "=t_ofs_amount/t_total_offer_amount", F_PCT, "ofs_share_of_offer"),
    ("Less: offer expenses borne by company", "=t_fresh_issue_expenses/10", F_CR, "fresh_expenses"),
    ("Net proceeds", "=t_net_proceeds/10", F_CR, "net_proceeds"),
    ("Total offer expenses as % of offer", "=t_total_offer_expenses/t_total_offer_amount", F_PCT, "offer_expenses_pct"),
    (None, None, None, None),
    ("Objects of the fresh issue", None, None, None),
    ("  Repay company borrowings", "=t_object_repay_company_debt/10", F_CR, "obj1"),
    ("  Repay subsidiary borrowings", "=t_object_repay_subsidiary_debt/10", F_CR, "obj2"),
    ("  New wiring harness plants (Jhajjar, Hosur)", "=t_object_new_plants/10", F_CR, "obj3"),
    ("  Acquisitions and general corporate purposes", "=t_object_gcp_and_inorganic/10", F_CR, "obj4"),
    ("  Total objects", "=SUM(B{obj1}:B{obj4})", F_CR, "objects_total"),
    (None, None, None, None),
    ("Debt repayment from proceeds", "=B{obj1}+B{obj2}", F_CR, "debt_repayment"),
    ("Borrowings, FY26", "=f_total_borrowings_FY2026/10", F_CR, "borrowings_fy26"),
    ("Repayment as % of FY26 borrowings", "=B{debt_repayment}/B{borrowings_fy26}", F_PCT, "debt_repayment_pct_fy26_borrowings"),
    ("Borrowings after repayment (pro forma FY26)", "=B{borrowings_fy26}-B{debt_repayment}", F_CR, "borrowings_after_repayment"),
    ("Net debt before issue (FY26, ex-leases)", "=Valuation!D8", F_CR, "net_debt_before"),
    ("Net debt after issue (pro forma)", "=Valuation!D9", F_CR, "net_debt_after"),
    (None, None, None, None),
    ("Fresh shares issued", "=t_fresh_issue_shares", F_SH, "fresh_shares"),
    ("Dilution to existing holders (fresh / post)", "=t_fresh_issue_shares/t_post_issue_shares", F_PCT, "dilution"),
    ("Promoter holding pre-issue", "=t_promoter_shares_pre/t_pre_issue_shares", F_PCT, "promoter_pct_pre"),
    ("Promoter holding post-issue", "=t_promoter_shares_post/t_post_issue_shares", F_PCT, "promoter_pct_post"),
    ("Promoter + group pre-issue", "=(t_promoter_shares_pre+t_promoter_group_shares_pre)/t_pre_issue_shares", F_PCT, "promoter_and_group_pct_pre"),
    ("Promoter + group post-issue", "=(t_promoter_shares_post+t_promoter_group_shares_post)/t_post_issue_shares", F_PCT, "promoter_and_group_pct_post"),
    ("Public float post-issue", "=1-B{promoter_and_group_pct_post}", F_PCT, "public_float_post"),
]
r = 4
rowof = {}
for lab, f, fmt, key in rows:
    if key:
        rowof[key] = r
    r += 1
r = 4
for lab, f, fmt, key in rows:
    if lab:
        put(ws, r, 1, lab, font=BOLD if f is None else BLACK)
    if f:
        put(ws, r, 2, f.format(**rowof), fmt, key=f"pd_{key}")
    r += 1
widths(ws, [46, 16])
# fix references to valuation rows
ws.cell(row=rowof["net_debt_before"], column=2, value=f"=Valuation!D{vrow['net_debt_pre']}")
ws.cell(row=rowof["net_debt_after"], column=2, value=f"=Valuation!D{vrow['net_debt_post']}")

# ------------------------------------------------------------------ Demand
ws = sheet("Demand", "Subscription, anchor book and money left on the table",
           "Day 3 = consolidated NSE+BSE book (excl. anchor). Registrar = basis of allotment, before rejections. Day 1-2 not published any more: needs manual entry.")
header(ws, 4, ["Category", "Day", "Date", "Shares offered", "Shares bid", "Multiple (source)", "Multiple (recalc)", "Source"])
r = 5
for row in m.read("subscription.csv"):
    put(ws, r, 1, row["category"], font=BLACK)
    put(ws, r, 2, row["day"], font=BLACK)
    put(ws, r, 3, row["date"], font=BLACK)
    for j, k in ((4, "shares_offered"), (5, "shares_bid"), (6, "multiple")):
        put(ws, r, j, m._num(row[k]), F_SH if j < 6 else '0.00"x"')
    if row["shares_offered"]:
        put(ws, r, 7, f"=E{r}/D{r}", '0.00"x"', key=f"dem_{row['day']}_{row['category']}")
    put(ws, r, 8, row["notes"] or row["source_url"], font=NOTE)
    r += 1
r += 1
header(ws, r, ["Anchor book and listing", "Value"])
r += 1
lst = {x["series"]: m._num(x["close"]) for x in m.read("post_listing.csv") if x["label"] == "Listing day"}
drows = [
    ("Anchor shares", "=t_anchor_shares", F_SH, "anchor_shares"),
    ("Net QIB shares allotted", "=t_net_qib_shares_allotted", F_SH, "net_qib"),
    ("Anchor as % of QIB portion", "=B{anchor_shares}/(B{anchor_shares}+B{net_qib})", F_PCT, "anchor_share_of_qib"),
    ("Anchor as % of total offer", "=B{anchor_shares}/t_total_offer_shares", F_PCT, "anchor_share_of_offer"),
    ("Domestic MF share of anchor book", "=t_anchor_mf_shares/t_anchor_shares", F_PCT, "anchor_mf_share"),
    ("Listing-day close (NSE)", lst["DHOOTTRANS"], F_PX, "listing_close"),
    ("Issue price", "=t_final_issue_price", F_PX, "issue_price"),
    ("Listing-day gain per share", "=B{listing_close}-B{issue_price}", F_PX, "listing_gain_per_share"),
    ("Listing-day gain %", "=B{listing_gain_per_share}/B{issue_price}", F_PCT, "listing_gain_pct"),
    ("Money left on the table (INR crore)", "=B{listing_gain_per_share}*t_total_offer_shares/1E7", F_CR, "money_left_on_table"),
    ("  of which fresh issue (company)", "=B{listing_gain_per_share}*t_fresh_issue_shares/1E7", F_CR, "money_left_fresh"),
    ("  of which OFS (selling shareholders)", "=B{listing_gain_per_share}*t_ofs_shares/1E7", F_CR, "money_left_ofs"),
]
dr = {}
for i, (_, _, _, key) in enumerate(drows):
    dr[key] = r + i
for lab, f, fmt, key in drows:
    put(ws, r, 1, lab, font=BLACK)
    put(ws, r, 2, f.format(**dr) if isinstance(f, str) else f, fmt, key=f"dem_{key}")
    r += 1
put(ws, r, 1, "Listing-day close source: data/post_listing.csv (NSE EOD, 2026-08-17).", font=NOTE)
widths(ws, [36, 26, 12, 15, 16, 16, 16, 70])

# ------------------------------------------------------------------ Aftermarket
ws = sheet("Aftermarket", "Aftermarket returns vs Nifty 50 and Nifty Auto",
           "Base = issue price (871) and index closes on 12-Aug-2026 (bid close). +90 and +180 days have not elapsed (built 2026-11-13).")
header(ws, 4, ["Checkpoint", "Date", "DHOOTTRANS", "NIFTY 50", "NIFTY AUTO", "Return: Dhoot", "Return: Nifty 50", "Return: Nifty Auto", "Excess vs Nifty 50", "Excess vs Nifty Auto"])
rows = m.read("post_listing.csv")
labels = ["Pricing (bid close)"] + m.CHECKPOINTS
r = 5
base_r = r
for lab in labels:
    px = {x["series"]: m._num(x["close"]) for x in rows if x["label"] == lab}
    dt = next(x["date"] or x["target_date"] for x in rows if x["label"] == lab)
    put(ws, r, 1, lab, font=BLACK)
    put(ws, r, 2, dt, font=BLACK)
    for j, s in enumerate(["DHOOTTRANS", "NIFTY 50", "NIFTY AUTO"], start=3):
        put(ws, r, j, px.get(s), F_PX)
    if lab != labels[0]:
        for j, c in enumerate("CDE", start=6):
            put(ws, r, j, f'=IF({c}{r}="","n/a",{c}{r}/{c}${base_r}-1)', F_PCT, key=f"am_{lab}_{c}")
        put(ws, r, 9, f'=IF(C{r}="","n/a",F{r}-G{r})', F_PCT, key=f"am_{lab}_x50")
        put(ws, r, 10, f'=IF(C{r}="","n/a",F{r}-H{r})', F_PCT, key=f"am_{lab}_xauto")
    r += 1
widths(ws, [24, 12, 13, 12, 12, 14, 14, 14, 16, 16])

# ------------------------------------------------------------------ Checks
ws = sheet("Checks", "Integrity checks (all should be TRUE)")
header(ws, 3, ["Check", "Result", "Difference"])
checks = [
    ("Pre-issue shares + fresh shares = post-issue shares", "=t_pre_issue_shares+t_fresh_issue_shares-t_post_issue_shares"),
    ("Fresh shares + OFS shares = total offer shares", "=t_fresh_issue_shares+t_ofs_shares-t_total_offer_shares"),
    ("Fresh + OFS amount = total offer amount", "=t_fresh_issue_amount+t_ofs_amount-t_total_offer_amount"),
    ("OFS shares by seller sum to OFS", "=t_ofs_shares_bc_asia_xv+t_ofs_shares_mangalam_capital-t_ofs_shares"),
    ("OFS shares x price = OFS amount", "=ROUND(t_ofs_shares*t_final_issue_price/1E6,2)-t_ofs_amount"),
    ("Gross fresh - expenses = net proceeds", "=t_gross_fresh_proceeds-t_fresh_issue_expenses-t_net_proceeds"),
    ("Objects sum to net proceeds", f"=ROUND({cells['pd_objects_total']}*10-t_net_proceeds,2)"),
    ("Promoter shares fall by BC Asia OFS", "=t_promoter_shares_pre-t_promoter_shares_post-t_ofs_shares_bc_asia_xv"),
    ("Promoter group shares fall by Mangalam OFS", "=t_promoter_group_shares_pre-t_promoter_group_shares_post-t_ofs_shares_mangalam_capital"),
    ("Offer categories sum to total offer shares", "=t_anchor_shares+t_net_qib_shares_allotted+t_retail_shares_reserved+t_nii_small_shares_reserved+t_nii_large_shares_reserved+t_employee_reservation_shares-t_total_offer_shares"),
    ("Anchor amount = anchor shares x price", "=ROUND(t_anchor_shares*t_final_issue_price/1E6-t_anchor_amount,4)"),
    ("Anchor categories sum to anchor shares", "=t_anchor_mf_shares+t_anchor_insurance_shares+t_anchor_fpi_shares+t_anchor_aif_shares+t_anchor_nbfc_shares-t_anchor_shares"),
    ("RHP net debt reconciles (FY24-FY26)", f"=ROUND(ABS({cells['fin_netdebt_check_B']})+ABS({cells['fin_netdebt_check_C']})+ABS({cells['fin_netdebt_check_D']}),2)"),
    ("RHP EBITDA reconciles (FY24-FY26)", f"=ROUND(ABS({cells['fin_ebitda_check_B']})+ABS({cells['fin_ebitda_check_C']})+ABS({cells['fin_ebitda_check_D']}),2)"),
    ("RHP cap-price P/E reproduced (35.70x)", f"=ROUND({cells['val_upper_band_pe_rhp_basis']},2)-t_rhp_cap_price_pe"),
    ("Sub-category NII bids sum to NII", "=Demand!E{a}+Demand!E{b}-Demand!E{c}"),
]
# locate NII rows on Demand for the final day
dws = wb["Demand"]
loc = {}
for rr in range(5, dws.max_row + 1):
    if dws.cell(rr, 2).value == "Day 3 (final)":
        loc[dws.cell(rr, 1).value] = rr
r = 4
for lab, f in checks:
    f = f.format(a=loc["NII >10 lakh"], b=loc["NII 2-10 lakh"], c=loc["NII"])
    put(ws, r, 1, lab, font=BLACK)
    put(ws, r, 3, f, '#,##0.00;(#,##0.00);0', key=f"chk_diff_{r}")
    put(ws, r, 2, f"=ABS(C{r})<0.005", key=f"chk_{r}")
    r += 1
put(ws, r + 1, 1, "All checks pass", font=BOLD)
put(ws, r + 1, 2, f"=AND(B4:B{r - 1})", key="chk_all")
widths(ws, [58, 10, 14])

wb.move_sheet("Checks", offset=0)
order = ["Terms", "Financials", "Peers", "Valuation", "Proceeds & Dilution", "Demand", "Aftermarket", "Checks"]
wb._sheets = [wb[n] for n in order]
wb.active = 3
OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
(ROOT / "output" / "cell_map.json").write_text(json.dumps(cells, indent=1))
print(f"wrote {OUT} ({len(cells)} mapped cells)")
