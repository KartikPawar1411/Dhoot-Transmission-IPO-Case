"""Dhoot Transmission IPO: valuation, proceeds, demand and aftermarket calculations.

Every input comes from the CSVs in ../data (each row carries its source).
Amounts in the CSVs are INR million, as reported; outputs are INR crore
(1 crore = 10 million) unless a key says otherwise.

Definitions used for both the IPO and the peers, so the multiples compare:
  EBITDA      = PBT before exceptional items + finance costs + D&A - other income
                (the RHP's own definition, p.457)
  Net debt    = borrowings (non-current + current) - cash - other bank balances.
                Lease liabilities and investments are excluded on both sides,
                because the peers' XBRL results do not tag lease liabilities.
  EV          = market cap + net debt + non-controlling interest
  P/E         = IPO: post-issue market cap / FY26 PAT attributable to owners
                Peers: close price / FY26 reported diluted EPS
  P/B         = IPO: post-issue market cap / (FY26 owners' equity + net fresh proceeds)
                Peers: market cap / FY26 owners' equity
"""
from __future__ import annotations

import csv
import math
from pathlib import Path
from statistics import median

DATA = Path(__file__).resolve().parent.parent / "data"
CR = 10.0  # INR million per crore

PEER_ORDER = ["MINDACORP", "UNOMINDA", "MSUMI", "SONACOMS"]
CHECKPOINTS = ["Listing day", "+30 days", "+90 days", "+180 days", "Latest (+87 days)"]
INDICES = ["NIFTY 50", "NIFTY AUTO"]


# ---------------------------------------------------------------- loading
def read(name: str) -> list[dict]:
    with open(DATA / name, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _num(v: str):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except ValueError:
        return v


def terms() -> dict:
    return {r["field"]: _num(r["value"]) for r in read("ipo_terms.csv")}


def financials() -> dict:
    out: dict = {}
    for r in read("financials.csv"):
        out.setdefault(r["metric"], {})[r["fiscal_year"]] = _num(r["value"])
    return out


def peers() -> dict:
    out: dict = {}
    for r in read("peers.csv"):
        p = out.setdefault(r["symbol"], {"name": r["peer"]})
        p[r["metric"]] = _num(r["value"])
    return out


# ---------------------------------------------------------------- (a) valuation
def ipo_valuation(price: float, t: dict | None = None, f: dict | None = None, post_shares: float | None = None) -> dict:
    t = t or terms()
    f = f or financials()
    fy = "FY2026"
    if post_shares is None:
        # Fresh issue amount is fixed in rupees; shares issued depend on price.
        post_shares = t["pre_issue_shares"] + math.floor(t["fresh_issue_amount"] * 1e6 / price)
    mcap = post_shares * price / 1e7  # crore
    net_debt_pre = (f["total_borrowings"][fy] - f["cash_and_equivalents"][fy] - f["other_bank_balances"][fy]) / CR
    net_debt_post = net_debt_pre - t["net_proceeds"] / CR
    nci = f["non_controlling_interest"][fy] / CR
    ev = mcap + net_debt_post + nci
    book_post = (f["equity_attributable_to_owners"][fy] + t["net_proceeds"]) / CR
    ebitda, rev, pat = (f[k][fy] / CR for k in ("ebitda", "revenue_from_operations", "pat_attributable_to_owners"))
    return {
        "price": price, "post_shares": post_shares, "market_cap": mcap,
        "net_debt_pre": net_debt_pre, "net_debt_post": net_debt_post, "nci": nci, "ev": ev,
        "pe": mcap / pat, "ev_ebitda": ev / ebitda, "ev_sales": ev / rev, "pb": mcap / book_post,
        "pe_rhp_basis": price / f["diluted_eps"][fy],
    }


def valuation_cases() -> dict:
    t, f = terms(), financials()
    return {
        "lower_band": ipo_valuation(t["price_band_low"], t, f),
        "upper_band": ipo_valuation(t["price_band_high"], t, f, post_shares=t["post_issue_shares"]),
        "final_price": ipo_valuation(t["final_issue_price"], t, f, post_shares=t["post_issue_shares"]),
    }


# ---------------------------------------------------------------- (b) peers
def peer_multiples(p: dict | None = None) -> dict:
    p = p or peers()
    out = {}
    for sym in PEER_ORDER:
        d = p[sym]
        mcap = d["close_price_2026_08_04"] * d["shares_outstanding"] / 1e7
        net_debt = (d["fy26_borrow_nc"] + d["fy26_borrow_c"] - d["fy26_cash"] - d["fy26_other_bank"]) / CR
        nci = (d.get("fy26_nci") or 0.0) / CR
        ev = mcap + net_debt + nci
        ebitda = (d["fy26_pbt_pre_exc"] + d["fy26_finance_costs"] + d["fy26_da"] - d["fy26_other_income"]) / CR
        rev = d["fy26_revenue"] / CR
        out[sym] = {
            "name": d["name"], "price": d["close_price_2026_08_04"], "shares": d["shares_outstanding"],
            "market_cap": mcap, "net_debt": net_debt, "nci": nci, "ev": ev, "ebitda": ebitda, "revenue": rev,
            "ebitda_margin": ebitda / rev,
            "pe": d["close_price_2026_08_04"] / d["fy26_diluted_eps"],
            "ev_ebitda": ev / ebitda, "ev_sales": ev / rev,
            "pb": mcap / (d["fy26_equity_owners"] / CR),
        }
    return out


MULTIPLES = ["pe", "ev_ebitda", "ev_sales", "pb"]


def peer_summary(pm: dict | None = None) -> dict:
    pm = pm or peer_multiples()
    s = {}
    for m in MULTIPLES:
        vals = [pm[k][m] for k in PEER_ORDER]
        s[m] = {"median": median(vals), "low": min(vals), "high": max(vals)}
    return s


def premium_to_peers(case: str = "final_price") -> dict:
    v = valuation_cases()[case]
    s = peer_summary()
    return {m: v[m] / s[m]["median"] - 1 for m in MULTIPLES}


# ---------------------------------------------------------------- (c) proceeds and dilution
def proceeds_and_dilution() -> dict:
    t, f = terms(), financials()
    objects = {
        "Repay company borrowings": t["object_repay_company_debt"],
        "Repay subsidiary borrowings": t["object_repay_subsidiary_debt"],
        "New wiring harness plants": t["object_new_plants"],
        "Acquisitions and general corporate purposes": t["object_gcp_and_inorganic"],
    }
    objects = {k: v / CR for k, v in objects.items()}
    gross = t["total_offer_amount"] / CR
    v = valuation_cases()["final_price"]
    debt_repay = (t["object_repay_company_debt"] + t["object_repay_subsidiary_debt"]) / CR
    borrow = f["total_borrowings"]["FY2026"] / CR
    return {
        "gross_offer": gross,
        "fresh_gross": t["gross_fresh_proceeds"] / CR,
        "fresh_expenses": t["fresh_issue_expenses"] / CR,
        "net_proceeds": t["net_proceeds"] / CR,
        "ofs": t["ofs_amount"] / CR,
        "ofs_share_of_offer": t["ofs_amount"] / t["total_offer_amount"],
        "fresh_share_of_offer": t["fresh_issue_amount"] / t["total_offer_amount"],
        "objects": objects,
        "objects_total": sum(objects.values()),
        "debt_repayment": debt_repay,
        "debt_repayment_pct_fy26_borrowings": debt_repay / borrow,
        "net_debt_before": v["net_debt_pre"],
        "net_debt_after": v["net_debt_post"],
        "borrowings_fy26": borrow,
        "borrowings_after_repayment": borrow - debt_repay,
        "fresh_shares": t["fresh_issue_shares"],
        "dilution": t["fresh_issue_shares"] / t["post_issue_shares"],
        "promoter_pct_pre": t["promoter_shares_pre"] / t["pre_issue_shares"],
        "promoter_pct_post": t["promoter_shares_post"] / t["post_issue_shares"],
        "promoter_and_group_pct_pre": (t["promoter_shares_pre"] + t["promoter_group_shares_pre"]) / t["pre_issue_shares"],
        "promoter_and_group_pct_post": (t["promoter_shares_post"] + t["promoter_group_shares_post"]) / t["post_issue_shares"],
        "public_float_post": 1 - (t["promoter_shares_post"] + t["promoter_group_shares_post"]) / t["post_issue_shares"],
        "offer_expenses_pct": t["total_offer_expenses"] / t["total_offer_amount"],
    }


# ---------------------------------------------------------------- (d) demand
def demand() -> dict:
    t = terms()
    sub = read("subscription.csv")
    final = {r["category"]: float(r["multiple"]) for r in sub if r["day"] == "Day 3 (final)"}
    registrar = {r["category"]: float(r["multiple"]) for r in sub if r["day"].startswith("Registrar")}
    listing = {r["series"]: _num(r["close"]) for r in read("post_listing.csv") if r["label"] == "Listing day"}
    gain = listing["DHOOTTRANS"] - t["final_issue_price"]
    qib_portion = t["anchor_shares"] + t["net_qib_shares_allotted"]
    return {
        "final_by_category": final,
        "registrar_by_category": registrar,
        "anchor_share_of_qib": t["anchor_shares"] / qib_portion,
        "anchor_share_of_offer": t["anchor_shares"] / t["total_offer_shares"],
        "anchor_mf_share": t["anchor_mf_shares"] / t["anchor_shares"],
        "listing_close": listing["DHOOTTRANS"],
        "listing_gain_per_share": gain,
        "listing_gain_pct": gain / t["final_issue_price"],
        "money_left_on_table": gain * t["total_offer_shares"] / 1e7,
        "money_left_fresh": gain * t["fresh_issue_shares"] / 1e7,
        "money_left_ofs": gain * t["ofs_shares"] / 1e7,
    }


# ---------------------------------------------------------------- (e) aftermarket
def aftermarket() -> dict:
    rows = read("post_listing.csv")
    base = {r["series"]: float(r["close"]) for r in rows if r["label"] == "Pricing (bid close)"}
    out = {}
    for label in CHECKPOINTS:
        px = {r["series"]: _num(r["close"]) for r in rows if r["label"] == label}
        if px.get("DHOOTTRANS") is None:
            out[label] = None
            continue
        rec = {"date": next(r["date"] for r in rows if r["label"] == label)}
        for s in ["DHOOTTRANS"] + INDICES:
            rec[s] = px[s] / base[s] - 1
        for s in INDICES:
            rec["excess_vs_" + s] = rec["DHOOTTRANS"] - rec[s]
        out[label] = rec
    return out


# ---------------------------------------------------------------- my analysis
# ANALYST VIEW (my judgement, not reported fact). Reasoning in docs/ipo_case.md.
NEW_ISSUE_DISCOUNT = {"low": 0.25, "high": 0.15}   # discount to peer median, applied to both methods
RECOMMENDED_BAND = (860.0, 910.0)                  # INR/share, my recommended price band


def my_price_range() -> dict:
    """Fundamental range: peer-median EV/EBITDA and P/E, each less a 15-25% new-issue
    discount, converted to a price on the post-issue share count; the two methods are
    averaged. The recommended band is then set with the book's demand in view."""
    t, f = terms(), financials()
    s = peer_summary()
    ebitda = f["ebitda"]["FY2026"] / CR
    pat = f["pat_attributable_to_owners"]["FY2026"] / CR
    v = valuation_cases()["final_price"]
    bridge = v["net_debt_post"] + v["nci"]
    shares = t["post_issue_shares"]
    out = {}
    for label, disc in NEW_ISSUE_DISCOUNT.items():
        ev_mult = s["ev_ebitda"]["median"] * (1 - disc)
        pe_mult = s["pe"]["median"] * (1 - disc)
        px_ev = (ev_mult * ebitda - bridge) * 1e7 / shares
        px_pe = pe_mult * pat * 1e7 / shares
        out[label] = {"discount": disc, "ev_ebitda": ev_mult, "pe": pe_mult,
                      "price_ev_ebitda": px_ev, "price_pe": px_pe, "price": (px_ev + px_pe) / 2}
    out["recommended_band"] = {k: ipo_valuation(p, t, f) for k, p in zip(("low", "high"), RECOMMENDED_BAND)}
    return out


def results() -> dict:
    f = financials()
    return {"terms": terms(), "financials": {k: v for k, v in f.items() if k != "stub_period"},
            "valuation": valuation_cases(), "peers": peer_multiples(), "peer_summary": peer_summary(),
            "premium": premium_to_peers(), "proceeds": proceeds_and_dilution(), "demand": demand(),
            "aftermarket": aftermarket(), "my_range": my_price_range()}


if __name__ == "__main__":
    import json
    out = DATA.parent / "output" / "results.json"
    out.write_text(json.dumps(results(), indent=1, default=str))
    print(f"wrote {out}")
