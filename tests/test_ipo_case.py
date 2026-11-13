"""Tests: sourcing, multiple recomputation, Excel-vs-Python parity, share reconciliation.

Run:  python -m pytest -q
The Excel parity tests need the workbook built (build_workbook.py) and recalculated
in Excel (recalc.ps1); they fail if cached values are missing.
"""
import csv
import json
import math
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
import ipo_case as m  # noqa: E402

DATA = ROOT / "data"
BOOK = ROOT / "output" / "Dhoot_IPO_Case.xlsx"
CELLS = ROOT / "output" / "cell_map.json"
FLAGS = ("NEEDS MANUAL ENTRY", "NOT YET ELAPSED", "NOT APPLICABLE")
VALUE_COL = {"ipo_terms.csv": "value", "financials.csv": "value", "peers.csv": "value", "subscription.csv": "multiple",
             "post_listing.csv": "close", "anchor_investors.csv": "shares_allocated"}


# ------------------------------------------------------------------ 1. every row has a source
@pytest.mark.parametrize("fname", sorted(VALUE_COL))
def test_every_row_has_a_source(fname):
    with open(DATA / fname, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    assert rows, fname
    for i, r in enumerate(rows, start=2):
        assert r["source_url"].startswith("https://"), f"{fname}:{i} missing source_url"
        assert r["source_page"].strip(), f"{fname}:{i} missing source_page"
        blank = r[VALUE_COL[fname]].strip() == ""
        if blank:
            assert any(flag in r.get("notes", "") for flag in FLAGS), f"{fname}:{i} blank value without a manual-entry flag"
        else:
            assert len(r["source_date"]) == 10 and r["source_date"][4] == "-", f"{fname}:{i} missing source_date"


def test_sources_are_primary():
    allowed = ("sebi.gov.in", "nseindia.com", "bseindia.com")
    for fname in VALUE_COL:
        with open(DATA / fname, newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                assert any(a in r["source_url"] for a in allowed), f"{fname}: non-primary source {r['source_url']}"


def test_peer_prices_are_on_announcement_date():
    t = m.terms()
    assert t["date_price_band_announcement"] == "2026-08-04"
    with open(DATA / "peers.csv", newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r["metric"] == "close_price_2026_08_04":
                assert r["source_date"] == "2026-08-04"


# ------------------------------------------------------------------ 2. multiples recompute correctly
def test_ipo_multiples_recompute_from_raw_inputs():
    mcap = 871 * 204548057 / 1e7
    assert mcap == pytest.approx(17816.1357647)
    net_debt_post = (8413.92 - 10842.82 - 115.66 - 13269.81) / 10
    ev = mcap + net_debt_post + 6.09 / 10
    v = m.valuation_cases()["final_price"]
    assert v["market_cap"] == pytest.approx(mcap)
    assert v["ev"] == pytest.approx(ev)
    assert v["pe"] == pytest.approx(mcap / 396.654)
    assert v["ev_ebitda"] == pytest.approx(ev / 710.989)
    assert v["ev_sales"] == pytest.approx(ev / 4524.955)
    assert v["pb"] == pytest.approx(mcap / ((24343.42 + 13269.81) / 10))


def test_rhp_cap_price_pe_reproduces():
    v = m.valuation_cases()["upper_band"]
    assert round(v["pe_rhp_basis"], 2) == m.terms()["rhp_cap_price_pe"]


def test_peer_eps_inputs_reproduce_rhp_peer_pe():
    # RHP p.163-165 uses BSE closes on 31-Jul-2026 and FY26 diluted EPS. Our EPS inputs must reproduce its P/Es.
    rhp = {"MINDACORP": (700.55, 46.49), "UNOMINDA": (1180.15, 56.87), "MSUMI": (40.65, 43.24), "SONACOMS": (768.80, 74.64)}
    p = m.peers()
    for sym, (px, pe) in rhp.items():
        assert round(px / p[sym]["fy26_diluted_eps"], 2) == pytest.approx(pe, abs=0.01), sym
    assert round(sum(v[1] for v in rhp.values()) / 4, 2) == m.terms()["rhp_peer_average_pe"]


def test_peer_multiples_recompute():
    p, pm = m.peers(), m.peer_multiples()
    d = p["UNOMINDA"]
    mcap = 1227.0 * 577467206 / 1e7
    ev = mcap + (10692.6 + 14680.3 - 2769.9 - 811.4) / 10 + 4309.0 / 10
    ebitda = (13950.8 + 1870.2 + 7035.6 - 344.3) / 10
    assert pm["UNOMINDA"]["ev_ebitda"] == pytest.approx(ev / ebitda)
    assert pm["UNOMINDA"]["pb"] == pytest.approx(mcap / (d["fy26_equity_owners"] / 10))
    s = m.peer_summary(pm)
    for k in m.MULTIPLES:
        vals = sorted(pm[x][k] for x in m.PEER_ORDER)
        assert s[k]["median"] == pytest.approx((vals[1] + vals[2]) / 2)
        assert s[k]["low"] == vals[0] and s[k]["high"] == vals[-1]
    prem = m.premium_to_peers()
    v = m.valuation_cases()["final_price"]
    for k in m.MULTIPLES:
        assert prem[k] == pytest.approx(v[k] / s[k]["median"] - 1)


def test_rhp_ebitda_and_net_debt_reconcile():
    f = m.financials()
    for fy in ("FY2024", "FY2025", "FY2026"):
        e = f["pbt_before_exceptional"][fy] + f["finance_costs"][fy] + f["depreciation_amortisation"][fy] - f["other_income"][fy]
        assert e == pytest.approx(f["ebitda"][fy], abs=0.01)
        nd = f["total_borrowings"][fy] + f["lease_liabilities"][fy] - f["cash_and_equivalents"][fy] - f["other_bank_balances"][fy]
        assert nd == pytest.approx(f["net_debt_rhp"][fy], abs=0.01)


def test_money_left_on_table():
    d = m.demand()
    assert d["money_left_on_table"] == pytest.approx((1187.05 - 871) * 35218047 / 1e7)
    assert d["money_left_fresh"] + d["money_left_ofs"] == pytest.approx(d["money_left_on_table"])


def test_subscription_multiples_recompute():
    for r in m.read("subscription.csv"):
        if r["shares_offered"]:
            calc = float(r["shares_bid"]) / float(r["shares_offered"])
            assert calc == pytest.approx(float(r["multiple"]), abs=0.006), (r["day"], r["category"])


# ------------------------------------------------------------------ 3. share counts reconcile
def test_share_counts_reconcile():
    t = m.terms()
    assert t["pre_issue_shares"] + t["fresh_issue_shares"] == t["post_issue_shares"]
    assert t["fresh_issue_shares"] + t["ofs_shares"] == t["total_offer_shares"]
    assert t["ofs_shares_bc_asia_xv"] + t["ofs_shares_mangalam_capital"] == t["ofs_shares"]
    assert t["promoter_shares_pre"] - t["promoter_shares_post"] == t["ofs_shares_bc_asia_xv"]
    assert t["promoter_group_shares_pre"] - t["promoter_group_shares_post"] == t["ofs_shares_mangalam_capital"]
    cats = ["anchor_shares", "net_qib_shares_allotted", "retail_shares_reserved", "nii_small_shares_reserved",
            "nii_large_shares_reserved", "employee_reservation_shares"]
    assert sum(t[c] for c in cats) == t["total_offer_shares"]
    assert t["fresh_issue_amount"] + t["ofs_amount"] == pytest.approx(t["total_offer_amount"])
    assert round(t["ofs_shares"] * t["final_issue_price"] / 1e6, 2) == t["ofs_amount"]


def test_anchor_register_reconciles():
    t = m.terms()
    rows = m.read("anchor_investors.csv")
    assert sum(int(r["shares_allocated"]) for r in rows) == t["anchor_shares"]
    assert sum(int(r["amount_inr"]) for r in rows) / 1e6 == pytest.approx(t["anchor_amount"])
    mf = sum(int(r["shares_allocated"]) for r in rows if r["category"] == "Mutual Fund")
    assert mf == t["anchor_mf_shares"]


def test_objects_sum_to_net_proceeds():
    p = m.proceeds_and_dilution()
    assert p["objects_total"] == pytest.approx(p["net_proceeds"])
    assert p["fresh_gross"] - p["fresh_expenses"] == pytest.approx(p["net_proceeds"])


# ------------------------------------------------------------------ 4. Excel cached values match Python
def _python_value(key):
    vc, pm, ps = m.valuation_cases(), m.peer_multiples(), m.peer_summary()
    pr, pd, dm, am, mv = m.premium_to_peers(), m.proceeds_and_dilution(), m.demand(), m.aftermarket(), m.my_price_range()
    if key.startswith("val_"):
        for case in vc:
            if key.startswith(f"val_{case}_"):
                return vc[case][key[len(f"val_{case}_"):]]
    if key.startswith("peer_"):
        _, a, b = key.split("_", 2)
        if a in ("median", "min", "max"):
            return ps[b][{"min": "low", "max": "high"}.get(a, a)] if b in ps else None
        return pm[a][b]
    if key.startswith("prem_"):
        return pr[key[5:]]
    if key.startswith("pd_"):
        k = key[3:]
        if k in ("obj1", "obj2", "obj3", "obj4"):
            return list(pd["objects"].values())[int(k[3]) - 1]
        return pd.get(k)
    if key.startswith("dem_"):
        return dm.get(key[4:])
    if key.startswith("am_"):
        _, label, col = key.split("_", 2)
        rec = am[label]
        if rec is None:
            return "n/a"
        return rec[{"C": "DHOOTTRANS", "D": "NIFTY 50", "E": "NIFTY AUTO", "x50": "excess_vs_NIFTY 50", "xauto": "excess_vs_NIFTY AUTO"}[col]]
    if key.startswith("view_"):
        _, side, k = key.split("_", 2)
        return mv[side][{"ev_mult": "ev_ebitda", "pe_mult": "pe"}.get(k, k)]
    if key.startswith("band_"):
        _, side, k = key.split("_", 2)
        return mv["recommended_band"][side][{"sh": "post_shares", "mc": "market_cap", "pe": "pe", "ev": "ev_ebitda"}[k]]
    return None


@pytest.fixture(scope="module")
def excel():
    if not BOOK.exists() or not CELLS.exists():
        pytest.fail("Build the workbook first: python build_workbook.py; powershell -File recalc.ps1")
    from openpyxl import load_workbook
    wb = load_workbook(BOOK, data_only=True)
    cmap = json.loads(CELLS.read_text())

    def get(ref):
        s, c = ref.rsplit("!", 1)
        return wb[s.strip("'")][c].value
    return cmap, get


def test_excel_has_cached_values(excel):
    cmap, get = excel
    missing = [k for k, ref in cmap.items() if get(ref) is None]
    assert not missing, f"no cached value (run recalc.ps1): {missing[:5]}"


def test_excel_matches_python(excel):
    cmap, get = excel
    compared = 0
    for key, ref in cmap.items():
        py = _python_value(key)
        if py is None or key.startswith(("chk", "fin_", "dem_Day", "dem_Registrar")):
            continue
        xl = get(ref)
        if isinstance(py, str):
            assert xl == py, key
        else:
            assert isinstance(xl, (int, float)) and math.isclose(xl, py, rel_tol=1e-9, abs_tol=1e-9), f"{key}: excel {xl} vs python {py}"
        compared += 1
    assert compared >= 120, compared


def test_excel_checks_all_true(excel):
    cmap, get = excel
    assert get(cmap["chk_all"]) is True
    for k, ref in cmap.items():
        if k.startswith("chk_") and not k.startswith("chk_diff") and k != "chk_all":
            assert get(ref) is True, k


def test_excel_has_no_hardcoded_outputs():
    from openpyxl import load_workbook
    wb = load_workbook(BOOK)
    cmap = json.loads(CELLS.read_text())
    inputs = {"dem_listing_close"}  # listing close is an input (NSE close), shown for reference
    for key, ref in cmap.items():
        if key in inputs:
            continue
        s, c = ref.rsplit("!", 1)
        v = wb[s.strip("'")][c].value
        assert isinstance(v, str) and v.startswith("="), f"{key} at {ref} is not a formula"
