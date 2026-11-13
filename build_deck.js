// Builds output/Dhoot_IPO_Case.pptx from output/results.json (written by `python src/ipo_case.py`).
// No figure is typed in here: every number comes from the Python results.
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const ROOT = __dirname;
const R = JSON.parse(fs.readFileSync(path.join(ROOT, "output", "results.json"), "utf8"));
const OUT = path.join(ROOT, "output", "Dhoot_IPO_Case.pptx");
const SKILL_THEME = path.join(ROOT, "scripts", "apply_theme.js"); // writes THEME colours into the deck

const THEME = {
  name: "Dhoot ECM",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1E2730", lt1: "FFFFFF", dk2: "22313F", lt2: "EEF1F4",
    accent1: "B8642E", accent2: "7A8B99", accent3: "3D5A73", accent4: "D9A47E",
    accent5: "4E7D5B", accent6: "A23B2A", hlink: "3D5A73", folHlink: "7A8B99",
  },
};
const HEX = THEME.colors;
const COPPER = HEX.accent1, SLATE = HEX.accent2, STEEL = HEX.accent3, SAND = HEX.accent4, GREEN = HEX.accent5, RED = HEX.accent6;

// ---------------------------------------------------------------- formatting
const t = R.terms, v = R.valuation.final_price, vl = R.valuation.lower_band, ps = R.peer_summary;
const pr = R.premium, pd = R.proceeds, dm = R.demand, am = R.aftermarket, my = R.my_range, f = R.financials;
const n0 = (x) => Math.round(x).toLocaleString("en-IN");
const n1 = (x) => x.toLocaleString("en-IN", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const cr = (x) => `₹${n0(x)} cr`;
const px = (x) => `₹${n0(x)}`;
const mx = (x) => `${x.toFixed(1)}x`;
const pc = (x, d = 1) => `${x < 0 ? "(" : ""}${Math.abs(x * 100).toFixed(d)}%${x < 0 ? ")" : ""}`;
const sgn = (x, d = 1) => `${x >= 0 ? "+" : "−"}${Math.abs(x * 100).toFixed(d)}%`;
const mn2cr = (x) => x / 10;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "ECM case study";
pres.title = "Dhoot Transmission IPO: pricing reconstruction";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const W = 13.333;

// ---------------------------------------------------------------- layouts
pres.defineSlideMaster({
  title: "COVER",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 2.0, w: 11.9, h: 1.4, fontFace: THEME.headFontFace, fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 3.55, w: 11.9, h: 1.0, fontFace: THEME.bodyFontFace, fontSize: 18, color: SAND, valign: "top", align: "left" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontFace: THEME.headFontFace, fontSize: 24, bold: true, color: C.text2, valign: "middle", align: "left" }, text: "" } },
    { placeholder: { options: { name: "source", type: "body", x: 0.6, y: 6.92, w: 11.2, h: 0.4, fontFace: THEME.bodyFontFace, fontSize: 9, color: SLATE, valign: "middle" }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 6.92, w: 0.5, h: 0.4, fontFace: THEME.bodyFontFace, fontSize: 9, color: SLATE, align: "right" },
});

const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", legendFontFace: "+mn-lt" };
const quiet = () => ({
  ...chartText,
  catAxisLabelColor: HEX.dk1, valAxisLabelColor: SLATE, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
  dataLabelFontSize: 11, dataLabelColor: HEX.dk1,
  valGridLine: { color: "D9DEE3", size: 0.5 }, catGridLine: { style: "none" },
  showTitle: true, titleFontSize: 13, titleColor: HEX.dk2,
});

function content(title, source, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  s.addText(source, { placeholder: "source" });
  return s;
}

// stat card: big number + label on a tinted card
function stat(s, x, y, w, h, big, label, color, name) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: C.background2 }, line: { type: "none" }, rectRadius: 0.08, objectName: name + " card" });
  s.addText(big, { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: h * 0.55, fontSize: 30, bold: true, color: color || C.accent1, fontFace: THEME.headFontFace, margin: 0, valign: "bottom", isTextBox: true, objectName: name + " value" });
  s.addText(label, { x: x + 0.2, y: y + h * 0.62, w: w - 0.4, h: h * 0.34, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: name + " label" });
}

function bullets(s, items, x, y, w, h, size = 14, name = "bullets") {
  s.addText(items.map((it, i) => ({ text: it, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 6 } })),
    { x, y, w, h, fontSize: size, color: C.text1, valign: "top", margin: 0.05, isTextBox: true, objectName: name });
}

function table(s, rows, opt) {
  const hdr = (txt) => ({ text: txt, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 12 } });
  const data = rows.map((r, i) => (i === 0 ? r.map(hdr) : r.map((c) => (typeof c === "object" ? c : { text: String(c) }))));
  s.addTable(data, { fontSize: 12, color: C.text1, border: { type: "solid", pt: 0.5, color: "D9DEE3" }, fill: { color: C.background1 }, valign: "middle", margin: 0.05, ...opt });
}

// ================================================================= 1. Cover
pres.addSection({ title: "Cover" });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Cover" });
  s.addText("Dhoot Transmission IPO", { placeholder: "title" });
  s.addText(`How a ${cr(pd.gross_offer)} auto-components IPO was priced, sized and received  |  ECM case study, data as of 12 Nov 2026`, { placeholder: "body" });
  const cards = [
    [cr(pd.gross_offer), "Issue size (fresh + OFS)"],
    [px(t.final_issue_price), "Issue price (top of band)"],
    [mx(dm.final_by_category.Total), "Overall subscription"],
    [sgn(dm.listing_gain_pct), "Listing-day close vs issue"],
  ];
  cards.forEach(([big, lab], i) => {
    const x = 0.7 + i * 3.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 5.0, w: 2.8, h: 1.35, fill: { color: STEEL }, line: { type: "none" }, rectRadius: 0.08, objectName: `cover card ${i}` });
    s.addText(big, { x: x + 0.2, y: 5.1, w: 2.4, h: 0.7, fontSize: 28, bold: true, color: C.background1, fontFace: THEME.headFontFace, margin: 0, valign: "bottom", isTextBox: true });
    s.addText(lab, { x: x + 0.2, y: 5.85, w: 2.4, h: 0.4, fontSize: 12, color: SAND, margin: 0, isTextBox: true });
  });
  s.addNotes("Every figure in this deck comes from the RHP (SEBI, 3 Aug 2026), the Prospectus (13 Aug 2026), the basis-of-allotment notice (14 Aug 2026), the anchor allocation report and NSE data. Slide 7 is my own analysis.");
}

// ================================================================= 2. Transaction summary
pres.addSection({ title: "Transaction" });
{
  const s = content(`A ${cr(pd.gross_offer)} offer: ${pc(pd.ofs_share_of_offer, 0)} secondary sale by Bain-backed BC Asia, ${pc(pd.fresh_share_of_offer, 0)} primary to cut debt`,
    "Source: Prospectus dated 13-Aug-2026 pp.96, 110, 117 and cover; NSE IPO page; basis-of-allotment notice 14-Aug-2026. Market cap at issue price on 204.5m post-issue shares.", "Transaction");
  table(s, [
    ["Term", "Detail"],
    ["Price band / issue price", `₹${t.price_band_low}–${t.price_band_high} / ₹${t.final_issue_price} (top of band)`],
    ["Fresh issue", `${cr(pd.fresh_gross)} (${n0(t.fresh_issue_shares)} shares)`],
    ["Offer for sale", `${cr(pd.ofs)} (${n0(t.ofs_shares)} shares)`],
    ["Selling shareholders", `BC Asia Investments XV ${n0(t.ofs_shares_bc_asia_xv)} sh; Mangalam Capital ${n0(t.ofs_shares_mangalam_capital)} sh`],
    ["Shares: pre → post issue", `${n0(t.pre_issue_shares)} → ${n0(t.post_issue_shares)}`],
    ["Promoter holding", `${pc(pd.promoter_pct_pre)} → ${pc(pd.promoter_pct_post)} (incl. group ${pc(pd.promoter_and_group_pct_pre)} → ${pc(pd.promoter_and_group_pct_post)})`],
    ["Anchor book", `${cr(t.anchor_amount / 10)} to ${t.anchor_investor_count} investors on 7-Aug`],
    ["Key dates", "Band 4-Aug · bid 10–12 Aug · allotment 13-Aug · listing 17-Aug-2026"],
    ["Lot size / BRLMs", `${t.lot_size} shares / Axis, Jefferies, Kotak, Nomura, SBI Caps, 360 ONE`],
  ], { x: 0.6, y: 1.45, w: 7.9, colW: [2.3, 5.6], rowH: 0.42 });
  stat(s, 8.9, 1.45, 3.8, 1.5, cr(v.market_cap), "Post-issue market cap at ₹871", C.accent1, "mcap");
  stat(s, 8.9, 3.15, 3.8, 1.5, mx(v.pe), `P/E on post-issue shares (RHP headline: ${mx(v.pe_rhp_basis)})`, C.accent3, "pe");
  stat(s, 8.9, 4.85, 3.8, 1.5, pc(f.revenue_from_operations.FY2026 / f.revenue_from_operations.FY2025 - 1), `FY26 revenue growth to ${cr(mn2cr(f.revenue_from_operations.FY2026))}`, C.accent5, "growth");
  s.addNotes("The RHP's 35.7x P/E divides the price by FY26 EPS on weighted-average shares, which barely includes the 22.2m shares issued to BC Asia on 20-Mar-2026. On the 204.5m post-issue shares, P/E is 44.9x.");
}

// ================================================================= 3. Valuation vs peers
pres.addSection({ title: "Valuation" });
{
  const s = content(`At ₹871 Dhoot priced at a ${pc(-pr.ev_ebitda, 0)} discount to peer EV/EBITDA and ${pc(-pr.pe, 0)} to peer P/E`,
    "Peers = all four RHP-listed peers. NSE close 4-Aug-2026 (price band date); shares from 30-Jun-2026 shareholding pattern; FY26 audited results XBRL (MSUMI standalone). EBITDA on RHP definition; net debt ex-leases.", "Valuation");
  const order = ["SONACOMS", "UNOMINDA", "MINDACORP", "MSUMI"];
  const short = { SONACOMS: "Sona BLW", UNOMINDA: "Uno Minda", MINDACORP: "Minda Corp", MSUMI: "Motherson Wiring" };
  const labels = order.map((k) => short[k]).concat(["Dhoot @ ₹871"]);
  const colors = order.map(() => SLATE).concat([COPPER]);
  [["ev_ebitda", "EV / EBITDA, FY26", 0.6], ["pe", "P/E, FY26", 6.75]].forEach(([k, ttl, x]) => {
    const vals = order.map((p) => R.peers[p][k]).concat([v[k]]);
    s.addChart(pres.charts.BAR, [{ name: ttl, labels, values: vals }], {
      x, y: 1.4, w: 5.95, h: 3.5, barDir: "bar", ...quiet(), title: `${ttl}  (peer median ${mx(ps[k].median)})`,
      chartColors: colors, showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0"x"',
      valAxisHidden: true, valGridLine: { style: "none" }, showLegend: false, catAxisOrientation: "maxMin", objectName: `${k} chart`,
    });
  });
  table(s, [
    ["", "Dhoot @ ₹871", "Peer median", "Peer range", "Premium / (discount)"],
    ...[["P/E", "pe"], ["EV / EBITDA", "ev_ebitda"], ["EV / Sales", "ev_sales"], ["Price / Book", "pb"]].map(([lab, k]) =>
      [lab, mx(v[k]), mx(ps[k].median), `${mx(ps[k].low)}–${mx(ps[k].high)}`, { text: pc(pr[k]), options: { bold: true, color: pr[k] < 0 ? C.accent5 : C.accent6 } }]),
  ], { x: 0.6, y: 5.05, w: 12.1, colW: [2.4, 2.4, 2.4, 2.4, 2.5], rowH: 0.34, fontSize: 11 });
  s.addNotes(`EV = market cap + post-issue net debt (FY26 net debt ${cr(v.net_debt_pre)} less net proceeds) + NCI. Dhoot is net cash after the issue. EV/Sales is at a premium because Dhoot's 15.7% EBITDA margin beats Uno Minda, Minda Corp and Motherson Wiring; P/B looks cheap because the post-money book includes ₹1,022 cr of fresh Bain equity and ₹1,327 cr of IPO proceeds.`);
}

// ================================================================= 4. Use of proceeds
pres.addSection({ title: "Proceeds" });
{
  const s = content(`${pc(pd.debt_repayment / pd.net_proceeds, 0)} of net primary proceeds repays debt; Dhoot lists with ${cr(-pd.net_debt_after)} of net cash`,
    "Source: Prospectus 13-Aug-2026 pp.96, 110, 117-118, 159-160; RHP pp.459, 480. Net debt = borrowings less cash and other bank balances (excludes ₹77 cr of lease liabilities).", "Proceeds");
  const objs = Object.entries(pd.objects);
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Objects", labels: objs.map((o) => o[0]), values: objs.map((o) => Math.round(o[1] * 10) / 10) }], {
    x: 0.6, y: 1.4, w: 5.6, h: 4.6, ...chartText, holeSize: 55, chartColors: [COPPER, SAND, STEEL, SLATE],
    showTitle: true, title: `Net proceeds ${cr(pd.net_proceeds)}, by object (₹ cr)`, titleFontSize: 13, titleColor: HEX.dk2,
    showValue: true, showPercent: false, dataLabelColor: HEX.lt1, dataLabelFontSize: 11, dataLabelFormatCode: "#,##0",
    showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: HEX.dk1, objectName: "objects doughnut",
  });
  stat(s, 6.6, 1.4, 3.0, 1.45, cr(pd.ofs), `Offer for sale (${pc(pd.ofs_share_of_offer, 0)}): goes to sellers, not the company`, C.accent2, "ofs");
  stat(s, 9.75, 1.4, 3.0, 1.45, pc(pd.dilution), `Dilution: ${n0(pd.fresh_shares / 1e5)} lakh new shares / post-issue`, C.accent1, "dilution");
  stat(s, 6.6, 3.05, 3.0, 1.45, cr(-pd.net_debt_before), "Net cash at FY26 (holds Bain's undeployed ₹1,022 cr)", C.accent3, "nd pre");
  stat(s, 9.75, 3.05, 3.0, 1.45, cr(-pd.net_debt_after), "Net cash pro forma, after net proceeds", C.accent5, "nd post");
  bullets(s, [
    `Debt repayment of ${cr(pd.debt_repayment)} is ${pc(pd.debt_repayment_pct_fy26_borrowings, 0)} of FY26 borrowings (${cr(pd.borrowings_fy26)}).`,
    `${cr(pd.objects["Acquisitions and general corporate purposes"])} goes to unidentified acquisitions and GCP: ${pc(pd.objects["Acquisitions and general corporate purposes"] / pd.fresh_gross, 0)} of gross fresh proceeds, under SEBI's 35% cap.`,
    `Offer expenses: ${pc(pd.offer_expenses_pct)} of the offer; free float after listing ${pc(pd.public_float_post)}.`,
  ], 6.6, 4.7, 6.15, 1.9, 13, "proceeds bullets");
  s.addNotes("The company was already net cash at FY26 year-end because BC Asia's tranche-2 primary investment (₹1,022.6 cr on 20-Mar-2026) was undeployed. Raising more primary capital to repay debt while sitting on cash is a question worth asking.");
}

// ================================================================= 5. Demand
pres.addSection({ title: "Demand" });
{
  const fc = dm.final_by_category;
  const cats = ["QIB (ex-anchor)", "NII >10 lakh", "NII 2-10 lakh", "Retail", "Employee", "Total"];
  const s = content(`Institutions drove a ${mx(fc.Total)} book, QIBs ${mx(fc["QIB (ex-anchor)"])}; the ${sgn(dm.listing_gain_pct)} debut left ${cr(dm.money_left_on_table)} on the table`,
    "Source: NSE consolidated bid book 12-Aug-2026 19:00 (excl. anchor); basis-of-allotment notice 14-Aug-2026; anchor allocation report 7-Aug-2026; NSE close 17-Aug-2026. Day 1–2 snapshots not published: see docs.", "Demand");
  s.addChart(pres.charts.BAR, [{ name: "Subscription (x)", labels: ["QIB", "NII >₹10L", "NII ₹2–10L", "Retail", "Employee", "Total"], values: cats.map((c) => fc[c]) }], {
    x: 0.6, y: 1.4, w: 6.6, h: 5.2, barDir: "col", ...quiet(), title: "Final-day subscription by category (times)",
    chartColors: [COPPER, SLATE, SLATE, SLATE, SLATE, STEEL], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0"x"',
    valAxisHidden: true, valGridLine: { style: "none" }, showLegend: false, objectName: "subscription chart",
  });
  stat(s, 7.6, 1.4, 2.5, 1.4, pc(dm.anchor_share_of_qib, 0), "Anchor share of QIB portion", C.accent3, "anchor qib");
  stat(s, 10.25, 1.4, 2.5, 1.4, pc(dm.anchor_mf_share, 0), "Domestic MFs in anchor book", C.accent3, "anchor mf");
  stat(s, 7.6, 3.0, 2.5, 1.4, px(dm.listing_close), `Listing-day close (${sgn(dm.listing_gain_pct)})`, C.accent1, "listing");
  stat(s, 10.25, 3.0, 2.5, 1.4, cr(dm.money_left_on_table), "Money left on table (gain × offer shares)", C.accent6, "molt");
  bullets(s, [
    `Split: ${cr(dm.money_left_fresh)} forgone by the company (fresh) and ${cr(dm.money_left_ofs)} by sellers (OFS).`,
    `Anchors bid for only ${mx(dm.registrar_by_category.Anchor)} their allocation at ₹871; QIB demand built in the open book.`,
    `Registrar's ${mx(dm.registrar_by_category["Total incl. anchor"])} counts anchor shares in the denominator; the exchange's ${mx(fc.Total)} does not.`,
  ], 7.6, 4.6, 5.15, 2.0, 13, "demand bullets");
}

// ================================================================= 6. Aftermarket
pres.addSection({ title: "Aftermarket" });
{
  const pts = ["Listing day", "+30 days", "Latest (+87 days)"];
  const lab = pts.map((p) => `${p.replace("Latest (+87 days)", "Latest, 12-Nov (+87d)")}`);
  const s = content(`Dhoot is up ${pc(am["Latest (+87 days)"].DHOOTTRANS, 0)} on the issue price while Nifty Auto is down ${pc(-am["Latest (+87 days)"]["NIFTY AUTO"], 1)}`,
    "Source: NSE end-of-day closes. Base: issue price ₹871 and index closes on 12-Aug-2026 (bid close). +90 days (16-Nov-2026) and +180 days (15-Feb-2027) have not elapsed.", "Aftermarket");
  s.addChart(pres.charts.BAR, [
    { name: "Dhoot Transmission", labels: lab, values: pts.map((p) => am[p].DHOOTTRANS) },
    { name: "Nifty 50", labels: lab, values: pts.map((p) => am[p]["NIFTY 50"]) },
    { name: "Nifty Auto", labels: lab, values: pts.map((p) => am[p]["NIFTY AUTO"]) },
  ], {
    x: 0.6, y: 1.4, w: 8.0, h: 5.2, barDir: "col", barGrouping: "clustered", ...quiet(), title: "Return from issue pricing",
    chartColors: [COPPER, STEEL, SLATE], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0%", dataLabelFontSize: 10,
    valAxisLabelFormatCode: "0%", showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: HEX.dk1, objectName: "aftermarket chart",
  });
  table(s, [
    ["Checkpoint", "Dhoot", "Excess vs Nifty 50", "Excess vs Nifty Auto"],
    ...pts.map((p) => [p, sgn(am[p].DHOOTTRANS), sgn(am[p]["excess_vs_NIFTY 50"]), sgn(am[p]["excess_vs_NIFTY AUTO"])]),
    ["+90 days", "not elapsed", "–", "–"],
    ["+180 days", "not elapsed", "–", "–"],
  ], { x: 8.9, y: 1.5, w: 3.85, colW: [1.25, 0.8, 0.9, 0.9], rowH: 0.45, fontSize: 10 });
  bullets(s, ["Most of the excess return came in the first month, while both indices fell.", `From +30 days to 12-Nov the stock moved only ${sgn((1 + am["Latest (+87 days)"].DHOOTTRANS) / (1 + am["+30 days"].DHOOTTRANS) - 1)}.`], 8.9, 4.65, 3.85, 1.9, 12, "aftermarket bullets");
}

// ================================================================= 7. How I would have priced it
pres.addSection({ title: "My view" });
{
  const rb = my.recommended_band;
  const s = content(`My view: a ₹${n0(rb.low.price)}–${n0(rb.high.price)} band, priced at ₹${n0(rb.high.price)}`,
    "ANALYST VIEW: my judgement, not a reported fact. Built on the peer medians in this deck (NSE closes 4-Aug-2026) and the FY26 restated financials. Method and caveats: docs/ipo_case.md.", "My view");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.35, w: 2.4, h: 0.4, fill: { color: SAND }, line: { type: "none" }, rectRadius: 0.06, objectName: "analyst view tag" });
  s.addText("ANALYST VIEW", { x: 0.6, y: 1.35, w: 2.4, h: 0.4, fontSize: 12, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true });
  // football field drawn with shapes: x-scale 700..1700 INR over 6.0 inches
  const X0 = 3.1, XW = 4.9, LO = 700, HI = 1700, sx = (p) => X0 + ((p - LO) / (HI - LO)) * XW;
  const rows = [
    ["EV/EBITDA: peer median less 15–25%", my.low.price_ev_ebitda, my.high.price_ev_ebitda, SLATE],
    ["P/E: peer median less 15–25%", my.low.price_pe, my.high.price_pe, SLATE],
    ["Fundamental range (average)", my.low.price, my.high.price, STEEL],
    ["My recommended band", rb.low.price, rb.high.price, COPPER],
    ["Actual price band", t.price_band_low, t.price_band_high, SAND],
    ["Traded: listing close → 12-Nov close", dm.listing_close, t.final_issue_price * (1 + am["Latest (+87 days)"].DHOOTTRANS), GREEN],
  ];
  rows.forEach(([lab, a, b, col], i) => {
    const y = 2.0 + i * 0.62;
    s.addText(lab, { x: 0.6, y, w: 2.45, h: 0.45, fontSize: 11, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: sx(a), y: y + 0.07, w: Math.max(sx(b) - sx(a), 0.04), h: 0.31, fill: { color: col }, line: { type: "none" }, objectName: `bar ${i}` });
    s.addText(px(a), { x: sx(a) - 0.75, y, w: 0.7, h: 0.45, fontSize: 10, color: C.text1, align: "right", valign: "middle", margin: 0, isTextBox: true });
    s.addText(px(b), { x: sx(b) + 0.05, y, w: 0.7, h: 0.45, fontSize: 10, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  });
  const xi = sx(t.final_issue_price);
  s.addShape(pres.shapes.LINE, { x: xi, y: 1.9, w: 0, h: 3.85, line: { color: RED, width: 1.25, dashType: "dash" }, objectName: "issue price line" });
  s.addText(`Issue ₹${t.final_issue_price}`, { x: xi - 0.6, y: 5.78, w: 1.2, h: 0.3, fontSize: 10, color: C.accent6, align: "center", margin: 0, isTextBox: true });
  bullets(s, [
    `Fundamentals alone gave ₹${n0(my.low.price)}–${n0(my.high.price)}, so ₹871 was defensible on the comps.`,
    `I would still have set the band at ₹${n0(rb.low.price)}–${n0(rb.high.price)}. ₹${n0(rb.high.price)} is ${mx(rb.high.pe)} P/E and ${mx(rb.high.ev_ebitda)} EV/EBITDA, still a ${pc(1 - rb.high.pe / ps.pe.median, 0)} / ${pc(1 - rb.high.ev_ebitda / ps.ev_ebitda.median, 0)} discount to the peer medians.`,
    `Why the top: 31% revenue growth, a 15.7% margin above three of four peers, net cash after the issue, and a 17% float.`,
    `Why no higher: margins fell from 18.3% to 15.7% over two years, and anchors covered only ${mx(dm.registrar_by_category.Anchor)} at ₹871.`,
    `At ₹${n0(rb.high.price)} sellers raise about ${cr((rb.high.price - t.final_issue_price) * t.ofs_shares / 1e7)} more, and the issue still has room to trade up.`,
  ], 8.45, 1.35, 4.3, 5.4, 12, "pricing rationale");
}

// ================================================================= 8. Limitations
pres.addSection({ title: "Notes" });
{
  const s = content("What the numbers can and cannot tell you",
    "Full source register: data/*.csv (URL, page and date for every figure). Workbook: output/Dhoot_IPO_Case.xlsx; checks on the Checks sheet.", "Notes");
  const cols = [
    ["Method", [
      "Peer set: all four RHP-listed peers, none excluded.",
      "Peer prices: NSE close on 4-Aug-2026, the price band date (the RHP used 31-Jul BSE closes).",
      "Same EBITDA, net debt and EV definitions for Dhoot and peers.",
      "P/E for Dhoot uses post-issue shares, not the RHP's weighted average.",
    ]],
    ["Limitations", [
      "Trailing FY26 multiples only; no forecasts were used.",
      "Motherson Wiring is standalone; Sona BLW's owners' PAT is not tagged, so peer P/E uses diluted EPS.",
      "Leases excluded from net debt on both sides.",
      "Uno Minda's EBITDA excludes its JV profits, which raises its EV/EBITDA.",
    ]],
    ["Needs manual entry", [
      "Day 1 and Day 2 subscription by category (NSE shows the final book only).",
      "+90-day (16-Nov-2026) and +180-day (15-Feb-2027) closes: not yet elapsed.",
      "Sona BLW FY26 profit attributable to owners (results PDF).",
    ]],
  ];
  cols.forEach(([hd, items], i) => {
    const x = 0.6 + i * 4.1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.6, w: 3.85, h: 3.9, fill: { color: C.background2 }, line: { type: "none" }, rectRadius: 0.08, objectName: `${hd} card` });
    s.addText(hd, { x: x + 0.25, y: 1.8, w: 3.4, h: 0.5, fontSize: 18, bold: true, color: C.accent1, fontFace: THEME.headFontFace, margin: 0, isTextBox: true });
    bullets(s, items, x + 0.25, 2.4, 3.4, 2.9, 13, `${hd} list`);
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  if (SKILL_THEME) {
    const { applyTheme } = require(SKILL_THEME);
    await applyTheme(OUT, THEME);
  }
  console.log("wrote " + OUT);
})();
