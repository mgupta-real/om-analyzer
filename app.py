import os, json, re, tempfile, io
import streamlit as st
import anthropic
import pdfplumber
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, HRFlowable, PageBreak
)

# ── Page setup ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="OM Analyzer",
    page_icon="🏢",
    layout="wide",
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
#MainMenu, footer, header {visibility: hidden;}
.stApp { background: #F5F4EF; }
.block-container { padding-top: 1.5rem !important; max-width: 1100px !important; }

section[data-testid="stSidebar"] { background: #1A1A18 !important; }
section[data-testid="stSidebar"] * { color: #C8C8B8 !important; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #D4B07A !important; }
section[data-testid="stSidebar"] .stMarkdown p { font-size: 13px; }

div[data-testid="metric-container"] {
    background: white; border: 1px solid #E0DED5;
    border-radius: 10px; padding: 12px 16px;
}
div[data-testid="metric-container"] label { color: #888 !important; font-size: 12px !important; }
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-size: 22px !important; color: #1A1A18 !important; font-weight: 500 !important;
}

.gold-header {
    background: #1A1A18; color: #D4B07A; padding: 6px 14px;
    border-radius: 6px; font-size: 13px; font-weight: 500;
    margin: 18px 0 8px; display: inline-block;
}
.flag-warn  { background:#FFF8EC; border-left:3px solid #D4A054; padding:10px 14px; border-radius:0 6px 6px 0; margin:6px 0; }
.flag-good  { background:#EAF5EE; border-left:3px solid #5AAA7A; padding:10px 14px; border-radius:0 6px 6px 0; margin:6px 0; }
.flag-info  { background:#EAF0FA; border-left:3px solid #5A8AC0; padding:10px 14px; border-radius:0 6px 6px 0; margin:6px 0; }
.flag-verify{ background:#F5F0FF; border-left:3px solid #9A7ACA; padding:10px 14px; border-radius:0 6px 6px 0; margin:6px 0; }
.flag-title { font-size:13px; font-weight:600; margin-bottom:3px; }
.flag-body  { font-size:12px; color:#555; line-height:1.5; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — PDF TEXT EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════
def extract_pdf_text(path: str) -> str:
    pages = []
    try:
        with pdfplumber.open(path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                if text:
                    pages.append(f"\n--- PAGE {i+1} ---\n{text}")
                for tbl in (page.extract_tables() or []):
                    if tbl:
                        rows = ["\t".join(str(c or "") for c in row) for row in tbl if row]
                        pages.append("[TABLE]\n" + "\n".join(rows) + "\n[/TABLE]")
        return "\n".join(pages)
    except Exception:
        try:
            from pypdf import PdfReader
            reader = PdfReader(path)
            return "\n".join(
                f"--- PAGE {i+1} ---\n{p.extract_text() or ''}"
                for i, p in enumerate(reader.pages)
            )
        except Exception as e:
            raise RuntimeError(f"Could not read PDF: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — AI ANALYSIS (MULTI-TURN AGENT)
# ══════════════════════════════════════════════════════════════════════════════

SYSTEM = """You are a senior multifamily real estate underwriting analyst with 20+ years of experience.
You read Offering Memoranda from any broker (JLL, CBRE, Marcus & Millichap, Cushman & Wakefield,
Newmark, Colliers, Berkadia, Walker & Dunlop, Eastdil, HFF, and boutique firms) and extract
comprehensive underwriting data.

Return ONLY a single valid JSON object — no markdown fences, no explanation text, nothing else.
For missing fields use null. For empty arrays use [].

SCHEMA:
{
  "broker": {"name": null, "agents": [], "date": null},
  "property": {
    "name": null, "address": null, "city": null, "state": null, "zip": null,
    "county": null, "msa": null, "units": null, "year_built": null,
    "rentable_sf": null, "avg_unit_sf": null, "buildings": null, "floors": null,
    "acres": null, "density": null, "occupancy_pct": null, "occupancy_date": null,
    "developer": null, "asset_class": null
  },
  "investment": {
    "market_rent": null, "market_rent_psf": null,
    "effective_rent": null, "effective_rent_psf": null,
    "comp_market_rent": null, "comp_effective_rent": null,
    "rent_gap": null, "annual_upside": null,
    "replacement_cost_note": null, "rent_growth_note": null,
    "renovation_tiers": [
      {"name": null, "units": null, "description": null,
       "appliances": null, "cabinets": null, "countertops": null,
       "flooring": null, "backsplash": null, "faucets": null,
       "lighting": null, "sinks": null, "premium": null}
    ],
    "exterior_opportunity": null,
    "additional_income": [{"name": null, "current": null, "proforma": null, "notes": null}],
    "amenities": [],
    "unit_features": [],
    "highlights": []
  },
  "demographics": {
    "pop_1mi": null, "pop_3mi": null, "pop_5mi": null, "pop_growth": null,
    "median_income_1mi": null, "median_income_3mi": null, "avg_income_3mi": null,
    "income_growth": null, "home_value": null, "home_value_area": null,
    "crime": null, "school_district": null, "elementary": null,
    "middle": null, "high_school": null, "college_pct": null, "median_age": null,
    "employers": [{"name": null, "drive": null, "employees": null, "sector": null, "notes": null}]
  },
  "unit_mix": [
    {"type": null, "plan": null, "count": null, "pct": null, "sf": null,
     "market_rent": null, "market_psf": null, "eff_rent": null, "eff_psf": null,
     "target_rent": null, "upside": null, "occupied": null, "vacant": null}
  ],
  "utilities": [
    {"name": null, "method": null, "paid_by": null,
     "reimbursement": null, "fee": null, "annual_income": null, "notes": null}
  ],
  "site": {
    "roof": null, "roof_age": null, "exterior": null, "foundation": null,
    "hvac": null, "plumbing": null, "wiring": null, "hot_water": null,
    "washer_dryer": null, "life_safety": null,
    "parking_open": null, "parking_reserved": null, "parking_covered": null,
    "parking_garage": null, "parking_total": null, "parking_ratio": null,
    "reserved_fee": null, "pet_yards": null, "storage": null, "notes": null
  },
  "rent_comps": [
    {"id": null, "name": null, "address": null, "city_state": null,
     "year_built": null, "units": null, "occupancy": null, "avg_sf": null,
     "total_market": null, "total_market_psf": null,
     "total_eff": null, "total_eff_psf": null,
     "by_bed": [
       {"type": null, "units": null, "sf": null,
        "market": null, "market_psf": null, "eff": null, "eff_psf": null}
     ]}
  ],
  "financials": {
    "periods": [],
    "income_lines": [
      {"item": null, "is_total": false, "is_subtotal": false, "is_deduction": false,
       "values": {}, "pct": {}, "note": null}
    ],
    "expense_lines": [
      {"item": null, "is_total": false, "is_subtotal": false,
       "values": {}, "per_unit": {}, "note": null}
    ],
    "noi": {}, "capex": {}, "cffo": {}, "expense_ratio": {}
  },
  "sale_comps": [
    {"name": null, "address": null, "city_state": null, "date": null,
     "year_built": null, "units": null, "price": null, "ppu": null,
     "ppsf": null, "cap_rate": null, "occupancy": null, "notes": null}
  ],
  "market": {
    "submarket": null, "sub_occupancy": null, "sub_rent": null,
    "sub_growth": null, "metro_inventory": null, "metro_occupancy": null,
    "pipeline": null, "absorption": null, "investment_vol": null
  },
  "flags": [
    {"category": null, "title": null, "detail": null}
  ]
}

RULES:
- Extract every number that exists anywhere in the text.
- For new construction (< 3 yrs): note single finish level, describe it fully.
- Capture every financial line item exactly as shown.
- Generate 5-8 flags covering risks, opportunities, and verify items.
- category values: Warning / Opportunity / Verify / Info
"""


def analyze_om(pdf_text: str, api_key: str, progress_cb=None) -> dict:
    client = anthropic.Anthropic(api_key=api_key)

      # Chunk large documents — keep at 140K to maximise context sent to Claude
    MAX = 140_000
    text = pdf_text[:MAX]
    if len(pdf_text) > MAX and progress_cb:
        progress_cb("Large OM detected — using first 140K characters (covers most OMs fully)")
 
    if progress_cb:
        progress_cb("Sending to Claude AI for analysis...")
 
    resp = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=16000,          # raised from 8000 — gives Claude enough room to finish the JSON
        system=SYSTEM,
        messages=[{"role": "user", "content": f"Analyze this OM:\n\n{text}"}]
    )
 
    raw = resp.content[0].text.strip()
    # Strip any accidental markdown fences
    raw = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("`").strip()
 
    if progress_cb:
        progress_cb("Parsing extracted data...")
 
    # Attempt 1 — clean parse
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass
 
    # Attempt 2 — extract JSON object if there is leading/trailing noise
    try:
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if m:
            return json.loads(m.group())
    except json.JSONDecodeError:
        pass
 
    # Attempt 3 — response was still truncated; close open structures and retry
    try:
        fixed = _fix_truncated_json(raw)
        return json.loads(fixed)
    except Exception:
        raise ValueError("Could not parse AI response. Please try again — this occasionally happens with very large OMs.")
 
 
def _fix_truncated_json(raw: str) -> str:
    """Close any open strings, arrays and objects left by a truncated JSON response."""
    # First pass — close any open string
    in_string = False
    escape_next = False
    for ch in raw:
        if escape_next:
            escape_next = False
            continue
        if ch == "\\":
            escape_next = True
            continue
        if ch == '"':
            in_string = not in_string
    if in_string:
        raw += '"'
 
    # Second pass — close unmatched brackets / braces
    opens = []
    in_str = False
    esc = False
    for ch in raw:
        if esc:
            esc = False
            continue
        if ch == "\\":
            esc = True
            continue
        if ch == '"':
            in_str = not in_str
            continue
        if not in_str:
            if ch in "{[":
                opens.append("}" if ch == "{" else "]")
            elif ch in "}]" and opens:
                opens.pop()
 
    raw += "".join(reversed(opens))
    return raw
# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — PDF REPORT GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

DARK   = colors.HexColor("#1A1A18")
DARK2  = colors.HexColor("#252522")
GOLD   = colors.HexColor("#B8965A")
GOLD_L = colors.HexColor("#D4B07A")
GOLD_P = colors.HexColor("#F5EDD8")
MID    = colors.HexColor("#777770")
ALT    = colors.HexColor("#F5F4EF")
BORDER = colors.HexColor("#E0DED5")
GREEN  = colors.HexColor("#1E7A4A")
GREEN_B= colors.HexColor("#EAF5EE")
RED    = colors.HexColor("#C0392B")
RED_B  = colors.HexColor("#FAEAEA")
BLUE   = colors.HexColor("#1B4F8A")
BLUE_B = colors.HexColor("#EAF0FA")
WARN   = colors.HexColor("#8A5A1B")
WARN_B = colors.HexColor("#FFF8EC")
PURPLE_B=colors.HexColor("#F5F0FF")
PURPLE = colors.HexColor("#6A3ABA")


def _v(val, fmt=None, suffix="", default="N/A"):
    if val is None or val == "":
        return default
    if fmt == "$":
        try: return f"${float(val):,.0f}{suffix}"
        except: return str(val)
    if fmt == "%":
        try: return f"{float(val):.1f}%"
        except: return str(val)
    if fmt == "n":
        try: return f"{int(float(val)):,}{suffix}"
        except: return str(val)
    return f"{val}{suffix}"


def _st(fontname="Helvetica", size=9, color=DARK, leading=13, space_before=0, space_after=2, align=0):
    return ParagraphStyle("x", fontName=fontname, fontSize=size, textColor=color,
                          leading=leading, spaceBefore=space_before, spaceAfter=space_after, alignment=align)


def _p(text, **kw):
    return Paragraph(str(text) if text else "N/A", _st(**kw))


def _hdr(story, title, num):
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=2, color=GOLD, spaceAfter=4))
    t = Table([[f"{num}. {title}"]], colWidths=[7.5*inch])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,-1), colors.white),
        ("TEXTCOLOR",  (0,0),(-1,-1), DARK),
        ("FONTNAME",   (0,0),(-1,-1), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0),(-1,-1), 13),
        ("TOPPADDING", (0,0),(-1,-1), 2),
        ("BOTTOMPADDING",(0,0),(-1,-1), 4),
    ]))
    story.append(t)
    story.append(Spacer(1, 4))


def _kv(rows, cw=None):
    cw = cw or [2.8*inch, 4.7*inch]
    data = [[_p(k, color=MID, size=8.5), _p(v or "N/A")] for k, v in rows]
    t = Table(data, colWidths=cw)
    t.setStyle(TableStyle([
        ("ROWBACKGROUNDS", (0,0),(-1,-1), [colors.white, ALT]),
        ("GRID",  (0,0),(-1,-1), 0.3, BORDER),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ("RIGHTPADDING",  (0,0),(-1,-1), 8),
        ("VALIGN", (0,0),(-1,-1), "TOP"),
    ]))
    return t


def _tbl(headers, rows, cw=None, hl_last=False):
    head = [_p(h, fontname="Helvetica-Bold", size=8, color=GOLD_L) for h in headers]
    data = [head] + [[_p(str(c) if c is not None else "—", size=8.5) for c in row] for row in rows]
    cw = cw or [7.5*inch/len(headers)]*len(headers)
    t = Table(data, colWidths=cw, repeatRows=1)
    ts = [
        ("BACKGROUND",(0,0),(-1,0), DARK),
        ("FONTSIZE",  (0,0),(-1,-1), 8),
        ("TOPPADDING",(0,0),(-1,-1), 5), ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LEFTPADDING",(0,0),(-1,-1), 6), ("RIGHTPADDING",(0,0),(-1,-1), 6),
        ("GRID",(0,0),(-1,-1), 0.3, BORDER),
        ("VALIGN",(0,0),(-1,-1), "TOP"),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [colors.white, ALT]),
    ]
    if hl_last and len(data) > 1:
        ts += [("BACKGROUND",(0,-1),(-1,-1), DARK2),
               ("TEXTCOLOR",(0,-1),(-1,-1), GOLD_L),
               ("FONTNAME",(0,-1),(-1,-1), "Helvetica-Bold")]
    t.setStyle(TableStyle(ts))
    return t


def _footer(prop, date):
    def draw(canvas, doc):
        canvas.saveState()
        w, _ = letter
        canvas.setStrokeColor(GOLD); canvas.setLineWidth(0.5)
        canvas.line(0.75*inch, 0.65*inch, w-0.75*inch, 0.65*inch)
        canvas.setFont("Helvetica", 7); canvas.setFillColor(MID)
        canvas.drawString(0.75*inch, 0.5*inch, f"{prop} — Underwriting Report")
        canvas.drawRightString(w-0.75*inch, 0.5*inch, f"Page {doc.page}  |  {date}")
        canvas.restoreState()
    return draw


def build_pdf(d: dict, filename: str) -> bytes:
    from datetime import datetime
    date = datetime.today().strftime("%B %d, %Y")
    prop = (d.get("property") or {}).get("name") or "Property"
    broker_d = d.get("broker") or {}
    inv = d.get("investment") or {}
    demo = d.get("demographics") or {}
    umix = d.get("unit_mix") or []
    utils = d.get("utilities") or []
    site = d.get("site") or {}
    rcomps = d.get("rent_comps") or []
    fin = d.get("financials") or {}
    scomps = d.get("sale_comps") or []
    flags = d.get("flags") or []
    mkt = d.get("market") or {}
    pd = d.get("property") or {}

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
        leftMargin=0.75*inch, rightMargin=0.75*inch,
        topMargin=0.75*inch, bottomMargin=0.9*inch,
        title=f"{prop} — Underwriting Report")
    story = []

    addr = f"{pd.get('address','')}, {pd.get('city','')}, {pd.get('state','')} {pd.get('zip','')}".strip(", ")

    def dark_row(content, bg=DARK, top=8, bot=8, left=32):
        t = Table([[content]], colWidths=[7.5*inch])
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),bg),
            ("TOPPADDING",(0,0),(-1,-1),top), ("BOTTOMPADDING",(0,0),(-1,-1),bot),
            ("LEFTPADDING",(0,0),(-1,-1),left)]))
        return t

    # Cover
    story.append(dark_row(_p("MULTIFAMILY UNDERWRITING REPORT", fontname="Helvetica", size=9, color=GOLD), top=40, bot=6))
    story.append(dark_row(_p(prop, fontname="Helvetica-Bold", size=26, color=colors.white), top=4, bot=8))
    meta = f"Units: {_v(pd.get('units'),'n')}  ·  Year Built: {_v(pd.get('year_built'))}  ·  Occupancy: {_v(pd.get('occupancy_pct'),'%')}  ·  Broker: {broker_d.get('name','N/A')}"
    story.append(dark_row(_p(meta, size=10, color=GOLD_L), bg=DARK2, top=10, bot=8))
    story.append(dark_row(_p(addr, size=11, color=MID), bg=DARK2, top=4, bot=24))
    story.append(Spacer(1,14))
    story.append(_p(f"Report Generated: {date}", color=MID, size=9))
    story.append(_p(f"Source: {filename}", color=MID, size=9))
    story.append(PageBreak())

    # 1 — Basic Details
    _hdr(story, "Basic Property Details", "1")
    story.append(_kv([
        ("Property Name", prop), ("Address", addr),
        ("County", pd.get("county")), ("MSA", pd.get("msa")),
        ("Total Units", _v(pd.get("units"),"n")),
        ("Year Built", _v(pd.get("year_built"))),
        ("Net Rentable Area", _v(pd.get("rentable_sf"),"n",suffix=" SF")),
        ("Avg Unit Size", _v(pd.get("avg_unit_sf"),"n",suffix=" SF")),
        ("Occupancy", f"{_v(pd.get('occupancy_pct'),'%')} as of {pd.get('occupancy_date','N/A')}"),
        ("Buildings / Stories", f"{_v(pd.get('buildings'),'n')} bldgs / {_v(pd.get('floors'))} stories"),
        ("Land Area", _v(pd.get("acres"),suffix=" acres")),
        ("Density", _v(pd.get("density"),suffix=" units/acre")),
        ("Developer", pd.get("developer")),
        ("Asset Class", pd.get("asset_class")),
        ("Broker", broker_d.get("name")),
        ("Listing Agents", ", ".join(broker_d.get("agents") or []) or "N/A"),
    ]))

    # 2 — Investment Highlights
    story.append(PageBreak())
    _hdr(story, "Investment Highlights & Opportunities", "2")

    story.append(_p("Rent Overview", fontname="Helvetica-Bold", size=11, space_before=6, space_after=4))
    story.append(_kv([
        ("Current Market Rent", f"{_v(inv.get('market_rent'),'$')} ({_v(inv.get('market_rent_psf'))} PSF)"),
        ("Current Effective Rent", f"{_v(inv.get('effective_rent'),'$')} ({_v(inv.get('effective_rent_psf'))} PSF)"),
        ("Comp Set Effective Rent", _v(inv.get("comp_effective_rent"),"$")),
        ("Rent Gap / Unit", _v(inv.get("rent_gap"),"$")),
        ("Annual Upside (total)", _v(inv.get("annual_upside"),"$")),
        ("Below Replacement Cost", inv.get("replacement_cost_note")),
    ]))

    story.append(_p("Renovation Tiers", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    tiers = inv.get("renovation_tiers") or []
    if tiers:
        story.append(_tbl(
            ["Tier","Units","Appliances","Cabinets","Countertops","Flooring","Premium vs Classic"],
            [[t.get("name","—"), _v(t.get("units"),"n"),
              t.get("appliances") or "—", t.get("cabinets") or "—",
              t.get("countertops") or "—", t.get("flooring") or "—",
              _v(t.get("premium"),"$") if t.get("premium") else "—"] for t in tiers],
            cw=[1.0*inch,0.6*inch,1.3*inch,1.1*inch,1.2*inch,1.2*inch,1.1*inch]
        ))
        for t in tiers:
            extras = [(k.title(), t.get(k)) for k in ("backsplash","faucets","lighting","sinks","description") if t.get(k)]
            if extras:
                story.append(_p(f"{t.get('name','Finish')} — additional specs", fontname="Helvetica-Bold", size=9.5, space_before=6))
                story.append(_kv(extras, cw=[1.4*inch,6.1*inch]))
    else:
        story.append(_p("No renovation tiers in OM. Property may be new construction with single finish level.", color=WARN))

    story.append(_p("Exterior Renovation Opportunity", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_p(inv.get("exterior_opportunity") or "Not disclosed in OM.", color=MID if not inv.get("exterior_opportunity") else DARK))

    story.append(_p("Additional Income Opportunities", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    add_inc = inv.get("additional_income") or []
    if add_inc:
        story.append(_tbl(
            ["Income Item","Current Annual","Pro Forma Annual","Notes"],
            [[i.get("name","—"), _v(i.get("current"),"$"), _v(i.get("proforma"),"$"), i.get("notes") or "—"] for i in add_inc],
            cw=[1.8*inch,1.2*inch,1.4*inch,3.1*inch]
        ))

    story.append(_p("Community Amenities", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    if inv.get("amenities"):
        story.append(_p("  ·  ".join(inv["amenities"])))
    story.append(_p("Unit Features", fontname="Helvetica-Bold", size=11, space_before=8, space_after=4))
    if inv.get("unit_features"):
        story.append(_p("  ·  ".join(inv["unit_features"])))
    if inv.get("highlights"):
        story.append(_p("Other Highlights", fontname="Helvetica-Bold", size=11, space_before=8, space_after=4))
        for h in inv["highlights"]:
            story.append(_p(f"• {h}"))

    # 3 — Demographics
    story.append(PageBreak())
    _hdr(story, "Demographics", "3")
    story.append(_tbl(
        ["Metric","1-Mile","3-Mile","5-Mile"],
        [
            ["Population",        _v(demo.get("pop_1mi"),"n"),  _v(demo.get("pop_3mi"),"n"),  _v(demo.get("pop_5mi"),"n")],
            ["Median HH Income",  _v(demo.get("median_income_1mi"),"$"), _v(demo.get("median_income_3mi"),"$"), "—"],
            ["Avg HH Income",     "—", _v(demo.get("avg_income_3mi"),"$"), "—"],
            ["Population Growth", demo.get("pop_growth") or "—","—","—"],
            ["Avg Home Value",    _v(demo.get("home_value"),"$"),"—","—"],
            ["College Educated",  "—", demo.get("college_pct") or "—","—"],
            ["Median Age",        "—", _v(demo.get("median_age")), "—"],
        ],
        cw=[2.4*inch,1.7*inch,1.7*inch,1.7*inch]
    ))
    story.append(_p("Schools", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_kv([
        ("School District",   demo.get("school_district")),
        ("Elementary",        demo.get("elementary")),
        ("Middle School",     demo.get("middle")),
        ("High School",       demo.get("high_school")),
    ]))
    story.append(_p("Crime Rate", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_p(demo.get("crime") or "⚠ Not provided in OM — source independently.", color=WARN if not demo.get("crime") else DARK))
    story.append(_p("Major Employers", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    employers = demo.get("employers") or []
    if employers:
        story.append(_tbl(
            ["Employer","Drive Time","Employees","Sector","Notes"],
            [[e.get("name","—"), e.get("drive") or "—", e.get("employees") or "—",
              e.get("sector") or "—", e.get("notes") or "—"] for e in employers],
            cw=[1.8*inch,0.85*inch,0.85*inch,1.2*inch,2.8*inch]
        ))

    # 4 — Unit Mix
    story.append(PageBreak())
    _hdr(story, "Unit Mix", "4")
    if umix:
        story.append(_tbl(
            ["Type","Plan","Units","Mix%","SF","Mkt Rent","Mkt PSF","Eff Rent","Eff PSF","Target Rent","Upside"],
            [[u.get("type","—"), u.get("plan") or "—", _v(u.get("count"),"n"),
              _v(u.get("pct"),"%"), _v(u.get("sf"),"n"),
              _v(u.get("market_rent"),"$"), f"${u.get('market_psf') or 0:.2f}",
              _v(u.get("eff_rent"),"$"), f"${u.get('eff_psf') or 0:.2f}",
              _v(u.get("target_rent"),"$"), _v(u.get("upside"),"$")] for u in umix],
            hl_last=True,
            cw=[0.9*inch,0.65*inch,0.5*inch,0.5*inch,0.5*inch,
                0.8*inch,0.65*inch,0.8*inch,0.65*inch,0.8*inch,0.65*inch]
        ))
        story.append(_p("* Target Rent = comparable market-supported upside.", color=MID, size=8))
    else:
        story.append(_p("No unit mix data found.", color=WARN))

    # 5 — Utilities
    story.append(PageBreak())
    _hdr(story, "Utility Information", "5")
    if utils:
        story.append(_tbl(
            ["Utility","Billing Method","Paid By","Reimbursement","Fee","Annual Income","Notes"],
            [[u.get("name","—"), u.get("method") or "—", u.get("paid_by") or "—",
              u.get("reimbursement") or "N/A", u.get("fee") or "—",
              _v(u.get("annual_income"),"$"), u.get("notes") or "—"] for u in utils],
            cw=[0.85*inch,1.1*inch,0.65*inch,1.2*inch,0.75*inch,0.85*inch,2.05*inch]
        ))
    else:
        story.append(_p("No utility data found.", color=WARN))

    # 6 — Site Info
    story.append(PageBreak())
    _hdr(story, "Site Information", "6")
    story.append(_kv([
        ("Roof / Age",        f"{site.get('roof') or 'N/A'}  —  {site.get('roof_age') or 'N/A'}"),
        ("Exterior",          site.get("exterior")),
        ("Foundation",        site.get("foundation")),
        ("HVAC",              site.get("hvac")),
        ("Plumbing",          site.get("plumbing")),
        ("Wiring",            site.get("wiring")),
        ("Hot Water",         site.get("hot_water")),
        ("Washer / Dryer",    site.get("washer_dryer")),
        ("Life Safety",       site.get("life_safety") or "Verify — not specified"),
        ("Construction",      site.get("notes")),
    ]))
    story.append(_p("Parking & Site Features", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_kv([
        ("Open Spaces",   _v(site.get("parking_open"),"n")),
        ("Reserved",      f"{_v(site.get('parking_reserved'),'n')} ({site.get('reserved_fee') or 'N/A'}/mo)"),
        ("Covered",       _v(site.get("parking_covered"),"n")),
        ("Garage",        site.get("parking_garage") or "None"),
        ("Total / Ratio", f"{_v(site.get('parking_total'),'n')} ({site.get('parking_ratio') or 'N/A'})"),
        ("Pet Yards",     site.get("pet_yards") or "N/A"),
        ("Storage",       site.get("storage") or "N/A"),
    ]))

    # 7 — Rent Comparables
    story.append(PageBreak())
    _hdr(story, "Rent Comparable Summary", "7")
    if rcomps:
        story.append(_tbl(
            ["#","Property","Address","Yr Built","Units","Occ%","Avg SF","Mkt Rent","Mkt PSF","Eff Rent","Eff PSF"],
            [[rc.get("id","—"), rc.get("name","—"),
              rc.get("address") or rc.get("city_state") or "—",
              _v(rc.get("year_built")), _v(rc.get("units"),"n"),
              _v(rc.get("occupancy"),"%"), _v(rc.get("avg_sf"),"n"),
              _v(rc.get("total_market"),"$"), f"${rc.get('total_market_psf') or 0:.2f}",
              _v(rc.get("total_eff"),"$"), f"${rc.get('total_eff_psf') or 0:.2f}"] for rc in rcomps],
            cw=[0.35*inch,1.3*inch,1.2*inch,0.55*inch,0.45*inch,0.45*inch,
                0.5*inch,0.7*inch,0.55*inch,0.7*inch,0.55*inch]
        ))
        story.append(_p("Rents by Bed Type", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
        detail = []
        for rc in rcomps:
            for b in (rc.get("by_bed") or []):
                detail.append([rc.get("name","—"), b.get("type","—"), _v(b.get("sf"),"n"),
                                _v(b.get("market"),"$"), f"${b.get('market_psf') or 0:.2f}",
                                _v(b.get("eff"),"$"), f"${b.get('eff_psf') or 0:.2f}"])
        if detail:
            story.append(_tbl(
                ["Property","Bed Type","SF","Market Rent","Mkt PSF","Eff Rent","Eff PSF"],
                detail, cw=[1.7*inch,0.75*inch,0.55*inch,0.95*inch,0.7*inch,0.95*inch,0.65*inch]
            ))
    else:
        story.append(_p("No rent comparable data found.", color=WARN))

    # 8 — Financial Analysis
    story.append(PageBreak())
    _hdr(story, "Financial Analysis — Operating Statement", "8")
    periods = fin.get("periods") or []
    inc_lines = fin.get("income_lines") or []
    exp_lines = fin.get("expense_lines") or []

    if periods and (inc_lines or exp_lines):
        p4 = periods[:4]
        cw_fin = [2.1*inch] + [0.85*inch]*len(p4) + [2.7*inch]

        def frow(item, key="values"):
            vals = item.get(key) or {}
            row = [item.get("item","—")]
            for p in p4:
                v = vals.get(p)
                pct = (item.get("pct") or item.get("per_unit") or {}).get(p,"")
                cell = _v(v,"$") if v is not None else "—"
                if pct: cell += f" ({pct})"
                row.append(cell)
            note = item.get("note") or ""
            row.append(note[:180] + ("…" if len(note)>180 else ""))
            return row

        all_rows = [["Line Item"]+p4+["Underwriting Note"]]
        smap = []

        all_rows.append(["INCOME"]+[""]*len(p4)+[""]); smap.append((len(all_rows)-1,"sec"))
        for item in inc_lines:
            all_rows.append(frow(item))
            if item.get("is_total"): smap.append((len(all_rows)-1,"tot"))
            elif item.get("is_subtotal"): smap.append((len(all_rows)-1,"sub"))

        all_rows.append(["EXPENSES"]+[""]*len(p4)+[""]); smap.append((len(all_rows)-1,"sec"))
        for item in exp_lines:
            all_rows.append(frow(item,"values"))
            if item.get("is_total"): smap.append((len(all_rows)-1,"tot"))
            elif item.get("is_subtotal"): smap.append((len(all_rows)-1,"sub"))

        noi = fin.get("noi") or {}
        all_rows.append(["NET OPERATING INCOME"]+[_v(noi.get(p),"$") for p in p4]+["Key metric"])
        smap.append((len(all_rows)-1,"tot"))

        capex = fin.get("capex") or {}
        all_rows.append(["  Capital Reserves"]+[_v(capex.get(p),"$") for p in p4]+["Market std $225–$300/unit"])

        cffo = fin.get("cffo") or {}
        all_rows.append(["CASH FLOW FROM OPERATIONS"]+[_v(cffo.get(p),"$") for p in p4]+["NOI minus capex"])
        smap.append((len(all_rows)-1,"tot"))

        t = Table(all_rows, colWidths=cw_fin, repeatRows=1)
        ts = [
            ("BACKGROUND",(0,0),(-1,0),DARK), ("TEXTCOLOR",(0,0),(-1,0),GOLD_L),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),7.5),
            ("TOPPADDING",(0,0),(-1,-1),4), ("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1),5), ("RIGHTPADDING",(0,0),(-1,-1),5),
            ("GRID",(0,0),(-1,-1),0.3,BORDER),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,ALT]),
        ]
        for ridx, stype in smap:
            if stype=="tot":   ts += [("BACKGROUND",(0,ridx),(-1,ridx),DARK2),("TEXTCOLOR",(0,ridx),(-1,ridx),GOLD_L),("FONTNAME",(0,ridx),(-1,ridx),"Helvetica-Bold")]
            elif stype=="sub": ts += [("BACKGROUND",(0,ridx),(-1,ridx),GOLD_P),("FONTNAME",(0,ridx),(-1,ridx),"Helvetica-Bold")]
            elif stype=="sec": ts += [("BACKGROUND",(0,ridx),(-1,ridx),GOLD_P),("FONTNAME",(0,ridx),(-1,ridx),"Helvetica-Bold"),("TEXTCOLOR",(0,ridx),(-1,ridx),WARN)]
        t.setStyle(TableStyle(ts))
        story.append(t)
    else:
        story.append(_p("No financial analysis data found.", color=WARN))

    # 9 — Sale Comps
    story.append(PageBreak())
    _hdr(story, "Sale Comparables", "9")
    if scomps:
        story.append(_tbl(
            ["Property","Address","Date","Yr Built","Units","Sale Price","$/Unit","$/SF","Cap Rate","Occ","Notes"],
            [[sc.get("name","—"), f"{sc.get('address') or ''} {sc.get('city_state') or ''}".strip() or "—",
              sc.get("date") or "—", _v(sc.get("year_built")), _v(sc.get("units"),"n"),
              _v(sc.get("price"),"$"), _v(sc.get("ppu"),"$"), _v(sc.get("ppsf"),"$"),
              sc.get("cap_rate") or "—", sc.get("occupancy") or "—", sc.get("notes") or "—"] for sc in scomps],
            cw=[1.2*inch,1.1*inch,0.6*inch,0.5*inch,0.45*inch,
                0.8*inch,0.7*inch,0.55*inch,0.6*inch,0.55*inch,0.85*inch]
        ))
    else:
        story.append(_p("⚠ Sale comps not in OM. Source from CoStar, RCA, or listing broker.", color=WARN))

    # 10 — Flags
    story.append(PageBreak())
    _hdr(story, "Underwriting Flags & Notes", "10")
    if flags:
        cat_map = {
            "Warning":     (RED_B,    RED),
            "Opportunity": (GREEN_B,  GREEN),
            "Verify":      (PURPLE_B, PURPLE),
            "Info":        (BLUE_B,   BLUE),
        }
        fdata = [["Category","Title","Detail"]]
        fstyle = [
            ("BACKGROUND",(0,0),(-1,0),DARK),("TEXTCOLOR",(0,0),(-1,0),GOLD_L),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),8),
            ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
            ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
            ("GRID",(0,0),(-1,-1),0.3,BORDER),("VALIGN",(0,0),(-1,-1),"TOP"),
        ]
        for i, f in enumerate(flags, 1):
            cat = f.get("category","Info")
            bg, tc = cat_map.get(cat, (BLUE_B, BLUE))
            fdata.append([cat, f.get("title",""), f.get("detail","")])
            fstyle += [("BACKGROUND",(0,i),(0,i),bg),("TEXTCOLOR",(0,i),(0,i),tc),
                       ("FONTNAME",(0,i),(0,i),"Helvetica-Bold")]
        ft = Table(fdata, colWidths=[1.0*inch,1.8*inch,4.7*inch])
        ft.setStyle(TableStyle(fstyle))
        story.append(ft)

    # Market overview
    if any(v for v in mkt.values() if v):
        story.append(_p("Market Overview", fontname="Helvetica-Bold", size=11, space_before=12, space_after=4))
        mkt_rows = [(k,v) for k,v in [
            ("Submarket",        mkt.get("submarket")),
            ("Sub. Occupancy",   mkt.get("sub_occupancy")),
            ("Sub. Avg Rent",    _v(mkt.get("sub_rent"),"$")),
            ("Sub. Rent Growth", mkt.get("sub_growth")),
            ("Metro Inventory",  mkt.get("metro_inventory")),
            ("Pipeline",         mkt.get("pipeline")),
            ("Absorption",       mkt.get("absorption")),
            ("Investment Vol.",  mkt.get("investment_vol")),
        ] if v]
        if mkt_rows:
            story.append(_kv(mkt_rows))

    story.append(Spacer(1,20))
    story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER))
    story.append(_p("DISCLAIMER: AI-generated from the uploaded OM. For internal underwriting reference only. Verify all figures independently. Powered by Anthropic Claude.", color=MID, size=7.5, space_before=6))

    doc.build(story, onFirstPage=_footer(prop, date), onLaterPages=_footer(prop, date))
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏢 OM Analyzer")
    st.markdown("---")
    st.markdown("**What this does**")
    st.markdown("""
Upload any Multifamily Offering Memorandum PDF and get a structured 10-section underwriting report.

**Covers:**
- Basic property details
- Investment highlights
- Demographics & schools
- Unit mix with rent upside
- Utility billing info
- Site & systems info
- Rent comparables
- Financial analysis
- Sale comparables
- Underwriting flags
""")
    st.markdown("---")
    st.markdown("**Supported Brokers**")
    st.markdown("JLL · CBRE · Marcus & Millichap · Cushman & Wakefield · Newmark · Colliers · Berkadia · Walker & Dunlop · and more")
    st.markdown("---")
    st.markdown("**Processing time:** 30–90 sec")
    st.markdown("**Max file size:** 50 MB")

# ── Main ──────────────────────────────────────────────────────────────────────
st.markdown("# 🏢 Multifamily OM Analyzer")
st.markdown("Upload an Offering Memorandum PDF → AI extracts all underwriting data → download a structured report.")
st.markdown("---")

# API key
api_key = st.secrets.get("ANTHROPIC_API_KEY", os.environ.get("ANTHROPIC_API_KEY", ""))
if not api_key:
    st.error("""
**API key not configured.**

- **Streamlit Cloud:** Go to your app → ⚙️ Settings → Secrets → add:
  ```
  ANTHROPIC_API_KEY = "sk-ant-..."
  ```
- **Local:** Create `.streamlit/secrets.toml` with the same line.
""")
    st.stop()
os.environ["ANTHROPIC_API_KEY"] = api_key

# Upload
uploaded = st.file_uploader(
    "Drop your OM PDF here",
    type=["pdf"],
    label_visibility="collapsed"
)

if uploaded is None:
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("**Step 1** — Upload a PDF using the box above")
    with col2:
        st.info("**Step 2** — Click Analyze (takes 30–90 sec)")
    with col3:
        st.info("**Step 3** — Download the underwriting PDF")
    st.stop()

# File info
size_mb = uploaded.size / 1024 / 1024
st.markdown(f"**File:** `{uploaded.name}` — {size_mb:.1f} MB")

if st.button("🔍  Analyze Offering Memorandum", type="primary", use_container_width=True):

    # ── Progress UI ───────────────────────────────────────────────────────────
    progress_bar = st.progress(0, text="Starting...")
    status_box   = st.empty()

    def set_progress(pct, msg):
        progress_bar.progress(pct, text=msg)
        status_box.markdown(f"*{msg}*")

    try:
        # Step 1 — Save & extract
        set_progress(10, "Extracting text from PDF…")
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(uploaded.getvalue())
            tmp_path = tmp.name

        pdf_text = extract_pdf_text(tmp_path)
        os.unlink(tmp_path)

        if not pdf_text or len(pdf_text.strip()) < 200:
            progress_bar.empty(); status_box.empty()
            st.error("Could not extract readable text. This PDF may be a scanned image — please use a text-based PDF.")
            st.stop()

        set_progress(30, f"Extracted {len(pdf_text):,} characters. Sending to Claude AI…")

        # Step 2 — Analyze
        def log(msg): set_progress(55, msg)
        data = analyze_om(pdf_text, api_key, log)

        set_progress(75, "Generating PDF report…")

        # Step 3 — Build PDF
        pdf_bytes = build_pdf(data, uploaded.name)

        set_progress(100, "Done!")
        progress_bar.empty()
        status_box.empty()

    except Exception as e:
        progress_bar.empty(); status_box.empty()
        st.error(f"**Error:** {e}")
        with st.expander("Full error"):
            import traceback; st.code(traceback.format_exc())
        st.stop()

    # ── Success banner ────────────────────────────────────────────────────────
    prop  = (data.get("property") or {}).get("name") or "Property"
    broker= (data.get("broker")   or {}).get("name") or "Unknown broker"
    st.success(f"✅  Report ready — **{prop}**  ·  Broker: {broker}")

    safe = re.sub(r"[^a-zA-Z0-9_\- ]", "", prop).strip().replace(" ","_")
    st.download_button(
        label="⬇️  Download Underwriting Report (PDF)",
        data=pdf_bytes,
        file_name=f"{safe}_Underwriting_Report.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

    # ── Summary metrics ───────────────────────────────────────────────────────
    pd_  = data.get("property") or {}
    inv_ = data.get("investment") or {}

    m1, m2, m3, m4, m5 = st.columns(5)
    with m1: st.metric("Units",        _v(pd_.get("units"),"n"))
    with m2: st.metric("Year Built",   _v(pd_.get("year_built")))
    with m3: st.metric("Occupancy",    _v(pd_.get("occupancy_pct"),"%"))
    with m4: st.metric("Eff. Rent",    _v(inv_.get("effective_rent"),"$"))
    with m5: st.metric("Rent Upside",  _v(inv_.get("rent_gap"),"$",suffix="/unit"))

    # ── Tabs for data preview ─────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Unit Mix", "Rent Comps", "Financials", "Demographics", "Flags"])

    with tab1:
        umix_ = data.get("unit_mix") or []
        if umix_:
            import pandas as pd
            df = pd.DataFrame([{
                "Type":        u.get("type"),
                "Units":       u.get("count"),
                "SF":          u.get("sf"),
                "Market Rent": u.get("market_rent"),
                "Eff. Rent":   u.get("eff_rent"),
                "Target Rent": u.get("target_rent"),
                "Upside/Unit": u.get("upside"),
            } for u in umix_])
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info("No unit mix data extracted.")

    with tab2:
        rcomps_ = data.get("rent_comps") or []
        if rcomps_:
            import pandas as pd
            rows = []
            for rc in rcomps_:
                rows.append({
                    "#":          rc.get("id"),
                    "Property":   rc.get("name"),
                    "City":       rc.get("city_state"),
                    "Year":       rc.get("year_built"),
                    "Units":      rc.get("units"),
                    "Occ%":       rc.get("occupancy"),
                    "Avg SF":     rc.get("avg_sf"),
                    "Mkt Rent":   rc.get("total_market"),
                    "Eff. Rent":  rc.get("total_eff"),
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        else:
            st.info("No comp data extracted.")

    with tab3:
        fin_ = data.get("financials") or {}
        periods_ = fin_.get("periods") or []
        inc_     = fin_.get("income_lines") or []
        exp_     = fin_.get("expense_lines") or []
        if periods_ and (inc_ or exp_):
            import pandas as pd
            rows = []
            for line in inc_ + exp_:
                row = {"Line Item": line.get("item","—")}
                for p in periods_:
                    row[p] = line.get("values",{}).get(p)
                row["Note"] = (line.get("note") or "")[:80]
                rows.append(row)
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        else:
            st.info("No financial data extracted.")

    with tab4:
        demo_ = data.get("demographics") or {}
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown('<div class="gold-header">Population & Income</div>', unsafe_allow_html=True)
            for label, val in [
                ("Population (1-mile)", _v(demo_.get("pop_1mi"),"n")),
                ("Population (3-mile)", _v(demo_.get("pop_3mi"),"n")),
                ("Median HH Income (1-mi)", _v(demo_.get("median_income_1mi"),"$")),
                ("Avg HH Income (3-mi)",    _v(demo_.get("avg_income_3mi"),"$")),
                ("Avg Home Value",          _v(demo_.get("home_value"),"$")),
            ]:
                st.markdown(f"**{label}:** {val}")
        with col_b:
            st.markdown('<div class="gold-header">Schools</div>', unsafe_allow_html=True)
            for label, val in [
                ("District",    demo_.get("school_district")),
                ("Elementary",  demo_.get("elementary")),
                ("Middle",      demo_.get("middle")),
                ("High School", demo_.get("high_school")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")

        employers_ = demo_.get("employers") or []
        if employers_:
            st.markdown('<div class="gold-header">Major Employers</div>', unsafe_allow_html=True)
            import pandas as pd
            st.dataframe(pd.DataFrame(employers_), use_container_width=True, hide_index=True)

    with tab5:
        flags_ = data.get("flags") or []
        if flags_:
            cat_css = {
                "Warning":     "flag-warn",
                "Opportunity": "flag-good",
                "Info":        "flag-info",
                "Verify":      "flag-verify",
            }
            for f in flags_:
                cat  = f.get("category","Info")
                css  = cat_css.get(cat,"flag-info")
                icon = {"Warning":"⚠️","Opportunity":"✅","Verify":"🔍","Info":"ℹ️"}.get(cat,"•")
                st.markdown(f"""
                <div class="{css}">
                  <div class="flag-title">{icon} [{cat}] {f.get('title','')}</div>
                  <div class="flag-body">{f.get('detail','')}</div>
                </div>""", unsafe_allow_html=True)
        else:
            st.info("No flags generated.")
