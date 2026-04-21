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
  "financing": {
    "offering_type": null,
    "debt_contact": null,
    "notes": null,
    "new_financing": {
      "loan_type": null,
      "lender": null,
      "loan_amount": null,
      "loan_to_value": null,
      "interest_rate": null,
      "rate_type": null,
      "amortization_years": null,
      "loan_term_years": null,
      "interest_only_period": null,
      "dscr": null,
      "recourse": null,
      "notes": null
    },
    "assumable_debt": {
      "loan_type": null,
      "lender": null,
      "loan_amount": null,
      "loan_to_value": null,
      "interest_rate": null,
      "rate_type": null,
      "amortization_years": null,
      "loan_term_years": null,
      "origination_date": null,
      "maturity_date": null,
      "monthly_payment": null,
      "annual_debt_service": null,
      "dscr": null,
      "prepayment_penalty": null,
      "recourse": null,
      "notes": null
    }
  },
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
  "value_add": {
    "scope": null,
    "total_cost": null,
    "cost_per_unit": null,
    "exterior_capex": null,
    "monthly_premium": null,
    "annual_premium": null,
    "roi_pct": null,
    "light_upgrade_items": [],
    "by_floor_plan": [
      {
        "type": null,
        "sf": null,
        "units": null,
        "inplace_rent": null,
        "inplace_psf": null,
        "rehab_cost": null,
        "premium": null,
        "post_rehab_rent": null,
        "post_rehab_psf": null
      }
    ]
  },
  "tax": {
    "parcel_id": null,
    "assessed_value": null,
    "millage_city": null,
    "millage_county": null,
    "millage_total": null,
    "tax_base": null,
    "solid_waste_fee": null,
    "total_tax": null,
    "abatement_program": null,
    "abatement_pct": null,
    "abatement_term_note": null,
    "abatement_annual_savings": null,
    "ami_pct": null,
    "max_allowable_rent": null,
    "avg_inplace_rent": null,
    "rent_headroom": null,
    "units_compliant": null,
    "pct_compliant": null
  },
  "demographics": {
    "pop_3mi": null, "pop_5mi": null,
    "pop_growth_3mi": null, "pop_growth_5mi": null,
    "pop_2030_3mi": null, "pop_2030_5mi": null,
    "median_income_3mi": null, "median_income_5mi": null,
    "median_income_2030_3mi": null, "median_income_2030_5mi": null,
    "income_growth_3mi": null, "income_growth_5mi": null,
    "renter_pct_3mi": null, "renter_pct_5mi": null,
    "college_pct_3mi": null, "college_pct_5mi": null,
    "white_collar_pct_3mi": null, "white_collar_pct_5mi": null,
    "home_value": null, "home_value_area": null,
    "crime": null, "school_district": null, "elementary": null,
    "middle": null, "high_school": null,
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
  "rent_comps_garden": [
    {"id": null, "name": null, "distance": null, "year_built": null, "rent": null, "notes": null}
  ],
  "rent_comps_townhouse": [
    {"id": null, "name": null, "distance": null, "year_built": null, "rent": null, "notes": null}
  ],
  "rent_comps": [
    {"id": null, "name": null, "address": null, "city_state": null,
     "distance": null, "year_built": null, "units": null, "occupancy": null, "avg_sf": null,
     "comp_type": null,
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
    "noi": {}, "noi_per_unit": {}, "capex": {}, "cffo": {}, "cffo_per_unit": {}, "expense_ratio": {}
  },
  "sale_comps": [
    {"name": null, "address": null, "city_state": null, "date": null,
     "year_built": null, "units": null, "price": null, "ppu": null,
     "ppsf": null, "cap_rate": null, "occupancy": null,
     "buyer": null, "seller": null, "notes": null}
  ],
  "market": {
    "submarket": null, "sub_occupancy": null, "sub_rent": null,
    "sub_growth": null, "metro_inventory": null, "metro_occupancy": null,
    "pipeline": null, "absorption": null, "investment_vol": null,
    "market_summary": null,
    "major_developments": [
      {"name": null, "description": null, "cost": null, "jobs": null, "timeline": null}
    ]
  },
  "affordability": {
    "current_rent": null,
    "avg_hh_income_3mi": null,
    "monthly_affordability_3x": null,
    "rent_headroom_3mi": null,
    "avg_hh_income_2030_3mi": null,
    "monthly_affordability_2030_3x": null,
    "rent_headroom_2030_3mi": null,
    "income_to_rent_ratio": null,
    "notes": null
  },
  "insurance": {
    "carrier": null,
    "annual_premium": null,
    "per_unit": null,
    "quote_source": null,
    "notes": null
  },
  "management": {
    "fee_pct": null,
    "fee_annual": null,
    "fee_per_unit": null,
    "current_manager": null,
    "proposed_manager": null,
    "notes": null
  },
  "replacement_cost": {
    "per_unit": null,
    "per_sf": null,
    "total": null,
    "source": null,
    "land_per_unit": null,
    "hard_cost_per_sf": null,
    "soft_cost_pct": null,
    "notes": null
  },
  "value_add": {
    "scope": null,
    "total_cost": null,
    "cost_per_unit": null,
    "exterior_capex": null,
    "monthly_premium": null,
    "annual_premium": null,
    "roi_pct": null,
    "light_upgrade_items": [],
    "by_floor_plan": [
      {
        "type": null,
        "sf": null,
        "units": null,
        "inplace_rent": null,
        "inplace_psf": null,
        "rehab_cost": null,
        "premium": null,
        "post_rehab_rent": null,
        "post_rehab_psf": null
      }
    ]
  },
  "flags": [
    {"category": null, "title": null, "detail": null}
  ]
}

CRITICAL EXTRACTION RULES — read every page carefully:

1. DEMOGRAPHICS: The OM tables show columns in order: 3-mile, 5-mile, County, Metro.
   - pop_3mi = the FIRST population column (smaller number, e.g. ~118,000)
   - pop_5mi = the SECOND population column (larger number, e.g. ~278,000)
   - median_income_3mi = first income column, median_income_5mi = second income column
   - Extract BOTH 3-mile and 5-mile data — do not mix them up.

2. FINANCIAL PERIODS: Extract ALL columns including "Pro Forma YR1" or "Proforma" column.
   periods array must include: ["T12", "T6 Ann", "T3 Ann", "T1 Ann", "Pro Forma YR1"]
   Every income and expense line must have values for ALL 5 periods.
   The "note" field for each line item must contain the full underwriting assumption text
   from the "Notes to Pro Forma" page (e.g. "Y1 vacancy of 5.00% assumed", "$45/u/mo at 95% occupancy").

3. RENT COMPS: The OM splits comps into TWO groups:
   - rent_comps_garden: garden apartment comps (Rise Avalon, Stone Gate, Biscayne, Vyne on Central, M Station)
   - rent_comps_townhouse: townhouse comps (Trailside NoDa, Reve Noda Junction, Radial at Mora Point, Towns at Eastland, Sablewood)
   Also populate rent_comps with all properties. Set comp_type = "Garden" or "Townhouse".
   Include the group average rent for each group.
   For EVERY comp, extract distance from subject (e.g. "0.7 miles") and year_built if shown.

4. VALUE-ADD: Extract the full "Proposed Value-Add by Floor Plan" table including:
   - Each floor plan row: type, sf, units, inplace_rent, inplace_psf, rehab_cost, premium, post_rehab_rent, post_rehab_psf
   - Summary totals: total_cost, cost_per_unit, exterior_capex, monthly_premium, annual_premium, roi_pct
   - If the OM shows a "Light Interior Upgrade" list instead of named tiers, populate light_upgrade_items as a
     plain array of strings (e.g. ["Install under cabinet lighting", "Update cabinet color", ...]).

5. TAX: Extract from the "Property Tax Value" page AND the "Tax Abatement" investment highlight:
   - parcel_id, assessed_value, millage_city, millage_county, millage_total
   - tax_base (tax before solid waste), solid_waste_fee, total_tax
   - abatement_program name/description, abatement_pct (100%), abatement_term_note
   - abatement_annual_savings = total_tax amount
   - ami_pct (80%), max_allowable_rent, avg_inplace_rent, rent_headroom, units_compliant, pct_compliant

6. OTHER INCOME: Extract ALL income line items individually from the OM:
   - Water/sewer submetering recapture ($112,594 proforma)
   - Trash fee ($25/unit/mo = $41,040/yr at 95%)
   - Pest control ($5/unit/mo = $8,208/yr at 95%)
   - Internet ($45/unit/mo at 95% = $73,872/yr)
   - Storage (55 units × $10/mo at 95% = $5,940/yr)
   - Parking/garages (14 × $100/mo + 16 × $25/mo at 95% = $20,520/yr)

7. FLAGS: Generate 6-10 flags. category must be one of: Warning / Opportunity / Verify / Info
   Include flags for: occupancy decline, HOA pool risk, insurance inconsistency, tax abatement opportunity,
   value-add ROI, pro forma vacancy assumption, bad debt assumption, Eastland Yards catalyst.

8. FINANCING: Extract all debt/financing terms from the OM into the split structure:
   - offering_type (e.g. "Free & Clear", "Assumable Debt", "All Cash")
   - debt_contact: name/info of financing contact if mentioned
   - new_financing: populate if the OM provides a new loan quote or financing guidance:
     loan_type, lender, loan_amount, loan_to_value, interest_rate, rate_type,
     amortization_years, loan_term_years, interest_only_period, dscr, recourse
   - assumable_debt: populate if the OM describes existing assumable debt:
     loan_type, lender, loan_amount, loan_to_value, interest_rate, rate_type,
     amortization_years, loan_term_years, origination_date, maturity_date,
     monthly_payment, annual_debt_service, dscr, prepayment_penalty, recourse
   - If property is Free & Clear with no debt, leave both sub-objects null
   - If all-cash purchase required, set offering_type = "All Cash"

9. SALE COMPS: For every sale comp, extract buyer and seller names if disclosed in the OM.

10. MARKET: Populate market_summary with a 2-3 sentence narrative of the market/submarket story.
    Populate major_developments array with each named development project (Amazon, Boeing, hospital
    expansions, mixed-use projects, etc.) including cost, jobs created, and timeline where mentioned.

11. AFFORDABILITY: If the OM shows a rent-to-income or affordability table, extract:
    - current_rent, avg_hh_income_3mi, monthly_affordability_3x (= income/12/3),
      rent_headroom_3mi (= affordability threshold minus current rent)
    - Same for 2030 projections if shown
    - income_to_rent_ratio = avg monthly income / current rent

12. INSURANCE: Extract insurer/carrier name, annual premium, per-unit cost, and quote source if mentioned.

13. MANAGEMENT: Extract management fee % of EGI, annual dollar amount, per-unit amount,
    current manager name, and any proposed manager mentioned.

14. REPLACEMENT COST: Extract per-unit, per-SF, total replacement cost, hard cost per SF,
    soft cost %, land cost per unit, and source (e.g. broker estimate, current construction comps).

15. Extract every number that exists anywhere in the text. Do not skip any table or data page.
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
    cw = cw or [2.2*inch, 5.3*inch]
    wrap_style = ParagraphStyle("kv", fontName="Helvetica", fontSize=9,
                                textColor=DARK, leading=13, wordWrap="CJK")
    key_style  = ParagraphStyle("kk", fontName="Helvetica", fontSize=8.5,
                                textColor=MID, leading=13)
    data = [[Paragraph(str(k), key_style),
             Paragraph(str(v) if v else "N/A", wrap_style)] for k, v in rows]
    t = Table(data, colWidths=cw)
    t.setStyle(TableStyle([
        ("ROWBACKGROUNDS", (0,0),(-1,-1), [colors.white, ALT]),
        ("GRID",  (0,0),(-1,-1), 0.3, BORDER),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ("RIGHTPADDING",  (0,0),(-1,-1), 8),
        ("VALIGN", (0,0),(-1,-1), "TOP"),
    ]))
    return t


def _tbl(headers, rows, cw=None, hl_last=False):
    head = [_p(h, fontname="Helvetica-Bold", size=8, color=GOLD_L) for h in headers]
    # Use Paragraphs for every cell so ReportLab wraps long text automatically
    data = [head] + [
        [Paragraph(str(c) if c is not None else "—",
                   ParagraphStyle("tc", fontName="Helvetica", fontSize=8,
                                  textColor=DARK, leading=11, wordWrap="CJK"))
         for c in row]
        for row in rows
    ]
    cw = cw or [7.5*inch/len(headers)]*len(headers)
    t = Table(data, colWidths=cw, repeatRows=1)
    ts = [
        ("BACKGROUND",(0,0),(-1,0), DARK),
        ("FONTSIZE",  (0,0),(-1,-1), 8),
        ("TOPPADDING",(0,0),(-1,-1), 5), ("BOTTOMPADDING",(0,0),(-1,-1), 6),
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
    rcomps   = d.get("rent_comps") or []
    rg       = d.get("rent_comps_garden") or []
    rth      = d.get("rent_comps_townhouse") or []
    fin      = d.get("financials") or {}
    scomps   = d.get("sale_comps") or []
    flags    = d.get("flags") or []
    mkt      = d.get("market") or {}
    pd       = d.get("property") or {}
    va       = d.get("value_add") or {}
    tax      = d.get("tax") or {}
    fin_info = d.get("financing") or {}
    afford   = d.get("affordability") or {}
    insur    = d.get("insurance") or {}
    mgmt     = d.get("management") or {}
    repl     = d.get("replacement_cost") or {}

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
        leftMargin=0.75*inch, rightMargin=0.75*inch,
        topMargin=0.75*inch, bottomMargin=0.9*inch,
        title=f"{prop} — Underwriting Report")
    story = []

    addr = f"{pd.get('address','')}, {pd.get('city','')}, {pd.get('state','')} {pd.get('zip','')}".strip(", ")

    def dark_row(content, bg=DARK, top=8, bot=8, left=28):
        t = Table([[content]], colWidths=[7.5*inch])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), bg),
            ("TOPPADDING",(0,0),(-1,-1), top),
            ("BOTTOMPADDING",(0,0),(-1,-1), bot),
            ("LEFTPADDING",(0,0),(-1,-1), left),
            ("RIGHTPADDING",(0,0),(-1,-1), left),
        ]))
        return t

    # Cover — wrapped in KeepTogether so it doesn't split across pages
    from reportlab.platypus import KeepTogether
    cover_blocks = [
        dark_row(_p("MULTIFAMILY UNDERWRITING REPORT",
                    fontname="Helvetica", size=9, color=GOLD), top=36, bot=6),
        dark_row(_p(prop, fontname="Helvetica-Bold", size=24,
                    color=colors.white, leading=30), top=4, bot=10),
        dark_row(_p(
            f"Units: {_v(pd.get('units'),'n')}  ·  Year Built: {_v(pd.get('year_built'))}"
            f"  ·  Occupancy: {_v(pd.get('occupancy_pct'),'%')}  ·  Broker: {broker_d.get('name','N/A')}",
            size=10, color=GOLD_L, leading=15), bg=DARK2, top=10, bot=8),
        dark_row(_p(addr, size=10, color=MID, leading=14), bg=DARK2, top=4, bot=20),
        Spacer(1, 12),
        _p(f"Report Generated: {date}", color=MID, size=9),
        _p(f"Source: {filename}", color=MID, size=9),
    ]
    story.append(KeepTogether(cover_blocks))
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
        ("Listing Agents", ", ".join(
            a.get("name", str(a)) if isinstance(a, dict) else str(a)
            for a in (broker_d.get("agents") or [])
        ) or "N/A"),
    ]))

    # 2 — Financing & Debt Terms
    story.append(PageBreak())
    _hdr(story, "Financing & Debt Terms", "2")

    story.append(_kv([
        ("Offering Type",  fin_info.get("offering_type")),
        ("Debt Contact",   fin_info.get("debt_contact")),
        ("Notes",          fin_info.get("notes")),
    ]))

    story.append(Spacer(1, 10))

    nf  = fin_info.get("new_financing") or {}
    asd = fin_info.get("assumable_debt") or {}

    # Side-by-side two-column table for New Financing vs Assumable Debt
    col_hdr_style = ParagraphStyle("fch", fontName="Helvetica-Bold", fontSize=9,
                                   textColor=GOLD_L, leading=13)
    key_style_fin = ParagraphStyle("fck", fontName="Helvetica", fontSize=8.5,
                                   textColor=MID, leading=13)
    val_style_fin = ParagraphStyle("fcv", fontName="Helvetica", fontSize=9,
                                   textColor=DARK, leading=13, wordWrap="CJK")

    def fin_cell(label, val):
        return [Paragraph(label, key_style_fin), Paragraph(str(val) if val else "N/A", val_style_fin)]

    new_rows = [
        fin_cell("Loan Type",          nf.get("loan_type")),
        fin_cell("Lender",             nf.get("lender")),
        fin_cell("Loan Amount",        _v(nf.get("loan_amount"),"$")),
        fin_cell("LTV",                f"{nf.get('loan_to_value') or 'N/A'}%"),
        fin_cell("Interest Rate",      f"{nf.get('interest_rate') or 'N/A'}%"),
        fin_cell("Rate Type",          nf.get("rate_type")),
        fin_cell("Loan Term",          nf.get("loan_term_years")),
        fin_cell("Amortization",       nf.get("amortization_years")),
        fin_cell("Interest-Only",      nf.get("interest_only_period")),
        fin_cell("DSCR",               nf.get("dscr")),
        fin_cell("Recourse",           nf.get("recourse")),
        fin_cell("Notes",              nf.get("notes")),
    ]
    asd_rows = [
        fin_cell("Loan Type",          asd.get("loan_type")),
        fin_cell("Lender",             asd.get("lender")),
        fin_cell("Loan Amount",        _v(asd.get("loan_amount"),"$")),
        fin_cell("LTV",                f"{asd.get('loan_to_value') or 'N/A'}%"),
        fin_cell("Interest Rate",      f"{asd.get('interest_rate') or 'N/A'}%"),
        fin_cell("Rate Type",          asd.get("rate_type")),
        fin_cell("Loan Term",          asd.get("loan_term_years")),
        fin_cell("Amortization",       asd.get("amortization_years")),
        fin_cell("Origination Date",   asd.get("origination_date")),
        fin_cell("Maturity Date",      asd.get("maturity_date")),
        fin_cell("Monthly Payment",    _v(asd.get("monthly_payment"),"$")),
        fin_cell("Annual Debt Svc",    _v(asd.get("annual_debt_service"),"$")),
        fin_cell("DSCR",               asd.get("dscr")),
        fin_cell("Prepayment Penalty", asd.get("prepayment_penalty")),
        fin_cell("Recourse",           asd.get("recourse")),
        fin_cell("Notes",              asd.get("notes")),
    ]

    # Pad shorter column so rows align
    spacer_cell = Paragraph(" ", val_style_fin)
    while len(new_rows) < len(asd_rows):
        new_rows.append([Paragraph(" ", key_style_fin), spacer_cell])
    while len(asd_rows) < len(new_rows):
        asd_rows.append([Paragraph(" ", key_style_fin), spacer_cell])

    # Build side-by-side table: [NF label | NF value | divider | ASD label | ASD value]
    fin_hdr = [
        Paragraph("NEW FINANCING", col_hdr_style),
        Paragraph("", col_hdr_style),
        Paragraph("", col_hdr_style),
        Paragraph("ASSUMABLE DEBT", col_hdr_style),
        Paragraph("", col_hdr_style),
    ]
    fin_data = [fin_hdr] + [
        [nr[0], nr[1], Paragraph("", val_style_fin), ar[0], ar[1]]
        for nr, ar in zip(new_rows, asd_rows)
    ]
    fin_tbl = Table(fin_data, colWidths=[1.3*inch, 2.1*inch, 0.2*inch, 1.3*inch, 2.1*inch], repeatRows=1)
    fin_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (1,0),  DARK),
        ("BACKGROUND",    (3,0), (4,0),  DARK2),
        ("BACKGROUND",    (2,0), (2,-1), colors.white),
        ("ROWBACKGROUNDS",(0,1), (1,-1), [colors.white, ALT]),
        ("ROWBACKGROUNDS",(3,1), (4,-1), [colors.white, ALT]),
        ("GRID",          (0,0), (1,-1), 0.3, BORDER),
        ("GRID",          (3,0), (4,-1), 0.3, BORDER),
        ("LINEAFTER",     (2,0), (2,-1), 0,   colors.white),
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING",   (0,0), (-1,-1), 7),
        ("RIGHTPADDING",  (0,0), (-1,-1), 7),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
    ]))
    story.append(fin_tbl)

    # 3 — Investment Highlights
    story.append(PageBreak())
    _hdr(story, "Investment Highlights & Opportunities", "3")

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

    # 4 — Proposed Value-Add
    story.append(PageBreak())
    _hdr(story, "Proposed Value-Add by Floor Plan", "4")
    plans = va.get("by_floor_plan") or []
    if plans:
        story.append(_tbl(
            ["Unit Type","SF","Units","In-Place Rent","In-Place PSF","Rehab Cost/Unit","Premium/Unit","Post-Rehab Rent","Post-Rehab PSF"],
            [[p.get("type","—"), _v(p.get("sf"),"n"), _v(p.get("units"),"n"),
              _v(p.get("inplace_rent"),"$"), f"${p.get('inplace_psf') or 0:.2f}",
              _v(p.get("rehab_cost"),"$"), _v(p.get("premium"),"$"),
              _v(p.get("post_rehab_rent"),"$"), f"${p.get('post_rehab_psf') or 0:.2f}"] for p in plans],
            cw=[1.4*inch,0.5*inch,0.45*inch,0.85*inch,0.7*inch,0.85*inch,0.75*inch,0.9*inch,0.75*inch]
        ))
    light_items = va.get("light_upgrade_items") or []
    if light_items:
        story.append(_p("Light Interior Upgrade Items", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
        for item in light_items:
            story.append(_p(f"• {item}"))
    story.append(_p("Capital Expenditures & ROI Summary", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_kv([
        ("Total Renovation Cost",     _v(va.get("total_cost"),"$")),
        ("Cost Per Unit (all-in)",     _v(va.get("cost_per_unit"),"$")),
        ("Exterior / Additional CapEx",_v(va.get("exterior_capex"),"$")),
        ("Monthly Rent Premium (total)",_v(va.get("monthly_premium"),"$")),
        ("Annual Rent Premium (total)", _v(va.get("annual_premium"),"$")),
        ("Return on Investment",        f"{va.get('roi_pct') or 'N/A'}%"),
        ("Value-Add Scope",             va.get("scope")),
    ]))

    # 5 — Property Tax & Abatement
    story.append(PageBreak())
    _hdr(story, "Property Tax & Tax Abatement Opportunity", "5")
    story.append(_p("Property Tax Detail", fontname="Helvetica-Bold", size=11, space_before=4, space_after=4))
    story.append(_kv([
        ("Parcel ID",             tax.get("parcel_id")),
        ("Assessed Market Value", _v(tax.get("assessed_value"),"$")),
        ("Millage — City",        tax.get("millage_city")),
        ("Millage — County",      tax.get("millage_county")),
        ("Total Millage Rate",    tax.get("millage_total")),
        ("Ad Valorem Tax",        _v(tax.get("tax_base"),"$")),
        ("Solid Waste / Dumpster Fees", _v(tax.get("solid_waste_fee"),"$")),
        ("Total Annual Tax Bill", _v(tax.get("total_tax"),"$")),
    ]))
    story.append(_p("NC Property Tax Abatement Program", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_kv([
        ("Program",               tax.get("abatement_program")),
        ("Abatement Percentage",  f"{tax.get('abatement_pct') or 'N/A'}%"),
        ("Commitment Term",       tax.get("abatement_term_note")),
        ("AMI Requirement",       f"{tax.get('ami_pct') or 'N/A'}% of Area Median Income"),
        ("Potential Annual Savings", _v(tax.get("abatement_annual_savings"),"$")),
    ]))
    story.append(_p("AMI Rent Compliance", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_tbl(
        ["Bedroom Type","# Units","Max Allowable Rent (2025)","Avg In-Place Rent","Headroom/Unit","Units Compliant","% Compliant"],
        [["2 Bedroom",
          _v(tax.get("units_compliant"),"n"),
          _v(tax.get("max_allowable_rent"),"$"),
          _v(tax.get("avg_inplace_rent"),"$"),
          _v(tax.get("rent_headroom"),"$"),
          _v(tax.get("units_compliant"),"n"),
          f"{tax.get('pct_compliant') or 'N/A'}%"]],
        cw=[1.0*inch,0.6*inch,1.4*inch,1.2*inch,0.9*inch,1.1*inch,0.9*inch]
    ))

    # 6 — Replacement Cost, Insurance & Management
    story.append(PageBreak())
    _hdr(story, "Replacement Cost, Insurance & Management", "6")

    story.append(_p("Replacement Cost", fontname="Helvetica-Bold", size=11, space_before=4, space_after=4))
    story.append(_kv([
        ("Cost Per Unit",       _v(repl.get("per_unit"),"$")),
        ("Cost Per SF",         _v(repl.get("per_sf"),"$")),
        ("Total Replacement",   _v(repl.get("total"),"$")),
        ("Land Per Unit",       _v(repl.get("land_per_unit"),"$")),
        ("Hard Cost Per SF",    _v(repl.get("hard_cost_per_sf"),"$")),
        ("Soft Cost %",         f"{repl.get('soft_cost_pct') or 'N/A'}%"),
        ("Source",              repl.get("source")),
        ("Notes",               repl.get("notes")),
    ]))

    story.append(_p("Insurance", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_kv([
        ("Carrier / Provider",  insur.get("carrier")),
        ("Annual Premium",      _v(insur.get("annual_premium"),"$")),
        ("Per Unit / Year",     _v(insur.get("per_unit"),"$")),
        ("Quote Source",        insur.get("quote_source")),
        ("Notes",               insur.get("notes")),
    ]))

    story.append(_p("Property Management", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_kv([
        ("Management Fee %",    f"{mgmt.get('fee_pct') or 'N/A'}% of EGI"),
        ("Annual Fee",          _v(mgmt.get("fee_annual"),"$")),
        ("Per Unit / Year",     _v(mgmt.get("fee_per_unit"),"$")),
        ("Current Manager",     mgmt.get("current_manager")),
        ("Proposed Manager",    mgmt.get("proposed_manager")),
        ("Notes",               mgmt.get("notes")),
    ]))

    # 7 — Affordability Analysis
    story.append(PageBreak())
    _hdr(story, "Affordability & Rent Growth Runway", "7")
    story.append(_kv([
        ("Current In-Place Rent",          _v(afford.get("current_rent"),"$")),
        ("Avg HH Income — 3-Mile (2025)",  _v(afford.get("avg_hh_income_3mi"),"$")),
        ("Monthly Affordability @ 3× Rent",_v(afford.get("monthly_affordability_3x"),"$")),
        ("Rent Headroom (3-Mile, 2025)",   _v(afford.get("rent_headroom_3mi"),"$")),
        ("Avg HH Income — 3-Mile (2030)",  _v(afford.get("avg_hh_income_2030_3mi"),"$")),
        ("Monthly Affordability @ 3× (2030)",_v(afford.get("monthly_affordability_2030_3x"),"$")),
        ("Rent Headroom (3-Mile, 2030)",   _v(afford.get("rent_headroom_2030_3mi"),"$")),
        ("Income-to-Rent Ratio",           afford.get("income_to_rent_ratio")),
        ("Notes",                          afford.get("notes")),
    ]))

    # 8 — Demographics
    story.append(PageBreak())
    _hdr(story, "Demographics", "8")
    story.append(_tbl(
        ["Metric","3-Mile Radius","5-Mile Radius"],
        [
            ["Population (2025)",        _v(demo.get("pop_3mi"),"n"),   _v(demo.get("pop_5mi"),"n")],
            ["Population (2030 proj.)",  _v(demo.get("pop_2030_3mi"),"n"), _v(demo.get("pop_2030_5mi"),"n")],
            ["Population Growth (5-yr)", demo.get("pop_growth_3mi") or "—", demo.get("pop_growth_5mi") or "—"],
            ["Median HH Income (2025)",  _v(demo.get("median_income_3mi"),"$"), _v(demo.get("median_income_5mi"),"$")],
            ["Median HH Income (2030)",  _v(demo.get("median_income_2030_3mi"),"$"), _v(demo.get("median_income_2030_5mi"),"$")],
            ["Income Growth (5-yr)",     demo.get("income_growth_3mi") or "—", demo.get("income_growth_5mi") or "—"],
            ["Renter-Occupied Units",    demo.get("renter_pct_3mi") or "—", demo.get("renter_pct_5mi") or "—"],
            ["Bachelor's Degree+",       demo.get("college_pct_3mi") or "—", demo.get("college_pct_5mi") or "—"],
            ["White-Collar Workers",     demo.get("white_collar_pct_3mi") or "—", demo.get("white_collar_pct_5mi") or "—"],
            ["Avg Home Value",           _v(demo.get("home_value"),"$"), "—"],
        ],
        cw=[2.8*inch,2.35*inch,2.35*inch]
    ))
    story.append(_p("Schools", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_kv([
        ("School District",   demo.get("school_district")),
        ("Elementary",        demo.get("elementary")),
        ("Middle School",     demo.get("middle")),
        ("High School",       demo.get("high_school")),
    ]))
    story.append(_p("Crime Rate", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    story.append(_p(demo.get("crime") or "Not provided in OM — source independently.", color=WARN if not demo.get("crime") else DARK))
    story.append(_p("Major Employers", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
    employers = demo.get("employers") or []
    if employers:
        story.append(_tbl(
            ["Employer","Drive Time","Employees","Sector","Notes"],
            [[e.get("name","—"), e.get("drive") or "—", e.get("employees") or "—",
              e.get("sector") or "—", e.get("notes") or "—"] for e in employers],
            cw=[1.8*inch,0.85*inch,0.85*inch,1.2*inch,2.8*inch]
        ))

    # 9 — Unit Mix
    story.append(PageBreak())
    _hdr(story, "Unit Mix", "9")
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

    # 10 — Utilities
    story.append(PageBreak())
    _hdr(story, "Utility Information", "10")
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

    # 11 — Site Info
    story.append(PageBreak())
    _hdr(story, "Site Information", "11")
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

    # 12 — Rent Comparables
    story.append(PageBreak())
    _hdr(story, "Rent Comparable Summary", "12")

    # Garden comps group
    if rg:
        story.append(_p("Two Bedroom Garden Comparables", fontname="Helvetica-Bold", size=11, space_before=4, space_after=4))
        story.append(_tbl(
            ["Property","Yr Built","Distance","Rent","vs. Subject","Notes"],
            [[c.get("name","—"), _v(c.get("year_built")), c.get("distance") or "—",
              _v(c.get("rent"),"$"),
              f"{'+' if (c.get('rent') or 0) >= 1393 else ''}{int((c.get('rent') or 0) - 1393):+,}" if c.get("rent") else "—",
              c.get("notes") or "—"] for c in rg],
            cw=[2.1*inch,0.7*inch,0.8*inch,0.9*inch,0.85*inch,2.15*inch]
        ))

    # Townhouse comps group
    if rth:
        story.append(_p("Two Bedroom Townhouse Comparables", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
        story.append(_tbl(
            ["Property","Yr Built","Distance","Rent","vs. Subject","Notes"],
            [[c.get("name","—"), _v(c.get("year_built")), c.get("distance") or "—",
              _v(c.get("rent"),"$"),
              f"{int((c.get('rent') or 0) - 1557):+,}" if c.get("rent") else "—",
              c.get("notes") or "—"] for c in rth],
            cw=[2.1*inch,0.7*inch,0.8*inch,0.9*inch,0.85*inch,2.15*inch]
        ))

    # Full comp detail if available
    if rcomps:
        story.append(_p("Full Comparable Detail", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
        story.append(_tbl(
            ["#","Property","Type","Yr Built","Distance","Units","Occ%","Mkt Rent","Eff Rent"],
            [[rc.get("id","—"), rc.get("name","—"),
              rc.get("comp_type") or "—",
              _v(rc.get("year_built")),
              rc.get("distance") or "—",
              _v(rc.get("units"),"n"),
              _v(rc.get("occupancy"),"%"),
              _v(rc.get("total_market"),"$"),
              _v(rc.get("total_eff"),"$")] for rc in rcomps],
            cw=[0.3*inch,1.6*inch,0.75*inch,0.55*inch,0.7*inch,0.5*inch,0.5*inch,0.8*inch,0.8*inch]
        ))
    elif not rg and not rth:
        story.append(_p("No rent comparable data found.", color=WARN))


    # 13 — Financial Analysis
    story.append(PageBreak())
    _hdr(story, "Financial Analysis — Operating Statement", "13")
    periods = fin.get("periods") or []
    inc_lines = fin.get("income_lines") or []
    exp_lines = fin.get("expense_lines") or []

    if periods and (inc_lines or exp_lines):
        # Use up to 5 periods; Pro Forma column gets a gold highlight
        p_all = periods[:5]
        # Narrow col widths: item 1.8", each period 0.72", note fills the rest
        note_w = max(1.2*inch, 7.5*inch - 1.8*inch - len(p_all)*0.72*inch)
        cw_fin = [1.8*inch] + [0.72*inch]*len(p_all) + [note_w]

        def frow(item, key="values"):
            vals = item.get(key) or {}
            row = [item.get("item","—")]
            for p in p_all:
                v = vals.get(p)
                cell = _v(v,"$") if v is not None else "—"
                row.append(cell)
            row.append(item.get("note") or "")
            return row

        note_style = ParagraphStyle("fn", fontName="Helvetica", fontSize=7,
                                    textColor=MID, leading=10, wordWrap="CJK")
        item_style = ParagraphStyle("fi", fontName="Helvetica", fontSize=7.5,
                                    textColor=DARK, leading=11, wordWrap="CJK")
        num_style  = ParagraphStyle("fnum", fontName="Helvetica", fontSize=7.5,
                                    textColor=DARK, leading=11, alignment=2)  # right-align numbers

        def para_row(row_data, is_tot=False, is_sub=False, is_sec=False):
            result = []
            for idx, cell in enumerate(row_data):
                txt = str(cell) if cell is not None else "—"
                if is_tot:
                    s = ParagraphStyle("x", fontName="Helvetica-Bold", fontSize=7.5,
                                       textColor=GOLD_L, leading=11,
                                       alignment=2 if idx > 0 and idx < len(row_data)-1 else 0)
                elif is_sub:
                    s = ParagraphStyle("x", fontName="Helvetica-Bold", fontSize=7.5,
                                       textColor=DARK, leading=11,
                                       alignment=2 if idx > 0 and idx < len(row_data)-1 else 0)
                elif is_sec:
                    s = ParagraphStyle("x", fontName="Helvetica-Bold", fontSize=8,
                                       textColor=WARN, leading=12)
                elif idx == 0:
                    s = item_style
                elif idx == len(row_data)-1:
                    s = note_style
                else:
                    s = ParagraphStyle("x", fontName="Helvetica", fontSize=7.5,
                                       textColor=DARK, leading=11, alignment=2)
                result.append(Paragraph(txt, s))
            return result

        # Build header row
        hdr_s = ParagraphStyle("fh", fontName="Helvetica-Bold", fontSize=7.5,
                               textColor=GOLD_L, leading=11)
        p_all_rows = [[Paragraph(h, hdr_s) for h in ["Line Item"]+p_all+["Underwriting Note"]]]
        p_smap = []

        sec_row = para_row(["INCOME"]+[""]*len(p_all)+[""], is_sec=True)
        p_all_rows.append(sec_row); p_smap.append((len(p_all_rows)-1,"sec"))

        for item in inc_lines:
            is_t = item.get("is_total",False)
            is_s = item.get("is_subtotal",False)
            p_all_rows.append(para_row(frow(item), is_tot=is_t, is_sub=is_s))
            if is_t:   p_smap.append((len(p_all_rows)-1,"tot"))
            elif is_s: p_smap.append((len(p_all_rows)-1,"sub"))

        sec_row = para_row(["EXPENSES"]+[""]*len(p_all)+[""], is_sec=True)
        p_all_rows.append(sec_row); p_smap.append((len(p_all_rows)-1,"sec"))

        for item in exp_lines:
            is_t = item.get("is_total",False)
            is_s = item.get("is_subtotal",False)
            p_all_rows.append(para_row(frow(item,"values"), is_tot=is_t, is_sub=is_s))
            if is_t:   p_smap.append((len(p_all_rows)-1,"tot"))
            elif is_s: p_smap.append((len(p_all_rows)-1,"sub"))

        noi = fin.get("noi") or {}
        p_all_rows.append(para_row(["NET OPERATING INCOME"]+[_v(noi.get(p),"$") for p in p_all]+["Key bottom-line metric"], is_tot=True))
        p_smap.append((len(p_all_rows)-1,"tot"))

        capex = fin.get("capex") or {}
        p_all_rows.append(para_row(["  Capital Reserves"]+[_v(capex.get(p),"$") for p in p_all]+["Market standard $225–$300/unit/yr"]))

        cffo = fin.get("cffo") or {}
        p_all_rows.append(para_row(["CASH FLOW FROM OPERATIONS"]+[_v(cffo.get(p),"$") for p in p_all]+["NOI minus capex reserves"], is_tot=True))
        p_smap.append((len(p_all_rows)-1,"tot"))

        # Adjust column widths: line item 2.0", periods 0.8" each, note gets remainder
        # cw already set above

        t = Table(p_all_rows, colWidths=cw_fin, repeatRows=1)
        ts = [
            ("BACKGROUND",(0,0),(-1,0), DARK),
            ("FONTSIZE",  (0,0),(-1,-1), 7.5),
            ("TOPPADDING",(0,0),(-1,-1), 4), ("BOTTOMPADDING",(0,0),(-1,-1), 5),
            ("LEFTPADDING",(0,0),(-1,-1), 5), ("RIGHTPADDING",(0,0),(-1,-1), 5),
            ("GRID",(0,0),(-1,-1), 0.3, BORDER),
            ("VALIGN",(0,0),(-1,-1), "TOP"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1), [colors.white, ALT]),
        ]
        for ridx, stype in p_smap:
            if stype=="tot":
                ts += [("BACKGROUND",(0,ridx),(-1,ridx),DARK2)]
            elif stype=="sub":
                ts += [("BACKGROUND",(0,ridx),(-1,ridx),GOLD_P)]
            elif stype=="sec":
                ts += [("BACKGROUND",(0,ridx),(-1,ridx),GOLD_P)]
        t.setStyle(TableStyle(ts))
        story.append(t)
    else:
        story.append(_p("No financial analysis data found.", color=WARN))

    # 14 — Sale Comps
    story.append(PageBreak())
    _hdr(story, "Sale Comparables", "14")
    if scomps:
        story.append(_tbl(
            ["Property","Address","Date","Yr Built","Units","Sale Price","$/Unit","$/SF","Cap Rate","Occ","Buyer","Seller","Notes"],
            [[sc.get("name","—"), f"{sc.get('address') or ''} {sc.get('city_state') or ''}".strip() or "—",
              sc.get("date") or "—", _v(sc.get("year_built")), _v(sc.get("units"),"n"),
              _v(sc.get("price"),"$"), _v(sc.get("ppu"),"$"), _v(sc.get("ppsf"),"$"),
              sc.get("cap_rate") or "—", sc.get("occupancy") or "—",
              sc.get("buyer") or "—", sc.get("seller") or "—",
              sc.get("notes") or "—"] for sc in scomps],
            cw=[0.9*inch,0.85*inch,0.55*inch,0.45*inch,0.4*inch,
                0.75*inch,0.65*inch,0.5*inch,0.55*inch,0.45*inch,0.85*inch,0.85*inch,0.7*inch]
        ))
    else:
        story.append(_p("⚠ Sale comps not in OM. Source from CoStar, RCA, or listing broker.", color=WARN))

    # 15 — Flags
    story.append(PageBreak())
    _hdr(story, "Underwriting Flags & Notes", "15")
    if flags:
        cat_map = {
            "Warning":     (RED_B,    RED),
            "Opportunity": (GREEN_B,  GREEN),
            "Verify":      (PURPLE_B, PURPLE),
            "Info":        (BLUE_B,   BLUE),
        }
        wrap_detail = ParagraphStyle("fd", fontName="Helvetica", fontSize=8,
                                     textColor=DARK, leading=12, wordWrap="CJK")
        wrap_title  = ParagraphStyle("ft", fontName="Helvetica-Bold", fontSize=8,
                                     textColor=DARK, leading=12, wordWrap="CJK")
        fdata = [[
            Paragraph("Category", ParagraphStyle("fh", fontName="Helvetica-Bold", fontSize=8, textColor=GOLD_L)),
            Paragraph("Title",    ParagraphStyle("fh", fontName="Helvetica-Bold", fontSize=8, textColor=GOLD_L)),
            Paragraph("Detail",   ParagraphStyle("fh", fontName="Helvetica-Bold", fontSize=8, textColor=GOLD_L)),
        ]]
        fstyle = [
            ("BACKGROUND",(0,0),(-1,0), DARK),
            ("FONTSIZE",  (0,0),(-1,-1), 8),
            ("TOPPADDING",(0,0),(-1,-1), 6), ("BOTTOMPADDING",(0,0),(-1,-1), 6),
            ("LEFTPADDING",(0,0),(-1,-1), 8), ("RIGHTPADDING",(0,0),(-1,-1), 8),
            ("GRID",(0,0),(-1,-1), 0.3, BORDER),
            ("VALIGN",(0,0),(-1,-1), "TOP"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1), [colors.white, ALT]),
        ]
        for i, f in enumerate(flags, 1):
            cat = f.get("category","Info")
            bg, tc = cat_map.get(cat, (BLUE_B, BLUE))
            cat_style = ParagraphStyle("fc", fontName="Helvetica-Bold", fontSize=8,
                                       textColor=tc, leading=12)
            fdata.append([
                Paragraph(cat, cat_style),
                Paragraph(f.get("title",""), wrap_title),
                Paragraph(f.get("detail",""), wrap_detail),   # full text, no truncation
            ])
            fstyle += [("BACKGROUND",(0,i),(0,i), bg)]
        # Column widths: Category=0.9", Title=1.6", Detail=5.0" — detail gets most space
        ft = Table(fdata, colWidths=[0.9*inch, 1.6*inch, 5.0*inch])
        ft.setStyle(TableStyle(fstyle))
        story.append(ft)

    # Market overview
    if any(v for v in mkt.values() if v):
        story.append(PageBreak())
        _hdr(story, "Market & Submarket Overview", "16")

        if mkt.get("market_summary"):
            story.append(_p("Market Summary", fontname="Helvetica-Bold", size=11, space_before=4, space_after=4))
            story.append(_p(mkt["market_summary"], color=DARK))

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
            story.append(_p("Submarket Metrics", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
            story.append(_kv(mkt_rows))

        devs = mkt.get("major_developments") or []
        if devs and any(d.get("name") for d in devs):
            story.append(_p("Major Economic Developments", fontname="Helvetica-Bold", size=11, space_before=10, space_after=4))
            story.append(_tbl(
                ["Development","Description","Est. Cost","Jobs","Timeline"],
                [[d.get("name","—"), d.get("description") or "—",
                  d.get("cost") or "—", d.get("jobs") or "—",
                  d.get("timeline") or "—"] for d in devs if d.get("name")],
                cw=[1.5*inch, 2.8*inch, 1.0*inch, 0.9*inch, 1.3*inch]
            ))

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
Upload any Multifamily Offering Memorandum PDF and get a structured 16-section underwriting report.

**Covers:**
- Basic property details
- Financing & debt terms *(new)*
- Investment highlights & renovation tiers
- Proposed value-add by floor plan
- Property tax & abatement
- Replacement cost, insurance & management *(new)*
- Affordability & rent growth runway *(new)*
- Demographics & schools
- Unit mix with rent upside
- Utility billing info
- Site & systems info
- Rent comparables (with distance & year built) *(new)*
- Financial analysis
- Sale comparables (with buyer/seller) *(new)*
- Underwriting flags
- Market & submarket overview with developments *(new)*
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
    va_  = data.get("value_add") or {}
    tax_ = data.get("tax") or {}

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    with m1: st.metric("Units",         _v(pd_.get("units"),"n"))
    with m2: st.metric("Year Built",    _v(pd_.get("year_built")))
    with m3: st.metric("Occupancy",     _v(pd_.get("occupancy_pct"),"%"))
    with m4: st.metric("Avg Rent",      _v(inv_.get("market_rent"),"$"))
    with m5: st.metric("Reno ROI",      f"{va_.get('roi_pct') or 'N/A'}%")
    with m6: st.metric("Tax Savings",   _v(tax_.get("abatement_annual_savings"),"$"))

    # ── Tabs for data preview ─────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "Unit Mix", "Value-Add", "Rent Comps", "Financials", "Tax & Abatement", "Demographics", "Financing", "Flags"
    ])

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
        import pandas as pd
        va_plans = va_.get("by_floor_plan") or []
        if va_plans:
            st.markdown('<div class="gold-header">Floor Plan Value-Add</div>', unsafe_allow_html=True)
            st.dataframe(pd.DataFrame([{
                "Type":            p.get("type"),
                "SF":              p.get("sf"),
                "Units":           p.get("units"),
                "In-Place Rent":   p.get("inplace_rent"),
                "Rehab Cost":      p.get("rehab_cost"),
                "Premium/Unit":    p.get("premium"),
                "Post-Rehab Rent": p.get("post_rehab_rent"),
                "Post-Rehab PSF":  p.get("post_rehab_psf"),
            } for p in va_plans]), use_container_width=True, hide_index=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Reno Cost",   _v(va_.get("total_cost"),"$"))
            st.metric("Cost Per Unit",     _v(va_.get("cost_per_unit"),"$"))
        with col2:
            st.metric("Annual Premium",    _v(va_.get("annual_premium"),"$"))
            st.metric("Monthly Premium",   _v(va_.get("monthly_premium"),"$"))
        with col3:
            st.metric("ROI",               f"{va_.get('roi_pct') or 'N/A'}%")
            st.metric("Exterior CapEx",    _v(va_.get("exterior_capex"),"$"))

    with tab3:
        import pandas as pd
        rg_  = data.get("rent_comps_garden") or []
        rth_ = data.get("rent_comps_townhouse") or []
        if rg_:
            st.markdown('<div class="gold-header">Two Bedroom Garden</div>', unsafe_allow_html=True)
            st.dataframe(pd.DataFrame([{"Property": c.get("name"), "Rent": c.get("rent"), "Notes": c.get("notes")} for c in rg_]),
                         use_container_width=True, hide_index=True)
        if rth_:
            st.markdown('<div class="gold-header">Two Bedroom Townhouse</div>', unsafe_allow_html=True)
            st.dataframe(pd.DataFrame([{"Property": c.get("name"), "Rent": c.get("rent"), "Notes": c.get("notes")} for c in rth_]),
                         use_container_width=True, hide_index=True)
        if not rg_ and not rth_:
            st.info("No comp data extracted.")

    with tab4:
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
                row["Note"] = (line.get("note") or "")[:120]
                rows.append(row)
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        else:
            st.info("No financial data extracted.")

    with tab5:
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown('<div class="gold-header">Property Tax Detail</div>', unsafe_allow_html=True)
            for label, val in [
                ("Parcel ID",           tax_.get("parcel_id")),
                ("Assessed Value",      _v(tax_.get("assessed_value"),"$")),
                ("Millage — City",      tax_.get("millage_city")),
                ("Millage — County",    tax_.get("millage_county")),
                ("Total Millage",       tax_.get("millage_total")),
                ("Ad Valorem Tax",      _v(tax_.get("tax_base"),"$")),
                ("Solid Waste Fee",     _v(tax_.get("solid_waste_fee"),"$")),
                ("Total Tax Bill",      _v(tax_.get("total_tax"),"$")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")
        with col_b:
            st.markdown('<div class="gold-header">Tax Abatement Program</div>', unsafe_allow_html=True)
            for label, val in [
                ("Program",             tax_.get("abatement_program")),
                ("Abatement %",         f"{tax_.get('abatement_pct') or 'N/A'}%"),
                ("Term",                tax_.get("abatement_term_note")),
                ("AMI Requirement",     f"{tax_.get('ami_pct') or 'N/A'}% AMI"),
                ("Max Allowable Rent",  _v(tax_.get("max_allowable_rent"),"$")),
                ("Avg In-Place Rent",   _v(tax_.get("avg_inplace_rent"),"$")),
                ("Headroom/Unit",       _v(tax_.get("rent_headroom"),"$")),
                ("Units Compliant",     f"{tax_.get('units_compliant') or 'N/A'} ({tax_.get('pct_compliant') or 'N/A'}%)"),
                ("Annual Tax Savings",  _v(tax_.get("abatement_annual_savings"),"$")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")

    with tab6:
        demo_ = data.get("demographics") or {}
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown('<div class="gold-header">3-Mile Radius</div>', unsafe_allow_html=True)
            for label, val in [
                ("Population (2025)",    _v(demo_.get("pop_3mi"),"n")),
                ("Population (2030)",    _v(demo_.get("pop_2030_3mi"),"n")),
                ("Pop. Growth (5-yr)",   demo_.get("pop_growth_3mi")),
                ("Median HH Income",     _v(demo_.get("median_income_3mi"),"$")),
                ("Income 2030 Proj.",    _v(demo_.get("median_income_2030_3mi"),"$")),
                ("Income Growth",        demo_.get("income_growth_3mi")),
                ("Renter-Occupied",      demo_.get("renter_pct_3mi")),
                ("College Educated",     demo_.get("college_pct_3mi")),
                ("White-Collar",         demo_.get("white_collar_pct_3mi")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")
        with col_b:
            st.markdown('<div class="gold-header">5-Mile Radius</div>', unsafe_allow_html=True)
            for label, val in [
                ("Population (2025)",    _v(demo_.get("pop_5mi"),"n")),
                ("Population (2030)",    _v(demo_.get("pop_2030_5mi"),"n")),
                ("Pop. Growth (5-yr)",   demo_.get("pop_growth_5mi")),
                ("Median HH Income",     _v(demo_.get("median_income_5mi"),"$")),
                ("Income 2030 Proj.",    _v(demo_.get("median_income_2030_5mi"),"$")),
                ("Income Growth",        demo_.get("income_growth_5mi")),
                ("Renter-Occupied",      demo_.get("renter_pct_5mi")),
                ("College Educated",     demo_.get("college_pct_5mi")),
                ("White-Collar",         demo_.get("white_collar_pct_5mi")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")
        employers_ = demo_.get("employers") or []
        if employers_:
            st.markdown('<div class="gold-header">Major Employers</div>', unsafe_allow_html=True)
            import pandas as pd
            st.dataframe(pd.DataFrame(employers_), use_container_width=True, hide_index=True)

    with tab7:
        fin_i = data.get("financing") or {}
        afford_i = data.get("affordability") or {}
        insur_i = data.get("insurance") or {}
        mgmt_i = data.get("management") or {}
        repl_i = data.get("replacement_cost") or {}

        col1, col2 = st.columns(2)
        st.markdown(f"**Offering Type:** {fin_i.get('offering_type') or 'N/A'}  &nbsp;|&nbsp;  **Debt Contact:** {fin_i.get('debt_contact') or 'N/A'}", unsafe_allow_html=True)
        if fin_i.get("notes"):
            st.markdown(f"*{fin_i['notes']}*")
        st.markdown("")

        nf_i  = fin_i.get("new_financing") or {}
        asd_i = fin_i.get("assumable_debt") or {}

        col1, col2 = st.columns(2)
        with col1:
            st.markdown('<div class="gold-header">New Financing</div>', unsafe_allow_html=True)
            for label, val in [
                ("Loan Type",        nf_i.get("loan_type")),
                ("Lender",           nf_i.get("lender")),
                ("Loan Amount",      _v(nf_i.get("loan_amount"),"$")),
                ("LTV",              f"{nf_i.get('loan_to_value') or 'N/A'}%"),
                ("Interest Rate",    f"{nf_i.get('interest_rate') or 'N/A'}%"),
                ("Rate Type",        nf_i.get("rate_type")),
                ("Loan Term",        nf_i.get("loan_term_years")),
                ("Amortization",     nf_i.get("amortization_years")),
                ("Interest Only",    nf_i.get("interest_only_period")),
                ("DSCR",             nf_i.get("dscr")),
                ("Recourse",         nf_i.get("recourse")),
                ("Notes",            nf_i.get("notes")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")

        with col2:
            st.markdown('<div class="gold-header">Assumable Debt</div>', unsafe_allow_html=True)
            for label, val in [
                ("Loan Type",          asd_i.get("loan_type")),
                ("Lender",             asd_i.get("lender")),
                ("Loan Amount",        _v(asd_i.get("loan_amount"),"$")),
                ("LTV",                f"{asd_i.get('loan_to_value') or 'N/A'}%"),
                ("Interest Rate",      f"{asd_i.get('interest_rate') or 'N/A'}%"),
                ("Rate Type",          asd_i.get("rate_type")),
                ("Loan Term",          asd_i.get("loan_term_years")),
                ("Amortization",       asd_i.get("amortization_years")),
                ("Origination Date",   asd_i.get("origination_date")),
                ("Maturity Date",      asd_i.get("maturity_date")),
                ("Monthly Payment",    _v(asd_i.get("monthly_payment"),"$")),
                ("Annual Debt Svc",    _v(asd_i.get("annual_debt_service"),"$")),
                ("DSCR",               asd_i.get("dscr")),
                ("Prepayment Penalty", asd_i.get("prepayment_penalty")),
                ("Recourse",           asd_i.get("recourse")),
                ("Notes",              asd_i.get("notes")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")

            st.markdown('<div class="gold-header">Insurance</div>', unsafe_allow_html=True)
            for label, val in [
                ("Carrier",       insur_i.get("carrier")),
                ("Annual Premium", _v(insur_i.get("annual_premium"),"$")),
                ("Per Unit/Year", _v(insur_i.get("per_unit"),"$")),
                ("Quote Source",  insur_i.get("quote_source")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")

        with col2:
            st.markdown('<div class="gold-header">Management</div>', unsafe_allow_html=True)
            for label, val in [
                ("Fee %",            f"{mgmt_i.get('fee_pct') or 'N/A'}% of EGI"),
                ("Annual Fee",       _v(mgmt_i.get("fee_annual"),"$")),
                ("Per Unit",         _v(mgmt_i.get("fee_per_unit"),"$")),
                ("Current Manager",  mgmt_i.get("current_manager")),
                ("Proposed Manager", mgmt_i.get("proposed_manager")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")

            st.markdown('<div class="gold-header">Replacement Cost</div>', unsafe_allow_html=True)
            for label, val in [
                ("Per Unit",        _v(repl_i.get("per_unit"),"$")),
                ("Per SF",          _v(repl_i.get("per_sf"),"$")),
                ("Hard Cost/SF",    _v(repl_i.get("hard_cost_per_sf"),"$")),
                ("Land/Unit",       _v(repl_i.get("land_per_unit"),"$")),
                ("Soft Cost %",     f"{repl_i.get('soft_cost_pct') or 'N/A'}%"),
                ("Source",          repl_i.get("source")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")

        st.markdown('<div class="gold-header">Affordability Analysis</div>', unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**2025**")
            for label, val in [
                ("Current Rent",           _v(afford_i.get("current_rent"),"$")),
                ("Avg HH Income (3-mi)",   _v(afford_i.get("avg_hh_income_3mi"),"$")),
                ("Affordability @ 3×",     _v(afford_i.get("monthly_affordability_3x"),"$")),
                ("Rent Headroom",          _v(afford_i.get("rent_headroom_3mi"),"$")),
                ("Income-to-Rent Ratio",   afford_i.get("income_to_rent_ratio")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")
        with col_b:
            st.markdown("**2030 Projection**")
            for label, val in [
                ("Avg HH Income (3-mi)",   _v(afford_i.get("avg_hh_income_2030_3mi"),"$")),
                ("Affordability @ 3×",     _v(afford_i.get("monthly_affordability_2030_3x"),"$")),
                ("Rent Headroom",          _v(afford_i.get("rent_headroom_2030_3mi"),"$")),
            ]:
                st.markdown(f"**{label}:** {val or 'N/A'}")

    with tab8:
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

