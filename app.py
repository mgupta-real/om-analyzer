import os, json, re, tempfile, io
from datetime import datetime
import streamlit as st
import anthropic
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1

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
    "additional_income": [
      {
        "name": null,
        "category": null,
        "fee_per_unit_per_month": null,
        "occupancy_assumption": null,
        "monthly_income": null,
        "current_annual": null,
        "proforma_annual": null,
        "calculation_detail": null,
        "notes": null
      }
    ],
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

6. OTHER INCOME: Extract EVERY other income line item from the OM into additional_income array.
   This includes but is not limited to:
   - Resident fees: admin fee, application fee, pet rent, pet fee (one-time), NSF fee,
     late fee, early termination fee, transfer fee, lease violation fee
   - Utility fees: water/sewer reimbursement, trash fee, valet trash, common area electric fee
   - Amenity fees: community amenity fee, package locker fee, reserved parking fee,
     storage fee, elevator building premium, garage fee
   - Service fees: pest control fee, internet/cable/bulk internet fee
   - Other: miscellaneous income, laundry income, renter's insurance reimbursement
   For each item extract: name, category (Resident Fee / Utility / Amenity / Service / Other),
   fee_per_unit_per_month, occupancy_assumption (e.g. "95%"), monthly_income,
   current_annual, proforma_annual, and calculation_detail (full description of how calculated).
   Do NOT skip any line item — extract every single one mentioned anywhere in the OM.

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
        model="claude-haiku-4-5-20251001",
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
# SECTION 3 — EXCEL REPORT GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

# ── Colour palette (hex strings for openpyxl, no leading #) ──────────────────
C_DARK   = "1A1A18"
C_DARK2  = "252522"
C_GOLD   = "B8965A"
C_GOLD_L = "D4B07A"
C_GOLD_P = "F5EDD8"
C_MID    = "777770"
C_ALT    = "F5F4EF"
C_WHITE  = "FFFFFF"
C_GREEN  = "EAF5EE"
C_RED    = "FAEAEA"
C_BLUE   = "EAF0FA"
C_PURPLE = "F5F0FF"
C_WARN   = "FFF8EC"



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


# ── openpyxl style helpers ────────────────────────────────────────────────────

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color=C_DARK, size=10, italic=False):
    return Font(name="Arial", bold=bold, color=color, size=size, italic=italic)

def _align(h="left", v="top", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border_thin():
    s = Side(style="thin", color="E0DED5")
    return Border(left=s, right=s, top=s, bottom=s)

def _border_bottom():
    s = Side(style="medium", color=C_GOLD)
    return Border(bottom=s)


class XLWriter:
    """Thin wrapper around an openpyxl worksheet for sequential row writing."""

    def __init__(self, ws):
        self.ws = ws
        self.row = 1
        self.ws.sheet_view.showGridLines = False

    @staticmethod
    def _auto_height(text, col_width, font_size=9, min_h=15, padding=6):
        """Estimate row height needed to display text without clipping."""
        if not text:
            return min_h
        chars_per_line = max(1, int(col_width * 1.3))
        text_str = str(text)
        lines = 0
        for segment in text_str.split("\n"):
            lines += max(1, -(-len(segment) // chars_per_line))
        line_height = font_size * 1.4
        return max(min_h, int(lines * line_height) + padding)

    def _set(self, cell, value, bold=False, bg=None, fg=C_DARK, size=10,
             align_h="left", wrap=True, italic=False, border=True, num_fmt=None):
        c = self.ws[cell]
        c.value = value
        c.font = _font(bold=bold, color=fg, size=size, italic=italic)
        c.alignment = _align(h=align_h, v="top", wrap=wrap)
        if bg:
            c.fill = _fill(bg)
        if border:
            c.border = _border_thin()
        if num_fmt:
            c.number_format = num_fmt

    def _fill_row(self, bg, col_start=1, col_end=13):
        """Fill background across a full row without merging."""
        for i in range(col_start, col_end + 1):
            col = get_column_letter(i)
            c = self.ws[f"{col}{self.row}"]
            c.fill = _fill(bg)

    def section(self, title):
        """Dark gold section header — title in col A, background fills A–M."""
        self.row += 1
        self._fill_row(C_DARK)
        self._set(f"A{self.row}", title, bold=True, bg=C_DARK, fg=C_GOLD_L,
                  size=11, border=False)
        self.ws.row_dimensions[self.row].height = 24
        self.row += 1

    def subsection(self, title):
        """Gold-tinted subsection label."""
        self._fill_row(C_GOLD_P)
        self._set(f"A{self.row}", title, bold=True, bg=C_GOLD_P, fg=C_DARK,
                  size=10, border=False)
        self.ws.row_dimensions[self.row].height = 20
        self.row += 1

    def kv(self, label, value, alt=False):
        """Key-value row: col A = label, col B = value (no merge)."""
        bg = C_ALT if alt else C_WHITE
        val_str = str(value) if value else "N/A"
        self._set(f"A{self.row}", label, fg=C_MID, bg=bg, size=9)
        self._set(f"B{self.row}", val_str, bg=bg, size=9, wrap=True)
        # Fill remaining columns with same background
        for col in ["C","D","E","F","G","H","I","J","K","L","M"]:
            c = self.ws[f"{col}{self.row}"]
            c.fill = _fill(bg)
        h = self._auto_height(val_str, col_width=40)
        self.ws.row_dimensions[self.row].height = h
        self.row += 1

    def blank(self, n=1):
        self.row += n

    def table_header(self, headers, col_start=1):
        """Write a dark header row for a data table."""
        for i, h in enumerate(headers):
            col = get_column_letter(col_start + i)
            self._set(f"{col}{self.row}", h, bold=True, bg=C_DARK,
                      fg=C_GOLD_L, size=9, align_h="center")
        self.ws.row_dimensions[self.row].height = 18
        self.row += 1

    def table_row(self, values, col_start=1, alt=False, col_widths=None):
        """Write a data row; auto-size height based on longest cell."""
        bg = C_ALT if alt else C_WHITE
        col_widths = col_widths or []
        max_h = 15
        for i, v in enumerate(values):
            col = get_column_letter(col_start + i)
            val_str = str(v) if v is not None else "—"
            self._set(f"{col}{self.row}", val_str, bg=bg, size=9, wrap=True)
            cw = col_widths[i] if i < len(col_widths) else 16
            max_h = max(max_h, self._auto_height(val_str, col_width=cw))
        self.ws.row_dimensions[self.row].height = max_h
        self.row += 1

    def side_by_side_headers(self, left_title, right_title):
        """Two dark headers side by side — left in A, right in F."""
        self._fill_row(C_DARK)
        self._set(f"A{self.row}", left_title, bold=True, bg=C_DARK,
                  fg=C_GOLD_L, size=10, border=False)
        self._set(f"F{self.row}", right_title, bold=True, bg=C_DARK2,
                  fg=C_GOLD_L, size=10, border=False)
        # Dark2 for right side cols F-I
        for col in ["G","H","I"]:
            self.ws[f"{col}{self.row}"].fill = _fill(C_DARK2)
        self.ws.row_dimensions[self.row].height = 20
        self.row += 1

    def side_by_side_kv(self, left_label, left_val, right_label, right_val, alt=False):
        bg = C_ALT if alt else C_WHITE
        lv = str(left_val) if left_val else "N/A"
        rv = str(right_val) if right_val else "N/A"
        self._set(f"A{self.row}", left_label, fg=C_MID, bg=bg, size=9)
        self._set(f"B{self.row}", lv, bg=bg, size=9, wrap=True)
        # Fill C-E with bg
        for col in ["C","D","E"]:
            self.ws[f"{col}{self.row}"].fill = _fill(bg)
        self._set(f"F{self.row}", right_label, fg=C_MID, bg=bg, size=9)
        self._set(f"G{self.row}", rv, bg=bg, size=9, wrap=True)
        # Fill H-I with bg
        for col in ["H","I"]:
            self.ws[f"{col}{self.row}"].fill = _fill(bg)
        h = max(
            self._auto_height(lv, col_width=40),
            self._auto_height(rv, col_width=40),
            15
        )
        self.ws.row_dimensions[self.row].height = h
        self.row += 1

    def set_col_widths(self, widths):
        for col_letter, w in widths.items():
            self.ws.column_dimensions[col_letter].width = w




def build_excel(d: dict, filename: str) -> bytes:
    date = datetime.today().strftime("%B %d, %Y")
    prop     = (d.get("property") or {}).get("name") or "Property"
    broker_d = d.get("broker") or {}
    inv      = d.get("investment") or {}
    demo     = d.get("demographics") or {}
    umix     = d.get("unit_mix") or []
    utils_   = d.get("utilities") or []
    site     = d.get("site") or {}
    rcomps   = d.get("rent_comps") or []
    rg       = d.get("rent_comps_garden") or []
    rth      = d.get("rent_comps_townhouse") or []
    fin      = d.get("financials") or {}
    scomps   = d.get("sale_comps") or []
    flags    = d.get("flags") or []
    mkt      = d.get("market") or {}
    pd_      = d.get("property") or {}
    va       = d.get("value_add") or {}
    tax      = d.get("tax") or {}
    fin_info = d.get("financing") or {}
    afford   = d.get("affordability") or {}
    insur    = d.get("insurance") or {}
    mgmt     = d.get("management") or {}
    repl     = d.get("replacement_cost") or {}

    wb = Workbook()
    ws = wb.active
    ws.title = "Underwriting Report"
    xl = XLWriter(ws)

    addr = f"{pd_.get('address','')}, {pd_.get('city','')}, {pd_.get('state','')} {pd_.get('zip','')}".strip(", ")

    # Set column widths for single sheet
    xl.set_col_widths({
        "A": 32, "B": 40, "C": 16, "D": 16,
        "E": 4,  "F": 32, "G": 40, "H": 16,
        "I": 16, "J": 16, "K": 16, "L": 16, "M": 16,
    })

    # ── Cover rows ────────────────────────────────────────────────────────────
    xl._fill_row(C_DARK)
    xl._set("A1", "MULTIFAMILY UNDERWRITING REPORT",
            bold=True, bg=C_DARK, fg=C_GOLD_L, size=12,
            align_h="left", border=False)
    ws.row_dimensions[1].height = 24
    xl.row = 2
    xl._fill_row(C_DARK)
    xl._set("A2", prop, bold=True, bg=C_DARK, fg=C_WHITE,
            size=16, align_h="left", border=False)
    ws.row_dimensions[2].height = 36
    xl.row = 3
    xl._fill_row(C_DARK2)
    xl._set("A3",
        f"Address: {addr}   ·   Broker: {broker_d.get('name','N/A')}   ·   "
        f"Units: {_v(pd_.get('units'),'n')}   ·   Year Built: {_v(pd_.get('year_built'))}   ·   "
        f"Report: {date}   ·   Source: {filename}",
        bg=C_DARK2, fg=C_GOLD_L, size=9, align_h="left", border=False, wrap=True)
    ws.row_dimensions[3].height = 18
    xl.row = 4

    # ── 1. Basic Property Details ─────────────────────────────────────────────
    xl.section("1. BASIC PROPERTY DETAILS")
    agents = ", ".join(
        a.get("name", str(a)) if isinstance(a, dict) else str(a)
        for a in (broker_d.get("agents") or [])
    ) or "N/A"
    for i, (k, v) in enumerate([
        ("Property Name",      prop),
        ("Address",            addr),
        ("County",             pd_.get("county")),
        ("MSA",                pd_.get("msa")),
        ("Total Units",        _v(pd_.get("units"),"n")),
        ("Year Built",         _v(pd_.get("year_built"))),
        ("Net Rentable Area",  _v(pd_.get("rentable_sf"),"n",suffix=" SF")),
        ("Avg Unit Size",      _v(pd_.get("avg_unit_sf"),"n",suffix=" SF")),
        ("Occupancy",          f"{_v(pd_.get('occupancy_pct'),'%')} as of {pd_.get('occupancy_date','N/A')}"),
        ("Buildings / Stories",f"{_v(pd_.get('buildings'),'n')} bldgs / {_v(pd_.get('floors'))} stories"),
        ("Land Area",          _v(pd_.get("acres"),suffix=" acres")),
        ("Density",            _v(pd_.get("density"),suffix=" units/acre")),
        ("Developer",          pd_.get("developer")),
        ("Asset Class",        pd_.get("asset_class")),
        ("Broker",             broker_d.get("name")),
        ("Listing Agents",     agents),
    ]):
        xl.kv(k, v, alt=bool(i % 2))

    # ── 2. Financing ──────────────────────────────────────────────────────────
    xl.section("2. FINANCING & DEBT TERMS")
    for i, (k, v) in enumerate([
        ("Offering Type",  fin_info.get("offering_type")),
        ("Debt Contact",   fin_info.get("debt_contact")),
        ("Notes",          fin_info.get("notes")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))
    xl.blank()

    nf  = fin_info.get("new_financing") or {}
    asd = fin_info.get("assumable_debt") or {}
    xl.side_by_side_headers("NEW FINANCING", "ASSUMABLE DEBT")

    nf_rows  = [
        ("Loan Type",      nf.get("loan_type")),
        ("Lender",         nf.get("lender")),
        ("Loan Amount",    _v(nf.get("loan_amount"),"$")),
        ("LTV",            f"{nf.get('loan_to_value') or 'N/A'}%"),
        ("Interest Rate",  f"{nf.get('interest_rate') or 'N/A'}%"),
        ("Rate Type",      nf.get("rate_type")),
        ("Loan Term",      nf.get("loan_term_years")),
        ("Amortization",   nf.get("amortization_years")),
        ("Interest-Only",  nf.get("interest_only_period")),
        ("DSCR",           nf.get("dscr")),
        ("Recourse",       nf.get("recourse")),
        ("Notes",          nf.get("notes")),
    ]
    asd_rows = [
        ("Loan Type",          asd.get("loan_type")),
        ("Lender",             asd.get("lender")),
        ("Loan Amount",        _v(asd.get("loan_amount"),"$")),
        ("LTV",                f"{asd.get('loan_to_value') or 'N/A'}%"),
        ("Interest Rate",      f"{asd.get('interest_rate') or 'N/A'}%"),
        ("Rate Type",          asd.get("rate_type")),
        ("Loan Term",          asd.get("loan_term_years")),
        ("Amortization",       asd.get("amortization_years")),
        ("Origination Date",   asd.get("origination_date")),
        ("Maturity Date",      asd.get("maturity_date")),
        ("Monthly Payment",    _v(asd.get("monthly_payment"),"$")),
        ("Annual Debt Svc",    _v(asd.get("annual_debt_service"),"$")),
        ("DSCR",               asd.get("dscr")),
        ("Prepayment Penalty", asd.get("prepayment_penalty")),
        ("Recourse",           asd.get("recourse")),
        ("Notes",              asd.get("notes")),
    ]
    max_r = max(len(nf_rows), len(asd_rows))
    nf_rows  += [("","")] * (max_r - len(nf_rows))
    asd_rows += [("","")] * (max_r - len(asd_rows))
    for i, ((lk, lv), (rk, rv)) in enumerate(zip(nf_rows, asd_rows)):
        xl.side_by_side_kv(lk, lv, rk, rv, alt=bool(i % 2))

    # ── 3. Investment Highlights ──────────────────────────────────────────────
    xl.section("3. INVESTMENT HIGHLIGHTS & OPPORTUNITIES")
    xl.subsection("Rent Overview")
    for i, (k, v) in enumerate([
        ("Current Market Rent",      f"{_v(inv.get('market_rent'),'$')} ({_v(inv.get('market_rent_psf'))} PSF)"),
        ("Current Effective Rent",   f"{_v(inv.get('effective_rent'),'$')} ({_v(inv.get('effective_rent_psf'))} PSF)"),
        ("Comp Set Effective Rent",  _v(inv.get("comp_effective_rent"),"$")),
        ("Rent Gap / Unit",          _v(inv.get("rent_gap"),"$")),
        ("Annual Upside (total)",    _v(inv.get("annual_upside"),"$")),
        ("Below Replacement Cost",   inv.get("replacement_cost_note")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))

    tiers = inv.get("renovation_tiers") or []
    if tiers:
        xl.blank()
        xl.subsection("Renovation Tiers")
        xl.table_header(["Tier","Units","Appliances","Cabinets","Countertops","Flooring","Premium"])
        for i, t in enumerate(tiers):
            xl.table_row([
                t.get("name","—"), _v(t.get("units"),"n"),
                t.get("appliances") or "—", t.get("cabinets") or "—",
                t.get("countertops") or "—", t.get("flooring") or "—",
                _v(t.get("premium"),"$") if t.get("premium") else "—"
            ], alt=bool(i % 2))

    light_items = inv.get("light_upgrade_items") or va.get("light_upgrade_items") or []
    if light_items:
        xl.blank()
        xl.subsection("Light Interior Upgrade Items")
        for i, item in enumerate(light_items):
            xl.kv(f"  {i+1}.", item, alt=bool(i % 2))

    add_inc = inv.get("additional_income") or []
    if add_inc:
        xl.blank()
        xl.subsection("Additional Income Opportunities")
        xl.table_header([
            "Income Item", "Category", "Fee/Unit/Mo",
            "Occupancy %", "Monthly Income",
            "Current Annual", "Pro Forma Annual", "Calculation Detail"
        ])
        for i, inc in enumerate(add_inc):
            xl.table_row([
                inc.get("name") or "—",
                inc.get("category") or "—",
                _v(inc.get("fee_per_unit_per_month"), "$") if inc.get("fee_per_unit_per_month") else "—",
                inc.get("occupancy_assumption") or "—",
                _v(inc.get("monthly_income"), "$") if inc.get("monthly_income") else "—",
                _v(inc.get("current_annual"), "$") if inc.get("current_annual") else "—",
                _v(inc.get("proforma_annual"), "$") if inc.get("proforma_annual") else "—",
                inc.get("calculation_detail") or inc.get("notes") or "—",
            ], alt=bool(i % 2))

    if inv.get("amenities"):
        xl.blank()
        xl.subsection("Community Amenities")
        xl.kv("Amenities", "  ·  ".join(inv["amenities"]))
    if inv.get("unit_features"):
        xl.blank()
        xl.subsection("Unit Features")
        xl.kv("Features", "  ·  ".join(inv["unit_features"]))
    if inv.get("highlights"):
        xl.blank()
        xl.subsection("Other Highlights")
        for i, h in enumerate(inv["highlights"]):
            xl.kv(f"  {i+1}.", h, alt=bool(i % 2))

    # ── 4. Value-Add ──────────────────────────────────────────────────────────
    xl.section("4. PROPOSED VALUE-ADD BY FLOOR PLAN")
    plans = va.get("by_floor_plan") or []
    if plans:
        xl.table_header(["Unit Type","SF","Units","In-Place Rent","In-Place PSF","Rehab Cost/Unit","Premium/Unit","Post-Rehab Rent","Post-Rehab PSF"])
        for i, p in enumerate(plans):
            xl.table_row([
                p.get("type","—"), _v(p.get("sf"),"n"), _v(p.get("units"),"n"),
                _v(p.get("inplace_rent"),"$"), f"${p.get('inplace_psf') or 0:.2f}",
                _v(p.get("rehab_cost"),"$"), _v(p.get("premium"),"$"),
                _v(p.get("post_rehab_rent"),"$"), f"${p.get('post_rehab_psf') or 0:.2f}"
            ], alt=bool(i % 2))
    xl.blank()
    xl.subsection("CapEx & ROI Summary")
    for i, (k, v) in enumerate([
        ("Total Renovation Cost",      _v(va.get("total_cost"),"$")),
        ("Cost Per Unit (all-in)",     _v(va.get("cost_per_unit"),"$")),
        ("Exterior / Additional CapEx",_v(va.get("exterior_capex"),"$")),
        ("Monthly Rent Premium",       _v(va.get("monthly_premium"),"$")),
        ("Annual Rent Premium",        _v(va.get("annual_premium"),"$")),
        ("Return on Investment",       f"{va.get('roi_pct') or 'N/A'}%"),
        ("Value-Add Scope",            va.get("scope")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))

    # ── 5. Tax & Abatement ────────────────────────────────────────────────────
    xl.section("5. PROPERTY TAX & TAX ABATEMENT")
    xl.subsection("Property Tax Detail")
    for i, (k, v) in enumerate([
        ("Parcel ID",             tax.get("parcel_id")),
        ("Assessed Market Value", _v(tax.get("assessed_value"),"$")),
        ("Millage — City",        tax.get("millage_city")),
        ("Millage — County",      tax.get("millage_county")),
        ("Total Millage Rate",    tax.get("millage_total")),
        ("Ad Valorem Tax",        _v(tax.get("tax_base"),"$")),
        ("Solid Waste / Fees",    _v(tax.get("solid_waste_fee"),"$")),
        ("Total Annual Tax Bill", _v(tax.get("total_tax"),"$")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))
    xl.blank()
    xl.subsection("Tax Abatement Program")
    for i, (k, v) in enumerate([
        ("Program",               tax.get("abatement_program")),
        ("Abatement Percentage",  f"{tax.get('abatement_pct') or 'N/A'}%"),
        ("Commitment Term",       tax.get("abatement_term_note")),
        ("AMI Requirement",       f"{tax.get('ami_pct') or 'N/A'}% of Area Median Income"),
        ("Potential Annual Savings", _v(tax.get("abatement_annual_savings"),"$")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))
    xl.blank()
    xl.subsection("AMI Rent Compliance")
    xl.table_header(["Bedroom Type","Units","Max Allowable Rent","Avg In-Place Rent","Headroom/Unit","Units Compliant","% Compliant"])
    xl.table_row(["2 Bedroom",
        _v(tax.get("units_compliant"),"n"), _v(tax.get("max_allowable_rent"),"$"),
        _v(tax.get("avg_inplace_rent"),"$"), _v(tax.get("rent_headroom"),"$"),
        _v(tax.get("units_compliant"),"n"), f"{tax.get('pct_compliant') or 'N/A'}%"])

    # ── 6. Replacement Cost, Insurance & Management ───────────────────────────
    xl.section("6. REPLACEMENT COST, INSURANCE & MANAGEMENT")
    xl.subsection("Replacement Cost")
    for i, (k, v) in enumerate([
        ("Cost Per Unit",      _v(repl.get("per_unit"),"$")),
        ("Cost Per SF",        _v(repl.get("per_sf"),"$")),
        ("Total Replacement",  _v(repl.get("total"),"$")),
        ("Land Per Unit",      _v(repl.get("land_per_unit"),"$")),
        ("Hard Cost Per SF",   _v(repl.get("hard_cost_per_sf"),"$")),
        ("Soft Cost %",        f"{repl.get('soft_cost_pct') or 'N/A'}%"),
        ("Source",             repl.get("source")),
        ("Notes",              repl.get("notes")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))
    xl.blank()
    xl.subsection("Insurance")
    for i, (k, v) in enumerate([
        ("Carrier / Provider", insur.get("carrier")),
        ("Annual Premium",     _v(insur.get("annual_premium"),"$")),
        ("Per Unit / Year",    _v(insur.get("per_unit"),"$")),
        ("Quote Source",       insur.get("quote_source")),
        ("Notes",              insur.get("notes")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))
    xl.blank()
    xl.subsection("Property Management")
    for i, (k, v) in enumerate([
        ("Management Fee %",   f"{mgmt.get('fee_pct') or 'N/A'}% of EGI"),
        ("Annual Fee",         _v(mgmt.get("fee_annual"),"$")),
        ("Per Unit / Year",    _v(mgmt.get("fee_per_unit"),"$")),
        ("Current Manager",    mgmt.get("current_manager")),
        ("Proposed Manager",   mgmt.get("proposed_manager")),
        ("Notes",              mgmt.get("notes")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))

    # ── 7. Affordability ──────────────────────────────────────────────────────
    xl.section("7. AFFORDABILITY & RENT GROWTH RUNWAY")
    for i, (k, v) in enumerate([
        ("Current In-Place Rent",              _v(afford.get("current_rent"),"$")),
        ("Avg HH Income — 3-Mile (2025)",      _v(afford.get("avg_hh_income_3mi"),"$")),
        ("Monthly Affordability @ 3× Rent",    _v(afford.get("monthly_affordability_3x"),"$")),
        ("Rent Headroom (3-Mile, 2025)",        _v(afford.get("rent_headroom_3mi"),"$")),
        ("Avg HH Income — 3-Mile (2030)",      _v(afford.get("avg_hh_income_2030_3mi"),"$")),
        ("Monthly Affordability @ 3× (2030)",  _v(afford.get("monthly_affordability_2030_3x"),"$")),
        ("Rent Headroom (3-Mile, 2030)",        _v(afford.get("rent_headroom_2030_3mi"),"$")),
        ("Income-to-Rent Ratio",               afford.get("income_to_rent_ratio")),
        ("Notes",                              afford.get("notes")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))

    # ── 8. Demographics ───────────────────────────────────────────────────────
    xl.section("8. DEMOGRAPHICS")
    xl.table_header(["Metric","3-Mile Radius","5-Mile Radius"])
    for i, row in enumerate([
        ("Population (2025)",        _v(demo.get("pop_3mi"),"n"),   _v(demo.get("pop_5mi"),"n")),
        ("Population (2030 proj.)",  _v(demo.get("pop_2030_3mi"),"n"), _v(demo.get("pop_2030_5mi"),"n")),
        ("Population Growth (5-yr)", demo.get("pop_growth_3mi") or "—", demo.get("pop_growth_5mi") or "—"),
        ("Median HH Income (2025)",  _v(demo.get("median_income_3mi"),"$"), _v(demo.get("median_income_5mi"),"$")),
        ("Median HH Income (2030)",  _v(demo.get("median_income_2030_3mi"),"$"), _v(demo.get("median_income_2030_5mi"),"$")),
        ("Income Growth (5-yr)",     demo.get("income_growth_3mi") or "—", demo.get("income_growth_5mi") or "—"),
        ("Renter-Occupied Units",    demo.get("renter_pct_3mi") or "—", demo.get("renter_pct_5mi") or "—"),
        ("Bachelor's Degree+",       demo.get("college_pct_3mi") or "—", demo.get("college_pct_5mi") or "—"),
        ("White-Collar Workers",     demo.get("white_collar_pct_3mi") or "—", demo.get("white_collar_pct_5mi") or "—"),
        ("Avg Home Value",           _v(demo.get("home_value"),"$"), "—"),
    ]):
        xl.table_row(list(row), alt=bool(i % 2))
    xl.blank()
    xl.subsection("Schools")
    for i, (k, v) in enumerate([
        ("School District", demo.get("school_district")),
        ("Elementary",      demo.get("elementary")),
        ("Middle School",   demo.get("middle")),
        ("High School",     demo.get("high_school")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))
    xl.blank()
    xl.subsection("Crime Rate")
    xl.kv("Crime", demo.get("crime") or "Not provided in OM — source independently.")
    employers = demo.get("employers") or []
    if employers:
        xl.blank()
        xl.subsection("Major Employers")
        xl.table_header(["Employer","Drive Time","Employees","Sector","Notes"])
        for i, e in enumerate(employers):
            xl.table_row([
                e.get("name","—"), e.get("drive") or "—", e.get("employees") or "—",
                e.get("sector") or "—", e.get("notes") or "—"
            ], alt=bool(i % 2))

    # ── 9. Unit Mix ───────────────────────────────────────────────────────────
    xl.section("9. UNIT MIX")
    if umix:
        xl.table_header(["Type","Plan","Units","Mix%","SF","Mkt Rent","Mkt PSF","Eff Rent","Eff PSF","Target Rent","Upside"])
        for i, u in enumerate(umix):
            xl.table_row([
                u.get("type","—"), u.get("plan") or "—", _v(u.get("count"),"n"),
                _v(u.get("pct"),"%"), _v(u.get("sf"),"n"),
                _v(u.get("market_rent"),"$"), f"${u.get('market_psf') or 0:.2f}",
                _v(u.get("eff_rent"),"$"), f"${u.get('eff_psf') or 0:.2f}",
                _v(u.get("target_rent"),"$"), _v(u.get("upside"),"$")
            ], alt=bool(i % 2))
    else:
        xl.kv("Note", "No unit mix data found.")

    # ── 10. Utilities ─────────────────────────────────────────────────────────
    xl.section("10. UTILITY INFORMATION")
    if utils_:
        xl.table_header(["Utility","Billing Method","Paid By","Reimbursement","Fee","Annual Income","Notes"])
        for i, u in enumerate(utils_):
            xl.table_row([
                u.get("name","—"), u.get("method") or "—", u.get("paid_by") or "—",
                u.get("reimbursement") or "N/A", u.get("fee") or "—",
                _v(u.get("annual_income"),"$"), u.get("notes") or "—"
            ], alt=bool(i % 2))
    else:
        xl.kv("Note", "No utility data found.")

    # ── 11. Site Info ─────────────────────────────────────────────────────────
    xl.section("11. SITE INFORMATION")
    for i, (k, v) in enumerate([
        ("Roof / Age",       f"{site.get('roof') or 'N/A'}  —  {site.get('roof_age') or 'N/A'}"),
        ("Exterior",         site.get("exterior")),
        ("Foundation",       site.get("foundation")),
        ("HVAC",             site.get("hvac")),
        ("Plumbing",         site.get("plumbing")),
        ("Wiring",           site.get("wiring")),
        ("Hot Water",        site.get("hot_water")),
        ("Washer / Dryer",   site.get("washer_dryer")),
        ("Life Safety",      site.get("life_safety") or "Verify — not specified"),
        ("Construction",     site.get("notes")),
    ]):
        xl.kv(k, v, alt=bool(i % 2))
    xl.blank()
    xl.subsection("Parking & Site Features")
    for i, (k, v) in enumerate([
        ("Open Spaces",    _v(site.get("parking_open"),"n")),
        ("Reserved",       f"{_v(site.get('parking_reserved'),'n')} ({site.get('reserved_fee') or 'N/A'}/mo)"),
        ("Covered",        _v(site.get("parking_covered"),"n")),
        ("Garage",         site.get("parking_garage") or "None"),
        ("Total / Ratio",  f"{_v(site.get('parking_total'),'n')} ({site.get('parking_ratio') or 'N/A'})"),
        ("Pet Yards",      site.get("pet_yards") or "N/A"),
        ("Storage",        site.get("storage") or "N/A"),
    ]):
        xl.kv(k, v, alt=bool(i % 2))

    # ── 12. Rent Comparables ──────────────────────────────────────────────────
    xl.section("12. RENT COMPARABLE SUMMARY")
    if rg:
        xl.subsection("Two Bedroom Garden Comparables")
        xl.table_header(["Property","Yr Built","Distance","Rent","vs. Subject","Notes"])
        for i, c in enumerate(rg):
            vs = f"{int((c.get('rent') or 0) - 1393):+,}" if c.get("rent") else "—"
            xl.table_row([
                c.get("name","—"), _v(c.get("year_built")), c.get("distance") or "—",
                _v(c.get("rent"),"$"), vs, c.get("notes") or "—"
            ], alt=bool(i % 2))
    if rth:
        xl.blank()
        xl.subsection("Two Bedroom Townhouse Comparables")
        xl.table_header(["Property","Yr Built","Distance","Rent","vs. Subject","Notes"])
        for i, c in enumerate(rth):
            vs = f"{int((c.get('rent') or 0) - 1557):+,}" if c.get("rent") else "—"
            xl.table_row([
                c.get("name","—"), _v(c.get("year_built")), c.get("distance") or "—",
                _v(c.get("rent"),"$"), vs, c.get("notes") or "—"
            ], alt=bool(i % 2))
    if rcomps:
        xl.blank()
        xl.subsection("Full Comparable Detail")
        xl.table_header(["#","Property","Type","Yr Built","Distance","Units","Occ%","Mkt Rent","Eff Rent"])
        for i, rc in enumerate(rcomps):
            xl.table_row([
                rc.get("id","—"), rc.get("name","—"), rc.get("comp_type") or "—",
                _v(rc.get("year_built")), rc.get("distance") or "—",
                _v(rc.get("units"),"n"), _v(rc.get("occupancy"),"%"),
                _v(rc.get("total_market"),"$"), _v(rc.get("total_eff"),"$")
            ], alt=bool(i % 2))
    if not rg and not rth and not rcomps:
        xl.kv("Note", "No rent comparable data found.")

    # ── 13. Financial Analysis ────────────────────────────────────────────────
    xl.section("13. FINANCIAL ANALYSIS — OPERATING STATEMENT")
    periods   = fin.get("periods") or []
    inc_lines = fin.get("income_lines") or []
    exp_lines = fin.get("expense_lines") or []

    if periods and (inc_lines or exp_lines):
        p_all = periods[:5]
        headers = ["Line Item"] + p_all + ["Underwriting Note"]
        xl.table_header(headers)

        def fin_row(item):
            vals = item.get("values") or {}
            row = [item.get("item","—")]
            for p in p_all:
                v = vals.get(p)
                row.append(_v(v,"$") if v is not None else "—")
            row.append(item.get("note") or "")
            return row

        for label, lines_list, is_section in [
            ("INCOME",   inc_lines, True),
            ("EXPENSES", exp_lines, True),
        ]:
            # Section label row — fill full row, title in col A
            xl._fill_row(C_GOLD_P)
            xl._set(f"A{xl.row}", label, bold=True, bg=C_GOLD_P, fg=C_DARK, size=9)
            xl.row += 1
            for i, item in enumerate(lines_list):
                is_t = item.get("is_total", False)
                bg = C_DARK2 if is_t else (C_GOLD_P if item.get("is_subtotal") else (C_ALT if i % 2 else C_WHITE))
                fg = C_GOLD_L if is_t else C_DARK
                row_vals = fin_row(item)
                for j, v in enumerate(row_vals):
                    col = get_column_letter(1 + j)
                    xl._set(f"{col}{xl.row}", v, bold=is_t, bg=bg, fg=fg, size=9,
                            align_h="right" if 0 < j < len(row_vals) - 1 else "left")
                xl.row += 1

        noi = fin.get("noi") or {}
        for label, source, is_tot in [
            ("NET OPERATING INCOME",      noi,                     True),
            ("  Capital Reserves",        fin.get("capex") or {}, False),
            ("CASH FLOW FROM OPERATIONS", fin.get("cffo") or {}, True),
        ]:
            bg = C_DARK2 if is_tot else C_WHITE
            fg = C_GOLD_L if is_tot else C_DARK
            row_vals = [label] + [_v(source.get(p),"$") for p in p_all] + [""]
            for j, v in enumerate(row_vals):
                col = get_column_letter(1 + j)
                xl._set(f"{col}{xl.row}", v, bold=is_tot, bg=bg, fg=fg, size=9,
                        align_h="right" if 0 < j < len(row_vals) - 1 else "left")
            xl.row += 1
    else:
        xl.kv("Note", "No financial analysis data found.")

    # ── 14. Sale Comparables ──────────────────────────────────────────────────
    xl.section("14. SALE COMPARABLES")
    if scomps:
        xl.table_header(["Property","Address","Date","Yr Built","Units","Sale Price","$/Unit","$/SF","Cap Rate","Occ","Buyer","Seller","Notes"])
        for i, sc in enumerate(scomps):
            xl.table_row([
                sc.get("name","—"),
                f"{sc.get('address') or ''} {sc.get('city_state') or ''}".strip() or "—",
                sc.get("date") or "—", _v(sc.get("year_built")), _v(sc.get("units"),"n"),
                _v(sc.get("price"),"$"), _v(sc.get("ppu"),"$"), _v(sc.get("ppsf"),"$"),
                sc.get("cap_rate") or "—", sc.get("occupancy") or "—",
                sc.get("buyer") or "—", sc.get("seller") or "—", sc.get("notes") or "—"
            ], alt=bool(i % 2))
    else:
        xl.kv("Note", "Sale comps not in OM. Source from CoStar, RCA, or listing broker.")

    # ── 15. Underwriting Flags ────────────────────────────────────────────────
    xl.section("15. UNDERWRITING FLAGS & NOTES")
    flag_bg = {"Warning": C_WARN, "Opportunity": C_GREEN, "Info": C_BLUE, "Verify": C_PURPLE}
    if flags:
        xl.table_header(["Category","Title","Detail"])
        for f in flags:
            cat = f.get("category","Info")
            bg  = flag_bg.get(cat, C_BLUE)
            detail = f.get("detail","")
            for j, v in enumerate([cat, f.get("title",""), detail]):
                col = get_column_letter(1 + j)
                xl._set(f"{col}{xl.row}", v, bg=bg, size=9, bold=(j == 0), wrap=True)
            # Auto height based on detail text in col C (~50 chars wide)
            h = max(20, xl._auto_height(detail, col_width=50))
            xl.ws.row_dimensions[xl.row].height = h
            xl.row += 1
    else:
        xl.kv("Note", "No flags generated.")

    # ── 16. Market Overview ───────────────────────────────────────────────────
    xl.section("16. MARKET & SUBMARKET OVERVIEW")
    if mkt.get("market_summary"):
        xl.subsection("Market Summary")
        xl.kv("Summary", mkt["market_summary"])
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
        xl.blank()
        xl.subsection("Submarket Metrics")
        for i, (k, v) in enumerate(mkt_rows):
            xl.kv(k, v, alt=bool(i % 2))
    devs = mkt.get("major_developments") or []
    if devs and any(d.get("name") for d in devs):
        xl.blank()
        xl.subsection("Major Economic Developments")
        xl.table_header(["Development","Description","Est. Cost","Jobs","Timeline"])
        for i, d in enumerate(devs):
            if d.get("name"):
                xl.table_row([
                    d.get("name","—"), d.get("description") or "—",
                    d.get("cost") or "—", d.get("jobs") or "—",
                    d.get("timeline") or "—"
                ], alt=bool(i % 2))

    # ── Disclaimer ────────────────────────────────────────────────────────────
    xl.blank()
    xl._fill_row(C_ALT)
    xl._set(f"A{xl.row}",
        "DISCLAIMER: AI-generated from the uploaded OM. For internal underwriting reference only. "
        "Verify all figures independently. Powered by Anthropic Claude.",
        bg=C_ALT, fg=C_MID, size=8, italic=True, border=False, wrap=True)
    xl.ws.row_dimensions[xl.row].height = 20

    buf = io.BytesIO()
    wb.save(buf)
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
Upload any Multifamily Offering Memorandum PDF and get a structured 16-sheet Excel underwriting report.

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
        st.info("**Step 3** — Download the underwriting Excel")
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

        set_progress(75, "Generating Excel report…")

        # Step 3 — Build Excel
        excel_bytes = build_excel(data, uploaded.name)

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
        label="⬇️  Download Underwriting Report (Excel)",
        data=excel_bytes,
        file_name=f"{safe}_Underwriting_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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

