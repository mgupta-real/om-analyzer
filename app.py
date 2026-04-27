"""
Multifamily OM Analyzer — Streamlit App (v2)
Upload any Offering Memorandum PDF → get a structured underwriting report as PDF.

Gap fixes vs v1 (from Forma at the Park analysis):
  1. Unit mix now captures three-tier rent (in-place / market / comp-supported)
  2. Rent comps: extract per-floorplan detail tables for each comp, not just summary
  3. Loan assumption / assumable debt terms extracted as a dedicated sub-section
  4. Sale comps absence handled gracefully (no error, clean "Not provided" output)
  5. Multiple NOI snapshots captured (T-12, 6-mo ann., 90-day ann., 30-day ann.)
  6. Value-add levers extracted as structured table (unit count + premium + annual upside)
  7. Underwriting flags enriched with capex age, economic vs physical occupancy gap,
     assumable debt flag, and chiller/major-system risk items
"""

import os
import io
import base64
import tempfile
import traceback

import streamlit as st

# ─── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Multifamily OM Analyzer",
    page_icon="🏢",
    layout="centered",
    initial_sidebar_state="expanded",
)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
#MainMenu, footer, header {visibility: hidden;}
.stApp { background: #F5F4EF; }
.block-container { max-width: 820px !important; padding-top: 2rem !important; }
.om-header {
    background: #1A1A18; border-radius: 12px;
    padding: 28px 32px; margin-bottom: 20px;
}
.om-header h1 { color: #D4B07A; font-size: 26px; font-weight: 500; margin: 0 0 4px 0; }
.om-header p  { color: #8C8C7A; font-size: 13px; margin: 0; }
.stButton>button {
    background: #1A1A18 !important; color: #D4B07A !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 600 !important; padding: 10px 24px !important;
    width: 100% !important;
}
</style>
""", unsafe_allow_html=True)

# ─── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="om-header">
  <h1>🏢 Multifamily OM Analyzer</h1>
  <p>Upload any broker Offering Memorandum PDF → AI extracts all underwriting data → download a structured report</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — PDF TEXT EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """Extract all text from a PDF, trying pdfplumber then pypdf as fallback."""
    text = ""
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            pages = []
            for i, page in enumerate(pdf.pages):
                page_text = page.extract_text() or ""
                # Also try to pull table text that might not appear in extract_text
                tables = page.extract_tables() or []
                table_text = ""
                for tbl in tables:
                    for row in tbl:
                        row_str = " | ".join(str(c) if c else "" for c in row)
                        table_text += row_str + "\n"
                pages.append(f"--- PAGE {i+1} ---\n{page_text}\n{table_text}")
            text = "\n".join(pages)
    except Exception:
        pass

    if len(text.strip()) < 200:
        try:
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(pdf_bytes))
            pages = []
            for i, page in enumerate(reader.pages):
                pages.append(f"--- PAGE {i+1} ---\n{page.extract_text() or ''}")
            text = "\n".join(pages)
        except Exception:
            pass

    return text


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — CLAUDE EXTRACTION PROMPT & SCHEMA
# ══════════════════════════════════════════════════════════════════════════════

EXTRACTION_PROMPT = """
You are an expert multifamily real estate underwriter. Analyze this Offering Memorandum and extract every data point available. Return ONLY a valid JSON object — no markdown, no preamble, no trailing text.

IMPORTANT RULES:
- Use null for any field not found in the document. Never invent numbers.
- For dollar amounts, return numbers only (no $ signs, no commas).
- For percentages, return as decimals (e.g. 0.065 for 6.5%) unless noted.
- For "not provided" sections (e.g. sale comps absent), return an empty array [], not null.
- Capture ALL rent tiers shown in unit mix tables (in-place, market, comp-supported).
- For each rent comp, capture the per-floorplan detail rows if shown, not just the summary.
- Capture ALL NOI/revenue snapshots (trailing periods + proforma).
- Capture ALL value-add levers with their individual unit counts and premiums.

Return this exact JSON schema:

{
  "property": {
    "name": null,
    "address": null,
    "city": null,
    "state": null,
    "zip": null,
    "year_built": null,
    "num_units": null,
    "num_buildings": null,
    "site_acres": null,
    "density_units_per_acre": null,
    "avg_unit_sf": null,
    "total_rentable_sf": null,
    "building_type": null,
    "foundation": null,
    "framing": null,
    "exterior": null,
    "roof": null,
    "ceiling_height": null,
    "countertops": null,
    "flooring": null,
    "appliances": null,
    "washer_dryer_units": null,
    "parking_spaces": null,
    "parking_ratio": null,
    "heating_cooling": null,
    "electric": null,
    "wiring": null,
    "water_heaters": null,
    "plumbing": null,
    "school_district": null,
    "schools": [],
    "tax_jurisdiction": null,
    "tax_id": null,
    "broker": null,
    "listing_broker_contacts": []
  },

  "offering": {
    "price": null,
    "price_per_unit": null,
    "price_per_sf": null,
    "cap_rate": null,
    "grm": null,
    "terms": null
  },

  "loan_assumption": {
    "available": false,
    "lender": null,
    "original_balance": null,
    "current_balance": null,
    "note_rate": null,
    "rate_type": null,
    "origination_date": null,
    "maturity_date": null,
    "term_months": null,
    "io_periods_months": null,
    "notes": null
  },

  "investment_highlights": {
    "summary": null,
    "key_bullets": []
  },

  "renovation": {
    "phases": [
      {
        "tier_name": null,
        "units_completed": null,
        "description": null,
        "features": [],
        "monthly_premium_achieved": null
      }
    ],
    "remaining_upside_units": null
  },

  "value_add_levers": [
    {
      "lever": null,
      "units": null,
      "monthly_premium": null,
      "annual_upside": null,
      "notes": null
    }
  ],

  "demographics": {
    "workforce_within_5mi": null,
    "businesses_within_5mi": null,
    "median_hh_income_5mi": null,
    "median_age_3mi": null,
    "population_1mi": null,
    "population_3mi": null,
    "population_5mi": null,
    "traffic_count_vpd": null,
    "traffic_road": null,
    "five_yr_avg_rent_growth": null,
    "five_yr_avg_occupancy": null,
    "nearby_employers": [],
    "nearby_amenities": []
  },

  "unit_mix": [
    {
      "type_code": null,
      "description": null,
      "num_units": null,
      "unit_sf": null,
      "total_sf": null,
      "rent_inplace": null,
      "rent_market": null,
      "rent_comp_supported": null,
      "rent_psf_inplace": null,
      "rent_psf_market": null,
      "rent_psf_comp_supported": null,
      "gpr_inplace": null,
      "gpr_market": null,
      "gpr_comp_supported": null
    }
  ],

  "utilities": {
    "electricity_billing": null,
    "water_billing": null,
    "gas_billing": null,
    "trash_billing": null,
    "trash_flat_fee": null,
    "pest_billing": null,
    "pest_flat_fee": null,
    "cable_internet": null,
    "notes": null
  },

  "amenities": {
    "unit_features": [],
    "community_features": [],
    "pet_policy": null,
    "pet_rent": null,
    "pet_deposit": null
  },

  "on_site_staff": {
    "total_employees": null,
    "roles": []
  },

  "rent_comps": {
    "summary": [
      {
        "map_num": null,
        "name": null,
        "address": null,
        "year_built": null,
        "num_units": null,
        "avg_unit_sf": null,
        "occupancy": null,
        "avg_rent_per_unit": null,
        "avg_rent_psf": null,
        "interior_finishes": []
      }
    ],
    "subject_vs_comp_avg": {
      "subject_avg_rent": null,
      "comp_avg_rent": null,
      "discount_per_unit": null,
      "discount_pct": null
    },
    "detail_by_comp": [
      {
        "comp_name": null,
        "floorplans": [
          {
            "description": null,
            "num_units": null,
            "unit_sf": null,
            "rent_per_unit": null,
            "rent_psf": null
          }
        ]
      }
    ]
  },

  "financial": {
    "noi_snapshots": [
      {
        "period": null,
        "gross_potential_rent": null,
        "loss_to_lease": null,
        "vacancy": null,
        "concessions": null,
        "other_rent_loss": null,
        "net_rental_income": null,
        "utility_reimbursement": null,
        "other_income": null,
        "gross_revenues": null,
        "total_expenses": null,
        "noi": null,
        "physical_occupancy": null,
        "economic_occupancy": null
      }
    ],
    "proforma_years": [
      {
        "year": null,
        "gpr": null,
        "total_economic_loss_pct": null,
        "net_rental_income": null,
        "gross_revenues": null,
        "total_expenses": null,
        "noi": null,
        "expense_per_unit": null
      }
    ],
    "key_expense_assumptions": {
      "management_fee_pct": null,
      "insurance_per_unit": null,
      "real_estate_tax_rate": null,
      "capex_reserve_per_unit": null,
      "utilities_per_unit": null
    },
    "historical_capex_total": null,
    "historical_capex_breakdown": []
  },

  "sale_comps": [
    {
      "name": null,
      "address": null,
      "sale_date": null,
      "num_units": null,
      "year_built": null,
      "sale_price": null,
      "price_per_unit": null,
      "cap_rate": null,
      "occupancy_at_sale": null,
      "notes": null
    }
  ],

  "market_overview": {
    "metro": null,
    "metro_population": null,
    "metro_job_growth_pct": null,
    "metro_unemployment": null,
    "submarket": null,
    "sub_occupancy": null,
    "sub_rent": null,
    "sub_growth": null,
    "pipeline_units": null,
    "absorption_units": null,
    "notes": null
  },

  "underwriting_flags": [
    {
      "category": null,
      "title": null,
      "detail": null
    }
  ]
}

Here is the full text of the Offering Memorandum:

{OM_TEXT}
"""


def _call_claude(prompt: str, api_key: str, model: str, max_tokens: int) -> str:
    """Raw Claude API call — returns response text and stop_reason."""
    import requests as req
    resp = req.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": model,
            "max_tokens": max_tokens,
            "messages": [{"role": "user", "content": prompt}],
        },
        timeout=240,
    )
    resp.raise_for_status()
    body = resp.json()
    text = body["content"][0]["text"].strip()
    stop_reason = body.get("stop_reason", "")
    return text, stop_reason


def _clean_json(raw: str) -> str:
    """Strip markdown fences and return clean JSON string."""
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else raw
        if raw.startswith("json"):
            raw = raw[4:]
    if raw.endswith("```"):
        raw = raw[:-3]
    return raw.strip()


# Slimmer fallback prompt used when the full schema overflows max_tokens
SLIM_PROMPT = """
You are a multifamily underwriter. Extract key data from this Offering Memorandum.
Return ONLY valid JSON. No markdown, no preamble. Use null for missing fields.
Dollar amounts as numbers only. Percentages as decimals (0.065 = 6.5%).
sale_comps should be [] if not present (never null).

{
  "property": {"name":null,"address":null,"city":null,"state":null,"zip":null,
    "year_built":null,"num_units":null,"num_buildings":null,"site_acres":null,
    "avg_unit_sf":null,"total_rentable_sf":null,"building_type":null,
    "foundation":null,"framing":null,"exterior":null,"roof":null,
    "heating_cooling":null,"electric":null,"parking_spaces":null,
    "parking_ratio":null,"school_district":null,"schools":[],"broker":null},
  "offering": {"price":null,"price_per_unit":null,"cap_rate":null,"terms":null},
  "loan_assumption": {"available":false,"lender":null,"current_balance":null,
    "note_rate":null,"rate_type":null,"maturity_date":null,"io_periods_months":null},
  "investment_highlights": {"summary":null,"key_bullets":[]},
  "renovation": {"phases":[],"remaining_upside_units":null},
  "value_add_levers": [],
  "demographics": {"workforce_within_5mi":null,"businesses_within_5mi":null,
    "median_hh_income_5mi":null,"median_age_3mi":null,"traffic_count_vpd":null,
    "traffic_road":null,"five_yr_avg_rent_growth":null,"five_yr_avg_occupancy":null,
    "nearby_employers":[]},
  "unit_mix": [{"type_code":null,"description":null,"num_units":null,"unit_sf":null,
    "rent_inplace":null,"rent_market":null,"rent_comp_supported":null,
    "gpr_inplace":null}],
  "utilities": {"electricity_billing":null,"water_billing":null,"gas_billing":null,
    "trash_billing":null,"trash_flat_fee":null,"pest_flat_fee":null,"cable_internet":null},
  "amenities": {"unit_features":[],"community_features":[],"pet_policy":null,
    "pet_rent":null,"pet_deposit":null},
  "on_site_staff": {"total_employees":null,"roles":[]},
  "rent_comps": {
    "summary": [{"map_num":null,"name":null,"year_built":null,"num_units":null,
      "avg_unit_sf":null,"occupancy":null,"avg_rent_per_unit":null,"avg_rent_psf":null}],
    "subject_vs_comp_avg": {"subject_avg_rent":null,"comp_avg_rent":null,
      "discount_per_unit":null},
    "detail_by_comp": []},
  "financial": {
    "noi_snapshots": [{"period":null,"gross_potential_rent":null,"net_rental_income":null,
      "gross_revenues":null,"total_expenses":null,"noi":null,
      "physical_occupancy":null,"economic_occupancy":null}],
    "proforma_years": [{"year":null,"gpr":null,"gross_revenues":null,
      "total_expenses":null,"noi":null}],
    "key_expense_assumptions": {"management_fee_pct":null,"insurance_per_unit":null,
      "real_estate_tax_rate":null,"capex_reserve_per_unit":null},
    "historical_capex_total":null,"historical_capex_breakdown":[]},
  "sale_comps": [],
  "market_overview": {"metro":null,"metro_population":null,"metro_job_growth_pct":null,
    "sub_occupancy":null,"sub_rent":null,"pipeline_units":null},
  "underwriting_flags": [{"category":null,"title":null,"detail":null}]
}

Offering Memorandum text:
{OM_TEXT}
"""


def analyze_om_with_claude(om_text: str, api_key: str, model: str = "claude-sonnet-4-20250514") -> dict:
    """Send OM text to Claude and return parsed JSON.

    Strategy:
      1. Try full schema with max_tokens=16000.
      2. If stop_reason=='max_tokens' (truncated output) or JSON parse fails,
         retry with slim schema at max_tokens=8000.
      3. If slim also fails, attempt to salvage partial JSON via a repair call.
    """
    import json

    # Truncate input to ~160k chars — leaves plenty of room for output tokens
    truncated = om_text[:160000]

    # ── Attempt 1: full schema ──────────────────────────────────────────────
    prompt_full = EXTRACTION_PROMPT.replace("{OM_TEXT}", truncated)
    raw, stop_reason = _call_claude(prompt_full, api_key, model, max_tokens=32000)

    if stop_reason != "max_tokens":
        try:
            return json.loads(_clean_json(raw))
        except json.JSONDecodeError:
            pass  # fall through to retry

    # ── Attempt 2: slim schema (shorter output) ────────────────────────────
    prompt_slim = SLIM_PROMPT.replace("{OM_TEXT}", truncated)
    raw2, stop_reason2 = _call_claude(prompt_slim, api_key, model, max_tokens=8000)

    if stop_reason2 != "max_tokens":
        try:
            return json.loads(_clean_json(raw2))
        except json.JSONDecodeError:
            pass

    # ── Attempt 3: repair the truncated JSON ──────────────────────────────
    # Ask Claude to complete/fix the broken JSON
    broken = raw2 if stop_reason2 == "max_tokens" else raw
    repair_prompt = (
        "The following JSON is incomplete or malformed because it was cut off. "
        "Please complete and fix it so it is valid JSON. "
        "Return ONLY the corrected JSON with no markdown or explanation.\n\n"
        + broken[:12000]
    )
    raw3, _ = _call_claude(repair_prompt, api_key, model, max_tokens=6000)
    return json.loads(_clean_json(raw3))


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — PDF REPORT GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

def generate_underwriting_pdf(data: dict, filename: str = "om_report.pdf") -> bytes:
    """Build a multi-section underwriting PDF from extracted data dict."""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        HRFlowable, KeepTogether
    )
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    from datetime import date

    # ── Color palette ──────────────────────────────────────────────────────
    DARK   = colors.HexColor("#1A1A18")
    GOLD   = colors.HexColor("#D4B07A")
    MID    = colors.HexColor("#6B6B5A")
    LIGHT  = colors.HexColor("#F5F4EF")
    BORDER = colors.HexColor("#DDDDD0")
    RED    = colors.HexColor("#C0392B")
    GREEN  = colors.HexColor("#27AE60")
    BLUE   = colors.HexColor("#2471A3")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        leftMargin=0.75*inch, rightMargin=0.75*inch,
        topMargin=0.9*inch, bottomMargin=0.75*inch,
    )

    styles = getSampleStyleSheet()

    def _s(name, **kw):
        base = styles["Normal"] if name not in styles else styles[name]
        return ParagraphStyle(name + "_custom_" + str(id(kw)), parent=base, **kw)

    S_H1    = _s("h1",   fontSize=18, textColor=DARK,  spaceAfter=4,  spaceBefore=10, fontName="Helvetica-Bold")
    S_H2    = _s("h2",   fontSize=12, textColor=DARK,  spaceAfter=3,  spaceBefore=8,  fontName="Helvetica-Bold")
    S_H3    = _s("h3",   fontSize=10, textColor=MID,   spaceAfter=2,  spaceBefore=5,  fontName="Helvetica-Bold")
    S_BODY  = _s("body", fontSize=9,  textColor=DARK,  spaceAfter=2,  leading=14)
    S_SMALL = _s("sm",   fontSize=7.5,textColor=MID,   spaceAfter=2,  leading=11)
    S_GOLD  = _s("gold", fontSize=9,  textColor=GOLD,  fontName="Helvetica-Bold")
    S_FLAG_R= _s("flr",  fontSize=8.5,textColor=RED,   fontName="Helvetica-Bold")
    S_FLAG_G= _s("flg",  fontSize=8.5,textColor=GREEN, fontName="Helvetica-Bold")
    S_FLAG_B= _s("flb",  fontSize=8.5,textColor=BLUE,  fontName="Helvetica-Bold")

    def _v(val, prefix="", suffix="", decimals=0, pct=False):
        if val is None or val == "":
            return "—"
        if pct:
            try:
                v = float(val)
                return f"{v*100:.1f}%"
            except Exception:
                return str(val)
        try:
            v = float(val)
            if decimals == 0:
                return f"{prefix}{v:,.0f}{suffix}"
            return f"{prefix}{v:,.{decimals}f}{suffix}"
        except Exception:
            return str(val)

    def _hr():
        return HRFlowable(width="100%", thickness=0.5, color=BORDER, spaceAfter=6)

    def _sec(title):
        return [
            Spacer(1, 10),
            _hr(),
            Paragraph(title.upper(), S_H2),
            Spacer(1, 4),
        ]

    def _kv_table(rows, col_widths=None):
        """Two-column key/value table."""
        if not rows:
            return []
        if col_widths is None:
            col_widths = [2.5*inch, 4.25*inch]
        data = [[Paragraph(str(k), S_H3), Paragraph(str(v), S_BODY)] for k, v in rows if v not in (None, "", "—")]
        if not data:
            return []
        tbl = Table(data, colWidths=col_widths, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("ROWBACKGROUNDS", (0,0), (-1,-1), [colors.white, LIGHT]),
            ("GRID",           (0,0), (-1,-1), 0.3, BORDER),
            ("VALIGN",         (0,0), (-1,-1), "TOP"),
            ("TOPPADDING",     (0,0), (-1,-1), 4),
            ("BOTTOMPADDING",  (0,0), (-1,-1), 4),
            ("LEFTPADDING",    (0,0), (-1,-1), 6),
        ]))
        return [tbl, Spacer(1, 6)]

    def _data_table(headers, rows, col_widths=None):
        """Multi-column data table with header row."""
        if not rows:
            return []
        all_rows = [[Paragraph(str(h), S_GOLD) for h in headers]] + \
                   [[Paragraph(str(c) if c is not None else "—", S_BODY) for c in row] for row in rows]
        if col_widths is None:
            avail = 7.0 * inch
            col_widths = [avail / len(headers)] * len(headers)
        tbl = Table(all_rows, colWidths=col_widths, hAlign="LEFT", repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",     (0,0), (-1,0), DARK),
            ("TEXTCOLOR",      (0,0), (-1,0), GOLD),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, LIGHT]),
            ("GRID",           (0,0), (-1,-1), 0.3, BORDER),
            ("VALIGN",         (0,0), (-1,-1), "TOP"),
            ("TOPPADDING",     (0,0), (-1,-1), 3),
            ("BOTTOMPADDING",  (0,0), (-1,-1), 3),
            ("LEFTPADDING",    (0,0), (-1,-1), 5),
        ]))
        return [tbl, Spacer(1, 8)]

    # ── Footer ─────────────────────────────────────────────────────────────
    prop = data.get("property", {})
    prop_name = prop.get("name") or "Unnamed Property"

    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(MID)
        canvas.drawString(0.75*inch, 0.45*inch, f"{prop_name}  |  AI-generated underwriting report  |  {date.today().isoformat()}")
        canvas.drawRightString(letter[0]-0.75*inch, 0.45*inch, f"Page {doc.page}")
        canvas.restoreState()

    # ══════════════════════════════════════════════════════════════════════
    story = []

    # ── Cover / Title ──────────────────────────────────────────────────────
    story.append(Spacer(1, 20))
    story.append(Paragraph(prop_name, S_H1))
    addr = ", ".join(filter(None, [prop.get("address"), prop.get("city"), prop.get("state"), prop.get("zip")]))
    if addr:
        story.append(Paragraph(addr, S_BODY))
    story.append(Paragraph(f"Underwriting Report  ·  Generated {date.today().strftime('%B %d, %Y')}", S_SMALL))
    story.append(Spacer(1, 6))
    story.append(_hr())

    # ── 1. Property Details ────────────────────────────────────────────────
    story += _sec("1. Property Details")
    story += _kv_table([
        ("Year Built",          prop.get("year_built")),
        ("# Units",             prop.get("num_units")),
        ("# Buildings",         prop.get("num_buildings")),
        ("Site Size",           _v(prop.get("site_acres"), suffix=" acres")),
        ("Density",             _v(prop.get("density_units_per_acre"), suffix=" units/acre", decimals=2)),
        ("Avg Unit Size",       _v(prop.get("avg_unit_sf"), suffix=" SF")),
        ("Total Rentable SF",   _v(prop.get("total_rentable_sf"), suffix=" SF")),
        ("Building Type",       prop.get("building_type")),
        ("Foundation",          prop.get("foundation")),
        ("Framing",             prop.get("framing")),
        ("Exterior",            prop.get("exterior")),
        ("Roof",                prop.get("roof")),
        ("Ceiling Height",      prop.get("ceiling_height")),
        ("Flooring",            prop.get("flooring")),
        ("Countertops",         prop.get("countertops")),
        ("Appliances",          prop.get("appliances")),
        ("W/D Units",           prop.get("washer_dryer_units")),
        ("Parking Spaces",      prop.get("parking_spaces")),
        ("Parking Ratio",       _v(prop.get("parking_ratio"), suffix="/unit", decimals=2)),
        ("Heating/Cooling",     prop.get("heating_cooling")),
        ("Electric",            prop.get("electric")),
        ("Wiring",              prop.get("wiring")),
        ("Water Heaters",       prop.get("water_heaters")),
        ("Plumbing",            prop.get("plumbing")),
        ("School District",     prop.get("school_district")),
        ("Schools",             ", ".join(prop.get("schools") or [])),
        ("Tax Jurisdiction",    prop.get("tax_jurisdiction")),
        ("Tax ID",              prop.get("tax_id")),
        ("Broker",              prop.get("broker")),
    ])

    # ── 2. Offering ────────────────────────────────────────────────────────
    off = data.get("offering", {})
    story += _sec("2. Offering")
    story += _kv_table([
        ("Asking Price",        _v(off.get("price"), "$")),
        ("Price / Unit",        _v(off.get("price_per_unit"), "$")),
        ("Price / SF",          _v(off.get("price_per_sf"), "$", decimals=2)),
        ("Cap Rate",            _v(off.get("cap_rate"), pct=True)),
        ("GRM",                 _v(off.get("grm"), decimals=2)),
        ("Terms",               off.get("terms")),
    ])

    # ── 3. Loan Assumption ─────────────────────────────────────────────────
    loan = data.get("loan_assumption", {})
    story += _sec("3. Loan Assumption")
    if loan.get("available"):
        story += _kv_table([
            ("Lender",              loan.get("lender")),
            ("Original Balance",    _v(loan.get("original_balance"), "$")),
            ("Current Balance",     _v(loan.get("current_balance"), "$")),
            ("Note Rate",           _v(loan.get("note_rate"), pct=True)),
            ("Rate Type",           loan.get("rate_type")),
            ("Origination Date",    loan.get("origination_date")),
            ("Maturity Date",       loan.get("maturity_date")),
            ("Term (Months)",       loan.get("term_months")),
            ("IO Period (Months)",  loan.get("io_periods_months")),
            ("Notes",               loan.get("notes")),
        ])
    else:
        story.append(Paragraph("No assumable debt identified in this OM.", S_BODY))

    # ── 4. Investment Highlights ───────────────────────────────────────────
    inv = data.get("investment_highlights", {})
    story += _sec("4. Investment Highlights")
    if inv.get("summary"):
        story.append(Paragraph(inv["summary"], S_BODY))
        story.append(Spacer(1, 4))
    for b in (inv.get("key_bullets") or []):
        story.append(Paragraph(f"• {b}", S_BODY))

    # ── 5. Renovation & Value-Add ──────────────────────────────────────────
    story += _sec("5. Renovation & Value-Add")

    ren = data.get("renovation", {})
    for phase in (ren.get("phases") or []):
        if phase.get("tier_name"):
            story.append(Paragraph(phase["tier_name"], S_H3))
        rows = []
        if phase.get("units_completed"):
            rows.append(("Units Completed", str(phase["units_completed"])))
        if phase.get("monthly_premium_achieved"):
            rows.append(("Monthly Premium", _v(phase["monthly_premium_achieved"], "$")))
        if phase.get("description"):
            rows.append(("Description", phase["description"]))
        story += _kv_table(rows)
        feats = phase.get("features") or []
        if feats:
            for f in feats:
                story.append(Paragraph(f"  · {f}", S_BODY))
        story.append(Spacer(1, 4))

    # Value-add levers table
    levers = data.get("value_add_levers") or []
    if levers:
        story.append(Paragraph("Value-Add Revenue Levers", S_H3))
        lever_rows = []
        total_annual = 0
        for lv in levers:
            ann = lv.get("annual_upside")
            try:
                total_annual += float(ann or 0)
            except Exception:
                pass
            lever_rows.append([
                lv.get("lever") or "—",
                _v(lv.get("units")),
                _v(lv.get("monthly_premium"), "$"),
                _v(ann, "$"),
            ])
        lever_rows.append(["TOTAL ANNUAL UPSIDE", "", "", _v(total_annual, "$")])
        story += _data_table(
            ["Lever", "Units", "Mo. Premium", "Annual Upside"],
            lever_rows,
            col_widths=[3.0*inch, 0.8*inch, 1.1*inch, 1.5*inch],
        )

    # ── 6. Demographics ────────────────────────────────────────────────────
    dem = data.get("demographics", {})
    story += _sec("6. Demographics & Location")
    story += _kv_table([
        ("Workforce (5 mi)",        _v(dem.get("workforce_within_5mi"))),
        ("Businesses (5 mi)",       _v(dem.get("businesses_within_5mi"))),
        ("Median HH Income (5 mi)", _v(dem.get("median_hh_income_5mi"), "$")),
        ("Median Age (3 mi)",       dem.get("median_age_3mi")),
        ("Population (1 mi)",       _v(dem.get("population_1mi"))),
        ("Population (3 mi)",       _v(dem.get("population_3mi"))),
        ("Population (5 mi)",       _v(dem.get("population_5mi"))),
        ("Traffic Count",           _v(dem.get("traffic_count_vpd"), suffix=" VPD") + (f" ({dem['traffic_road']})" if dem.get("traffic_road") else "")),
        ("5-Yr Avg Rent Growth",    _v(dem.get("five_yr_avg_rent_growth"), pct=True)),
        ("5-Yr Avg Occupancy",      _v(dem.get("five_yr_avg_occupancy"), pct=True)),
        ("Major Employers",         ", ".join(dem.get("nearby_employers") or [])),
    ])

    # ── 7. Unit Mix ────────────────────────────────────────────────────────
    story += _sec("7. Unit Mix")
    mix = data.get("unit_mix") or []
    if mix:
        # Detect if comp-supported rent is available
        has_comp = any(u.get("rent_comp_supported") for u in mix)
        headers = ["Type", "Description", "Units", "SF", "In-Place Rent", "Market Rent"]
        widths  = [0.55*inch, 1.1*inch, 0.5*inch, 0.55*inch, 0.95*inch, 0.95*inch]
        if has_comp:
            headers.append("Comp-Supp. Rent")
            widths.append(1.05*inch)
        headers += ["GPR In-Place"]
        widths  += [1.0*inch]

        mix_rows = []
        for u in mix:
            row = [
                u.get("type_code") or "—",
                u.get("description") or "—",
                _v(u.get("num_units")),
                _v(u.get("unit_sf")),
                _v(u.get("rent_inplace"), "$"),
                _v(u.get("rent_market"), "$"),
            ]
            if has_comp:
                row.append(_v(u.get("rent_comp_supported"), "$"))
            row.append(_v(u.get("gpr_inplace"), "$"))
            mix_rows.append(row)

        story += _data_table(headers, mix_rows, col_widths=widths)

        # Totals/averages row summary
        total_units = sum(float(u.get("num_units") or 0) for u in mix)
        total_gpr   = sum(float(u.get("gpr_inplace") or 0) for u in mix)
        avg_sf_vals = [u.get("unit_sf") for u in mix if u.get("unit_sf")]
        avg_sf      = sum(float(x) for x in avg_sf_vals) / len(avg_sf_vals) if avg_sf_vals else None
        story.append(Paragraph(
            f"Total Units: {_v(total_units)}  ·  Avg SF: {_v(avg_sf, decimals=0)}  ·  Total GPR (In-Place): {_v(total_gpr, '$')}",
            S_SMALL
        ))

    # ── 8. Utilities & Billing ─────────────────────────────────────────────
    util = data.get("utilities", {})
    story += _sec("8. Utilities & Billing")
    story += _kv_table([
        ("Electricity",     util.get("electricity_billing")),
        ("Water/Sewer",     util.get("water_billing")),
        ("Gas",             util.get("gas_billing")),
        ("Trash",           (util.get("trash_billing") or "") + (f" (${util['trash_flat_fee']}/mo)" if util.get("trash_flat_fee") else "")),
        ("Pest",            (util.get("pest_billing") or "") + (f" (${util['pest_flat_fee']}/mo)" if util.get("pest_flat_fee") else "")),
        ("Cable/Internet",  util.get("cable_internet")),
        ("Notes",           util.get("notes")),
    ])

    # ── 9. Amenities ──────────────────────────────────────────────────────
    amen = data.get("amenities", {})
    story += _sec("9. Amenities")
    unit_feats = amen.get("unit_features") or []
    comm_feats = amen.get("community_features") or []
    if unit_feats:
        story.append(Paragraph("Unit Features", S_H3))
        story.append(Paragraph("  ·  ".join(unit_feats), S_BODY))
    if comm_feats:
        story.append(Paragraph("Community Features", S_H3))
        story.append(Paragraph("  ·  ".join(comm_feats), S_BODY))
    story += _kv_table([
        ("Pet Policy",   amen.get("pet_policy")),
        ("Pet Rent",     _v(amen.get("pet_rent"), "$", "/mo")),
        ("Pet Deposit",  amen.get("pet_deposit")),
    ])

    # ── 10. Rent Comparables ───────────────────────────────────────────────
    story += _sec("10. Rent Comparables")
    rc = data.get("rent_comps", {})
    svsc = rc.get("subject_vs_comp_avg", {})
    if svsc.get("subject_avg_rent") or svsc.get("comp_avg_rent"):
        story += _kv_table([
            ("Subject Avg Rent",       _v(svsc.get("subject_avg_rent"), "$")),
            ("Comp Set Avg Rent",      _v(svsc.get("comp_avg_rent"), "$")),
            ("Discount per Unit",      _v(svsc.get("discount_per_unit"), "$")),
            ("Discount %",             _v(svsc.get("discount_pct"), pct=True)),
        ])

    # Summary comp table
    comps = rc.get("summary") or []
    if comps:
        story.append(Paragraph("Comp Summary", S_H3))
        comp_rows = []
        for c in comps:
            comp_rows.append([
                c.get("name") or "—",
                c.get("year_built") or "—",
                _v(c.get("num_units")),
                _v(c.get("avg_unit_sf")),
                _v(c.get("occupancy"), pct=True),
                _v(c.get("avg_rent_per_unit"), "$"),
                _v(c.get("avg_rent_psf"), "$", decimals=2),
            ])
        story += _data_table(
            ["Property", "Built", "Units", "Avg SF", "Occ", "Avg Rent", "Rent/SF"],
            comp_rows,
            col_widths=[2.2*inch, 0.6*inch, 0.6*inch, 0.7*inch, 0.6*inch, 0.85*inch, 0.7*inch],
        )

    # Per-comp floorplan detail
    detail_comps = rc.get("detail_by_comp") or []
    if detail_comps:
        story.append(Paragraph("Comp Floorplan Detail", S_H3))
        for dc in detail_comps:
            if dc.get("comp_name"):
                story.append(Paragraph(dc["comp_name"], S_H3))
            fp_rows = []
            for fp in (dc.get("floorplans") or []):
                fp_rows.append([
                    fp.get("description") or "—",
                    _v(fp.get("num_units")),
                    _v(fp.get("unit_sf")),
                    _v(fp.get("rent_per_unit"), "$"),
                    _v(fp.get("rent_psf"), "$", decimals=2),
                ])
            if fp_rows:
                story += _data_table(
                    ["Floorplan", "Units", "SF", "Rent/Unit", "Rent/SF"],
                    fp_rows,
                    col_widths=[2.0*inch, 0.8*inch, 0.9*inch, 1.1*inch, 0.9*inch],
                )

    # ── 11. Financial Analysis ─────────────────────────────────────────────
    story += _sec("11. Financial Analysis")
    fin = data.get("financial", {})

    # NOI Snapshots
    snaps = fin.get("noi_snapshots") or []
    if snaps:
        story.append(Paragraph("NOI Snapshots", S_H3))
        snap_rows = []
        for s in snaps:
            snap_rows.append([
                s.get("period") or "—",
                _v(s.get("gross_potential_rent"), "$"),
                _v(s.get("net_rental_income"), "$"),
                _v(s.get("gross_revenues"), "$"),
                _v(s.get("total_expenses"), "$"),
                _v(s.get("noi"), "$"),
                _v(s.get("physical_occupancy"), pct=True),
                _v(s.get("economic_occupancy"), pct=True),
            ])
        story += _data_table(
            ["Period", "GPR", "Net Rental", "Gross Rev", "Expenses", "NOI", "Phys Occ", "Econ Occ"],
            snap_rows,
            col_widths=[1.15*inch, 0.9*inch, 0.9*inch, 0.9*inch, 0.85*inch, 0.9*inch, 0.7*inch, 0.7*inch],
        )

    # Proforma
    pf_years = fin.get("proforma_years") or []
    if pf_years:
        story.append(Paragraph("5-Year Proforma", S_H3))
        pf_rows = []
        for y in pf_years:
            pf_rows.append([
                str(y.get("year") or "—"),
                _v(y.get("gpr"), "$"),
                _v(y.get("total_economic_loss_pct"), pct=True),
                _v(y.get("gross_revenues"), "$"),
                _v(y.get("total_expenses"), "$"),
                _v(y.get("noi"), "$"),
                _v(y.get("expense_per_unit"), "$"),
            ])
        story += _data_table(
            ["Year", "GPR", "Econ Loss", "Gross Rev", "Expenses", "NOI", "Exp/Unit"],
            pf_rows,
            col_widths=[0.55*inch, 1.0*inch, 0.85*inch, 1.0*inch, 1.0*inch, 1.0*inch, 0.85*inch],
        )

    # Key expense assumptions
    kea = fin.get("key_expense_assumptions") or {}
    story += _kv_table([
        ("Mgmt Fee",            _v(kea.get("management_fee_pct"), pct=True)),
        ("Insurance/Unit",      _v(kea.get("insurance_per_unit"), "$")),
        ("RE Tax Rate",         _v(kea.get("real_estate_tax_rate"), pct=True)),
        ("CapEx Reserve/Unit",  _v(kea.get("capex_reserve_per_unit"), "$")),
        ("Utilities/Unit",      _v(kea.get("utilities_per_unit"), "$")),
    ])

    # Historical capex
    if fin.get("historical_capex_total"):
        story.append(Paragraph(f"Historical CapEx Invested: {_v(fin['historical_capex_total'], '$')}", S_BODY))
    capex_bkdn = fin.get("historical_capex_breakdown") or []
    if capex_bkdn:
        cx_rows = []
        for item in capex_bkdn:
            if isinstance(item, dict):
                cx_rows.append([item.get("category","—"), _v(item.get("amount"), "$")])
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                cx_rows.append([str(item[0]), _v(item[1], "$")])
        if cx_rows:
            story += _data_table(["CapEx Category", "Amount"], cx_rows,
                                  col_widths=[4.5*inch, 2.0*inch])

    # ── 12. Sale Comparables ───────────────────────────────────────────────
    story += _sec("12. Sale Comparables")
    sale_comps = data.get("sale_comps") or []
    if sale_comps:
        sc_rows = []
        for sc in sale_comps:
            sc_rows.append([
                sc.get("name") or "—",
                sc.get("sale_date") or "—",
                _v(sc.get("num_units")),
                sc.get("year_built") or "—",
                _v(sc.get("sale_price"), "$"),
                _v(sc.get("price_per_unit"), "$"),
                _v(sc.get("cap_rate"), pct=True),
            ])
        story += _data_table(
            ["Property", "Sale Date", "Units", "Built", "Price", "$/Unit", "Cap Rate"],
            sc_rows,
            col_widths=[1.9*inch, 0.85*inch, 0.6*inch, 0.6*inch, 1.0*inch, 0.85*inch, 0.75*inch],
        )
    else:
        story.append(Paragraph("Sale comparables were not included in this Offering Memorandum.", S_SMALL))

    # ── 13. Market Overview ────────────────────────────────────────────────
    mkt = data.get("market_overview", {})
    story += _sec("13. Market Overview")
    story += _kv_table([
        ("Metro",               mkt.get("metro")),
        ("Metro Population",    _v(mkt.get("metro_population"))),
        ("Job Growth (YoY)",    _v(mkt.get("metro_job_growth_pct"), pct=True)),
        ("Unemployment",        _v(mkt.get("metro_unemployment"), pct=True)),
        ("Submarket",           mkt.get("submarket")),
        ("Sub. Occupancy",      _v(mkt.get("sub_occupancy"), pct=True)),
        ("Sub. Avg Rent",       _v(mkt.get("sub_rent"), "$")),
        ("Sub. Rent Growth",    mkt.get("sub_growth")),
        ("Pipeline Units",      _v(mkt.get("pipeline_units"))),
        ("Absorption Units",    _v(mkt.get("absorption_units"))),
        ("Notes",               mkt.get("notes")),
    ])

    # ── 14. Underwriting Flags ─────────────────────────────────────────────
    story += _sec("14. Underwriting Flags")
    flags = data.get("underwriting_flags") or []
    if flags:
        for flag in flags:
            cat   = (flag.get("category") or "Note").upper()
            title = flag.get("title") or ""
            detail= flag.get("detail") or ""
            if cat in ("RISK", "WARNING", "CAUTION"):
                s = S_FLAG_R
            elif cat in ("OPPORTUNITY", "UPSIDE"):
                s = S_FLAG_G
            else:
                s = S_FLAG_B
            story.append(Paragraph(f"[{cat}]  {title}", s))
            if detail:
                story.append(Paragraph(detail, S_BODY))
            story.append(Spacer(1, 4))
    else:
        story.append(Paragraph("No underwriting flags generated.", S_SMALL))

    # ── Disclaimer ─────────────────────────────────────────────────────────
    story.append(Spacer(1, 16))
    story.append(_hr())
    story.append(Paragraph(
        "DISCLAIMER: This report was AI-generated from the uploaded Offering Memorandum. "
        "For internal underwriting reference only. Verify all figures independently before making investment decisions. "
        "Powered by Anthropic Claude.",
        S_SMALL
    ))

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏢 OM Analyzer v2")
    st.markdown("---")
    st.markdown("**10 sections extracted:**")
    st.markdown("""
1. Property Details
2. Offering Terms
3. Loan Assumption *(new)*
4. Investment Highlights
5. Renovation + Value-Add Levers *(enhanced)*
6. Demographics
7. Unit Mix — 3-tier rents *(enhanced)*
8. Utilities & Billing
9. Amenities
10. Rent Comps + Floorplan Detail *(enhanced)*
11. Financial Analysis — multi-period NOI *(enhanced)*
12. Sale Comps *(graceful absent handling)*
13. Market Overview
14. Underwriting Flags *(enriched)*
""")
    st.markdown("---")
    st.markdown("**Supported brokers:** JLL · CBRE · Marcus & Millichap · Cushman · Newmark · Colliers · Berkadia · Walker & Dunlop · Northmarq · and more")
    st.markdown("---")
    st.markdown("**Max file size:** 50 MB  \n**Processing time:** 30–90 sec")

# ── API Key ───────────────────────────────────────────────────────────────────
api_key = st.secrets.get("ANTHROPIC_API_KEY", os.environ.get("ANTHROPIC_API_KEY", ""))
if not api_key:
    st.error("""
**API key not configured.**

- **Streamlit Cloud:** App → ⚙️ Settings → Secrets → add:
  ```
  ANTHROPIC_API_KEY = "sk-ant-..."
  ```
- **Local:** Create `.streamlit/secrets.toml` with the same line.
""")
    st.stop()

# ── Upload ────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload Offering Memorandum (PDF)",
    type=["pdf"],
    help="Any broker OM in PDF format. Text-based PDFs only (scanned/image PDFs are not supported).",
)

if uploaded:
    st.success(f"✅ Uploaded: **{uploaded.name}** ({uploaded.size/1024:.0f} KB)")

    if st.button("🔍 Analyze OM"):
        pdf_bytes = uploaded.read()

        with st.spinner("Extracting text from PDF…"):
            try:
                om_text = extract_text_from_pdf(pdf_bytes)
                if len(om_text.strip()) < 100:
                    st.error("Could not extract text. This may be a scanned/image-only PDF.")
                    st.stop()
                st.caption(f"Extracted {len(om_text):,} characters from {uploaded.name}")
            except Exception as e:
                st.error(f"PDF extraction failed: {e}")
                st.stop()

        with st.spinner("Analyzing with Claude AI (30–90 sec — auto-retries if output is large)…"):
            try:
                result = analyze_om_with_claude(om_text, api_key)
            except Exception as e:
                st.error(f"Claude API call failed: {e}\n\n{traceback.format_exc()}")
                st.stop()

        with st.spinner("Generating PDF report…"):
            try:
                pdf_out = generate_underwriting_pdf(result, filename=uploaded.name)
            except Exception as e:
                st.error(f"PDF generation failed: {e}\n\n{traceback.format_exc()}")
                st.stop()

        st.success("✅ Report ready!")

        # ── Download button ──────────────────────────────────────────────
        base_name = uploaded.name.replace(".pdf", "").replace(".PDF", "")
        out_name  = f"{base_name}_underwriting_report.pdf"
        st.download_button(
            label="📥 Download Underwriting Report (PDF)",
            data=pdf_out,
            file_name=out_name,
            mime="application/pdf",
        )

        # ── Quick preview of key extracted values ────────────────────────
        st.markdown("---")
        st.markdown("### Quick Preview")

        prop = result.get("property", {})
        off  = result.get("offering", {})
        fin  = result.get("financial", {})

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Units",          prop.get("num_units") or "—")
            st.metric("Year Built",     prop.get("year_built") or "—")
            st.metric("Avg Unit SF",    f"{prop.get('avg_unit_sf') or '—'}")
        with col2:
            snaps = fin.get("noi_snapshots") or []
            t12 = next((s for s in snaps if "trailing" in str(s.get("period","")).lower() or "t-12" in str(s.get("period","")).lower()), None)
            noi_display = f"${float(t12['noi']):,.0f}" if t12 and t12.get("noi") else "—"
            st.metric("T-12 NOI",       noi_display)
            loan = result.get("loan_assumption", {})
            if loan.get("available"):
                st.metric("Assume Rate",  f"{float(loan.get('note_rate',0))*100:.2f}%" if loan.get("note_rate") else "—")
                st.metric("Loan Balance", f"${float(loan.get('current_balance',0)):,.0f}" if loan.get("current_balance") else "—")
            else:
                st.metric("Loan Assumption", "None")
        with col3:
            mix    = result.get("unit_mix") or []
            n_types = len(mix)
            avg_ip  = sum(float(u.get("rent_inplace") or 0) for u in mix) / len(mix) if mix else None
            avg_mkt = sum(float(u.get("rent_market") or 0) for u in mix) / len(mix) if mix else None
            st.metric("Floor Plan Types",  n_types or "—")
            st.metric("Avg In-Place Rent", f"${avg_ip:,.0f}" if avg_ip else "—")
            st.metric("Avg Market Rent",   f"${avg_mkt:,.0f}" if avg_mkt else "—")

        # Flags preview
        flags = result.get("underwriting_flags") or []
        if flags:
            st.markdown("**Underwriting Flags:**")
            for f in flags[:6]:
                cat = (f.get("category") or "NOTE").upper()
                icon = "🔴" if cat in ("RISK","WARNING","CAUTION") else ("🟢" if cat in ("OPPORTUNITY","UPSIDE") else "🔵")
                st.markdown(f"{icon} **[{cat}]** {f.get('title','')} — {f.get('detail','')}")

        # Raw JSON expander (useful for debugging)
        with st.expander("🔧 Raw extracted JSON (debug)"):
            import json
            st.code(json.dumps(result, indent=2), language="json")

else:
    st.info("👆 Upload an Offering Memorandum PDF to get started.")

    st.markdown("""
---
**How it works:**
1. Upload any multifamily OM PDF from any broker
2. Claude reads the full document and extracts all underwriting data
3. A structured 14-section PDF report is generated for download

**What's new in v2:**
- Three-tier rent columns in unit mix (in-place / market / comp-supported)
- Per-floorplan detail tables extracted for each rent comparable
- Assumable debt / loan assumption captured as its own section
- Multiple NOI snapshots (T-12, 6-month annualized, 90-day, 30-day, proforma)
- Value-add levers table with unit counts, premiums, and annual upside totals
- Sale comps absence handled gracefully — no errors, clean "not provided" message
- Richer underwriting flags (capex age, econ vs physical occ gap, rate lock quality)
""")
