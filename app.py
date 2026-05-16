import os, json, re, tempfile, io
from datetime import datetime
import streamlit as st
import anthropic
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="OM Analyzer · RealVal",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ══════════════════════════════════════════════════════════════════════════════
# AUTH GATE — must come right after set_page_config
# ══════════════════════════════════════════════════════════════════════════════
from auth import login_page, logout
if "user" not in st.session_state:
    login_page()
    st.stop()
# ══════════════════════════════════════════════════════════════════════════════


import base64, pathlib

# ─── Logo loader (caches base64 of the brand logo for embedding) ──────────────
@st.cache_data
def _load_logo_b64():
    """Load the combined RealVal nav logo (RV mark + 'REALVAL' wordmark together)
    as base64. Falls back to other available logo assets if the preferred one
    is missing."""
    try:
        # Prefer the combined logo (mark + wordmark, white/teal on dark)
        p = pathlib.Path(__file__).parent / "assets" / "realval_logo_navwhite.png"
        if not p.exists():
            p = pathlib.Path(__file__).parent / "assets" / "realval_logo_nav.png"
        if not p.exists():
            p = pathlib.Path(__file__).parent / "assets" / "realval_logo_transparent.png"
        if not p.exists():
            p = pathlib.Path(__file__).parent / "assets" / "realval_logo.png"
        if p.exists():
            return base64.b64encode(p.read_bytes()).decode("ascii")
    except Exception:
        pass
    return ""

_LOGO_B64 = _load_logo_b64()

st.markdown(f"""
<style>
/* ═════════════════════════════════════════════════════════════════════════
   REALVAL OM INTELLIGENCE — Streamlit theme
   Brand:  black + teal (#02A9A1) + warm white
   ═══════════════════════════════════════════════════════════════════════ */

/* Hide Streamlit chrome */
#MainMenu, footer, header {{ visibility: hidden; height: 0 !important; }}

/* Base — clean white app surface */
html, body, [class*="css"], .stApp, .main {{
    background-color: #FAFAFA !important;
    color: #0A0A0A !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Inter', 'Segoe UI', sans-serif !important;
}}

.block-container {{
    padding: 0 !important;
    max-width: 100% !important;
    background: #FAFAFA !important;
}}

/* ═════════════════════════════════════════════════════════════════════════
   NAVBAR — black, slim, logo-only top-left, clean nav, user chip top-right
   ═══════════════════════════════════════════════════════════════════════ */
.rv-nav {{
    background: #0A0A0A;
    padding: 0 36px;
    height: 64px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 1px solid #1A1A1A;
}}
.rv-nav-left {{ display: flex; align-items: center; gap: 14px; }}
.rv-brand,
.rv-brand:link,
.rv-brand:visited,
.rv-brand:hover,
.rv-brand:active {{
    display: flex; align-items: center; gap: 10px;
    text-decoration: none !important;
    border-bottom: none !important;
}}
.rv-brand * {{ text-decoration: none !important; }}
.rv-logo-full {{
    height: 32px; width: auto; display: block;
}}
.rv-mark {{
    height: 32px; width: auto; display: block;
}}
.rv-mark-fallback {{
    color: #FFFFFF; font-weight: 500; font-size: 18px;
    letter-spacing: 1.5px;
}}
.rv-wordmark {{
    color: #FFFFFF;
    font-size: 18px;
    font-weight: 500;
    letter-spacing: 2.5px;
    line-height: 1;
}}
.rv-nav-tagline {{
    color: #02A9A1;
    font-size: 11px;
    padding: 3px 10px;
    border: 1px solid #02A9A1;
    border-radius: 4px;
    letter-spacing: 0.4px;
    text-transform: uppercase;
    font-weight: 500;
    margin-left: 8px;
}}
.rv-nav-right {{ display: flex; align-items: center; gap: 24px; }}
.rv-nav-link {{
    color: #9CA3AF;
    font-size: 13px;
    text-decoration: none;
}}
.rv-nav-link.active {{ color: #FFFFFF; border-bottom: 1px solid #02A9A1; padding-bottom: 4px; }}
.rv-user-chip {{
    display: flex; align-items: center; gap: 10px;
    color: #D1D5DB; font-size: 12px;
}}
.rv-user-chip-avatar {{
    width: 30px; height: 30px; border-radius: 50%;
    background: #1F1F1F; color: #02A9A1;
    display: flex; align-items: center; justify-content: center;
    font-size: 11px; font-weight: 500;
    border: 1px solid #2A2A2A;
    cursor: pointer;
    transition: border-color 0.15s, background 0.15s;
}}
.rv-user-chip-avatar:hover {{
    border-color: #02A9A1;
    background: #2A2A2A;
}}

/* ═════════════════════════════════════════════════════════════════════════
   PROFILE DROPDOWN — click avatar → menu with email + Sign out
   Uses native <details>/<summary> so no JS needed; closes on outside click
   ═══════════════════════════════════════════════════════════════════════ */
.rv-profile {{
    position: relative;
    display: inline-block;
}}
.rv-profile-trigger {{
    list-style: none;
    cursor: pointer;
    display: flex;
    align-items: center;
}}
.rv-profile-trigger::-webkit-details-marker {{ display: none; }}
.rv-profile-trigger::marker {{ display: none; content: ""; }}
details[open] .rv-user-chip-avatar {{
    border-color: #02A9A1;
    background: #2A2A2A;
}}
.rv-profile-menu {{
    position: absolute;
    top: calc(100% + 8px);
    right: 0;
    min-width: 240px;
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 10px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.12), 0 2px 4px rgba(0,0,0,0.06);
    z-index: 9999;
    padding: 4px 0;
    overflow: hidden;
}}
.rv-profile-header {{
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 12px 14px 10px;
}}
.rv-profile-avatar-lg {{
    width: 36px; height: 36px; border-radius: 50%;
    background: #02A9A1; color: #FFFFFF;
    display: flex; align-items: center; justify-content: center;
    font-size: 13px; font-weight: 500;
    flex-shrink: 0;
}}
.rv-profile-info {{ min-width: 0; flex: 1; }}
.rv-profile-name {{
    color: #0A0A0A;
    font-size: 13px;
    font-weight: 500;
    line-height: 1.3;
    text-transform: capitalize;
}}
.rv-profile-email {{
    color: #6B7280;
    font-size: 11px;
    line-height: 1.3;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}}
.rv-profile-divider {{
    height: 1px;
    background: #F3F4F6;
    margin: 4px 0;
}}
.rv-profile-item,
.rv-profile-item:visited,
.rv-profile-item:hover,
.rv-profile-item:active {{
    display: block;
    padding: 9px 14px;
    color: #374151 !important;
    font-size: 13px;
    text-decoration: none !important;
    cursor: pointer;
    transition: background 0.1s;
}}
.rv-profile-item:hover {{ background: #F9FAFB; }}
.rv-profile-item svg {{ color: inherit; stroke: currentColor; }}
.rv-profile-item-danger,
.rv-profile-item-danger:visited,
.rv-profile-item-danger:hover,
.rv-profile-item-danger:active {{
    color: #B91C1C !important;
}}
.rv-profile-item-danger:hover {{ background: #FEF2F2 !important; color: #991B1B !important; }}

/* ═════════════════════════════════════════════════════════════════════════
   HERO — clean upload area, centered, gold-standard whitespace
   ═══════════════════════════════════════════════════════════════════════ */
.rv-hero {{
    padding: 56px 24px 40px;
    text-align: center;
}}
.rv-hero-title {{
    color: #0A0A0A;
    font-size: 26px;
    font-weight: 500;
    letter-spacing: -0.3px;
    margin: 0 0 10px;
}}
.rv-hero-sub {{
    color: #6B7280;
    font-size: 14px;
    line-height: 1.6;
    max-width: 560px;
    margin: 0 auto;
}}

/* Streamlit file uploader — restyle to match dashed teal dropzone */
[data-testid="stFileUploader"] {{
    max-width: 580px;
    margin: 24px auto 0;
}}
[data-testid="stFileUploader"] > section {{
    background: #FFFFFF !important;
    border: 1.5px dashed #02A9A1 !important;
    border-radius: 12px !important;
    padding: 36px !important;
    text-align: center;
    transition: border-color 0.15s, background 0.15s;
}}
[data-testid="stFileUploader"] > section:hover {{
    background: #F0FBF9 !important;
    border-color: #018A82 !important;
}}
[data-testid="stFileUploader"] section button {{
    background: transparent !important;
    color: #02A9A1 !important;
    border: 1px solid #02A9A1 !important;
    border-radius: 6px !important;
    padding: 6px 18px !important;
    font-size: 12px !important;
    font-weight: 500 !important;
    box-shadow: none !important;
}}
[data-testid="stFileUploader"] small {{
    color: #6B7280 !important;
    font-size: 12px !important;
}}

/* ═════════════════════════════════════════════════════════════════════════
   STEP CARDS — below upload, three compact cards
   ═══════════════════════════════════════════════════════════════════════ */
.rv-steps {{
    max-width: 580px;
    margin: 32px auto 0;
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 12px;
}}
.rv-step {{
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 8px;
    padding: 14px 16px;
}}
.rv-step-num {{
    color: #02A9A1;
    font-size: 11px;
    font-weight: 500;
    letter-spacing: 0.5px;
    margin-bottom: 4px;
}}
.rv-step-title {{
    color: #0A0A0A;
    font-size: 13px;
    font-weight: 500;
    margin-bottom: 2px;
}}
.rv-step-desc {{
    color: #6B7280;
    font-size: 11px;
    line-height: 1.5;
}}

/* ═════════════════════════════════════════════════════════════════════════
   FILE INFO BAR — shown after upload, with the analyze button next to it
   ═══════════════════════════════════════════════════════════════════════ */
.rv-file-bar {{
    max-width: 720px;
    margin: 20px auto 16px;
    padding: 14px 20px;
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}}
.rv-file-meta {{ color: #0A0A0A; font-size: 13px; }}
.rv-file-meta b {{ font-weight: 500; }}
.rv-file-size {{ color: #6B7280; font-size: 12px; }}

/* ═════════════════════════════════════════════════════════════════════════
   STREAMLIT BUTTONS — primary CTA = teal, others = ghost
   ═══════════════════════════════════════════════════════════════════════ */
.stButton > button {{
    background: transparent !important;
    color: #4B5563 !important;
    border: 1px solid #D1D5DB !important;
    border-radius: 6px !important;
    padding: 7px 16px !important;
    font-size: 12px !important;
    font-weight: 500 !important;
    transition: all 0.15s !important;
    box-shadow: none !important;
}}
.stButton > button:hover {{
    border-color: #02A9A1 !important;
    color: #02A9A1 !important;
    background: #F0FBF9 !important;
}}
.stButton > button[kind="primary"] {{
    background: #02A9A1 !important;
    color: #FFFFFF !important;
    border: 1px solid #02A9A1 !important;
    padding: 10px 28px !important;
    font-size: 14px !important;
    font-weight: 500 !important;
}}
.stButton > button[kind="primary"]:hover {{
    background: #018A82 !important;
    border-color: #018A82 !important;
}}

.stDownloadButton > button {{
    background: #02A9A1 !important;
    color: #FFFFFF !important;
    border: 1px solid #02A9A1 !important;
    border-radius: 6px !important;
    padding: 10px 22px !important;
    font-weight: 500 !important;
    box-shadow: none !important;
}}
.stDownloadButton > button:hover {{
    background: #018A82 !important;
    border-color: #018A82 !important;
}}

/* ═════════════════════════════════════════════════════════════════════════
   CUSTOMIZE EXPANDER — replaces the right-panel checkboxes
   ═══════════════════════════════════════════════════════════════════════ */
[data-testid="stExpander"] {{
    max-width: 720px;
    margin: 0 auto 20px;
    background: #FFFFFF !important;
    border: 1px solid #E5E7EB !important;
    border-radius: 8px !important;
    box-shadow: none !important;
}}
[data-testid="stExpander"] summary {{
    padding: 12px 18px !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    color: #0A0A0A !important;
}}
[data-testid="stExpander"] summary:hover {{ background: #F9FAFB !important; }}
[data-testid="stExpander"] [data-testid="stExpanderDetails"] {{
    padding: 4px 18px 18px !important;
    border-top: 1px solid #F3F4F6 !important;
}}
.rv-cust-group-title {{
    color: #02A9A1;
    font-size: 11px;
    font-weight: 500;
    letter-spacing: 0.6px;
    text-transform: uppercase;
    margin: 12px 0 6px;
}}

/* Checkbox styling
   DOM: label > span(checkbox-square) + input + div(label-text)
   We need to (a) keep the square teal when checked, (b) keep the text plain */
[data-testid="stCheckbox"] {{ margin-bottom: 4px; }}
[data-testid="stCheckbox"] label > div {{
    background: transparent !important;
}}
[data-testid="stCheckbox"] label > div p {{
    background: transparent !important;
    color: #374151 !important;
    font-size: 12.5px !important;
    margin: 0 !important;
    line-height: 1.4 !important;
}}

/* ═════════════════════════════════════════════════════════════════════════
   RESULTS — metric strip + tabs + download bar
   ═══════════════════════════════════════════════════════════════════════ */
.rv-results-header {{
    max-width: 1200px; margin: 28px auto 12px;
    padding: 0 24px;
}}
.rv-prop-name {{ color: #0A0A0A; font-size: 22px; font-weight: 500; margin: 0; }}
.rv-prop-meta {{ color: #6B7280; font-size: 13px; margin-top: 4px; }}

.rv-metric {{
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 8px;
    padding: 14px 16px;
    height: 100%;
}}
.rv-metric-label {{
    color: #6B7280;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 6px;
}}
.rv-metric-value {{ color: #0A0A0A; font-size: 18px; font-weight: 500; }}

/* Sticky download bar */
.rv-dl-bar {{
    max-width: 1200px;
    margin: 16px auto;
    padding: 16px 24px;
    background: #FFFFFF;
    border: 1px solid #02A9A1;
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 2px 8px rgba(2, 169, 161, 0.08);
}}
.rv-dl-label {{ color: #0A0A0A; font-size: 13px; font-weight: 500; }}
.rv-dl-sub {{ color: #6B7280; font-size: 11px; margin-top: 2px; }}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {{
    background: transparent;
    gap: 0;
    border-bottom: 1px solid #E5E7EB;
    padding: 0 24px;
}}
.stTabs [data-baseweb="tab"] {{
    background: transparent !important;
    color: #6B7280 !important;
    font-size: 13px !important;
    padding: 10px 16px !important;
    border-bottom: 2px solid transparent !important;
}}
.stTabs [aria-selected="true"] {{
    color: #0A0A0A !important;
    border-bottom-color: #02A9A1 !important;
    font-weight: 500 !important;
}}

/* Streamlit alerts */
.stAlert {{ border-radius: 8px !important; max-width: 1200px; margin: 12px auto !important; }}

/* Progress bar */
.stProgress > div > div > div > div {{ background: #02A9A1 !important; }}

/* ═════════════════════════════════════════════════════════════════════════
   FLAG CARDS — in-app Flags tab rendering
   Each flag = colored left border + soft tinted bg + bold title + body text.
   Color-coded by category: warn=amber, good=green, verify=purple, info=blue
   ═══════════════════════════════════════════════════════════════════════ */
.flag-warn, .flag-good, .flag-verify, .flag-info {{
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-left: 4px solid #9CA3AF;
    border-radius: 8px;
    padding: 14px 18px;
    margin: 0 0 12px 0;
}}
.flag-warn   {{ border-left-color: #D97706; background: #FFFBEB; }}   /* amber */
.flag-good   {{ border-left-color: #059669; background: #F0FDF4; }}   /* green */
.flag-verify {{ border-left-color: #7C3AED; background: #F5F3FF; }}   /* purple */
.flag-info   {{ border-left-color: #02A9A1; background: #F0FBF9; }}   /* brand teal */

.flag-title {{
    color: #0A0A0A;
    font-size: 14px;
    font-weight: 600;
    margin-bottom: 6px;
    line-height: 1.3;
}}
.flag-body {{
    color: #374151;
    font-size: 13px;
    line-height: 1.55;
}}

/* Footer */
.rv-footer {{
    margin-top: 60px;
    padding: 16px 36px;
    background: #F4F4F4;
    border-top: 1px solid #E5E7EB;
    display: flex;
    justify-content: space-between;
    font-size: 11px;
    color: #6B7280;
}}
.rv-footer a {{ color: #02A9A1; text-decoration: none; }}

/* Scrollbar */
::-webkit-scrollbar {{ width: 8px; }}
::-webkit-scrollbar-track {{ background: #F4F4F4; }}
::-webkit-scrollbar-thumb {{ background: #D1D5DB; border-radius: 4px; }}
::-webkit-scrollbar-thumb:hover {{ background: #9CA3AF; }}
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
# SECTION 2 — AI ANALYSIS
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
      "loan_type": null, "lender": null, "loan_amount": null, "loan_to_value": null,
      "interest_rate": null, "rate_type": null, "amortization_years": null,
      "loan_term_years": null, "interest_only_period": null, "dscr": null,
      "recourse": null, "notes": null
    },
    "assumable_debt": {
      "loan_type": null, "lender": null, "loan_amount": null, "loan_to_value": null,
      "interest_rate": null, "rate_type": null, "amortization_years": null,
      "loan_term_years": null, "interest_only_period": null,
      "origination_date": null, "maturity_date": null,
      "monthly_payment": null, "annual_debt_service": null, "dscr": null,
      "prepayment_penalty": null, "recourse": null, "notes": null
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
      {"name": null, "category": null, "fee_per_unit_per_month": null,
       "occupancy_assumption": null, "monthly_income": null,
       "current_annual": null, "proforma_annual": null,
       "calculation_detail": null, "notes": null}
    ],
    "amenities": [], "unit_features": [], "highlights": []
  },
  "value_add": {
    "scope": null, "total_cost": null, "cost_per_unit": null, "exterior_capex": null,
    "monthly_premium": null, "annual_premium": null, "roi_pct": null,
    "light_upgrade_items": [],
    "by_floor_plan": [
      {"type": null, "sf": null, "units": null, "inplace_rent": null, "inplace_psf": null,
       "rehab_cost": null, "premium": null, "post_rehab_rent": null, "post_rehab_psf": null}
    ]
  },
  "value_add_levers": [
    {"lever": null, "units": null, "monthly_premium": null, "annual_upside": null, "notes": null}
  ],
  "tax": {
    "parcel_id": null, "assessed_value": null, "millage_city": null,
    "millage_county": null, "millage_total": null, "tax_base": null,
    "solid_waste_fee": null, "total_tax": null, "abatement_program": null,
    "abatement_pct": null, "abatement_term_note": null, "abatement_annual_savings": null,
    "ami_pct": null, "max_allowable_rent": null, "avg_inplace_rent": null,
    "rent_headroom": null, "units_compliant": null, "pct_compliant": null
  },
  "demographics": {
    "pop_1mi": null, "pop_3mi": null, "pop_5mi": null,
    "pop_growth_1mi": null, "pop_growth_3mi": null, "pop_growth_5mi": null,
    "pop_2030_1mi": null, "pop_2030_3mi": null, "pop_2030_5mi": null,
    "median_income_1mi": null, "median_income_3mi": null, "median_income_5mi": null,
    "median_income_2030_1mi": null, "median_income_2030_3mi": null, "median_income_2030_5mi": null,
    "income_growth_1mi": null, "income_growth_3mi": null, "income_growth_5mi": null,
    "renter_pct_1mi": null, "renter_pct_3mi": null, "renter_pct_5mi": null,
    "college_pct_1mi": null, "college_pct_3mi": null, "college_pct_5mi": null,
    "white_collar_pct_1mi": null, "white_collar_pct_3mi": null, "white_collar_pct_5mi": null,
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
     "comp_type": null, "total_market": null, "total_market_psf": null,
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
    "current_rent": null, "avg_hh_income_3mi": null,
    "monthly_affordability_3x": null, "rent_headroom_3mi": null,
    "avg_hh_income_2030_3mi": null, "monthly_affordability_2030_3x": null,
    "rent_headroom_2030_3mi": null, "income_to_rent_ratio": null, "notes": null
  },
  "insurance": {
    "carrier": null, "annual_premium": null, "per_unit": null,
    "quote_source": null, "notes": null
  },
  "management": {
    "fee_pct": null, "fee_annual": null, "fee_per_unit": null,
    "current_manager": null, "proposed_manager": null, "notes": null
  },
  "replacement_cost": {
    "per_unit": null, "per_sf": null, "total": null, "source": null,
    "land_per_unit": null, "land_total": null,
    "hard_cost_per_sf": null, "hard_cost_per_unit": null, "hard_cost_total": null,
    "soft_cost_pct": null, "soft_cost_per_unit": null, "soft_cost_total": null,
    "direct_replacement_per_unit": null, "direct_replacement_per_sf": null, "direct_replacement_total": null,
    "developer_fee_pct": null, "developer_fee_per_unit": null, "developer_fee_total": null,
    "gc_fee_pct": null, "gc_fee_per_unit": null, "gc_fee_total": null,
    "gross_replacement_per_unit": null, "gross_replacement_per_sf": null, "gross_replacement_total": null,
    "notes": null, "source": null
  },
  "investment_highlights": [
    "string narrative point about this property's investment thesis"
  ],
  "concession_burnoff": {
    "has_concessions": null,
    "monthly_income_current": null,
    "monthly_income_projected": null,
    "burnoff_timeline": null,
    "notes": null
  },
  "flags": [
    {"category": null, "title": null, "detail": null}
  ]
}

CRITICAL EXTRACTION RULES:
1. DEMOGRAPHICS: columns order is 1-mile, 3-mile, 5-mile, County, Metro. Extract all three radii.
   If 1-mile data is not present in the OM, leave 1-mile fields as null.
2. FINANCIAL PERIODS: Extract ALL column headers exactly as they appear (e.g. "T-12 Actual", "T-9 Annualized",
   "T-6 Annualized", "T-3 Annualized", "T-1 Annualized", "Year 1 Pro Forma", "Pro Forma YR1", "Year 0",
   "YR1", "YR2", "FY1", "F-3 Proforma Income", "Current Rents Proforma", "2nd Generation Leases Proforma",
   "Year 2", "Year 3", etc.). Do NOT rename or standardize them — preserve the exact label from the OM.

   *** CRITICAL PRIORITY ORDER (this is the #1 most important rule for financials) ***
   When the OM has BOTH an Operating Statement (with trailing periods + a Pro Forma column) AND a separate
   multi-year Cash Flow projection table, you MUST extract the Operating Statement columns first, in this
   priority order:
     (a) Pro Forma / Year 1 Pro Forma  ← ALWAYS INCLUDE — this is the most important column for buyers
     (b) T-12 Actual (or T12)           ← ALWAYS INCLUDE
     (c) T-3 Annualized                 ← include if present
     (d) T-6 Annualized                 ← include if present
     (e) T-1 Annualized                 ← include if present
     (f) T-9 Annualized                 ← include only if room remains
     (g) Year 2, Year 3 from Cash Flow  ← include only if room remains AFTER above

   The Pro Forma / Year 1 Pro Forma column is NEVER allowed to be dropped. If the Operating Statement has
   a "Year 1 Pro Forma" or "Pro Forma" column (typically the rightmost column on the operating statement page),
   that column MUST appear in the periods array and every income_line and expense_line MUST have its value
   under that period key.

   If the OM ONLY has a multi-year Cash Flow table (no separate trailing operating statement), then extract
   Year 1 through Year 5 as periods.

   Cap total periods at 7 columns. Every income and expense line must have values for ALL period columns present.
   Mark is_total=true for EGI, Total Expenses, NOI, CFFO rows.
   Mark is_subtotal=true for subtotal rows (Net Rental Income, Rental Collections, Total Controllable,
   Total Non-Controllable, NOI Before Reserves, etc.).
   The note field must contain the full underwriting assumption text/footnote from the OM.

3. RENT COMPS: Use rent_comps_garden for garden/flat apartment comps, rent_comps_townhouse for townhouse comps.
   If the OM does not split by type, put all comps in rent_comps (full detail array) only.
   Set comp_type = "Garden", "Townhouse", "Flat", "Mid-Rise", "High-Rise", or as labeled in the OM.
   Include distance and year_built for every comp. Include avg occupancy if shown.
   Include a "Subject" row and "Average" row if the OM shows them.

4. VALUE-ADD: Extract the full floor plan table. If no floor plan table exists (only light upgrade list),
   leave by_floor_plan as [] and populate light_upgrade_items instead.
   For rows with N/A rent/cost, use null — never use 0.
   REVENUE LEVERS: If the OM contains a structured revenue/value-add levers table (e.g. rows listing
   "Continue Interior Renovation", "Push Rents to Market", "Install Package Lockers", "Bulk Cable/Internet",
   "Valet Trash", "Smart Home Tech", "Covered Parking", "Washer/Dryer Equipment" etc.), extract EVERY row
   into value_add_levers with: lever (name), units (unit count), monthly_premium ($/unit/month),
   annual_upside (total annual $ upside), notes. Include the grand total row as the last entry with
   lever="TOTAL ANNUAL REVENUE UPSIDE".

5. TAX: Extract parcel, millage, abatement program, AMI compliance if present.
   If no abatement program exists, set all abatement fields to null.

6. OTHER INCOME: Extract EVERY other income line from the OM into additional_income array.
   Include fees, utility reimbursements, parking, internet, laundry, storage, pet, admin, etc.
   For each: name, category, fee_per_unit_per_month, occupancy_assumption,
   monthly_income, current_annual, proforma_annual, calculation_detail.

7. FLAGS: Generate 6-10 flags relevant to THIS specific property. category = Warning / Opportunity / Verify / Info.
   Base flags on actual data found — occupancy trends, expense anomalies, market catalysts,
   value-add ROI, vacancy assumptions, bad debt, tax risks, financing terms, etc.
   Do NOT fabricate flags for things not mentioned in the OM.

8. FINANCING: Extract offering_type, debt_contact, new_financing, assumable_debt fully.
   If All Cash, set offering_type="All Cash" and leave financing sub-objects null.
   If Free & Clear with a soft quote provided, populate new_financing with the quoted terms.
   Convert percentage strings to decimals (e.g. "72%" → 0.72, "5.75%" → 0.0575).
   For assumable_debt, always extract interest_only_period (e.g. "60 months", "5 years") if stated.

9. SALE COMPS: Extract buyer and seller names if disclosed. Include cap rate, $/unit, $/SF.

10. MARKET: Populate market_summary with a 2-3 sentence narrative.
    Populate major_developments with every named project including cost, jobs, timeline.

11. AFFORDABILITY: Extract rent-to-income table if present. Calculate:
    monthly_affordability_3x = avg_hh_income_3mi / 12 / 3
    rent_headroom = monthly_affordability_3x - current_rent

12. INSURANCE: Carrier, annual premium, per-unit cost, quote source.

13. MANAGEMENT: Fee % of EGI, annual $, per-unit, current and proposed manager names.

14. REPLACEMENT COST: Extract the full table if present. Fields: land_per_unit, land_total,
    hard_cost_per_sf, hard_cost_per_unit, hard_cost_total, soft_cost_pct, soft_cost_per_unit, soft_cost_total,
    direct_replacement_per_unit, direct_replacement_per_sf, direct_replacement_total,
    developer_fee_pct, developer_fee_per_unit, developer_fee_total,
    gc_fee_pct, gc_fee_per_unit, gc_fee_total,
    gross_replacement_per_unit, gross_replacement_per_sf, gross_replacement_total.
    Also set per_unit = gross_replacement_per_unit and total = gross_replacement_total.
    If no breakdown exists, populate just per_unit, per_sf, total, source, notes.

15. DEMOGRAPHICS: Extract 1-mile, 3-mile and 5-mile data separately. Never mix them.
    If demographic data is not in the OM, set all demographic fields to null.

16. Extract every number that exists anywhere in the OM. Do not skip any table or data page.

17. INVESTMENT HIGHLIGHTS: Extract the 6-10 bullet-point highlights from the executive summary / investment profile
    into investment_highlights as an array of strings. These are the broker's key selling points.

18. CONCESSION BURNOFF: If the OM contains a concession burn-off analysis or timeline, extract:
    monthly_income_current (current monthly income before burnoff), monthly_income_projected (projected after burnoff),
    burnoff_timeline (e.g. "April-May 2026"), and a notes narrative.

19. COLLECTIONS SUMMARY: If a trailing monthly collections table exists (T-6 or similar), include each month
    as a period column in periods array (e.g. "Jun", "Jul", "Aug", "Sep", "Oct", "Nov") so the operating
    statement correctly reflects the trailing period detail. If trailing months AND proforma columns both exist,
    prioritize the proforma columns but include T-3 or T-2 actuals as well.
"""


def analyze_om(pdf_text: str, api_key: str, progress_cb=None) -> dict:
    import httpx
    MAX = 180_000
    text = pdf_text[:MAX]
    if len(pdf_text) > MAX and progress_cb:
        progress_cb("Large OM — using first 180K characters")
    if progress_cb:
        progress_cb("Sending to Claude AI for analysis...")

    # Use context manager so the http client is closed cleanly each call.
    # Timeout shortened to 180s so it fails fast instead of letting the
    # Streamlit Cloud websocket time out first.
    def _call(prompt_text, max_tok, system_text):
        with httpx.Client(timeout=httpx.Timeout(180.0, connect=15.0)) as hc:
            client = anthropic.Anthropic(api_key=api_key, http_client=hc)
            resp = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=max_tok,
                system=system_text,
                messages=[{"role": "user", "content": prompt_text}]
            )
        raw = resp.content[0].text.strip()
        raw = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("`").strip()
        return raw, resp.stop_reason

    def _parse(raw):
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            pass
        try:
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            if m:
                return json.loads(m.group())
        except json.JSONDecodeError:
            pass
        try:
            return json.loads(_fix_truncated_json(raw))
        except Exception:
            pass
        return None

    raw, stop_reason = _call(
        f"Analyze this OM:\n\n{text}",
        max_tok=32000,
        system_text=SYSTEM
    )
    if progress_cb:
        progress_cb("Parsing extracted data...")

    result = _parse(raw)
    if result is not None:
        return result

    if stop_reason == "max_tokens" or len(raw) > 100:
        if progress_cb:
            progress_cb("Response truncated — retrying completion...")
        raw2, _ = _call(
            f"Complete this truncated JSON so it is fully valid. Output ONLY the completed JSON:\n\n{raw}",
            max_tok=16000,
            system_text="You are a JSON completion assistant. Output ONLY valid JSON, nothing else."
        )
        for candidate in [raw + raw2, raw2]:
            result = _parse(candidate)
            if result is not None:
                return result

    raise ValueError("Could not parse AI response. Try uploading a smaller OM or check your API key.")


def _fix_truncated_json(raw: str) -> str:
    in_string = escape_next = False
    for ch in raw:
        if escape_next: escape_next = False; continue
        if ch == "\\": escape_next = True; continue
        if ch == '"': in_string = not in_string
    if in_string:
        raw += '"'
    opens, in_str, esc = [], False, False
    for ch in raw:
        if esc: esc = False; continue
        if ch == "\\": esc = True; continue
        if ch == '"': in_str = not in_str; continue
        if not in_str:
            if ch in "{[": opens.append("}" if ch == "{" else "]")
            elif ch in "}]" and opens: opens.pop()
    return raw + "".join(reversed(opens))


def _v(val, fmt=None, suffix="", default="N/A"):
    if val is None or val == "": return default
    if fmt == "$":
        try: return f"${float(val):,.0f}{suffix}"
        except: return str(val)
    if fmt == "%":
        try:
            f = float(val)
            # Auto-detect decimals: 0.72 → 72%, 5.75 → 5.75%
            if 0 < abs(f) < 1.0:
                f = f * 100
            # 2 decimals if not a whole number, else 1
            if abs(f - round(f)) < 0.01:
                return f"{f:.1f}%"
            return f"{f:.2f}%"
        except: return str(val)
    if fmt == "n":
        try: return f"{int(float(val)):,}{suffix}"
        except: return str(val)
    return f"{val}{suffix}"

def _pct(val, suffix="", default="N/A"):
    if val is None or val == "": return default
    try:
        f = float(val)
        if abs(f) <= 2.0:
            f = f * 100
        result = f"{f:.2f}%"
        return result + suffix if suffix else result
    except:
        s = str(val).strip()
        return s if s.endswith("%") else f"{s}%"

def _psf(val):
    if val is None or val == "" or val == "N/A": return "—"
    try:
        f = float(val)
        return "—" if f == 0 else f"${f:.2f}"
    except: return "—"

def _is_num(v):
    if v is None or str(v).strip().lower() in ("", "n/a", "na", "—", "-"): return False
    try: float(str(v).replace("$","").replace(",","")); return True
    except: return False


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — EXCEL REPORT GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

C_HDR      = "FF44546A"
C_HDR2     = "FF2C3644"
# ══════════════════════════════════════════════════════════════════════════════
# CLEAN CORPORATE STYLING SYSTEM
# White + grey + deep-navy accent. Bottom-borders only. Real Excel numbers
# where possible (auto-detected from formatted strings).
# ══════════════════════════════════════════════════════════════════════════════

# ── Palette ──────────────────────────────────────────────────────────────────
C_NAVY        = "FF1E3A5F"  # Primary accent (section headers, totals text)
C_NAVY_LIGHT  = "FF2C4F7E"  # Slightly lighter navy (subheaders' bottom rule)
C_TEXT        = "FF1F2937"  # Charcoal body text
C_TEXT_MUTED  = "FF6B7280"  # Secondary / labels
C_DIVIDER     = "FFE5E7EB"  # Light grey divider lines
C_SUB_FILL    = "FFF3F4F6"  # Sub-header band background
C_ROW_ALT     = "FFFAFAFA"  # Zebra striping (very subtle)
C_WHITE       = "FFFFFFFF"
C_TOTAL_FILL  = "FFEFF3F8"  # Subtotal band (very pale blue-grey)
C_NEG         = "FFB91C1C"  # Red for negative numbers

# Flag colors (kept subtle, just left-border accent in implementation)
C_WARN_BG     = "FFFEF3C7"  # Soft amber
C_WARN_BAR    = "FFD97706"
C_GOOD_BG     = "FFD1FAE5"  # Soft green
C_GOOD_BAR    = "FF059669"
C_INFO_BG     = "FFDBEAFE"  # Soft blue
C_INFO_BAR    = "FF2563EB"
C_VER_BG      = "FFE9D5FF"  # Soft purple
C_VER_BAR     = "FF7C3AED"

# ── Legacy aliases (kept so the existing build_excel call sites still work) ──
C_HDR      = C_NAVY
C_HDR2     = C_NAVY
C_HDR3     = C_NAVY
C_SUB_HDR  = C_SUB_FILL
C_HDR_TEXT = C_WHITE
C_AMBER    = C_NAVY        # totals text on light fill — now navy not amber
C_BLUE_IN  = C_TEXT        # data text — charcoal not blue
C_LABEL    = C_TEXT_MUTED
C_BODY     = C_TEXT
C_ALT      = C_ROW_ALT
C_SUBTOTAL = C_TOTAL_FILL
C_WARN     = C_WARN_BG
C_GREEN_L  = C_GOOD_BG
C_BLUE_L   = C_INFO_BG
C_PURPLE_L = C_VER_BG
C_BORDER   = C_DIVIDER


# ── Borders: bottom-only, light grey ─────────────────────────────────────────
def _fill(c): return PatternFill("solid", fgColor=c)

_NO_BORDER     = Border()
_BOTTOM_BORDER = Border(bottom=Side(style="thin", color=C_DIVIDER))
_THICK_BOTTOM  = Border(bottom=Side(style="medium", color=C_NAVY))
_TOP_RULE      = Border(top=Side(style="medium", color=C_NAVY))

def _border():
    """Legacy compat — returns the standard subtle bottom rule."""
    return _BOTTOM_BORDER


# ── Number-format detection ──────────────────────────────────────────────────
_RE_CURRENCY = re.compile(r"^\$?-?\(?[\d,]+(\.\d+)?\)?$")
_RE_PERCENT  = re.compile(r"^-?\(?\d+(\.\d+)?\)?%$")
_RE_INT      = re.compile(r"^-?\(?[\d,]+\)?$")

def _parse_numeric(s):
    """If s looks like a formatted number, return (float_value, excel_format).
    Else return (None, None).

    Special cases:
    - 4-digit years (1900-2100) → integer format with no comma ('0')
    - Percent strings (e.g. '5.75%') → decimal value with '0.00%' format
    - Currency strings → '$#,##0' (or '$#,##0.00' if decimal in source)
    """
    if s is None: return None, None
    if not isinstance(s, str): s = str(s)
    t = s.strip()
    if t in ("", "—", "-", "N/A", "n/a", "NA"): return None, None
    neg = t.startswith("(") and t.endswith(")")
    if neg: t = t[1:-1]

    # Percent
    if _RE_PERCENT.match(t):
        try:
            raw = float(t.rstrip("%").replace(",", ""))
            f = raw / 100
            # Use 2 decimals if the percent isn't a whole number
            fmt = '0.0%' if abs(raw - round(raw)) < 0.05 else '0.00%'
            return (-f if neg else f), fmt
        except: return None, None

    # Currency with $ sign
    if _RE_CURRENCY.match(t) and "$" in t:
        try:
            f = float(t.replace("$", "").replace(",", ""))
            has_dec = "." in t
            return (-f if neg else f), ('$#,##0.00' if has_dec else '$#,##0;($#,##0)')
        except: return None, None

    # Plain integer (no $, no %, no decimal)
    if _RE_INT.match(t):
        try:
            f = float(t.replace(",", ""))
            # Special: 4-digit years render without comma
            if 1900 <= f <= 2100 and f == int(f) and "," not in t:
                return (-f if neg else f), '0'
            return (-f if neg else f), '#,##0;(#,##0)'
        except: return None, None

    # Numeric with decimal but no $
    if _RE_CURRENCY.match(t):
        try:
            f = float(t.replace(",", ""))
            return (-f if neg else f), '#,##0.00;(#,##0.00)'
        except: return None, None

    return None, None


# ── Core cell writer ─────────────────────────────────────────────────────────
def _sc(ws, row, col, val, bold=False, bg=None, fg=None,
        size=9, ha="left", wrap=True, italic=False, num_fmt=None):
    """Set a cell. Auto-detects numeric strings and writes them as real numbers
    with a number format so clients can sort/filter/formula. Pass num_fmt to
    override detection. Pass val=None or '—' to render an em-dash placeholder."""
    c = ws.cell(row=row, column=col)
    final_fg = fg if fg else C_TEXT

    # Try numeric promotion
    if val is not None and val != "" and val != "—" and num_fmt is None:
        num_val, detected_fmt = _parse_numeric(val)
        if num_val is not None:
            c.value = num_val
            c.number_format = detected_fmt
            # Negative numbers in red (data rows only — not totals/headers)
            if num_val < 0 and not bold:
                final_fg = C_NEG
        else:
            c.value = val
    elif num_fmt is not None:
        c.value = val if val not in (None, "", "—") else None
        c.number_format = num_fmt
    else:
        c.value = val if val not in (None, "") else "—"

    c.font = Font(name="Calibri", bold=bold, color=final_fg, size=size, italic=italic)
    c.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    # Default: no border. Specific row helpers set bottom rules where needed.
    c.border = _NO_BORDER
    return c


def _fr(ws, row, ncols, bg):
    """Fill an entire row with a background color (no border)."""
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill("solid", fgColor=bg)


# ── Section header: navy bar, white text, generous height ────────────────────
def _sec(ws, row, title, n=14):
    _fr(ws, row, n, C_NAVY)
    c = _sc(ws, row, 1, title, bold=True, bg=C_NAVY, fg=C_WHITE,
            size=11, ha="left", wrap=False)
    # Merge across all cols so a long title doesn't wrap onto two visual lines
    try:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    except Exception:
        pass
    ws.row_dimensions[row].height = 26
    return row + 1


# ── Column header row: light grey band with navy bottom rule ────────────────
def _thdr(ws, row, headers, n=14):
    _fr(ws, row, n, C_SUB_FILL)
    for i, h in enumerate(headers):
        c = _sc(ws, row, 1 + i, h, bold=True, bg=C_SUB_FILL,
                fg=C_NAVY, size=9, ha="center")
        c.border = _THICK_BOTTOM
    ws.row_dimensions[row].height = 22
    return row + 1


# ── Sub-header inside a section (e.g. "Renovation Tiers") ───────────────────
def _shdr(ws, row, title, n=14):
    _fr(ws, row, n, C_WHITE)
    c = _sc(ws, row, 1, title.upper(), bold=True, bg=C_WHITE,
            fg=C_NAVY, size=9, ha="left", wrap=False)
    try:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    except Exception:
        pass
    # Letter-spacing effect via thin bottom rule
    for col in range(1, n + 1):
        ws.cell(row=row, column=col).border = _BOTTOM_BORDER
    ws.row_dimensions[row].height = 20
    return row + 1


# ── Key/value row: bold label left, value right ─────────────────────────────
def _kv(ws, row, label, value, alt=False, n=14):
    # Alt zebra is so subtle we drop it entirely for KV rows (cleaner)
    _fr(ws, row, n, C_WHITE)
    c1 = _sc(ws, row, 1, label, bold=True, fg=C_TEXT_MUTED, bg=C_WHITE,
             size=9, ha="left")
    c1.border = _BOTTOM_BORDER
    c2 = _sc(ws, row, 2, value if value not in (None, "") else "—",
             fg=C_TEXT, bg=C_WHITE, size=9, ha="left", wrap=True)
    c2.border = _BOTTOM_BORDER
    # Extend bottom rule across remaining cols for a clean line
    for col in range(3, n + 1):
        ws.cell(row=row, column=col).border = _BOTTOM_BORDER
    ws.row_dimensions[row].height = 18
    return row + 1


# ── Data row: subtle zebra, bottom rule, right-align numeric columns ────────
def _drow(ws, row, vals, alt=False, als=None, h=17, n=None, cs=1):
    bg = C_ROW_ALT if alt else C_WHITE
    nc = n or (cs + len(vals) - 1)
    _fr(ws, row, nc, bg)
    for i, v in enumerate(vals):
        ha = als[i] if als and i < len(als) else "left"
        c = _sc(ws, row, cs + i, v, bg=bg, fg=C_TEXT,
                size=9, ha=ha, wrap=(ha == "left"))
        c.border = _BOTTOM_BORDER
    # Also apply bottom rule on any unused columns for clean line continuation
    for col in range(cs + len(vals), nc + 1):
        ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=bg)
        ws.cell(row=row, column=col).border = _BOTTOM_BORDER
    ws.row_dimensions[row].height = h
    return row + 1


# ── Subtotal row: pale-blue fill, bold navy text, top rule ──────────────────
def _subtrow(ws, row, vals, n=14):
    _fr(ws, row, n, C_TOTAL_FILL)
    for i, v in enumerate(vals):
        ha = "left" if i == 0 else ("left" if i == len(vals) - 1 else "right")
        c = _sc(ws, row, 1 + i, v, bold=True, bg=C_TOTAL_FILL,
                fg=C_NAVY, size=9, ha=ha)
        c.border = _BOTTOM_BORDER
    ws.row_dimensions[row].height = 19
    return row + 1


# ── Total row: white fill, bold navy text, thick navy top rule ──────────────
def _totrow(ws, row, vals, n=14, col=None):
    _fr(ws, row, n, C_WHITE)
    for i, v in enumerate(vals):
        ha = "left" if i == 0 else ("left" if i == len(vals) - 1 else "right")
        c = _sc(ws, row, 1 + i, v, bold=True, bg=C_WHITE,
                fg=C_NAVY, size=10, ha=ha)
        c.border = Border(
            top=Side(style="medium", color=C_NAVY),
            bottom=Side(style="thin", color=C_NAVY)
        )
    ws.row_dimensions[row].height = 22
    return row + 1


# ── NOI / hero row: navy fill, white text — used sparingly for key totals ───
def _noirow(ws, row, vals, n=14):
    _fr(ws, row, n, C_NAVY)
    for i, v in enumerate(vals):
        ha = "left" if i == 0 else ("left" if i == len(vals) - 1 else "right")
        _sc(ws, row, 1 + i, v, bold=True, bg=C_NAVY,
            fg=C_WHITE, size=10, ha=ha)
    ws.row_dimensions[row].height = 24
    return row + 1


# ── Vertical spacer ─────────────────────────────────────────────────────────
def _sp(ws, row, h=10):
    ws.row_dimensions[row].height = h
    return row + 1


# ── Cover block: large title + thin navy underline + subtitle ───────────────
def _cover(ws, row, line1, line2, n=14):
    # Big title — merge across all columns
    _fr(ws, row, n, C_WHITE)
    _sc(ws, row, 1, line1, bold=True, bg=C_WHITE,
        fg=C_NAVY, size=18, ha="left", wrap=False)
    try: ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    except Exception: pass
    ws.row_dimensions[row].height = 32
    row += 1
    # Subtitle (the property/broker/date strip) — also merged
    _fr(ws, row, n, C_WHITE)
    _sc(ws, row, 1, line2, bold=False, bg=C_WHITE,
        fg=C_TEXT_MUTED, size=10, ha="left", wrap=False)
    try: ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    except Exception: pass
    ws.row_dimensions[row].height = 18
    row += 1
    # Thin navy rule across the page
    _fr(ws, row, n, C_WHITE)
    for col in range(1, n + 1):
        ws.cell(row=row, column=col).border = Border(
            top=Side(style="medium", color=C_NAVY)
        )
    ws.row_dimensions[row].height = 4
    row += 1
    return _sp(ws, row, h=14)


def _setup_sheet(ws, freeze_at="A6", tab_color=None,
                 fit_to_width=True, landscape=True, margins=True):
    """Pro polish per worksheet: no gridlines, tab color, print-fit-to-width,
    landscape orientation, and clean margins.
    NOTE: Frozen panes removed per user request — sheet scrolls normally."""
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = True
    # freeze_panes intentionally NOT set — full sheet scroll
    if tab_color:
        ws.sheet_properties.tabColor = tab_color
    if fit_to_width:
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    if landscape:
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    if margins:
        ws.page_margins.left = 0.4
        ws.page_margins.right = 0.4
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5
        ws.page_margins.header = 0.3
        ws.page_margins.footer = 0.3
    ws.print_options.horizontalCentered = True
    # Footer with filename + page numbers
    ws.oddFooter.left.text  = "&\"Calibri,Italic\"&8&K6B7280 Confidential — Internal Use Only"
    ws.oddFooter.right.text = "&\"Calibri,Regular\"&8&K6B7280 Page &P of &N"


def build_excel(d: dict, filename: str, sections: dict = None) -> bytes:
    S = sections or {}
    def _on(key): return S.get(key, True)
    date   = datetime.today().strftime("%B %d, %Y")
    prop   = (d.get("property") or {}).get("name") or "Property"
    broker = (d.get("broker") or {}).get("name") or "N/A"
    pd_    = d.get("property") or {}
    inv    = d.get("investment") or {}
    va     = d.get("value_add") or {}
    tax    = d.get("tax") or {}
    fin    = d.get("financials") or {}
    fin_i  = d.get("financing") or {}
    insur  = d.get("insurance") or {}
    mgmt   = d.get("management") or {}
    repl   = d.get("replacement_cost") or {}
    afford = d.get("affordability") or {}
    demo   = d.get("demographics") or {}
    mkt    = d.get("market") or {}
    umix   = d.get("unit_mix") or []
    utils_ = d.get("utilities") or []
    site   = d.get("site") or {}
    rg     = d.get("rent_comps_garden") or []
    rth    = d.get("rent_comps_townhouse") or []
    rcomps = d.get("rent_comps") or []
    scomps = d.get("sale_comps") or []
    flags  = d.get("flags") or []

    addr = f"{pd_.get('address','')}, {pd_.get('city','')}, {pd_.get('state','')} {pd_.get('zip','')}".strip(", ")
    subtitle = f"{prop}  ·  {addr}  ·  {_v(pd_.get('units'),'n')} Units  ·  {_v(pd_.get('year_built'))} Built  ·  {broker}  ·  {date}"

    wb = Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet("Financials")
    _setup_sheet(ws1, freeze_at="A6", tab_color="1E3A5F")
    _n_periods = len((d.get("financials") or {}).get("periods") or [])
    _data_w = 14 if _n_periods >= 6 else (16 if _n_periods >= 4 else 18)
    _notes_w = 30 if _n_periods >= 6 else 36
    _fin_cols = {"A": 32}
    for _i, _c in enumerate("BCDEFGHIJK"):
        if _i < _n_periods:
            _fin_cols[_c] = _data_w
        elif _i == _n_periods:
            _fin_cols[_c] = _notes_w
        else:
            _fin_cols[_c] = 12
    for col, w in _fin_cols.items():
        ws1.column_dimensions[col].width = w

    r = 1
    r = _cover(ws1, r, f"MULTIFAMILY UNDERWRITING REPORT  ·  FINANCIALS", subtitle)

    if _on('deal'):
        r = _sec(ws1, r, "A.  DEAL SUMMARY & PROPERTY DETAILS")
        agents = ", ".join(
            a.get("name", str(a)) if isinstance(a, dict) else str(a)
            for a in ((d.get("broker") or {}).get("agents") or [])
        ) or "N/A"
        for i, (k, v) in enumerate([
            ("Property Name",      prop),
            ("Address",            addr),
            ("County / MSA",       f"{pd_.get('county') or 'N/A'}  ·  {pd_.get('msa') or 'N/A'}"),
            ("Total Units",        _v(pd_.get("units"), "n")),
            ("Year Built",         _v(pd_.get("year_built"))),
            ("Net Rentable SF",    f"{_v(pd_.get('rentable_sf'), 'n')} SF  (Avg {_v(pd_.get('avg_unit_sf'), 'n')} SF/unit)"),
            ("Buildings / Stories",f"{_v(pd_.get('buildings'), 'n')} buildings  ·  {_v(pd_.get('floors'))} stories"),
            ("Land Area",          f"{_v(pd_.get('acres'))} acres  ({_v(pd_.get('density'))} units/acre)"),
            ("Asset Class",        pd_.get("asset_class")),
            ("Occupancy",          f"{_v(pd_.get('occupancy_pct'), '%')}  as of {pd_.get('occupancy_date') or 'N/A'}"),
            ("Offering Type",      fin_i.get("offering_type")),
            ("Broker / Agents",    f"{broker}  ·  {agents}"),
            ("Debt Contact",       fin_i.get("debt_contact")),
        ]):
            r = _kv(ws1, r, k, v, alt=bool(i % 2))
        r = _sp(ws1, r)

    if _on('unitmix'):
        r = _sec(ws1, r, "B.  UNIT MIX")
        if umix:
            r = _thdr(ws1, r, ["Unit Type","Plan","Units","Mix %","SF / Unit",
                                "In-Place Rent","In-Place PSF","Mkt Rent","Mkt PSF","Target Rent","Upside / Unit"])
            ra = ["left","left","center","center","right","right","right","right","right","right","right"]
            # Compute Mix % directly from count/total to guarantee totals = 100%.
            # The AI-returned "pct" field can be unreliable.
            _total_units = sum(int(u.get("count") or 0) for u in umix) or 1
            _total_sf_x_units = 0
            _running_units = 0
            for i, u in enumerate(umix):
                _cnt = int(u.get("count") or 0)
                _running_units += _cnt
                _mix_pct = (_cnt / _total_units) * 100 if _total_units else 0
                _sf_val = u.get("sf") or 0
                try: _total_sf_x_units += float(_sf_val) * _cnt
                except: pass
                r = _drow(ws1, r, [
                    _v(u.get("type","—")), u.get("plan") or "—", _v(u.get("count"),"n"),
                    f"{_mix_pct:.1f}%", _v(u.get("sf"),"n"),
                    _v(u.get("market_rent"),"$"), _psf(u.get('market_psf')),
                    _v(u.get("eff_rent"),"$"), _psf(u.get('eff_psf')),
                    _v(u.get("target_rent"),"$"), _v(u.get("upside"),"$"),
                ], alt=bool(i % 2), als=ra)
            # TOTAL row — verifies units sum and shows weighted-avg SF
            _avg_sf = _total_sf_x_units / _total_units if _total_units else 0
            r = _totrow(ws1, r, [
                "TOTAL / AVG", "",
                f"{_total_units}", "100.0%",
                f"{int(_avg_sf):,}",
                "", "", "", "", "", ""
            ])
        else:
            r = _kv(ws1, r, "Note", "No unit mix data found.")
        r = _sp(ws1, r)

        periods   = fin.get("periods") or []
    if _on('opstat'):
        inc_lines = fin.get("income_lines") or []
        exp_lines = fin.get("expense_lines") or []

        if periods and (inc_lines or exp_lines):
            p_all = periods[:7]
            r = _sec(ws1, r, f"C.  OPERATING STATEMENT  ({'  ·  '.join(p_all)})")
            r = _thdr(ws1, r, ["Line Item"] + p_all + ["Underwriting Notes"])
            f6 = ["left"] + ["right"] * len(p_all) + ["left"]

            def _fin_row(item):
                vals = item.get("values") or {}
                return [item.get("item", "—")] + [
                    _v(vals.get(p), "$") if vals.get(p) is not None else "—" for p in p_all
                ] + [item.get("note") or ""]

            rendered_totals = set()

            r = _shdr(ws1, r, "INCOME")
            for i, item in enumerate(inc_lines):
                name = (item.get("item") or "").lower()
                if item.get("is_total"):
                    rendered_totals.add(name)
                    if "before reserve" in name or "pre reserve" in name:
                        r = _subtrow(ws1, r, _fin_row(item))
                    else:
                        r = _noirow(ws1, r, _fin_row(item))
                elif item.get("is_subtotal"):
                    r = _subtrow(ws1, r, _fin_row(item))
                else:
                    r = _drow(ws1, r, _fin_row(item), alt=bool(i % 2), als=f6)

            r = _shdr(ws1, r, "EXPENSES")
            for i, item in enumerate(exp_lines):
                name = (item.get("item") or "").lower()
                if item.get("is_total"):
                    rendered_totals.add(name)
                    if "before reserve" in name or "pre reserve" in name:
                        r = _subtrow(ws1, r, _fin_row(item))
                    else:
                        r = _totrow(ws1, r, _fin_row(item))
                elif item.get("is_subtotal"):
                    r = _subtrow(ws1, r, _fin_row(item))
                else:
                    r = _drow(ws1, r, _fin_row(item), alt=bool(i % 2), als=f6)

            noi_d = fin.get("noi") or {}
            noi_already = any("net operating income" in t for t in rendered_totals)
            if not noi_already and any(_is_num(noi_d.get(p)) for p in p_all):
                r = _noirow(ws1, r, ["NET OPERATING INCOME"] + [
                    _v(noi_d.get(p), "$") if _is_num(noi_d.get(p)) else "—" for p in p_all
                ] + [""])

            capex_d = fin.get("capex") or {}
            if any(_is_num(capex_d.get(p)) for p in p_all):
                r = _drow(ws1, r, ["  Capital Reserves"] + [
                    _v(capex_d.get(p), "$") if _is_num(capex_d.get(p)) else "—" for p in p_all
                ] + [""], als=f6)

            cffo_d = fin.get("cffo") or {}
            if any(_is_num(cffo_d.get(p)) for p in p_all):
                r = _noirow(ws1, r, ["CASH FLOW FROM OPERATIONS"] + [
                    _v(cffo_d.get(p), "$") if _is_num(cffo_d.get(p)) else "—" for p in p_all
                ] + [""])
        else:
            r = _sec(ws1, r, "C.  OPERATING STATEMENT")
            r = _kv(ws1, r, "Note", "No financial data found in OM.")
        r = _sp(ws1, r)
    if _on('valueadd'):

        r = _sec(ws1, r, "D.  PROPOSED VALUE-ADD BY FLOOR PLAN")
        plans = va.get("by_floor_plan") or []
        if plans:
            r = _thdr(ws1, r, ["Unit Type","Units","SF","In-Place Rent","In-Place PSF",
                                "Rehab Cost / Unit","Premium / Unit","Post-Rehab Rent","Post-Rehab PSF"])
            ra2 = ["left","center","right","right","right","right","right","right","right"]
            for i, p in enumerate(plans):
                r = _drow(ws1, r, [
                    p.get("type","—"), _v(p.get("units"),"n"), _v(p.get("sf"),"n"),
                    _v(p.get("inplace_rent"),"$"), _psf(p.get('inplace_psf')),
                    _v(p.get("rehab_cost"),"$"), _v(p.get("premium"),"$"),
                    _v(p.get("post_rehab_rent"),"$"), _psf(p.get('post_rehab_psf')),
                ], alt=bool(i % 2), als=ra2)

        tiers = inv.get("renovation_tiers") or []
        if tiers:
            r = _sp(ws1, r)
            r = _shdr(ws1, r, "Renovation Tiers")
            r = _thdr(ws1, r, ["Tier","Units","Appliances","Cabinets","Countertops","Flooring","Premium"])
            for i, t in enumerate(tiers):
                r = _drow(ws1, r, [
                    t.get("name","—"), _v(t.get("units"),"n"),
                    t.get("appliances") or "—", t.get("cabinets") or "—",
                    t.get("countertops") or "—", t.get("flooring") or "—",
                    _v(t.get("premium"),"$") if t.get("premium") else "—",
                ], alt=bool(i % 2))

        light = inv.get("light_upgrade_items") or va.get("light_upgrade_items") or []
        if light:
            r = _sp(ws1, r)
            r = _shdr(ws1, r, "Interior Upgrade Scope")
            for i, item in enumerate(light):
                bg = C_ALT if bool(i % 2) else C_WHITE
                _fr(ws1, r, 14, bg)
                _sc(ws1, r, 1, f"{i+1}.  {item}", bg=bg, fg=C_BODY, size=9)
                ws1.row_dimensions[r].height = 15; r += 1

        r = _sp(ws1, r)
        r = _shdr(ws1, r, "CapEx & ROI Summary")
        for i, (k, v) in enumerate([
            ("Total Renovation Cost",       _v(va.get("total_cost"), "$")),
            ("Cost Per Unit (all-in)",      _v(va.get("cost_per_unit"), "$")),
            ("Exterior / Additional CapEx", _v(va.get("exterior_capex"), "$")),
            ("Monthly Rent Premium",        _v(va.get("monthly_premium"), "$")),
            ("Annual Rent Premium",         _v(va.get("annual_premium"), "$")),
            ("Return on Investment",        _pct(va.get('roi_pct'))),
            ("Value-Add Scope",             va.get("scope")),
        ]):
            r = _kv(ws1, r, k, v, alt=bool(i % 2))

        levers = d.get("value_add_levers") or []
        if levers:
            r = _sp(ws1, r)
            r = _shdr(ws1, r, "Revenue Upside Levers")
            r = _thdr(ws1, r, ["Value-Add Lever", "Units", "Mo. Premium / Unit", "Annual Upside", "Notes"])
            total_auto = 0
            for i, lv in enumerate(levers):
                is_total = "TOTAL" in (lv.get("lever") or "").upper()
                ann = lv.get("annual_upside")
                try:
                    if not is_total:
                        total_auto += float(ann or 0)
                except Exception:
                    pass
                row_data = [
                    lv.get("lever") or "—",
                    _v(lv.get("units"), "n") if not is_total else "",
                    _v(lv.get("monthly_premium"), "$") if not is_total else "",
                    _v(ann, "$"),
                    lv.get("notes") or "",
                ]
                if is_total:
                    r = _noirow(ws1, r, row_data)
                else:
                    r = _drow(ws1, r, row_data, alt=bool(i % 2),
                              als=["left", "center", "right", "right", "left"])
            if levers and "TOTAL" not in (levers[-1].get("lever") or "").upper() and total_auto > 0:
                r = _noirow(ws1, r, ["TOTAL ANNUAL REVENUE UPSIDE", "", "", _v(total_auto, "$"), ""])

        r = _sp(ws1, r)
    if _on('financing'):

        r = _sec(ws1, r, "E.  FINANCING & DEBT TERMS")
        r = _shdr(ws1, r, "Offering & Contact")
        for i, (k, v) in enumerate([
            ("Offering Type", fin_i.get("offering_type")),
            ("Debt Contact",  fin_i.get("debt_contact")),
            ("Notes",         fin_i.get("notes")),
        ]):
            r = _kv(ws1, r, k, v, alt=bool(i % 2))

        nf  = fin_i.get("new_financing") or {}
        asd = fin_i.get("assumable_debt") or {}

        r = _shdr(ws1, r, "New Financing")
        r = _thdr(ws1, r, ["Loan Type","Lender","Loan Amount","LTV","Interest Rate",
                            "Rate Type","Loan Term","Amortization","Interest-Only","DSCR","Recourse","Notes"])
        if any(nf.get(k) for k in ["loan_type","lender","loan_to_value","interest_rate"]):
            r = _drow(ws1, r, [
                nf.get("loan_type") or "—", nf.get("lender") or "—",
                _v(nf.get("loan_amount"), "$"), _pct(nf.get('loan_to_value')),
                _pct(nf.get('interest_rate')), nf.get("rate_type") or "—",
                nf.get("loan_term_years") or "—", nf.get("amortization_years") or "—",
                nf.get("interest_only_period") or "—", nf.get("dscr") or "—",
                nf.get("recourse") or "—", nf.get("notes") or "—",
            ], als=["left","left","right","right","right","center","center","center","center","center","center","left"])
        else:
            _fr(ws1, r, 14, C_ALT)
            offering = (fin_i.get("offering_type") or "").lower()
            if "all cash" in offering:
                msg = "All Cash offering — no financing available. Buyer must close without debt."
            else:
                msg = "No new financing terms provided in OM. Contact debt broker for quote."
            _sc(ws1, r, 1, msg, bg=C_ALT, fg=C_BODY, size=9, italic=True)
            ws1.row_dimensions[r].height = 15; r += 1

        r = _shdr(ws1, r, "Assumable Debt")
        r = _thdr(ws1, r, ["Loan Type","Lender","Loan Amount","LTV","Interest Rate","Rate Type",
                            "Loan Term","Amortization","Interest-Only","Origination","Maturity","Monthly Pmt","Annual DS","DSCR","Recourse"])
        if any(asd.get(k) for k in ["loan_type","lender","loan_to_value","interest_rate"]):
            r = _drow(ws1, r, [
                asd.get("loan_type") or "—", asd.get("lender") or "—",
                _v(asd.get("loan_amount"), "$"), _pct(asd.get("loan_to_value")),
                _pct(asd.get("interest_rate")), asd.get("rate_type") or "—",
                asd.get("loan_term_years") or "—", asd.get("amortization_years") or "—",
                asd.get("interest_only_period") or "—",
                asd.get("origination_date") or "—", asd.get("maturity_date") or "—",
                _v(asd.get("monthly_payment"), "$"), _v(asd.get("annual_debt_service"), "$"),
                asd.get("dscr") or "—", asd.get("recourse") or "—",
            ], als=["left","left","right","right","right","center","center","center","center","center","center","right","right","center","center"])
        else:
            _fr(ws1, r, 14, C_ALT)
            offering = (fin_i.get("offering_type") or "").lower()
            if "all cash" in offering:
                msg = "All Cash offering — no assumable debt. No existing financing on the property."
            else:
                msg = "None — property offered free and clear. No existing assumable debt."
            _sc(ws1, r, 1, msg, bg=C_ALT, fg=C_BODY, size=9, italic=True)
            ws1.row_dimensions[r].height = 15; r += 1
        r = _sp(ws1, r)
    # NOTE: Underwriting Flags moved to dedicated "Flags" tab (ws4) for readability
    if _on('tax'):

        r = _sec(ws1, r, "G.  PROPERTY TAX & TAX ABATEMENT")
        r = _shdr(ws1, r, "Property Tax Detail")
        for i, (k, v) in enumerate([
            ("Parcel ID",             tax.get("parcel_id")),
            ("Assessed Market Value", _v(tax.get("assessed_value"), "$")),
            ("Millage — City",        tax.get("millage_city")),
            ("Millage — County",      tax.get("millage_county")),
            ("Total Millage Rate",    tax.get("millage_total")),
            ("Ad Valorem Tax",        _v(tax.get("tax_base"), "$")),
            ("Solid Waste / Fees",    _v(tax.get("solid_waste_fee"), "$")),
            ("Total Annual Tax Bill", _v(tax.get("total_tax"), "$")),
        ]):
            r = _kv(ws1, r, k, v, alt=bool(i % 2))
        r = _shdr(ws1, r, "Tax Abatement Program")
        for i, (k, v) in enumerate([
            ("Program",               tax.get("abatement_program")),
            ("Abatement %",           _pct(tax.get('abatement_pct'))),
            ("Commitment Term",       tax.get("abatement_term_note")),
            ("AMI Requirement",       f"{_pct(tax.get('ami_pct'))} of Area Median Income"),
            ("Annual Tax Savings",    _v(tax.get("abatement_annual_savings"), "$")),
            ("Max Allowable Rent",    _v(tax.get("max_allowable_rent"), "$")),
            ("Avg In-Place Rent",     _v(tax.get("avg_inplace_rent"), "$")),
            ("Headroom / Unit",       _v(tax.get("rent_headroom"), "$")),
        ]):
            r = _kv(ws1, r, k, v, alt=bool(i % 2))
        r = _sp(ws1, r)
    if _on('repl'):

        r = _sec(ws1, r, "H.  REPLACEMENT COST  ·  INSURANCE  ·  MANAGEMENT")
        r = _shdr(ws1, r, "Replacement Cost")
        has_detail = any(repl.get(k) for k in ["hard_cost_per_sf","land_per_unit","gross_replacement_per_unit"])
        if has_detail:
            repl_rows = [
                ("Component",                "Per Unit",                                           "Per SF",                                               "Total"),
                ("Land",                     _v(repl.get("land_per_unit"),"$"),                   "—",                                                    _v(repl.get("land_total"),"$")),
                ("Hard Costs",               _v(repl.get("hard_cost_per_unit"),"$"),              _v(repl.get("hard_cost_per_sf"),"$"),                   _v(repl.get("hard_cost_total"),"$")),
                ("Soft Costs",               _v(repl.get("soft_cost_per_unit"),"$"),              f"{repl.get('soft_cost_pct') or '—'}% of hard",         _v(repl.get("soft_cost_total"),"$")),
            ]
            r = _thdr(ws1, r, ["Component","Per Unit","Per SF","Total Cost"])
            for i, row in enumerate(repl_rows[1:]):
                r = _drow(ws1, r, list(row), alt=bool(i%2), als=["left","right","right","right"])
            if repl.get("direct_replacement_per_unit"):
                r = _subtrow(ws1, r, ["Direct Replacement Cost", _v(repl.get("direct_replacement_per_unit"),"$"), _v(repl.get("direct_replacement_per_sf"),"$"), _v(repl.get("direct_replacement_total"),"$")])
            dev_rows = []
            if repl.get("developer_fee_per_unit"):
                dev_rows.append(("Developer Fee", _v(repl.get("developer_fee_per_unit"),"$"), f"{repl.get('developer_fee_pct') or '—'}% of project", _v(repl.get("developer_fee_total"),"$")))
            if repl.get("gc_fee_per_unit"):
                dev_rows.append(("GC Fee",        _v(repl.get("gc_fee_per_unit"),"$"), f"{repl.get('gc_fee_pct') or '—'}% of hard costs", _v(repl.get("gc_fee_total"),"$")))
            for i, row in enumerate(dev_rows):
                r = _drow(ws1, r, list(row), alt=bool(i%2), als=["left","right","right","right"])
            if repl.get("gross_replacement_per_unit"):
                r = _noirow(ws1, r, ["Gross Replacement Cost", _v(repl.get("gross_replacement_per_unit"),"$"), _v(repl.get("gross_replacement_per_sf"),"$"), _v(repl.get("gross_replacement_total"),"$")])
            if repl.get("notes"):
                r = _kv(ws1, r, "Notes", repl.get("notes"))
        else:
            for i, (k, v) in enumerate([
                ("Cost Per Unit",     _v(repl.get("per_unit") or repl.get("gross_replacement_per_unit"), "$")),
                ("Cost Per SF",       _v(repl.get("per_sf") or repl.get("gross_replacement_per_sf"), "$")),
                ("Total Replacement", _v(repl.get("total") or repl.get("gross_replacement_total"), "$")),
                ("Land Per Unit",     _v(repl.get("land_per_unit"), "$")),
                ("Hard Cost Per SF",  _v(repl.get("hard_cost_per_sf"), "$")),
                ("Soft Cost %",       _pct(repl.get('soft_cost_pct'))),
                ("Source",            repl.get("source")),
                ("Notes",             repl.get("notes")),
            ]):
                r = _kv(ws1, r, k, v, alt=bool(i % 2))
        r = _shdr(ws1, r, "Insurance")
        for i, (k, v) in enumerate([
            ("Carrier / Provider", insur.get("carrier")),
            ("Annual Premium",     _v(insur.get("annual_premium"), "$")),
            ("Per Unit / Year",    _v(insur.get("per_unit"), "$")),
            ("Quote Source",       insur.get("quote_source")),
            ("Notes",              insur.get("notes")),
        ]):
            r = _kv(ws1, r, k, v, alt=bool(i % 2))
        r = _shdr(ws1, r, "Property Management")
        for i, (k, v) in enumerate([
            ("Management Fee %",   f"{_pct(mgmt.get('fee_pct'))} of EGI"),
            ("Annual Fee",         _v(mgmt.get("fee_annual"), "$")),
            ("Per Unit / Year",    _v(mgmt.get("fee_per_unit"), "$")),
            ("Current Manager",    mgmt.get("current_manager")),
            ("Proposed Manager",   mgmt.get("proposed_manager")),
            ("Notes",              mgmt.get("notes")),
        ]):
            r = _kv(ws1, r, k, v, alt=bool(i % 2))
        r = _sp(ws1, r)

        _fr(ws1, r, 14, C_ALT)
        _sc(ws1, r, 1, "AI-generated from broker OM. Internal use only. Verify all figures independently. Powered by Anthropic Claude.",
            bg=C_ALT, fg="FF888880", size=8, italic=True)
        ws1.row_dimensions[r].height = 13

    ws2 = wb.create_sheet("Comparables")
    _setup_sheet(ws2, freeze_at="A6", tab_color="1E3A5F")
    for col, w in {"A":30,"B":14,"C":10,"D":10,"E":10,"F":12,
                   "G":12,"H":12,"I":12,"J":10,"K":32}.items():
        ws2.column_dimensions[col].width = w

    r = 1
    r = _cover(ws2, r, f"COMPARABLES — RENT & SALE  |  {prop}", subtitle, n=11)

    ra_c = ["left","center","center","center","center","center","right","right","right","right","left"]
    def _subj_rent_for(comp_type):
        ct = (comp_type or "").lower()
        matching = [u for u in umix if ct in (u.get("plan") or u.get("type") or "").lower()]
        pool = matching if matching else umix
        rents = [float(u["eff_rent"]) for u in pool if u.get("eff_rent") and _is_num(u.get("eff_rent"))]
        return int(sum(rents) / len(rents)) if rents else 0

    def _render_comp_group(ws, r, comps, section_title, comp_type_label, n=11):
        subj_rent = _subj_rent_for(comp_type_label)
        r = _sec(ws, r, section_title, n=n)
        r = _thdr(ws, r, ["Property","Type","Yr Built","Distance","Units",
                           "Occ %","Market Rent","Eff Rent","Rent PSF","vs. Subject","Notes"], n=n)
        for i, c in enumerate(comps):
            rent_val = c.get("rent") or c.get("total_market") or c.get("total_eff")
            try: rent_num = float(str(rent_val).replace("$","").replace(",",""))
            except: rent_num = 0
            vs = f"{int(rent_num - subj_rent):+,}" if rent_num and subj_rent else "—"
            name = c.get("name","—")
            is_highlight = any(k in name.lower() for k in ("subject","average","avg","index"))
            row_data = [name, c.get("comp_type") or comp_type_label,
                        _v(c.get("year_built")), c.get("distance") or "—",
                        _v(c.get("units"),"n") if c.get("units") else "—",
                        _pct(c.get("occupancy")) if c.get("occupancy") else "—",
                        _v(rent_val,"$"), _v(c.get("total_eff") or rent_val,"$"),
                        _psf(c.get("total_market_psf") or c.get("total_eff_psf")),
                        vs, c.get("notes") or "—"]
            if is_highlight:
                r = _totrow(ws, r, row_data, n=n)
            else:
                r = _drow(ws, r, row_data, alt=bool(i % 2), als=ra_c, n=n)
        r = _sp(ws, r)
        return r

    if _on('rentcomps'):
        if rg:
            r = _render_comp_group(ws2, r, rg, "A.  GARDEN RENT COMPARABLES", "Garden")

    if _on('rentcomps'):
        if rth:
            sec_letter = "B" if rg else "A"
            r = _render_comp_group(ws2, r, rth, f"{sec_letter}.  TOWNHOUSE RENT COMPARABLES", "Townhouse")

    if _on('rentcomps'):
        if rcomps:
            sec_full = "A" if (not rg and not rth) else ("C" if (rg and rth) else "B")
            r = _sec(ws2, r, f"{sec_full}.  FULL COMPARABLE DETAIL", n=11)
            r = _thdr(ws2, r, ["#","Property","Type","Yr Built","Distance","Units",
                                "Occ %","Mkt Rent","Eff Rent","Avg SF","Notes"], n=11)
            for i, rc in enumerate(rcomps):
                r = _drow(ws2, r, [
                    rc.get("id","—"), rc.get("name","—"), rc.get("comp_type") or "—",
                    _v(rc.get("year_built")), rc.get("distance") or "—",
                    _v(rc.get("units"),"n"), _pct(rc.get("occupancy")),
                    _v(rc.get("total_market"),"$"), _v(rc.get("total_eff"),"$"),
                    _v(rc.get("avg_sf"),"n"), "—",
                ], alt=bool(i % 2), als=["center","left","center","center","center","center",
                                          "center","right","right","right","left"], n=11)
            r = _sp(ws2, r)

        if not rg and not rth and not rcomps:
            r = _sec(ws2, r, "A.  RENT COMPARABLES", n=11)
            r = _kv(ws2, r, "Note", "No rent comparable data found in OM. Source from listing broker or CoStar.", n=11)
            r = _sp(ws2, r)

    if _on('salecomps'):
        r = _sec(ws2, r, "D.  SALE COMPARABLES", n=11)
        if scomps:
            r = _thdr(ws2, r, ["Property","Date","Yr Built","Units","Sale Price",
                                "$/Unit","$/SF","Cap Rate","Occ","Buyer","Seller"], n=11)
            for i, sc in enumerate(scomps):
                r = _drow(ws2, r, [
                    sc.get("name","—"), sc.get("date") or "—", _v(sc.get("year_built")),
                    _v(sc.get("units"),"n"), _v(sc.get("price"),"$"),
                    _v(sc.get("ppu"),"$"), _v(sc.get("ppsf"),"$"),
                    sc.get("cap_rate") or "—", sc.get("occupancy") or "—",
                    sc.get("buyer") or "—", sc.get("seller") or "—",
                ], alt=bool(i % 2), n=11,
                   als=["left","center","center","center","right","right","right","center","center","left","left"])
        else:
            _fr(ws2, r, 11, C_ALT)
            _sc(ws2, r, 1, "Sale comps not provided in OM. Source from CoStar, RCA, or listing broker.",
                bg=C_ALT, fg=C_BODY, size=9, italic=True, wrap=True)
            ws2.row_dimensions[r].height = 20; r += 1
        r = _sp(ws2, r)

    _fr(ws2, r, 11, C_ALT)
    _sc(ws2, r, 1, "AI-generated. Internal use only. Verify all figures independently. Powered by Anthropic Claude.",
        bg=C_ALT, fg="FF888880", size=8, italic=True)
    ws2.row_dimensions[r].height = 13

    ws3 = wb.create_sheet("Demographics")
    _setup_sheet(ws3, freeze_at="A6", tab_color="1E3A5F")
    for col, w in {"A":36,"B":20,"C":20,"D":20,"E":14,"F":42}.items():
        ws3.column_dimensions[col].width = w

    r = 1
    r = _cover(ws3, r, f"DEMOGRAPHICS & MARKET OVERVIEW  |  {prop}", subtitle, n=6)

    if _on('addinc'):
        add_inc = inv.get("additional_income") or []
        if add_inc:
            r = _sec(ws3, r, "A.  ADDITIONAL INCOME", n=6)
            r = _thdr(ws3, r, ["Income Item","Category","Fee/Unit/Mo","Occ %",
                                "Monthly Income","Current Annual","Pro Forma Annual","Calculation Detail"], n=6)
            for i, inc in enumerate(add_inc):
                r = _drow(ws3, r, [
                    inc.get("name") or "—", inc.get("category") or "—",
                    _v(inc.get("fee_per_unit_per_month"), "$") if inc.get("fee_per_unit_per_month") else "—",
                    inc.get("occupancy_assumption") or "—",
                    _v(inc.get("monthly_income"), "$") if inc.get("monthly_income") else "—",
                    _v(inc.get("current_annual"), "$") if inc.get("current_annual") else "—",
                    _v(inc.get("proforma_annual"), "$") if inc.get("proforma_annual") else "—",
                    inc.get("calculation_detail") or inc.get("notes") or "—",
                ], alt=bool(i % 2), n=6,
                   als=["left","left","right","center","right","right","right","left"])
            r = _sp(ws3, r)

    if _on('utilities'):
        r = _sec(ws3, r, "B.  UTILITY INFORMATION", n=6)
        if utils_:
            r = _thdr(ws3, r, ["Utility","Billing Method","Paid By","Reimbursement","Annual Income","Notes"], n=6)
            for i, u in enumerate(utils_):
                r = _drow(ws3, r, [
                    u.get("name","—"), u.get("method") or "—", u.get("paid_by") or "—",
                    u.get("reimbursement") or "N/A",
                    _v(u.get("annual_income"), "$"), u.get("notes") or "—",
                ], alt=bool(i % 2), als=["left","center","center","left","right","left"], n=6)
        else:
            r = _kv(ws3, r, "Note", "No utility data found in OM.", n=6)
        r = _sp(ws3, r)

    if _on('pop'):
        r = _sec(ws3, r, "C.  POPULATION & INCOME DEMOGRAPHICS", n=6)
        has_demo = any(demo.get(k) for k in ["pop_1mi","pop_3mi","pop_5mi","median_income_3mi"])
        if not has_demo:
            _fr(ws3, r, 6, C_WARN)
            _sc(ws3, r, 1, "Detailed 1-mile / 3-mile / 5-mile radius demographic data not provided in this OM. Source independently from CoStar, Esri, or census.gov.",
                bg=C_WARN, fg=C_LABEL, size=9, bold=True, wrap=True)
            ws3.row_dimensions[r].height = 30; r += 1; r = _sp(ws3, r)

        r = _thdr(ws3, r, ["Metric","1-Mile Radius","3-Mile Radius","5-Mile Radius","Notes"], n=6)
        for i, (metric, v1, v3, v5, note) in enumerate([
            ("Population (2025)",        _v(demo.get("pop_1mi"),"n"),  _v(demo.get("pop_3mi"),"n"),  _v(demo.get("pop_5mi"),"n"),  ""),
            ("Population (2030 Proj.)",  _v(demo.get("pop_2030_1mi"),"n"), _v(demo.get("pop_2030_3mi"),"n"), _v(demo.get("pop_2030_5mi"),"n"), ""),
            ("Population Growth (5-yr)", demo.get("pop_growth_1mi") or "—", demo.get("pop_growth_3mi") or "—", demo.get("pop_growth_5mi") or "—", ""),
            ("Median HH Income (2025)",  _v(demo.get("median_income_1mi"),"$"), _v(demo.get("median_income_3mi"),"$"), _v(demo.get("median_income_5mi"),"$"), ""),
            ("Median HH Income (2030)",  _v(demo.get("median_income_2030_1mi"),"$"), _v(demo.get("median_income_2030_3mi"),"$"), _v(demo.get("median_income_2030_5mi"),"$"), ""),
            ("Income Growth (5-yr)",     demo.get("income_growth_1mi") or "—", demo.get("income_growth_3mi") or "—", demo.get("income_growth_5mi") or "—", ""),
            ("Renter-Occupied Units",    demo.get("renter_pct_1mi") or "—", demo.get("renter_pct_3mi") or "—", demo.get("renter_pct_5mi") or "—", ""),
            ("Bachelor's Degree+",       demo.get("college_pct_1mi") or "—", demo.get("college_pct_3mi") or "—", demo.get("college_pct_5mi") or "—", ""),
            ("White-Collar Workers",     demo.get("white_collar_pct_1mi") or "—", demo.get("white_collar_pct_3mi") or "—", demo.get("white_collar_pct_5mi") or "—", ""),
            ("Avg Home Value",           "—", _v(demo.get("home_value"),"$"), "—", ""),
        ]):
            bg = C_ALT if bool(i % 2) else C_WHITE
            _fr(ws3, r, 6, bg)
            _sc(ws3, r, 1, metric, bold=True, fg=C_LABEL, bg=bg, size=9)
            _sc(ws3, r, 2, v1, fg=C_BLUE_IN, bg=bg, size=9, ha="right")
            _sc(ws3, r, 3, v3, fg=C_BLUE_IN, bg=bg, size=9, ha="right")
            _sc(ws3, r, 4, v5, fg=C_BLUE_IN, bg=bg, size=9, ha="right")
            _sc(ws3, r, 5, note, fg=C_BODY, bg=bg, size=9, wrap=True)
            ws3.cell(row=r, column=6).fill = PatternFill("solid", fgColor=bg)
            ws3.row_dimensions[r].height = 16; r += 1
        r = _sp(ws3, r)

    if _on('afford'):
        r = _sec(ws3, r, "D.  AFFORDABILITY & RENT GROWTH RUNWAY", n=6)
        r = _thdr(ws3, r, ["Metric","2025 (3-Mile)","2030 Proj. (3-Mile)","Notes"], n=6)
        for i, (metric, v25, v30, note) in enumerate([
            ("Current In-Place Rent",          _v(afford.get("current_rent"),"$"), "—", "Subject effective rent"),
            ("Avg HH Income",                  _v(afford.get("avg_hh_income_3mi"),"$"), _v(afford.get("avg_hh_income_2030_3mi"),"$"), ""),
            ("Monthly Affordability (3× rule)",_v(afford.get("monthly_affordability_3x"),"$"), _v(afford.get("monthly_affordability_2030_3x"),"$"), "Income ÷ 12 ÷ 3"),
            ("Rent Headroom",                  _v(afford.get("rent_headroom_3mi"),"$"), _v(afford.get("rent_headroom_2030_3mi"),"$"), "Threshold minus current rent"),
            ("Income-to-Rent Ratio",           afford.get("income_to_rent_ratio") or "—", "—", ""),
        ]):
            bg = C_ALT if bool(i % 2) else C_WHITE
            _fr(ws3, r, 6, bg)
            _sc(ws3, r, 1, metric, bold=True, fg=C_LABEL, bg=bg, size=9)
            _sc(ws3, r, 2, v25, fg=C_BLUE_IN, bg=bg, size=9, ha="right")
            _sc(ws3, r, 3, v30, fg=C_BLUE_IN, bg=bg, size=9, ha="right")
            _sc(ws3, r, 4, note, fg=C_BODY, bg=bg, size=9)
            ws3.cell(row=r, column=5).fill = PatternFill("solid", fgColor=bg)
            ws3.cell(row=r, column=6).fill = PatternFill("solid", fgColor=bg)
            ws3.row_dimensions[r].height = 16; r += 1
        r = _sp(ws3, r)

    if _on('schools'):
        r = _sec(ws3, r, "E.  SCHOOLS, CRIME & QUALITY OF LIFE", n=6)
        r = _shdr(ws3, r, "Assigned Schools  (source: greatschools.org)", n=6)
        for i, (k, v) in enumerate([
            ("School District", demo.get("school_district") or "Not provided — verify via district map"),
            ("Elementary",      demo.get("elementary") or "Not provided — verify via district map"),
            ("Middle School",   demo.get("middle") or "Not provided — verify via district map"),
            ("High School",     demo.get("high_school") or "Not provided — verify via district map"),
        ]):
            r = _kv(ws3, r, k, v, alt=bool(i % 2), n=6)
        r = _shdr(ws3, r, "Crime  (source: crimegrade.org)", n=6)
        r = _kv(ws3, r, "Crime Data", demo.get("crime") or "Not provided in OM — source independently at crimegrade.org", n=6)
        r = _sp(ws3, r)

    if _on('utilities'):
        r = _sec(ws3, r, "F.  SITE & CONSTRUCTION INFORMATION", n=6)
        r = _shdr(ws3, r, "Physical Plant", n=6)
        for i, (k, v) in enumerate([
            ("Roof / Age",     f"{site.get('roof') or 'N/A'}  —  {site.get('roof_age') or 'N/A'}"),
            ("Exterior",       site.get("exterior")),
            ("Foundation",     site.get("foundation")),
            ("HVAC",           site.get("hvac")),
            ("Plumbing",       site.get("plumbing")),
            ("Wiring",         site.get("wiring")),
            ("Hot Water",      site.get("hot_water")),
            ("Washer / Dryer", site.get("washer_dryer")),
            ("Life Safety",    site.get("life_safety") or "Verify — not specified"),
            ("Construction",   site.get("notes")),
        ]):
            if v: r = _kv(ws3, r, k, v, alt=bool(i % 2), n=6)
        r = _shdr(ws3, r, "Parking & Site Features", n=6)
        for i, (k, v) in enumerate([
            ("Open Spaces",  _v(site.get("parking_open"), "n")),
            ("Covered",      _v(site.get("parking_covered"), "n")),
            ("Garage",       site.get("parking_garage") or "None"),
            ("Total / Ratio",f"{_v(site.get('parking_total'), 'n')} ({site.get('parking_ratio') or 'N/A'})"),
            ("Pet Yards",    site.get("pet_yards") or "N/A"),
            ("Storage",      site.get("storage") or "N/A"),
        ]):
            r = _kv(ws3, r, k, v, alt=bool(i % 2), n=6)
        r = _sp(ws3, r)

    if _on('employers'):
        employers = demo.get("employers") or []
        r = _sec(ws3, r, "G.  MAJOR EMPLOYERS & ECONOMIC DRIVERS", n=6)
        if employers and any(e.get("name") for e in employers):
            r = _thdr(ws3, r, ["Employer / Institution","Drive Time","Employees","Sector","Notes"], n=6)
            for i, e in enumerate(employers):
                if e.get("name"):
                    r = _drow(ws3, r, [
                        e.get("name","—"), e.get("drive") or "—", e.get("employees") or "—",
                        e.get("sector") or "—", e.get("notes") or "—",
                    ], alt=bool(i % 2), als=["left","center","center","left","left"], h=22, n=6)
        else:
            r = _kv(ws3, r, "Note", "No employer data found in OM — source from CoStar or broker.", n=6)
        r = _sp(ws3, r)

    if _on('market'):
        r = _sec(ws3, r, "H.  MARKET & SUBMARKET OVERVIEW", n=6)
        if mkt.get("market_summary"):
            r = _shdr(ws3, r, "Market Narrative", n=6)
            _fr(ws3, r, 6, C_WHITE)
            _sc(ws3, r, 1, mkt["market_summary"], bg=C_WHITE, fg=C_BODY, size=9, wrap=True)
            ws3.row_dimensions[r].height = 45; r += 1; r = _sp(ws3, r)

        r = _shdr(ws3, r, "Submarket Metrics", n=6)
        for i, (k, v) in enumerate([
            ("Submarket",         mkt.get("submarket")),
            ("Sub. Occupancy",    mkt.get("sub_occupancy")),
            ("Sub. Avg Rent",     _v(mkt.get("sub_rent"), "$")),
            ("Sub. Rent Growth",  mkt.get("sub_growth")),
            ("Metro Inventory",   mkt.get("metro_inventory")),
            ("Metro Occupancy",   mkt.get("metro_occupancy")),
        ]):
            if v: r = _kv(ws3, r, k, v, alt=bool(i % 2), n=6)

    if _on('market'):
        r = _sp(ws3, r)
        r = _sec(ws3, r, "I.  SUPPLY & DEMAND", n=6)
        r = _thdr(ws3, r, ["Metric","Value","Notes"], n=6)
        supply_rows = [
            ("Pipeline Units (Under Construction)", _v(mkt.get("pipeline"), "n"),    "Units currently under construction in submarket"),
            ("Absorption (Units / Year)",           _v(mkt.get("absorption"), "n"),   "Annual net absorption in submarket"),
            ("Investment Volume",                   mkt.get("investment_vol") or "—", "Total multifamily investment volume"),
        ]
        for i, (k, v, note) in enumerate(supply_rows):
            if v and v != "N/A":
                bg = C_ALT if bool(i % 2) else C_WHITE
                _fr(ws3, r, 6, bg)
                _sc(ws3, r, 1, k,    bold=True, fg=C_LABEL,   bg=bg, size=9)
                _sc(ws3, r, 2, v,    fg=C_BLUE_IN, bg=bg, size=9, ha="right")
                _sc(ws3, r, 3, note, fg=C_BODY,    bg=bg, size=9, wrap=True)
                ws3.cell(row=r, column=4).fill = PatternFill("solid", fgColor=bg)
                ws3.cell(row=r, column=5).fill = PatternFill("solid", fgColor=bg)
                ws3.cell(row=r, column=6).fill = PatternFill("solid", fgColor=bg)
                ws3.row_dimensions[r].height = 16; r += 1

        devs = mkt.get("major_developments") or []
        if devs and any(d.get("name") for d in devs):
            r = _sp(ws3, r)
            r = _shdr(ws3, r, "Pipeline Developments", n=6)
            r = _thdr(ws3, r, ["Development","Description","Est. Cost","Jobs","Timeline"], n=6)
            for i, dev in enumerate(devs):
                if dev.get("name"):
                    r = _drow(ws3, r, [
                        dev.get("name","—"), dev.get("description") or "—",
                        dev.get("cost") or "—", dev.get("jobs") or "—",
                        dev.get("timeline") or "—",
                    ], alt=bool(i % 2), als=["left","left","right","center","left"], h=25, n=6)
        r = _sp(ws3, r)

    _fr(ws3, r, 6, C_ALT)
    _sc(ws3, r, 1, "AI-generated. Internal use only. Verify all figures independently. Powered by Anthropic Claude.",
        bg=C_ALT, fg="FF888880", size=8, italic=True)
    ws3.row_dimensions[r].height = 13

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 4 — UNDERWRITING FLAGS (dedicated sheet, properly sized columns)
    # ══════════════════════════════════════════════════════════════════════════
    if _on('flags'):
        ws4 = wb.create_sheet("Flags")
        _setup_sheet(ws4, freeze_at="A6", tab_color="1E3A5F")
        # Generous Detail column so flag text reads naturally
        ws4.column_dimensions["A"].width = 14   # Category
        ws4.column_dimensions["B"].width = 32   # Title
        ws4.column_dimensions["C"].width = 80   # Detail — wide
        # Trailing columns kept narrow so colored band doesn't extend forever
        for _letter in "DEFG":
            ws4.column_dimensions[_letter].width = 2

        r = 1
        r = _cover(ws4, r, f"UNDERWRITING FLAGS  |  {prop}", subtitle, n=3)
        r = _sec(ws4, r, "RISK FLAGS & UPSIDE OBSERVATIONS", n=3)
        r = _thdr(ws4, r, ["Category", "Flag", "Detail"], n=3)

        flag_bg_map = {
            "Warning": C_WARN, "Caution": C_WARN, "Risk": C_WARN,
            "Opportunity": C_GREEN_L, "Upside": C_GREEN_L,
            "Info": C_BLUE_L,
            "Verify": C_PURPLE_L, "Verification": C_PURPLE_L
        }
        flag_fg_map = {
            "Warning": "FF92400E", "Caution": "FF92400E", "Risk": "FF92400E",
            "Opportunity": "FF065F46", "Upside": "FF065F46",
            "Info": "FF1E3A8A",
            "Verify": "FF5B21B6", "Verification": "FF5B21B6"
        }

        if flags:
            for f in flags:
                cat = f.get("category", "Info")
                bg  = flag_bg_map.get(cat, C_WHITE)
                fg  = flag_fg_map.get(cat, C_TEXT)
                # Category cell — colored chip
                _sc(ws4, r, 1, cat.upper(), bold=True, bg=bg, fg=fg,
                    size=9, ha="center", wrap=False)
                ws4.cell(row=r, column=1).border = _BOTTOM_BORDER
                # Title cell — bold dark text on white
                _sc(ws4, r, 2, f.get("title", ""), bold=True, bg=C_WHITE,
                    fg=C_NAVY, size=10, ha="left", wrap=True)
                ws4.cell(row=r, column=2).border = _BOTTOM_BORDER
                # Detail cell — body text wrapped on white
                _sc(ws4, r, 3, f.get("detail", ""), bg=C_WHITE,
                    fg=C_TEXT, size=9, ha="left", wrap=True)
                ws4.cell(row=r, column=3).border = _BOTTOM_BORDER
                # Estimate height based on detail length
                detail_len = len(f.get("detail", ""))
                approx_lines = max(1, -(-detail_len // 90))   # 90 chars/line at width 80
                ws4.row_dimensions[r].height = max(36, approx_lines * 15 + 8)
                r += 1
            r = _sp(ws4, r)
        else:
            r = _kv(ws4, r, "Note", "No underwriting flags generated.", n=3)

        # Footer
        _fr(ws4, r, 3, C_WHITE)
        for col in range(1, 4):
            ws4.cell(row=r, column=col).border = _BOTTOM_BORDER
        _sc(ws4, r, 1,
            "AI-generated from broker OM. Internal use only. Verify all figures independently. Powered by Anthropic Claude.",
            bg=C_WHITE, fg=C_TEXT_MUTED, size=8, italic=True, wrap=False)
        try: ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        except Exception: pass
        ws4.row_dimensions[r].height = 16

    # ── Auto-adjust row heights based on content length & column width ──
    def _auto_fit_rows(ws):
        col_widths = {}
        for letter, dim in ws.column_dimensions.items():
            if dim.width:
                col_widths[letter] = dim.width
        for row_cells in ws.iter_rows():
            r_idx = row_cells[0].row
            existing = ws.row_dimensions[r_idx].height
            if r_idx == 1:
                continue
            max_lines = 1
            for cell in row_cells:
                val = cell.value
                if val is None or val == "":
                    continue
                text = str(val)
                col_letter = cell.column_letter
                col_w = col_widths.get(col_letter, 12)
                chars_per_line = max(int(col_w * 1.1), 8)
                lines_for_cell = 0
                for segment in text.split("\n"):
                    seg_len = len(segment)
                    if seg_len == 0:
                        lines_for_cell += 1
                    else:
                        lines_for_cell += max(1, -(-seg_len // chars_per_line))
                if lines_for_cell > max_lines:
                    max_lines = lines_for_cell
            target = min(max(14, max_lines * 14 + 2), 80)
            if existing is None or target > existing:
                ws.row_dimensions[r_idx].height = target

    for sheet in wb.worksheets:
        # Flags sheet has manually-tuned row heights based on detail length;
        # skip auto-fit (which caps at 80pt and would clip long flag detail rows).
        if sheet.title == "Flags":
            continue
        _auto_fit_rows(sheet)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    pass

# ═══════════════════════════════════════════════════════════════════════════
# NEW NAVBAR — RealVal logo (clickable → therealval.com), tagline, user chip
# ═══════════════════════════════════════════════════════════════════════════
_user_email = ""
if st.session_state.get("user"):
    try:
        _user_email = st.session_state["user"].email or ""
    except Exception:
        _user_email = ""
_avatar_initials = "".join(p[0].upper() for p in (_user_email.split("@")[0] or "U").replace(".", " ").split()[:2]) or "U"

_logo_img_tag = (
    f'<img src="data:image/png;base64,{_LOGO_B64}" alt="RealVal" class="rv-logo-full" />'
    if _LOGO_B64 else
    '<span class="rv-mark-fallback">RealVal</span>'
)

st.markdown(f"""
<div class="rv-nav">
  <div class="rv-nav-left">
    <a href="https://therealval.com/" target="_blank" rel="noopener" class="rv-brand">
      {_logo_img_tag}
    </a>
    <span class="rv-nav-tagline">OM Intelligence</span>
  </div>
  <div class="rv-nav-right">
    <span class="rv-nav-link active">Analyze</span>
    <details class="rv-profile">
      <summary class="rv-profile-trigger" aria-label="Account menu">
        <div class="rv-user-chip-avatar">{_avatar_initials}</div>
      </summary>
      <div class="rv-profile-menu">
        <div class="rv-profile-header">
          <div class="rv-profile-avatar-lg">{_avatar_initials}</div>
          <div class="rv-profile-info">
            <div class="rv-profile-name">{_user_email.split('@')[0] if _user_email else 'User'}</div>
            <div class="rv-profile-email">{_user_email}</div>
          </div>
        </div>
        <div class="rv-profile-divider"></div>
        <a href="?signout=1" target="_top" class="rv-profile-item rv-profile-item-danger"
           onclick="event.stopPropagation();try{{window.top.location.href=(window.top.location.pathname||'/')+'?signout=1';}}catch(e){{window.location.href='?signout=1';}}return false;">
          <svg width="14" height="14" viewBox="0 0 16 16" fill="none" style="vertical-align:middle;margin-right:8px;pointer-events:none;">
            <path d="M6 14H3.5C2.67 14 2 13.33 2 12.5v-9C2 2.67 2.67 2 3.5 2H6" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
            <path d="M10.5 11L14 8L10.5 5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
            <path d="M14 8H6.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
          </svg>
          Sign out
        </a>
      </div>
    </details>
  </div>
</div>
""", unsafe_allow_html=True)

# Handle sign-out triggered by the profile-menu link (?signout=1)
if st.query_params.get("signout") == "1":
    st.query_params.clear()
    logout()

# ═══════════════════════════════════════════════════════════════════════════
# SECTION FLAGS — same keys/state as before, just defaulted on
# ═══════════════════════════════════════════════════════════════════════════
_keys = ["deal","unitmix","opstat","valueadd","financing","flags","tax","repl",
         "rentcomps","addinc2","salecomps","addinc","utilities","pop","afford",
         "schools","employers","market"]
for k in _keys:
    if "sel_"+k not in st.session_state:
        st.session_state["sel_"+k] = True

def _cb_selall():
    for k in _keys:
        st.session_state["sel_"+k] = True

def _cb_desall():
    for k in _keys:
        st.session_state["sel_"+k] = False

# Sidebar reserved for future (History, Settings). Kept collapsed by default.
# Sign-out moved to the profile dropdown menu in the navbar.

# ═══════════════════════════════════════════════════════════════════════════
# HERO — single centred upload area with title, subtitle, and dropzone
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="rv-hero">
  <div class="rv-hero-title">Drop your Offering Memorandum</div>
  <div class="rv-hero-sub">
    PDF only · CBRE, JLL, Marcus &amp; Millichap, Cushman &amp; Wakefield, Newmark, Berkadia and others · ~60 seconds to analyze
  </div>
</div>
""", unsafe_allow_html=True)

api_key = st.secrets.get("ANTHROPIC_API_KEY", os.environ.get("ANTHROPIC_API_KEY", ""))
if not api_key:
    st.error("""
**API key not configured.**
- **Streamlit Cloud:** Go to ⚙️ Settings → Secrets → add: `ANTHROPIC_API_KEY = "sk-ant-..."`
- **Local:** Create `.streamlit/secrets.toml` with the same line.
""")
    st.stop()
os.environ["ANTHROPIC_API_KEY"] = api_key

uploaded = st.file_uploader("Drop your OM PDF here", type=["pdf"], label_visibility="collapsed")

# ═══════════════════════════════════════════════════════════════════════════
# CUSTOMIZE REPORT — expander replaces the right-column wall of checkboxes
# ═══════════════════════════════════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════════════════════
# CUSTOMIZE REPORT — staged changes pattern
# Checkboxes write to tmp_sel_* keys; Save copies them to the real sel_* keys.
# Until Save is clicked, the actual report selection is unchanged.
# ═══════════════════════════════════════════════════════════════════════════
def _cb(key, label):
    # Bind to tmp_* widget so the real sel_* stays until Save commits
    tmp_key = "tmp_" + key
    if tmp_key not in st.session_state:
        st.session_state[tmp_key] = st.session_state.get(key, True)
    st.checkbox(label, key=tmp_key)

def _commit_selections():
    """Copy staged tmp_sel_* values into the real sel_* keys."""
    for _kk in _keys:
        _real = "sel_" + _kk
        _tmp  = "tmp_sel_" + _kk
        if _tmp in st.session_state:
            st.session_state[_real] = st.session_state[_tmp]
    st.session_state["_customize_saved_flash"] = True

def _stage_selall():
    for _kk in _keys:
        st.session_state["tmp_sel_" + _kk] = True

def _stage_desall():
    for _kk in _keys:
        st.session_state["tmp_sel_" + _kk] = False

# Count of staged selections (what the user is choosing right now)
_n_staged = sum(
    st.session_state.get("tmp_sel_" + k,
                          st.session_state.get("sel_" + k, True))
    for k in _keys
)
# Count of committed selections (what will actually run)
_n_committed = sum(st.session_state.get("sel_" + k, True) for k in _keys)
# Detect unsaved diff
_unsaved = any(
    st.session_state.get("tmp_sel_" + k,
                          st.session_state.get("sel_" + k, True))
    != st.session_state.get("sel_" + k, True)
    for k in _keys
)

_label_suffix = (
    f"{_n_staged} of 18 staged"
    + ("  ·  unsaved changes" if _unsaved else "")
)

with st.expander(f"⚙  Customize report  ·  {_label_suffix}", expanded=_unsaved):
    cust_c1, cust_c2, cust_c3 = st.columns(3)
    with cust_c1:
        st.markdown('<div class="rv-cust-group-title">Financials Tab</div>', unsafe_allow_html=True)
        _cb("sel_deal",      "Deal summary & property details")
        _cb("sel_unitmix",   "Unit mix with rent upside")
        _cb("sel_opstat",    "Operating statement (all periods)")
        _cb("sel_valueadd",  "Value-add by floor plan & revenue levers")
        _cb("sel_financing", "Financing & debt terms")
        _cb("sel_tax",       "Property tax & abatement")
        _cb("sel_repl",      "Replacement cost, insurance & management")
    with cust_c2:
        st.markdown('<div class="rv-cust-group-title">Comparables Tab</div>', unsafe_allow_html=True)
        _cb("sel_rentcomps", "Garden & townhouse rent comps")
        _cb("sel_addinc2",   "Additional income opportunities")
        _cb("sel_salecomps", "Sale comparables with buyer/seller")
        st.markdown('<div class="rv-cust-group-title" style="margin-top:18px;">Flags Tab</div>', unsafe_allow_html=True)
        _cb("sel_flags",     "Underwriting flags")
    with cust_c3:
        st.markdown('<div class="rv-cust-group-title">Demographics Tab</div>', unsafe_allow_html=True)
        _cb("sel_addinc",    "Additional income breakdown")
        _cb("sel_utilities", "Utilities & site information")
        _cb("sel_pop",       "Population & income (1-/3-/5-mile)")
        _cb("sel_afford",    "Affordability analysis")
        _cb("sel_schools",   "Schools, crime & quality of life")
        _cb("sel_employers", "Major employers & economic drivers")
        _cb("sel_market",    "Market, submarket & supply/demand")

    st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)
    sa1, sa2, sa3, sa4 = st.columns([0.18, 0.18, 0.30, 0.34])
    with sa1: st.button("✓  Select all",   key="btn_sel", on_click=_stage_selall, use_container_width=True)
    with sa2: st.button("✕  Deselect all", key="btn_des", on_click=_stage_desall, use_container_width=True)
    with sa4:
        st.button(
            "💾  Save changes" if _unsaved else "✓  Saved",
            key="btn_save_cust",
            type="primary" if _unsaved else "secondary",
            on_click=_commit_selections,
            use_container_width=True,
            disabled=not _unsaved,
        )

    # Flash success message after save
    if st.session_state.pop("_customize_saved_flash", False):
        st.markdown(
            '<div style="margin-top:8px;color:#02A9A1;font-size:12px;">✓ Selections saved.</div>',
            unsafe_allow_html=True
        )

sel_deal=st.session_state["sel_deal"]; sel_unitmix=st.session_state["sel_unitmix"]
sel_opstat=st.session_state["sel_opstat"]; sel_valueadd=st.session_state["sel_valueadd"]
sel_financing=st.session_state["sel_financing"]; sel_flags=st.session_state["sel_flags"]
sel_tax=st.session_state["sel_tax"]; sel_repl=st.session_state["sel_repl"]
sel_rentcomps=st.session_state["sel_rentcomps"]; sel_addinc2=st.session_state["sel_addinc2"]
sel_salecomps=st.session_state["sel_salecomps"]; sel_addinc=st.session_state["sel_addinc"]
sel_utilities=st.session_state["sel_utilities"]; sel_pop=st.session_state["sel_pop"]
sel_afford=st.session_state["sel_afford"]; sel_schools=st.session_state["sel_schools"]
sel_employers=st.session_state["sel_employers"]; sel_market=st.session_state["sel_market"]

# ═══════════════════════════════════════════════════════════════════════════
# MAIN BODY — original logic preserved; outer column wrapper removed
# ═══════════════════════════════════════════════════════════════════════════
if True:

    # ══════════════════════════════════════════════════════════════════════
    # ── DEMO STATE INJECTOR — for UI screenshots only ──
    # Triggers when ?demo=blueridge is in URL; loads a saved analysis into
    # session_state so the results UI can be screenshotted without re-running.
    # Safe to leave in: silent no-op unless the param is set.
    # ══════════════════════════════════════════════════════════════════════
    try:
        _qp = st.query_params
        if _qp.get("demo") == "blueridge" and "analysis_data" not in st.session_state:
            import pickle, pathlib
            _demo_path = pathlib.Path(__file__).parent / "assets" / "demo_blueridge.pkl"
            if _demo_path.exists():
                _demo = pickle.loads(_demo_path.read_bytes())
                for _dkey, _dval in _demo.items():
                    st.session_state[_dkey] = _dval
                # Mark the upload as "matched" so the auto-clear below doesn't wipe it
                st.session_state["uploaded_file_id"] = "demo:blueridge"
    except Exception:
        pass

    # ── Detect new file upload and clear previous results ─────────────────
    _current_file_id = None
    if uploaded is not None:
        _current_file_id = f"{uploaded.name}_{uploaded.size}"
    elif st.query_params.get("demo") == "blueridge":
        _current_file_id = "demo:blueridge"   # don't wipe injected demo state
    _last_file_id = st.session_state.get("uploaded_file_id")
    if _current_file_id != _last_file_id:
        for _k in ("analysis_data", "analysis_excel_bytes", "analysis_filename",
                   "analysis_prop_name", "analysis_broker_name"):
            st.session_state.pop(_k, None)
        st.session_state["uploaded_file_id"] = _current_file_id

    if uploaded is None and "analysis_data" not in st.session_state:
        st.markdown("""
<div class="rv-steps">
  <div class="rv-step">
    <div class="rv-step-num">01</div>
    <div class="rv-step-title">Upload OM</div>
    <div class="rv-step-desc">Any broker PDF</div>
  </div>
  <div class="rv-step">
    <div class="rv-step-num">02</div>
    <div class="rv-step-title">AI extracts</div>
    <div class="rv-step-desc">Financials · comps · flags</div>
  </div>
  <div class="rv-step">
    <div class="rv-step-num">03</div>
    <div class="rv-step-title">Export</div>
    <div class="rv-step-desc">Excel ready to share</div>
  </div>
</div>
""", unsafe_allow_html=True)
        st.stop()

    if uploaded is not None:
        size_mb = uploaded.size / 1024 / 1024
        st.markdown(f"""
<div class="rv-file-bar">
  <div>
    <div class="rv-file-meta">📄 &nbsp;<b>{uploaded.name}</b></div>
    <div class="rv-file-size">{size_mb:.1f} MB &nbsp;·&nbsp; ready to analyze</div>
  </div>
</div>
""", unsafe_allow_html=True)

        # Center the Analyze button
        _bc1, _bc2, _bc3 = st.columns([0.3, 0.4, 0.3])
        with _bc2:
            _analyze_clicked = st.button("Analyze Offering Memorandum",
                                          type="primary",
                                          use_container_width=True,
                                          key="btn_analyze_main")
        # Breathing room below the Analyze button
        st.markdown('<div style="height:48px;"></div>', unsafe_allow_html=True)
    else:
        _analyze_clicked = False

    if _analyze_clicked:
        progress_bar = st.progress(0, text="Starting...")
        status_box   = st.empty()

        def set_progress(pct, msg):
            progress_bar.progress(pct, text=msg)
            status_box.markdown(
                f"<div style='font-size:12px;color:#5A8FAA;margin-top:4px;'>{msg}</div>",
                unsafe_allow_html=True)

        try:
            set_progress(10, "Extracting text from PDF…")
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(uploaded.getvalue())
                tmp_path = tmp.name
            pdf_text = extract_pdf_text(tmp_path)
            os.unlink(tmp_path)

            if not pdf_text or len(pdf_text.strip()) < 200:
                progress_bar.empty(); status_box.empty()
                st.error("Could not extract readable text. This PDF may be scanned — please use a text-based PDF.")
                st.stop()

            set_progress(30, f"Extracted {len(pdf_text):,} characters. Sending to Claude AI…")
            def log(msg): set_progress(55, msg)
            data = analyze_om(pdf_text, api_key, log)

            set_progress(75, "Generating Excel report…")
            sections = {
                "deal":      sel_deal,     "unitmix":   sel_unitmix,
                "opstat":    sel_opstat,   "valueadd":  sel_valueadd,
                "financing": sel_financing,"flags":     sel_flags,
                "tax":       sel_tax,      "repl":      sel_repl,
                "rentcomps": sel_rentcomps,"addinc2":   sel_addinc2,
                "salecomps": sel_salecomps,"addinc":    sel_addinc,
                "utilities": sel_utilities,"pop":       sel_pop,
                "afford":    sel_afford,   "schools":   sel_schools,
                "employers": sel_employers,"market":    sel_market,
            }
            excel_bytes = build_excel(data, uploaded.name, sections=sections)
            set_progress(100, "Done!")
            progress_bar.empty(); status_box.empty()

            # ── Cache results in session_state so they survive reruns ──
            st.session_state["analysis_data"]        = data
            st.session_state["analysis_excel_bytes"] = excel_bytes
            st.session_state["analysis_filename"]    = uploaded.name
            st.session_state["analysis_prop_name"]   = (data.get("property") or {}).get("name") or "Property"
            st.session_state["analysis_broker_name"] = (data.get("broker")   or {}).get("name") or "Unknown broker"
            # Clear any stale error from a previous failed run
            st.session_state.pop("last_error", None)

        except Exception as e:
            progress_bar.empty(); status_box.empty()
            import traceback
            tb = traceback.format_exc()
            # Persist to session_state so it survives reruns / websocket drops.
            # The error renders OUTSIDE this button block so it can't be hidden
            # by st.stop() or by the websocket dying mid-traceback.
            st.session_state["last_error"] = f"{type(e).__name__}: {e}\n\n{tb}"

    # ══════════════════════════════════════════════════════════════════════
    # ── Persistent error display — shows even after a websocket drop ──
    # ══════════════════════════════════════════════════════════════════════
    if "last_error" in st.session_state:
        st.error("⚠️ The last analysis hit an error:")
        st.code(st.session_state["last_error"], language="python")
        if st.button("Dismiss error", key="dismiss_err"):
            del st.session_state["last_error"]
            st.rerun()

    # ══════════════════════════════════════════════════════════════════════
    # ── Results block — renders whenever cached analysis exists ──
    # Lives OUTSIDE the analyze button so download clicks don't wipe it
    # ══════════════════════════════════════════════════════════════════════
    if "analysis_data" in st.session_state:
        data         = st.session_state["analysis_data"]
        excel_bytes  = st.session_state["analysis_excel_bytes"]
        prop_name    = st.session_state["analysis_prop_name"]
        broker_name  = st.session_state["analysis_broker_name"]

        st.markdown(f"""
<div style="padding: 0 44px;">
<div class="rv-success">
  <div class="rv-success-icon">✅</div>
  <div class="rv-success-text">Report ready &nbsp;·&nbsp; <b>{prop_name}</b> &nbsp;·&nbsp; Broker: {broker_name}</div>
</div>
</div>
""", unsafe_allow_html=True)

        safe = re.sub(r"[^a-zA-Z0-9_\- ]", "", prop_name).strip().replace(" ", "_")

        # ── Build "all summary tabs" multi-sheet workbook (Fix 4) ──
        def _build_summary_tabs_workbook(d: dict) -> bytes:
            """One xlsx with every in-app summary tab as a separate sheet,
            styled to match the main Underwriting Report (cover row, navy
            section headers, light subheaders, alternating rows, accent NOI/total rows)."""
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter as _col

            # ── Use the SAME palette tokens as the main report ──
            HDR      = C_NAVY
            HDR2     = C_NAVY
            HDR3     = C_NAVY
            SUB_HDR  = C_SUB_FILL
            HDR_TXT  = C_WHITE
            AMBER    = C_WHITE   # totals on navy = white
            BLUE_IN  = C_TEXT
            LBL      = C_TEXT_MUTED
            BODY     = C_TEXT
            WHITE    = C_WHITE
            ALT      = C_ROW_ALT
            BORDER   = C_DIVIDER

            _bord_bottom = Border(bottom=Side(style="thin", color=BORDER))
            _bord_thick  = Border(bottom=Side(style="medium", color=HDR))
            _bord_top    = Border(top=Side(style="medium", color=HDR))
            _bord_none   = Border()

            prop_x   = (d.get("property") or {}).get("name") or "Property"
            broker_x = (d.get("broker") or {}).get("name") or "N/A"
            pd_x     = d.get("property") or {}
            addr_x = (
                f"{pd_x.get('address','')}, {pd_x.get('city','')}, "
                f"{pd_x.get('state','')} {pd_x.get('zip','')}"
            ).strip(", ")
            date_x   = datetime.today().strftime("%B %d, %Y")
            subtitle_x = (
                f"{prop_x}  ·  {addr_x}  ·  {_v(pd_x.get('units'),'n')} Units  ·  "
                f"{_v(pd_x.get('year_built'))} Built  ·  {broker_x}  ·  {date_x}"
            )

            wb_s = Workbook()
            wb_s.remove(wb_s.active)

            def _set_cell(ws, row, col, val, *, bold=False, bg=None, fg=BODY,
                          size=9, ha="left", italic=False, wrap=True, border=None):
                """Set a cell with numeric auto-detection + clean borders."""
                c = ws.cell(row=row, column=col)
                final_fg = fg

                # Auto-promote numeric strings to real Excel numbers
                if val is not None and val != "" and val != "—":
                    num_val, fmt = _parse_numeric(val)
                    if num_val is not None:
                        c.value = num_val
                        c.number_format = fmt
                        if num_val < 0 and not bold:
                            final_fg = C_NEG
                    else:
                        c.value = val
                else:
                    c.value = val if val not in (None, "") else "—"

                c.font = Font(name="Calibri", bold=bold, color=final_fg,
                              size=size, italic=italic)
                c.alignment = Alignment(horizontal=ha, vertical="center",
                                        wrap_text=wrap)
                if bg:
                    c.fill = PatternFill("solid", fgColor=bg)
                c.border = border if border is not None else _bord_none
                return c

            def _fill_row(ws, row, ncols, bg):
                for c in range(1, ncols + 1):
                    ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=bg)

            def _cover_block(ws, line1, line2, ncols):
                # Big navy title on white — merged across cols
                _fill_row(ws, 1, ncols, WHITE)
                _set_cell(ws, 1, 1, line1, bold=True, bg=WHITE, fg=HDR,
                          size=18, ha="left", wrap=False)
                try: ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
                except Exception: pass
                ws.row_dimensions[1].height = 32

                # Subtitle in muted grey — merged
                _fill_row(ws, 2, ncols, WHITE)
                _set_cell(ws, 2, 1, line2, bold=False, bg=WHITE, fg=LBL,
                          size=10, ha="left", wrap=False)
                try: ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
                except Exception: pass
                ws.row_dimensions[2].height = 18

                # Thin navy rule
                _fill_row(ws, 3, ncols, WHITE)
                for col in range(1, ncols + 1):
                    ws.cell(row=3, column=col).border = _bord_top
                ws.row_dimensions[3].height = 4

                # Spacer
                ws.row_dimensions[4].height = 12
                return 5

            def _section_header(ws, row, title, ncols):
                _fill_row(ws, row, ncols, HDR)
                _set_cell(ws, row, 1, title, bold=True, bg=HDR, fg=HDR_TXT,
                          size=11, ha="left", wrap=False)
                try: ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
                except Exception: pass
                ws.row_dimensions[row].height = 26
                return row + 1

            def _column_headers(ws, row, headers):
                _fill_row(ws, row, len(headers), SUB_HDR)
                for i, h in enumerate(headers):
                    _set_cell(ws, row, 1 + i, h, bold=True, bg=SUB_HDR,
                              fg=HDR, size=9, ha="center", border=_bord_thick)
                ws.row_dimensions[row].height = 22
                return row + 1

            def _data_row(ws, row, vals, alts=None, alt_bg=False):
                bg = ALT if alt_bg else WHITE
                _fill_row(ws, row, len(vals), bg)
                for i, v in enumerate(vals):
                    ha = alts[i] if alts and i < len(alts) else "left"
                    _set_cell(ws, row, 1 + i, v, bg=bg, fg=BODY,
                              size=9, ha=ha, wrap=(ha == "left"),
                              border=_bord_bottom)
                ws.row_dimensions[row].height = 17
                return row + 1

            def _accent_row(ws, row, vals, ncols, color=None):
                """Highlighted total row — navy fill, white bold text."""
                fill_color = color or HDR
                _fill_row(ws, row, ncols, fill_color)
                for i, v in enumerate(vals):
                    ha = "left" if i == 0 else "right"
                    _set_cell(ws, row, 1 + i, v, bold=True, bg=fill_color,
                              fg=HDR_TXT, size=10, ha=ha)
                ws.row_dimensions[row].height = 24
                return row + 1

            def _kv_row(ws, row, label, value, ncols, alt_bg=False):
                _fill_row(ws, row, ncols, WHITE)
                _set_cell(ws, row, 1, label, bold=True, bg=WHITE, fg=LBL,
                          size=9, border=_bord_bottom)
                _set_cell(ws, row, 2, value, bg=WHITE, fg=BODY, size=9,
                          ha="left", border=_bord_bottom)
                # Extend bottom rule across remaining cols for clean line
                for col in range(3, ncols + 1):
                    ws.cell(row=row, column=col).border = _bord_bottom
                ws.row_dimensions[row].height = 18
                return row + 1

            def _autosize_columns(ws, widths):
                for col_letter, w in widths.items():
                    ws.column_dimensions[col_letter].width = w

            def _add_footer(ws, row, ncols):
                ws.row_dimensions[row].height = 10
                row += 1
                _fill_row(ws, row, ncols, WHITE)
                # Thin top rule above footer
                for col in range(1, ncols + 1):
                    ws.cell(row=row, column=col).border = _bord_bottom
                _set_cell(ws, row, 1,
                          "AI-generated. Internal use only. Verify all figures "
                          "independently. Powered by Anthropic Claude.",
                          bg=WHITE, fg=LBL, size=8, italic=True, wrap=False,
                          border=_bord_bottom)
                try: ws.merge_cells(start_row=row, start_column=1,
                                    end_row=row, end_column=ncols)
                except Exception: pass
                ws.row_dimensions[row].height = 16

            # ────────────────────────────────────────────────────────────────
            # Sheet 1 — Unit Mix
            # ────────────────────────────────────────────────────────────────
            ws = wb_s.create_sheet("Unit Mix")
            _setup_sheet(ws, freeze_at="A6", tab_color="1E3A5F")
            _autosize_columns(ws, {"A":18,"B":18,"C":10,"D":10,"E":10,
                                    "F":14,"G":14,"H":14,"I":14})
            r = _cover_block(ws, "UNIT MIX SUMMARY", subtitle_x, 9)
            r = _section_header(ws, r, "A.  UNIT MIX BY FLOOR PLAN", 9)
            umix_x = d.get("unit_mix") or []
            if umix_x:
                r = _column_headers(ws, r, ["Type","Plan","Units","Mix %","SF",
                                             "Market Rent","Eff. Rent",
                                             "Target Rent","Upside/Unit"])
                aligns = ["left","left","center","center","right",
                          "right","right","right","right"]
                # Compute mix % from count/total to guarantee sum = 100%
                _total_units = sum(int(u.get("count") or 0) for u in umix_x) or 1
                _total_sf_x_units = 0
                for i, u in enumerate(umix_x):
                    _cnt = int(u.get("count") or 0)
                    _mix_pct = (_cnt / _total_units) * 100 if _total_units else 0
                    _sf_val = u.get("sf") or 0
                    try: _total_sf_x_units += float(_sf_val) * _cnt
                    except: pass
                    r = _data_row(ws, r, [
                        u.get("type"), u.get("plan"),
                        _v(u.get("count"), "n"), f"{_mix_pct:.1f}%",
                        _v(u.get("sf"), "n"),
                        _v(u.get("market_rent"), "$"), _v(u.get("eff_rent"), "$"),
                        _v(u.get("target_rent"), "$"), _v(u.get("upside"), "$"),
                    ], alts=aligns, alt_bg=bool(i % 2))
                # TOTAL row
                _avg_sf = _total_sf_x_units / _total_units if _total_units else 0
                _fill_row(ws, r, 9, HDR)
                for _i, _val in enumerate([
                    "TOTAL / AVG", "", f"{_total_units}", "100.0%",
                    f"{int(_avg_sf):,}", "", "", "", ""
                ]):
                    _ha = "left" if _i == 0 else ("center" if _i in (2, 3) else "right")
                    _set_cell(ws, r, 1 + _i, _val, bold=True, bg=HDR,
                              fg=HDR_TXT, size=10, ha=_ha)
                ws.row_dimensions[r].height = 24
                r += 1
            else:
                r = _kv_row(ws, r, "Note", "No unit mix data extracted.", 9)
            _add_footer(ws, r, 9)

            # ────────────────────────────────────────────────────────────────
            # Sheet 2 — Value-Add
            # ────────────────────────────────────────────────────────────────
            ws = wb_s.create_sheet("Value-Add")
            _setup_sheet(ws, freeze_at="A6", tab_color="1E3A5F")
            _autosize_columns(ws, {"A":18,"B":10,"C":10,"D":14,"E":14,"F":14,"G":16,"H":30})
            r = _cover_block(ws, "VALUE-ADD SUMMARY", subtitle_x, 8)

            r = _section_header(ws, r, "A.  VALUE-ADD BY FLOOR PLAN", 8)
            va_x       = d.get("value_add") or {}
            va_plans_x = va_x.get("by_floor_plan") or []
            if va_plans_x:
                r = _column_headers(ws, r, ["Type","Units","SF","In-Place Rent",
                                             "Rehab Cost","Premium/Unit","Post-Rehab Rent",""])
                aligns = ["left","center","right","right","right","right","right","left"]
                for i, p in enumerate(va_plans_x):
                    r = _data_row(ws, r, [
                        p.get("type"), _v(p.get("units"), "n"), _v(p.get("sf"), "n"),
                        _v(p.get("inplace_rent"), "$"), _v(p.get("rehab_cost"), "$"),
                        _v(p.get("premium"), "$"), _v(p.get("post_rehab_rent"), "$"), "",
                    ], alts=aligns, alt_bg=bool(i % 2))
            else:
                r = _kv_row(ws, r, "Note", "No floor plan value-add data extracted.", 8)

            ws.row_dimensions[r].height = 6; r += 1
            r = _section_header(ws, r, "B.  REVENUE UPSIDE LEVERS", 8)
            levers_x = d.get("value_add_levers") or []
            if levers_x:
                r = _column_headers(ws, r, ["Lever","Units","Mo. Premium","Annual Upside",
                                             "Notes","","",""])
                aligns = ["left","center","right","right","left","left","left","left"]
                for i, lv in enumerate(levers_x):
                    is_total = "TOTAL" in (lv.get("lever") or "").upper()
                    row_data = [
                        lv.get("lever"),
                        "" if is_total else _v(lv.get("units"), "n"),
                        "" if is_total else _v(lv.get("monthly_premium"), "$"),
                        _v(lv.get("annual_upside"), "$"),
                        lv.get("notes") or "", "", "", "",
                    ]
                    if is_total:
                        r = _accent_row(ws, r, row_data, 8)
                    else:
                        r = _data_row(ws, r, row_data, alts=aligns, alt_bg=bool(i % 2))
            else:
                r = _kv_row(ws, r, "Note", "No revenue lever data extracted.", 8)

            ws.row_dimensions[r].height = 6; r += 1
            r = _section_header(ws, r, "C.  CAPEX & ROI SUMMARY", 8)
            for i, (k, v) in enumerate([
                ("Total Renovation Cost",       _v(va_x.get("total_cost"), "$")),
                ("Cost Per Unit (all-in)",      _v(va_x.get("cost_per_unit"), "$")),
                ("Exterior / Additional CapEx", _v(va_x.get("exterior_capex"), "$")),
                ("Monthly Rent Premium",        _v(va_x.get("monthly_premium"), "$")),
                ("Annual Rent Premium",         _v(va_x.get("annual_premium"), "$")),
                ("Return on Investment",        _pct(va_x.get('roi_pct'))),
            ]):
                r = _kv_row(ws, r, k, v, 8, alt_bg=bool(i % 2))
            _add_footer(ws, r, 8)

            # ────────────────────────────────────────────────────────────────
            # Sheet 3 — Rent Comps
            # ────────────────────────────────────────────────────────────────
            ws = wb_s.create_sheet("Rent Comps")
            _setup_sheet(ws, freeze_at="A6", tab_color="1E3A5F")
            _autosize_columns(ws, {"A":24,"B":12,"C":10,"D":10,"E":12,"F":14,"G":14,"H":10,"I":24})
            r = _cover_block(ws, "RENT COMPARABLES SUMMARY", subtitle_x, 9)

            rg_x     = d.get("rent_comps_garden")    or []
            rth_x    = d.get("rent_comps_townhouse") or []
            rcomps_x = d.get("rent_comps")           or []

            if rg_x:
                r = _section_header(ws, r, "A.  GARDEN RENT COMPARABLES", 9)
                r = _column_headers(ws, r, ["Property","Type","Year","Distance","Units",
                                             "Mkt Rent","Eff Rent","Avg SF","Notes"])
                for i, c in enumerate(rg_x):
                    r = _data_row(ws, r, [
                        c.get("name"), "Garden", _v(c.get("year_built")),
                        c.get("distance") or "—", _v(c.get("units"),"n"),
                        _v(c.get("rent"),"$"), _v(c.get("rent"),"$"),
                        _v(c.get("avg_sf"),"n"), c.get("notes") or "",
                    ], alts=["left","center","center","center","center","right","right","right","left"],
                       alt_bg=bool(i % 2))
                ws.row_dimensions[r].height = 6; r += 1

            if rth_x:
                sec_letter = "B" if rg_x else "A"
                r = _section_header(ws, r, f"{sec_letter}.  TOWNHOUSE RENT COMPARABLES", 9)
                r = _column_headers(ws, r, ["Property","Type","Year","Distance","Units",
                                             "Mkt Rent","Eff Rent","Avg SF","Notes"])
                for i, c in enumerate(rth_x):
                    r = _data_row(ws, r, [
                        c.get("name"), "Townhouse", _v(c.get("year_built")),
                        c.get("distance") or "—", _v(c.get("units"),"n"),
                        _v(c.get("rent"),"$"), _v(c.get("rent"),"$"),
                        _v(c.get("avg_sf"),"n"), c.get("notes") or "",
                    ], alts=["left","center","center","center","center","right","right","right","left"],
                       alt_bg=bool(i % 2))
                ws.row_dimensions[r].height = 6; r += 1

            if rcomps_x:
                sec_full = "A" if (not rg_x and not rth_x) else ("C" if (rg_x and rth_x) else "B")
                r = _section_header(ws, r, f"{sec_full}.  FULL COMPARABLE DETAIL", 9)
                r = _column_headers(ws, r, ["Property","Type","Year Built","Units","Occupancy",
                                             "Mkt Rent","Eff Rent","Avg SF","Notes"])
                for i, c in enumerate(rcomps_x):
                    r = _data_row(ws, r, [
                        c.get("name"), c.get("comp_type") or "—",
                        _v(c.get("year_built")), _v(c.get("units"), "n"),
                        _pct(c.get("occupancy")) if c.get("occupancy") else "—",
                        _v(c.get("total_market"), "$"), _v(c.get("total_eff"), "$"),
                        _v(c.get("avg_sf"), "n"), "",
                    ], alts=["left","center","center","center","center","right","right","right","left"],
                       alt_bg=bool(i % 2))
                ws.row_dimensions[r].height = 6; r += 1

            if not rg_x and not rth_x and not rcomps_x:
                r = _section_header(ws, r, "A.  RENT COMPARABLES", 9)
                r = _kv_row(ws, r, "Note", "No rent comp data extracted.", 9)

            _add_footer(ws, r, 9)

            # ────────────────────────────────────────────────────────────────
            # Sheet 4 — Financials
            # ────────────────────────────────────────────────────────────────
            ws = wb_s.create_sheet("Financials")
            _setup_sheet(ws, freeze_at="A6", tab_color="1E3A5F")
            fin_x     = d.get("financials") or {}
            periods_x = (fin_x.get("periods") or [])[:7]
            inc_x     = fin_x.get("income_lines")  or []
            exp_x     = fin_x.get("expense_lines") or []
            noi_x     = fin_x.get("noi") or {}
            n_p = max(len(periods_x), 1)
            cols_fin = {"A": 32}
            for i, c_letter in enumerate("BCDEFGH"):
                if i < n_p:
                    cols_fin[c_letter] = 16
            cols_fin[chr(ord("A") + 1 + n_p)] = 36   # Notes column
            _autosize_columns(ws, cols_fin)
            ncols_fin = 1 + n_p + 1
            r = _cover_block(ws, "FINANCIALS SUMMARY", subtitle_x, ncols_fin)

            if periods_x and (inc_x or exp_x):
                title = f"A.  OPERATING STATEMENT  ({'  ·  '.join(periods_x)})"
                r = _section_header(ws, r, title, ncols_fin)
                r = _column_headers(ws, r, ["Line Item"] + periods_x + ["Notes"])
                aligns = ["left"] + ["right"] * n_p + ["left"]

                def _line_row(line):
                    vals = line.get("values") or {}
                    return [line.get("item") or "—"] + [
                        _v(vals.get(p), "$") if vals.get(p) is not None else "—"
                        for p in periods_x
                    ] + [line.get("note") or ""]

                rendered_totals = set()

                # INCOME
                _fill_row(ws, r, ncols_fin, SUB_HDR)
                _set_cell(ws, r, 1, "INCOME", bold=True, bg=SUB_HDR, fg=LBL, size=9)
                ws.row_dimensions[r].height = 17; r += 1
                for i, line in enumerate(inc_x):
                    name = (line.get("item") or "").lower()
                    if line.get("is_total"):
                        rendered_totals.add(name)
                        r = _accent_row(ws, r, _line_row(line), ncols_fin)
                    elif line.get("is_subtotal"):
                        r = _accent_row(ws, r, _line_row(line), ncols_fin, color=HDR2)
                    else:
                        r = _data_row(ws, r, _line_row(line), alts=aligns, alt_bg=bool(i % 2))

                # EXPENSES
                _fill_row(ws, r, ncols_fin, SUB_HDR)
                _set_cell(ws, r, 1, "EXPENSES", bold=True, bg=SUB_HDR, fg=LBL, size=9)
                ws.row_dimensions[r].height = 17; r += 1
                for i, line in enumerate(exp_x):
                    name = (line.get("item") or "").lower()
                    if line.get("is_total"):
                        rendered_totals.add(name)
                        r = _accent_row(ws, r, _line_row(line), ncols_fin, color=HDR2)
                    elif line.get("is_subtotal"):
                        r = _accent_row(ws, r, _line_row(line), ncols_fin, color=HDR2)
                    else:
                        r = _data_row(ws, r, _line_row(line), alts=aligns, alt_bg=bool(i % 2))

                noi_already = any("net operating income" in t for t in rendered_totals)
                if not noi_already and any(_is_num(noi_x.get(p)) for p in periods_x):
                    r = _accent_row(ws, r, ["NET OPERATING INCOME"] + [
                        _v(noi_x.get(p), "$") if _is_num(noi_x.get(p)) else "—"
                        for p in periods_x
                    ] + [""], ncols_fin)
            else:
                r = _section_header(ws, r, "A.  OPERATING STATEMENT", ncols_fin)
                r = _kv_row(ws, r, "Note", "No financial data extracted.", ncols_fin)

            _add_footer(ws, r, ncols_fin)

            # ────────────────────────────────────────────────────────────────
            # Sheet 5 — Tax & Abatement
            # ────────────────────────────────────────────────────────────────
            ws = wb_s.create_sheet("Tax & Abatement")
            _setup_sheet(ws, freeze_at="A6", tab_color="1E3A5F")
            _autosize_columns(ws, {"A":32,"B":40})
            r = _cover_block(ws, "TAX & ABATEMENT SUMMARY", subtitle_x, 2)

            tax_x = d.get("tax") or {}
            tax_field_rows = [
                ("Parcel ID",             tax_x.get("parcel_id")),
                ("Assessed Market Value", _v(tax_x.get("assessed_value"), "$")),
                ("Millage — City",        tax_x.get("millage_city")),
                ("Millage — County",      tax_x.get("millage_county")),
                ("Total Millage Rate",    tax_x.get("millage_total")),
                ("Ad Valorem Tax",        _v(tax_x.get("tax_base"), "$")),
                ("Solid Waste / Fees",    _v(tax_x.get("solid_waste_fee"), "$")),
                ("Total Annual Tax Bill", _v(tax_x.get("total_tax"), "$")),
            ]
            tax_field_rows = [(k, v) for k, v in tax_field_rows if v not in (None, "", "N/A")]
            if tax_field_rows:
                r = _section_header(ws, r, "A.  PROPERTY TAX DETAIL", 2)
                for i, (k, v) in enumerate(tax_field_rows):
                    r = _kv_row(ws, r, k, v, 2, alt_bg=bool(i % 2))

            ab_rows = [
                ("Program",            tax_x.get("abatement_program")),
                ("Abatement %",        _pct(tax_x.get('abatement_pct')) if tax_x.get("abatement_pct") else None),
                ("Commitment Term",    tax_x.get("abatement_term_note")),
                ("Annual Tax Savings", _v(tax_x.get("abatement_annual_savings"), "$") if tax_x.get("abatement_annual_savings") else None),
                ("Max Allowable Rent", _v(tax_x.get("max_allowable_rent"), "$") if tax_x.get("max_allowable_rent") else None),
            ]
            ab_rows = [(k, v) for k, v in ab_rows if v not in (None, "", "N/A")]
            if ab_rows:
                ws.row_dimensions[r].height = 6; r += 1
                r = _section_header(ws, r, "B.  TAX ABATEMENT PROGRAM", 2)
                for i, (k, v) in enumerate(ab_rows):
                    r = _kv_row(ws, r, k, v, 2, alt_bg=bool(i % 2))

            if not tax_field_rows and not ab_rows:
                r = _section_header(ws, r, "A.  PROPERTY TAX", 2)
                r = _kv_row(ws, r, "Note", "No tax data extracted.", 2)

            _add_footer(ws, r, 2)

            # ────────────────────────────────────────────────────────────────
            # Sheet 6 — Demographics
            # ────────────────────────────────────────────────────────────────
            ws = wb_s.create_sheet("Demographics")
            _setup_sheet(ws, freeze_at="A6", tab_color="1E3A5F")
            _autosize_columns(ws, {"A":34,"B":18,"C":18,"D":18})
            r = _cover_block(ws, "DEMOGRAPHICS SUMMARY", subtitle_x, 4)
            r = _section_header(ws, r, "A.  POPULATION & INCOME", 4)
            r = _column_headers(ws, r, ["Metric","1-Mile Radius","3-Mile Radius","5-Mile Radius"])

            demo_x = d.get("demographics") or {}
            metric_pairs = [
                ("Population (2025)",         "pop_1mi", "pop_3mi", "pop_5mi", "n"),
                ("Population (2030 Proj.)",   "pop_2030_1mi", "pop_2030_3mi", "pop_2030_5mi", "n"),
                ("Population Growth (5-yr)",  "pop_growth_1mi", "pop_growth_3mi", "pop_growth_5mi", None),
                ("Median HH Income (2025)",   "median_income_1mi", "median_income_3mi", "median_income_5mi", "$"),
                ("Median HH Income (2030)",   "median_income_2030_1mi", "median_income_2030_3mi", "median_income_2030_5mi", "$"),
                ("Income Growth (5-yr)",      "income_growth_1mi", "income_growth_3mi", "income_growth_5mi", None),
                ("Renter %",                  "renter_pct_1mi", "renter_pct_3mi", "renter_pct_5mi", None),
                ("Bachelor's Degree+",        "college_pct_1mi", "college_pct_3mi", "college_pct_5mi", None),
                ("White-Collar %",            "white_collar_pct_1mi", "white_collar_pct_3mi", "white_collar_pct_5mi", None),
            ]
            aligns_demo = ["left", "right", "right", "right"]
            for i, (label, k1, k3, k5, fmt) in enumerate(metric_pairs):
                v1 = _v(demo_x.get(k1), fmt) if fmt else (demo_x.get(k1) or "—")
                v3 = _v(demo_x.get(k3), fmt) if fmt else (demo_x.get(k3) or "—")
                v5 = _v(demo_x.get(k5), fmt) if fmt else (demo_x.get(k5) or "—")
                r = _data_row(ws, r, [label, v1, v3, v5], alts=aligns_demo, alt_bg=bool(i % 2))
            _add_footer(ws, r, 4)

            # ────────────────────────────────────────────────────────────────
            # Sheet 7 — Financing
            # ────────────────────────────────────────────────────────────────
            ws = wb_s.create_sheet("Financing")
            _setup_sheet(ws, freeze_at="A6", tab_color="1E3A5F")
            _autosize_columns(ws, {"A":32,"B":40})
            r = _cover_block(ws, "FINANCING SUMMARY", subtitle_x, 2)

            fin_i_x = d.get("financing") or {}
            asd_i_x = fin_i_x.get("assumable_debt") or {}
            nf_i_x  = fin_i_x.get("new_financing")  or {}

            r = _section_header(ws, r, "A.  OFFERING & DEBT CONTACT", 2)
            for i, (k, v) in enumerate([
                ("Offering Type", fin_i_x.get("offering_type")),
                ("Debt Contact",  fin_i_x.get("debt_contact")),
                ("Notes",         fin_i_x.get("notes")),
            ]):
                r = _kv_row(ws, r, k, v, 2, alt_bg=bool(i % 2))

            if any(asd_i_x.get(k) for k in ["lender","loan_amount","interest_rate","loan_type"]):
                ws.row_dimensions[r].height = 6; r += 1
                r = _section_header(ws, r, "B.  ASSUMABLE DEBT", 2)
                rows_asd = [
                    ("Lender",            asd_i_x.get("lender")),
                    ("Loan Type",         asd_i_x.get("loan_type")),
                    ("Loan Amount",       _v(asd_i_x.get("loan_amount"), "$")),
                    ("Interest Rate",     _pct(asd_i_x.get("interest_rate"))),
                    ("Rate Type",         asd_i_x.get("rate_type")),
                    ("LTV",               _pct(asd_i_x.get("loan_to_value"))),
                    ("Interest-Only",     asd_i_x.get("interest_only_period")),
                    ("Maturity Date",     asd_i_x.get("maturity_date")),
                    ("DSCR",              asd_i_x.get("dscr")),
                ]
                for i, (k, v) in enumerate(rows_asd):
                    r = _kv_row(ws, r, k, v, 2, alt_bg=bool(i % 2))

            if any(nf_i_x.get(k) for k in ["lender","loan_amount","interest_rate","loan_type"]):
                ws.row_dimensions[r].height = 6; r += 1
                sec = "C" if any(asd_i_x.get(k) for k in ["lender","loan_amount","interest_rate","loan_type"]) else "B"
                r = _section_header(ws, r, f"{sec}.  NEW FINANCING", 2)
                rows_nf = [
                    ("Lender",        nf_i_x.get("lender")),
                    ("Loan Type",     nf_i_x.get("loan_type")),
                    ("Loan Amount",   _v(nf_i_x.get("loan_amount"), "$")),
                    ("Interest Rate", _pct(nf_i_x.get("interest_rate"))),
                    ("Rate Type",     nf_i_x.get("rate_type")),
                    ("LTV",           _pct(nf_i_x.get("loan_to_value"))),
                    ("Interest-Only", nf_i_x.get("interest_only_period")),
                ]
                for i, (k, v) in enumerate(rows_nf):
                    r = _kv_row(ws, r, k, v, 2, alt_bg=bool(i % 2))

            _add_footer(ws, r, 2)

            # ────────────────────────────────────────────────────────────────
            # Sheet 8 — Flags (dedicated, with wide Detail column for readability)
            # ────────────────────────────────────────────────────────────────
            ws = wb_s.create_sheet("Flags")
            _setup_sheet(ws, freeze_at="A6", tab_color="1E3A5F")
            # Wider Detail column so flag text reads naturally
            _autosize_columns(ws, {"A":14, "B":32, "C":80,
                                    "D":2, "E":2, "F":2})
            r = _cover_block(ws, "UNDERWRITING FLAGS", subtitle_x, 3)
            r = _section_header(ws, r, "RISK FLAGS & UPSIDE OBSERVATIONS", 3)
            r = _column_headers(ws, r, ["Category", "Flag", "Detail"])

            flags_x = d.get("flags") or []
            if flags_x:
                flag_bg_map = {
                    "Warning": C_WARN, "Caution": C_WARN, "Risk": C_WARN,
                    "Opportunity": C_GREEN_L, "Upside": C_GREEN_L,
                    "Info": C_BLUE_L,
                    "Verify": C_PURPLE_L, "Verification": C_PURPLE_L
                }
                flag_fg_map = {
                    "Warning": "FF92400E", "Caution": "FF92400E", "Risk": "FF92400E",
                    "Opportunity": "FF065F46", "Upside": "FF065F46",
                    "Info": "FF1E3A8A",
                    "Verify": "FF5B21B6", "Verification": "FF5B21B6"
                }
                for fl in flags_x:
                    cat = fl.get("category", "Info")
                    bg  = flag_bg_map.get(cat, WHITE)
                    fg  = flag_fg_map.get(cat, BODY)
                    # Colored category chip
                    _set_cell(ws, r, 1, cat.upper(), bold=True, bg=bg,
                              fg=fg, size=9, ha="center", wrap=False,
                              border=_bord_bottom)
                    # Bold navy title on white
                    _set_cell(ws, r, 2, fl.get("title", ""), bold=True,
                              bg=WHITE, fg=HDR, size=10, ha="left",
                              wrap=True, border=_bord_bottom)
                    # Detail body text on white, wrapped
                    _set_cell(ws, r, 3, fl.get("detail", ""),
                              bg=WHITE, fg=BODY, size=9, ha="left",
                              wrap=True, border=_bord_bottom)
                    # Tune row height by detail length
                    detail_len = len(fl.get("detail", ""))
                    approx_lines = max(1, -(-detail_len // 90))
                    ws.row_dimensions[r].height = max(36, approx_lines * 15 + 8)
                    r += 1
            else:
                r = _kv_row(ws, r, "Note", "No flags generated.", 3)

            _add_footer(ws, r, 3)

            # ── Auto-fit row heights on every sheet ──
            for sheet in wb_s.worksheets:
                # Flags sheet has manually-tuned row heights for long detail text;
                # skip auto-fit so heights don't get clipped at 80pt.
                if sheet.title == "Flags":
                    continue
                col_widths = {}
                for letter, dim in sheet.column_dimensions.items():
                    if dim.width:
                        col_widths[letter] = dim.width
                for row_cells in sheet.iter_rows():
                    r_idx = row_cells[0].row
                    if r_idx == 1:
                        continue
                    existing = sheet.row_dimensions[r_idx].height
                    max_lines = 1
                    for cell in row_cells:
                        val = cell.value
                        if val is None or val == "":
                            continue
                        text = str(val)
                        col_w = col_widths.get(cell.column_letter, 12)
                        chars_per_line = max(int(col_w * 1.1), 8)
                        lines_for_cell = 0
                        for segment in text.split("\n"):
                            seg_len = len(segment)
                            if seg_len == 0:
                                lines_for_cell += 1
                            else:
                                lines_for_cell += max(1, -(-seg_len // chars_per_line))
                        if lines_for_cell > max_lines:
                            max_lines = lines_for_cell
                    target = min(max(14, max_lines * 14 + 2), 80)
                    if existing is None or target > existing:
                        sheet.row_dimensions[r_idx].height = target

            buf_s = io.BytesIO()
            wb_s.save(buf_s)
            return buf_s.getvalue()

        # ── Download buttons row: full report + summary tabs ──
        dlc1, dlc2 = st.columns(2)
        with dlc1:
            st.download_button(
                label="⬇️  Download Full Underwriting Report (.xlsx)",
                data=excel_bytes,
                file_name=f"{safe}_Underwriting_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_full_report",
            )
        with dlc2:
            summary_xlsx = _build_summary_tabs_workbook(data)
            st.download_button(
                label="📊  Download All Summary Tabs (.xlsx)",
                data=summary_xlsx,
                file_name=f"{safe}_Summary_Tabs.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_summary_tabs",
            )

        st.markdown("<div style='margin-top:24px;'>", unsafe_allow_html=True)
        pd_  = data.get("property")   or {}
        inv_ = data.get("investment") or {}
        va_  = data.get("value_add")  or {}
        tax_ = data.get("tax")        or {}
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        with m1: st.metric("Units",       _v(pd_.get("units"), "n"))
        with m2: st.metric("Year Built",  _v(pd_.get("year_built")))
        with m3: st.metric("Occupancy",   _pct(pd_.get("occupancy_pct")))
        with m4: st.metric("Avg Rent",    _v(inv_.get("market_rent"), "$"))
        with m5: st.metric("Reno ROI",    _pct(va_.get("roi_pct")))
        with m6: st.metric("Tax Savings", _v(tax_.get("abatement_annual_savings"), "$"))
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='margin-top:20px;'>", unsafe_allow_html=True)
        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
            "Unit Mix", "Value-Add", "Rent Comps", "Financials",
            "Tax & Abatement", "Demographics", "Financing", "Flags"
        ])

        with tab1:
            umix_ = data.get("unit_mix") or []
            if umix_:
                import pandas as pd
                st.dataframe(pd.DataFrame([{
                    "Type": u.get("type"), "Units": u.get("count"), "SF": u.get("sf"),
                    "Market Rent": u.get("market_rent"), "Eff. Rent": u.get("eff_rent"),
                    "Target Rent": u.get("target_rent"), "Upside/Unit": u.get("upside"),
                } for u in umix_]), use_container_width=True, hide_index=True)
            else:
                st.info("No unit mix data extracted.")

        with tab2:
            import pandas as pd
            va_plans = va_.get("by_floor_plan") or []
            if va_plans:
                st.markdown('<div class="gold-header">Floor Plan Value-Add</div>', unsafe_allow_html=True)
                st.dataframe(pd.DataFrame([{
                    "Type": p.get("type"), "SF": p.get("sf"), "Units": p.get("units"),
                    "In-Place Rent": p.get("inplace_rent"), "Rehab Cost": p.get("rehab_cost"),
                    "Premium/Unit": p.get("premium"), "Post-Rehab Rent": p.get("post_rehab_rent"),
                } for p in va_plans]), use_container_width=True, hide_index=True)
            levers_ = data.get("value_add_levers") or []
            if levers_:
                st.markdown('<div class="gold-header">Revenue Upside Levers</div>', unsafe_allow_html=True)
                st.dataframe(pd.DataFrame([{
                    "Lever": lv.get("lever"), "Units": lv.get("units"),
                    "Mo. Premium": lv.get("monthly_premium"), "Annual Upside": lv.get("annual_upside"),
                    "Notes": lv.get("notes"),
                } for lv in levers_]), use_container_width=True, hide_index=True)
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Reno Cost", _v(va_.get("total_cost"), "$"))
                st.metric("Cost Per Unit",   _v(va_.get("cost_per_unit"), "$"))
            with col2:
                st.metric("Annual Premium",  _v(va_.get("annual_premium"), "$"))
                st.metric("Monthly Premium", _v(va_.get("monthly_premium"), "$"))
            with col3:
                st.metric("ROI",             _pct(va_.get("roi_pct")))
                st.metric("Exterior CapEx",  _v(va_.get("exterior_capex"), "$"))

        with tab3:
            import pandas as pd
            rg_     = data.get("rent_comps_garden")    or []
            rth_    = data.get("rent_comps_townhouse") or []
            rcomps_ = data.get("rent_comps")           or []
            if rg_:
                st.markdown('<div class="gold-header">Garden Comps</div>', unsafe_allow_html=True)
                st.dataframe(pd.DataFrame([{"Property": c.get("name"), "Rent": c.get("rent"), "Notes": c.get("notes")} for c in rg_]),
                             use_container_width=True, hide_index=True)
            if rth_:
                st.markdown('<div class="gold-header">Townhouse Comps</div>', unsafe_allow_html=True)
                st.dataframe(pd.DataFrame([{"Property": c.get("name"), "Rent": c.get("rent"), "Notes": c.get("notes")} for c in rth_]),
                             use_container_width=True, hide_index=True)
            if rcomps_:
                st.markdown('<div class="gold-header">Full Comp Detail</div>', unsafe_allow_html=True)
                st.dataframe(pd.DataFrame([{
                    "Property": c.get("name"), "Type": c.get("comp_type"),
                    "Built": c.get("year_built"), "Units": c.get("units"),
                    "Occ": _pct(c.get("occupancy")), "Mkt Rent": _v(c.get("total_market"), "$"),
                    "Avg SF": c.get("avg_sf"),
                } for c in rcomps_]), use_container_width=True, hide_index=True)
            if not rg_ and not rth_ and not rcomps_:
                st.info("No rent comp data extracted.")

        with tab4:
            fin_ = data.get("financials") or {}
            periods_ = fin_.get("periods") or []
            inc_     = fin_.get("income_lines") or []
            exp_     = fin_.get("expense_lines") or []
            noi_     = fin_.get("noi") or {}
            if periods_ and (inc_ or exp_):
                import pandas as pd
                rows = []
                for line in inc_ + exp_:
                    row = {"Line Item": line.get("item")}
                    for p in periods_:
                        row[p] = _v(line.get("values", {}).get(p), "$")
                    rows.append(row)
                if noi_:
                    row = {"Line Item": "NET OPERATING INCOME"}
                    for p in periods_:
                        row[p] = _v(noi_.get(p), "$")
                    rows.append(row)
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            else:
                st.info("No financial data extracted.")

        with tab5:
            tax_ = data.get("tax") or {}
            import pandas as pd
            rows = [(k.replace("_"," ").title(), v) for k, v in tax_.items() if v and k != "abatement_program"]
            if rows:
                st.markdown('<div class="gold-header">Property Tax Detail</div>', unsafe_allow_html=True)
                st.dataframe(pd.DataFrame(rows, columns=["Field","Value"]), use_container_width=True, hide_index=True)
            if tax_.get("abatement_program"):
                st.markdown('<div class="gold-header">Tax Abatement Program</div>', unsafe_allow_html=True)
                st.info(tax_["abatement_program"])
            if not rows and not tax_.get("abatement_program"):
                st.info("No tax data extracted.")

        with tab6:
            demo_ = data.get("demographics") or {}
            import pandas as pd
            cols_d = ["population_3mi","population_5mi","hh_income_3mi","hh_income_5mi",
                      "median_age","renter_pct","college_pct","white_collar"]
            demo_rows = [(k.replace("_"," ").title(), demo_.get(k)) for k in cols_d if demo_.get(k)]
            if demo_rows:
                st.dataframe(pd.DataFrame(demo_rows, columns=["Metric","Value"]), use_container_width=True, hide_index=True)
            else:
                st.info("No demographics data extracted.")

        with tab7:
            fin_i_ = data.get("financing") or {}
            nf_i_  = fin_i_.get("new_financing")  or {}
            asd_i_ = fin_i_.get("assumable_debt") or {}
            import pandas as pd
            rows_f = []
            if asd_i_.get("lender"):
                rows_f += [
                    ("Type",          "Assumable Debt"),
                    ("Lender",        asd_i_.get("lender")),
                    ("Loan Amount",   _v(asd_i_.get("loan_amount"), "$")),
                    ("Interest Rate", _pct(asd_i_.get("interest_rate"))),
                    ("Rate Type",     asd_i_.get("rate_type")),
                    ("IO Period",     asd_i_.get("interest_only_period")),
                    ("Maturity Date", asd_i_.get("maturity_date")),
                ]
            if nf_i_.get("lender") or nf_i_.get("loan_type"):
                rows_f += [
                    ("Type",          "New Financing"),
                    ("Lender",        nf_i_.get("lender")),
                    ("Loan Amount",   _v(nf_i_.get("loan_amount"), "$")),
                    ("Interest Rate", _pct(nf_i_.get("interest_rate"))),
                    ("LTV",           _pct(nf_i_.get("loan_to_value"))),
                    ("IO Period",     nf_i_.get("interest_only_period")),
                ]
            if rows_f:
                st.dataframe(pd.DataFrame([(k,v) for k,v in rows_f if v],
                             columns=["Field","Value"]), use_container_width=True, hide_index=True)
            else:
                st.info(f"Offering type: {fin_i_.get('offering_type') or 'Not specified'}")

        with tab8:
            flags_ = data.get("flags") or []
            if flags_:
                for fl in flags_:
                    cat  = (fl.get("category") or "").lower()
                    cls  = "flag-warn"   if cat in ("warning","caution","risk") else \
                           "flag-good"   if cat in ("opportunity","upside") else \
                           "flag-verify" if cat == "verify" else "flag-info"
                    icon = "⚠️" if cls=="flag-warn" else "\u2705" if cls=="flag-good" else "\U0001f50d" if cls=="flag-verify" else "\u2139️"
                    st.markdown(f"""
<div class="{cls}">
  <div class="flag-title">{icon} &nbsp;{fl.get('title','')}</div>
  <div class="flag-body">{fl.get('detail','')}</div>
</div>""", unsafe_allow_html=True)
            else:
                st.info("No underwriting flags extracted.")

        st.markdown("</div>", unsafe_allow_html=True)
