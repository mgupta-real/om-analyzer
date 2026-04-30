from supabase import create_client
import streamlit as st


def init_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)


def login_page():
    st.markdown("""
<style>
/* ── Reset & base ── */
#MainMenu, footer, header { visibility: hidden; }
html, body, [class*="css"], .stApp, .main {
    background-color: #0A1628 !important;
    color: #E0E6EF !important;
}
.block-container {
    padding: 0 !important;
    max-width: 100% !important;
    background: #0A1628 !important;
}

/* ── Top navbar ── */
.om-navbar {
    background: #0B1929;
    border-bottom: 1px solid #1E3148;
    padding: 0 48px;
    height: 110px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.om-nav-left { display: flex; align-items: center; gap: 18px; }
.om-nav-icon {
    width: 64px; height: 64px;
    background: #1DC9A4; border-radius: 16px;
    display: flex; align-items: center; justify-content: center;
    font-weight: 800; font-size: 22px; color: #0D1B2A;
    flex-shrink: 0;
}
.om-nav-title { font-size: 26px; font-weight: 700; color: #FFFFFF; line-height: 1.2; }
.om-nav-sub {
    font-size: 11px; font-weight: 500; color: #5A8FAA;
    letter-spacing: 0.14em; text-transform: uppercase; margin-top: 4px;
}
.om-secure-badge {
    border: 1.5px solid #1DC9A4; color: #1DC9A4;
    border-radius: 20px; padding: 7px 20px;
    font-size: 11px; font-weight: 700; letter-spacing: 0.1em;
    text-transform: uppercase;
}

/* ── Card top ── */
.om-card-top {
    background: #0F1E30;
    border: 1px solid #1A3250;
    border-radius: 16px 16px 0 0;
    padding: 36px 40px 28px;
    text-align: center;
    margin-top: 60px;
}
.om-card-icon {
    width: 64px; height: 64px;
    background: #1DC9A4; border-radius: 16px;
    display: inline-flex; align-items: center; justify-content: center;
    font-weight: 800; font-size: 22px; color: #0D1B2A;
    margin-bottom: 16px;
}
.om-card-title { font-size: 22px; font-weight: 700; color: #FFFFFF; margin-bottom: 6px; }
.om-card-sub { font-size: 13px; color: #5A8FAA; }

/* ── Form wrap ── */
.om-form-wrap {
    background: #0A1628;
    border: 1px solid #1A3250;
    border-top: none;
    border-radius: 0 0 16px 16px;
    padding: 28px 40px 32px;
}

/* ── Input overrides ── */
.stTextInput > div > div > input {
    background: #0F1E30 !important;
    border: 1px solid #1E3A55 !important;
    border-radius: 10px !important;
    color: #E0E6EF !important;
    font-size: 14px !important;
    padding: 12px 16px !important;
    height: 48px !important;
}
.stTextInput > div > div > input::placeholder { color: #3A6080 !important; }
.stTextInput > div > div > input:focus {
    border-color: #1DC9A4 !important;
    box-shadow: 0 0 0 2px rgba(29,201,164,0.15) !important;
}
.stTextInput label {
    color: #7A9AB8 !important;
    font-size: 13px !important;
    font-weight: 500 !important;
}

/* ── Sign In button ── */
.stButton > button {
    background: #1DC9A4 !important;
    color: #0D1B2A !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    padding: 13px 0 !important;
    width: 100% !important;
    margin-top: 6px !important;
    letter-spacing: 0.02em !important;
}
.stButton > button:hover { background: #18B090 !important; }

.stAlert {
    background: #0F2133 !important;
    border-color: #1E3148 !important;
    color: #C0D0E0 !important;
    border-radius: 8px !important;
}
</style>
""", unsafe_allow_html=True)

    # ── Navbar ──
    st.markdown("""
<div class="om-navbar">
  <div class="om-nav-left">
    <div class="om-nav-icon">OM</div>
    <div>
      <div class="om-nav-title">OM Analyzer</div>
      <div class="om-nav-sub">RealVal &nbsp;·&nbsp; Multifamily Underwriting Intelligence</div>
    </div>
  </div>
  <div class="om-secure-badge">🔒 &nbsp;SECURE ACCESS</div>
</div>
""", unsafe_allow_html=True)

    # ── Card top (pure HTML, centered via columns) ──
    _, center, _ = st.columns([1, 1.6, 1])
    with center:
        st.markdown("""
<div class="om-card-top">
  <div class="om-card-icon">OM</div>
  <div class="om-card-title">Sign In to OM Analyzer</div>
  <div class="om-card-sub">Authorized analysts only</div>
</div>
<div class="om-form-wrap">
""", unsafe_allow_html=True)

        email    = st.text_input("Email", placeholder="analyst@yourfirm.com",
                                 key="login_email", label_visibility="visible")
        password = st.text_input("Password", type="password",
                                 placeholder="Enter password...",
                                 key="login_password", label_visibility="visible")

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

        if st.button("Sign In →", use_container_width=True, key="login_btn"):
            if not email or not password:
                st.error("Please enter both email and password.")
            else:
                try:
                    supabase = init_supabase()
                    response = supabase.auth.sign_in_with_password({
                        "email": email,
                        "password": password
                    })
                    st.session_state["user"] = response.user
                    st.session_state["access_token"] = response.session.access_token
                    st.rerun()
                except Exception:
                    st.error("Invalid email or password. Contact your administrator for access.")

        st.markdown("""
<div style="margin-top:20px; padding-top:16px; border-top:1px solid #1A3250;
            text-align:center; font-size:11px; color:#2A4A60; line-height:1.8;">
  Access is by invitation only.<br>
  Contact your administrator to request access.
</div>
</div>
""", unsafe_allow_html=True)


def logout():
    st.session_state.clear()
    st.rerun()
