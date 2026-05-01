from supabase import create_client
import streamlit as st


def init_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)


def login_page():
    st.markdown("""
<style>
#MainMenu, footer, header { visibility: hidden; }
html, body, [class*="css"], .stApp, .main {
    background-color: #0A1628 !important;
    color: #E0E6EF !important;
}
.block-container {
    padding: 0 !important;
    max-width: 480px !important;
    margin: auto !important;
    background: #0A1628 !important;
}
/* Navbar */
.om-navbar {
    background: #0B1929;
    border-bottom: 1px solid #1E3148;
    padding: 0 48px;
    height: 100px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    width: 100%;
    box-sizing: border-box;
}
.om-nav-left { display: flex; align-items: center; gap: 16px; }
.om-nav-icon {
    width: 58px; height: 58px;
    background: #1DC9A4; border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    font-weight: 800; font-size: 20px; color: #0D1B2A;
}
.om-nav-title { font-size: 24px; font-weight: 700; color: #FFFFFF; line-height: 1.2; }
.om-nav-sub {
    font-size: 10px; color: #5A8FAA;
    letter-spacing: 0.14em; text-transform: uppercase; margin-top: 3px;
}
.om-secure-badge {
    border: 1.5px solid #1DC9A4; color: #1DC9A4;
    border-radius: 20px; padding: 6px 18px;
    font-size: 11px; font-weight: 700; letter-spacing: 0.1em;
    text-transform: uppercase;
}
/* Card */
.om-card-top {
    background: #0F1E30;
    border: 1px solid #1A3250;
    border-radius: 16px 16px 0 0;
    padding: 32px 36px 24px;
    text-align: center;
    margin-top: 56px;
}
.om-card-icon {
    width: 60px; height: 60px;
    background: #1DC9A4; border-radius: 14px;
    display: inline-flex; align-items: center; justify-content: center;
    font-weight: 800; font-size: 20px; color: #0D1B2A;
    margin-bottom: 14px;
}
.om-card-title { font-size: 21px; font-weight: 700; color: #FFFFFF; margin-bottom: 5px; }
.om-card-sub { font-size: 13px; color: #5A8FAA; }
.om-form-wrap {
    background: #0A1628;
    border: 1px solid #1A3250;
    border-top: none;
    border-radius: 0 0 16px 16px;
    padding: 24px 36px 28px;
}
/* Input fields */
.stTextInput > div > div > input {
    background: #0F1E30 !important;
    border: 1px solid #1E3A55 !important;
    border-radius: 10px !important;
    color: #E0E6EF !important;
    font-size: 14px !important;
    padding: 12px 14px !important;
    height: 46px !important;
}
.stTextInput > div > div > input::placeholder { color: #3A6080 !important; }
.stTextInput > div > div > input:focus {
    border-color: #1DC9A4 !important;
    box-shadow: 0 0 0 2px rgba(29,201,164,0.15) !important;
    outline: none !important;
}
.stTextInput label {
    color: #7A9AB8 !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    margin-bottom: 2px !important;
}
/* Button */
.stButton > button {
    background: #1DC9A4 !important;
    color: #0D1B2A !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    padding: 12px 0 !important;
    width: 100% !important;
    margin-top: 4px !important;
}
.stButton > button:hover { background: #18B090 !important; }
/* Alerts */
div[data-testid="stAlert"] {
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

    # ── Centered login card ──
    _, col, _ = st.columns([1, 1.4, 1])
    with col:
        st.markdown("""
<div class="om-card-top">
  <div class="om-card-icon">OM</div>
  <div class="om-card-title">Sign In to OM Analyzer</div>
  <div class="om-card-sub">Authorized analysts only</div>
</div>
<div class="om-form-wrap">
""", unsafe_allow_html=True)

        email    = st.text_input("Email", placeholder="analyst@yourfirm.com", key="login_email")
        password = st.text_input("Password", type="password", placeholder="Enter password...", key="login_password")

        st.markdown("<div style='height:2px'></div>", unsafe_allow_html=True)

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
<div style="margin-top:18px; padding-top:14px; border-top:1px solid #1A3250;
            text-align:center; font-size:11px; color:#2A4A60; line-height:1.9;">
  Access is by invitation only.<br>
  Contact your administrator to request access.
</div>
</div>
""", unsafe_allow_html=True)


def logout():
    st.session_state.clear()
    st.rerun()
