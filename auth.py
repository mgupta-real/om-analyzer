from supabase import create_client
import streamlit as st

def init_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

def login_page():
    st.title("🔐 Login")

    email = st.text_input("Email")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if not email or not password:
            st.error("Please enter both email and password.")
            return

        supabase = init_supabase()
        try:
            response = supabase.auth.sign_in_with_password({
                "email": email,
                "password": password
            })
            st.session_state["user"] = response.user
            st.session_state["access_token"] = response.session.access_token
            st.rerun()
        except Exception as e:
            st.error("Invalid email or password.")

def logout():
    st.session_state.clear()
    st.rerun()
