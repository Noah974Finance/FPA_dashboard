import streamlit as st
import time
import requests
import urllib.parse
import extra_streamlit_components as stx

def supabase_request(method, table, params=None, json_data=None):
    url = st.secrets["supabase"]["url"] + "/rest/v1/" + table
    headers = {
        "apikey": st.secrets["supabase"]["key"],
        "Authorization": f"Bearer {st.secrets['supabase']['key']}",
        "Content-Type": "application/json",
        "Prefer": "return=representation"
    }
    if method == "GET":
        return requests.get(url, headers=headers, params=params).json()
    elif method == "POST":
        return requests.post(url, headers=headers, json=json_data).json()
    elif method == "PATCH":
        return requests.patch(url, headers=headers, params=params, json=json_data).json()

def get_cookie_manager():
    return stx.CookieManager(key="auth_cookies")

def get_oauth_config():
    try:
        client_id = st.secrets["google"]["client_id"]
        client_secret = st.secrets["google"]["client_secret"]
        redirect_uri = st.secrets["google"].get("redirect_uri", "http://localhost:8501")
        return True, client_id, client_secret, redirect_uri
    except Exception:
        return False, None, None, None

def auto_login_google_user(email):
    res = supabase_request("GET", "users", params={"select": "id", "email": f"eq.{email}"})
    
    if not res:
        # Create user if they don't exist
        supabase_request("POST", "users", json_data={
            "email": email,
            "password_hash": "google_oauth",
            "is_verified": 1
        })
        
    st.session_state.authenticated = True
    st.session_state.user_email = email
    get_cookie_manager().set("auth_token", email, key="set_login_cookie_google", expires_at=None)

def handle_google_oauth():
    enabled, client_id, client_secret, redirect_uri = get_oauth_config()
    if not enabled: return
    
    if "code" in st.query_params:
        code = st.query_params["code"]
        token_url = "https://oauth2.googleapis.com/token"
        data = {
            "code": code,
            "client_id": client_id,
            "client_secret": client_secret,
            "redirect_uri": redirect_uri,
            "grant_type": "authorization_code",
        }
        res = requests.post(token_url, data=data)
        if res.status_code == 200:
            access_token = res.json().get("access_token")
            user_info_url = "https://www.googleapis.com/oauth2/v2/userinfo"
            headers = {"Authorization": f"Bearer {access_token}"}
            user_res = requests.get(user_info_url, headers=headers)
            if user_res.status_code == 200:
                email = user_res.json().get("email")
                if email:
                    auto_login_google_user(email)
                    st.query_params.clear()
                    st.rerun()
                else:
                    st.error("Impossible de lire l'e-mail depuis Google.")
            else:
                st.error("Échec de la récupération du profil Google.")
        else:
            st.error("Échec de l'authentification Google.")
        st.query_params.clear()

def render_auth_ui():
    cookie_manager = get_cookie_manager()
    auth_token = cookie_manager.get(cookie="auth_token")
    
    if st.session_state.get("force_logout", False):
        cookie_manager.delete("auth_token", key="delete_cookie_auth")
        st.session_state.force_logout = False
        auth_token = None

    if auth_token and not st.session_state.get("authenticated", False):
        st.session_state.authenticated = True
        st.session_state.user_email = auth_token
        time.sleep(0.1)
        st.rerun()

    enabled, client_id, client_secret, redirect_uri = get_oauth_config()
    
    if enabled:
        handle_google_oauth()
        
    st.markdown("""
    <div style="text-align:center;padding:40px 0 20px">
      <div style="font-size:48px">🔒</div>
      <h1 style="font-size:2rem;font-weight:800;color:var(--text-color);margin:8px 0">Authentification Sécurisée</h1>
      <p style="color:var(--text-color);opacity:0.7">Connectez-vous pour accéder au Dashboard FP&A</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        if enabled:
            auth_url = f"https://accounts.google.com/o/oauth2/v2/auth?response_type=code&client_id={client_id}&redirect_uri={urllib.parse.quote(redirect_uri)}&scope=openid%20email%20profile"
            st.link_button("🌐 Continuer avec Google", auth_url, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("👤 Continuer en tant qu'invité", use_container_width=True):
            import uuid
            st.session_state.authenticated = True
            st.session_state.user_email = f"guest_{uuid.uuid4().hex[:8]}"
            st.rerun()
