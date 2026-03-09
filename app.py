"""
TDR Portal – Single-app entry. Choose Synthetic data (TDR mapping) or Production data (LVT TDR Delivery).
Navigation uses session state so it works on Streamlit Cloud (no switch_page / page_link).
"""
import streamlit as st

# ---------------------------------------------------------------------------
# VERSION: Update this before every push so the footer shows the correct release.
# Current: 3.12  →  Bump (e.g. to 3.13) whenever you push changes.
# ---------------------------------------------------------------------------
PORTAL_VERSION = "3.12"
CREATED_BY = "Ravikumar Garg"
CREATED_BY_EMAIL = "ravikumg@amdocs.com"


def _version_label():
    """Show only the app version (e.g. v3.7). Bump PORTAL_VERSION in this file when you release."""
    return f"v{PORTAL_VERSION}"

st.set_page_config(page_title="TDR Portal", page_icon="📋", layout="centered", initial_sidebar_state="expanded")

# Which view we're showing (portal | synthetic | production)
if "portal_view" not in st.session_state:
    st.session_state.portal_view = "portal"

st.markdown(
    """
    <style>
    .stApp { background: linear-gradient(160deg, #e0f2f1 0%, #f1f5f9 50%, #fef3c7 100%) !important; min-height: 100vh; }
    .block-container { padding: 2rem 1.5rem 5rem !important; max-width: 760px !important; }
    #portal-footer { position: fixed; bottom: 0; left: 0; right: 0; text-align: center; background: rgba(241,245,249,0.95); padding: 4px 0; font-size: 0.65rem; line-height: 1.35; color: #64748b; backdrop-filter: blur(6px); border-top: 1px solid rgba(15, 118, 110, 0.12); }
    .portal-hero { background: linear-gradient(135deg, #0f766e 0%, #0d9488 35%, #0ea5e9 100%) !important; color: #fff; padding: 28px 28px; border-radius: 16px; margin-bottom: 28px; text-align: center; box-shadow: 0 10px 40px rgba(15, 118, 110, 0.35), 0 0 0 1px rgba(255,255,255,0.1) inset; }
    .portal-hero h1 { font-size: 1.75rem; font-weight: 800; margin: 0 0 8px 0; letter-spacing: -0.02em; }
    .portal-hero .sub { font-size: 0.85rem; opacity: 0.95; }
    .portal-hero .ver { display: inline-block; font-size: 0.7rem; opacity: 0.9; background: rgba(255,255,255,0.2); padding: 4px 10px; border-radius: 20px; margin-top: 10px; }
    .portal-sub { font-size: 0.95rem; color: #475569; font-weight: 600; margin-bottom: 20px; }
    .portal-card { background: #fff; border-radius: 14px; padding: 24px; margin-bottom: 20px; box-shadow: 0 4px 20px rgba(0,0,0,0.06), 0 0 0 1px rgba(0,0,0,0.04); transition: transform 0.2s, box-shadow 0.2s; border-left: 4px solid #0d9488; }
    .portal-card:hover { box-shadow: 0 8px 28px rgba(0,0,0,0.1), 0 0 0 1px rgba(0,0,0,0.06); }
    .portal-card.bulk { border-left-color: #ea580c; }
    .stButton > button[kind="primary"] { background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%) !important; color: #fff !important; border: none !important; border-radius: 12px !important; padding: 12px 24px !important; font-weight: 600 !important; box-shadow: 0 4px 14px rgba(13, 148, 136, 0.4) !important; transition: transform 0.2s, box-shadow 0.2s !important; }
    .stButton > button[kind="primary"]:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(13, 148, 136, 0.45) !important; }
    [data-testid="column"]:first-of-type .stButton > button[kind="primary"] { background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%) !important; }
    [data-testid="column"]:last-of-type .stButton > button[kind="primary"] { background: linear-gradient(135deg, #ea580c 0%, #c2410c 100%) !important; box-shadow: 0 4px 14px rgba(234, 88, 12, 0.35) !important; }
    [data-testid="column"]:last-of-type .stButton > button[kind="primary"]:hover { box-shadow: 0 6px 20px rgba(234, 88, 12, 0.4) !important; }
    [data-testid="column"] { background: #fff; border-radius: 14px; padding: 22px 20px !important; margin: 0 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.06), 0 0 0 1px rgba(0,0,0,0.04); border-left: 4px solid #0d9488; }
    [data-testid="column"]:last-of-type { border-left-color: #ea580c; }
    section[data-testid="stSidebar"] .stButton:first-of-type > button { background: #0d9488 !important; color: #fff !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

def _clear_view_state():
    """Clear result and detection state for the current view so it resets when re-entered."""
    if st.session_state.portal_view == "synthetic":
        for key in ("tdr_result", "_detected", "_last_detection_key"):
            st.session_state.pop(key, None)
    elif st.session_state.portal_view == "production":
        st.session_state.pop("lvt_result", None)
        st.session_state.pop("tdr_list_result", None)
        st.session_state.pop("cap_validation_result", None)
        st.session_state.pop("cap_download", None)
        st.session_state.pop("cap_removed_bytes", None)
        st.session_state.pop("cap_highlighted_bytes", None)
        st.session_state.pop("cap_validation_key", None)


# ----- Back button in sidebar and sub-view content -----
if st.session_state.portal_view != "portal":
    with st.sidebar:
        if st.button("← Back to TDR Portal", key="back_to_portal", type="secondary", use_container_width=True):
            _clear_view_state()
            st.session_state.portal_view = "portal"
            st.rerun()

    if st.session_state.portal_view == "synthetic":
        from streamlit_views.synthetic import render_synthetic
        render_synthetic()
    elif st.session_state.portal_view == "production":
        from streamlit_views.production import render_production
        render_production()
    st.markdown(
        f"""
        <div id="portal-footer">
            <div>Created by: {CREATED_BY}</div>
            <div>email — {CREATED_BY_EMAIL}</div>
            <div>v{PORTAL_VERSION}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

# ----- Portal home -----
st.markdown(
    """
    <div class="portal-hero">
        <h1>TDR Portal</h1>
        <div class="sub">Choose your tool: TDR wise mapping or Bulk data mapping</div>
    </div>
    <div class="portal-sub">Select an option below to open the respective tool.</div>
    """,
    unsafe_allow_html=True,
)

col1, col2 = st.columns(2)

with col1:
    if st.button("📋 **TDR wise mapping**", use_container_width=True, type="primary", key="btn_synthetic"):
        st.session_state.portal_view = "synthetic"
        st.rerun()
    st.caption("TDR-wise mapping for synthetic data. Inputs needed: TDR data sheets, device details, and LVT report.")

with col2:
    if st.button("📋 **Bulk data mapping**", use_container_width=True, type="primary", key="btn_production"):
        st.session_state.portal_view = "production"
        st.rerun()
    st.caption("Bulk mapping for both production and synthetic data, with INSERT query creation for BAN Master table. Inputs needed: TDR data, LVT report, and capability reports.")

st.markdown(
    f"""
    <div id="portal-footer">
        <div>Created by: {CREATED_BY}</div>
        <div>email — {CREATED_BY_EMAIL}</div>
        <div>{_version_label()}</div>
    </div>
    """,
    unsafe_allow_html=True,
)
