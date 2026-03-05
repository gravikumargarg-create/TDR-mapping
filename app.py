"""
TDR Portal – Single-app entry. Choose Synthetic data (TDR mapping) or Production data (LVT TDR Delivery).
Navigation uses session state so it works on Streamlit Cloud (no switch_page / page_link).
"""
import streamlit as st

st.set_page_config(page_title="TDR Portal", page_icon="📋", layout="centered", initial_sidebar_state="expanded")

# Which view we're showing (portal | synthetic | production)
if "portal_view" not in st.session_state:
    st.session_state.portal_view = "portal"

st.markdown(
    """
    <style>
    .stApp { background: #f1f5f9 !important; }
    .block-container { padding: 2rem 1.5rem !important; max-width: 720px !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----- Back button and sub-view content -----
if st.session_state.portal_view != "portal":
    if st.button("← Back to TDR Portal", key="back_to_portal", type="secondary"):
        st.session_state.portal_view = "portal"
        st.rerun()
    st.sidebar.markdown("**TDR Portal**")
    if st.sidebar.button("← Back to portal", key="sidebar_back"):
        st.session_state.portal_view = "portal"
        st.rerun()

    if st.session_state.portal_view == "synthetic":
        from streamlit_views.synthetic import render_synthetic
        render_synthetic()
    elif st.session_state.portal_view == "production":
        from streamlit_views.production import render_production
        render_production()
    st.stop()

# ----- Portal home -----
st.markdown(
    """
    <div style="
        background: linear-gradient(90deg, #0f766e 0%, #0d9488 50%, #14b8a6 100%);
        color: #fff; padding: 20px 24px; border-radius: 10px; margin-bottom: 24px;
        box-shadow: 0 2px 8px rgba(15, 118, 110, 0.3); text-align: center;
    ">
        <div style="font-size: 1.35rem; font-weight: 700; margin-bottom: 6px;">TDR Portal</div>
        <div style="font-size: 0.8rem; opacity: 0.95;">Choose your tool: Synthetic data or Production data</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("**Select an option below to open the respective tool.**")
st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    if st.button("📋 **Synthetic data** – TDR mapping sheet creation", use_container_width=True, type="primary", key="btn_synthetic"):
        st.session_state.portal_view = "synthetic"
        st.rerun()
    st.caption("Upload TDR data + LVT report → mapping and TDR-wise reports.")

with col2:
    if st.button("📋 **Production data** – LVT TDR Delivery", use_container_width=True, type="primary", key="btn_production"):
        st.session_state.portal_view = "production"
        st.rerun()
    st.caption("Upload LVT + data Excel files → report + INSERT SQL (no DB).")

st.markdown("<div style='height: 24px;'></div>", unsafe_allow_html=True)
st.markdown(
    "---  \n*Synthetic*: TDR mapping for synthetic/test data.  \n*Production*: LVT TDR mapping and INSERT SQL for production (run SQL manually)."
)
