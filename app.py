"""
TDR Portal – Single-app entry. Choose Synthetic data (TDR mapping) or Production data (LVT TDR Delivery).
Navigation uses session state so it works on Streamlit Cloud (no switch_page / page_link).
"""
import streamlit as st

# Version: update only when pushing. Minor (GUI/small fixes): 2.1 → 2.2 → 2.3. Major (new section/feature): 3.0, 4.0, …
PORTAL_VERSION = "2.1"
VERSION_NOTE = "Production: Full bulk or TDR list only"
CREATED_BY = "Ravikumar Garg"
CREATED_BY_EMAIL = "ravikumg@amdocs.com"

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
        <div style="font-size: 0.65rem; opacity: 0.7; margin-top: 4px;">v""" + PORTAL_VERSION + """ — """ + VERSION_NOTE + """</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("**Select an option below to open the respective tool.**")
st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

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

st.markdown("<div style='height: 24px;'></div>", unsafe_allow_html=True)
st.markdown(
    f"""
    <div style="text-align: center;">
        <hr style="margin: 16px 0;" />
        <div style="font-size: 0.9rem; color: #334155;"><em>TDR wise mapping</em>: TDR-wise mapping for synthetic/test data.</div>
        <div style="font-size: 0.9rem; color: #334155;"><em>Bulk data mapping</em>: LVT TDR mapping and INSERT SQL for BAN Master table (run SQL manually).</div>
        <div style="height: 32px;"></div>
        <div style="color: #64748b; font-size: 0.8rem;">
            <div>Created by</div>
            <div>{CREATED_BY}</div>
            <div>email — {CREATED_BY_EMAIL}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
