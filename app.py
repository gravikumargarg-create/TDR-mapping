"""
TDR Portal – Choose Synthetic data (TDR mapping) or Production data (LVT TDR Delivery).
"""
import streamlit as st

st.set_page_config(page_title="TDR Portal", page_icon="📋", layout="centered", initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
    .stApp { background: #f1f5f9 !important; }
    .block-container { padding: 2rem 1.5rem !important; max-width: 720px !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

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
    if hasattr(st, "page_link"):
        st.page_link("pages/1_Synthetic_TDR.py", label="Synthetic data – TDR mapping sheet creation", icon="📋")
    else:
        st.markdown("[**Synthetic data – TDR mapping sheet creation**](pages/1_Synthetic_TDR)")
    st.caption("Upload TDR data + LVT report → mapping and TDR-wise reports.")

with col2:
    if hasattr(st, "page_link"):
        st.page_link("pages/2_Production_LVT.py", label="Production data – LVT TDR Delivery", icon="📋")
    else:
        st.markdown("[**Production data – LVT TDR Delivery**](pages/2_Production_LVT)")
    st.caption("Upload LVT + data Excel files → report + INSERT SQL (no DB).")

st.markdown("<div style='height: 24px;'></div>", unsafe_allow_html=True)
st.markdown(
    "---  \n*Synthetic*: TDR mapping for synthetic/test data.  \n*Production*: LVT TDR mapping and INSERT SQL for production (run SQL manually).  \n\nYou can also use the **sidebar** to open **1_Synthetic_TDR** or **2_Production_LVT** directly."
)
