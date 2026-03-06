"""Production data – LVT TDR Delivery. Uses same view as app.py Production (radio + full/tdr_only)."""
import os
import sys

import streamlit as st

# Ensure app root is on path so lvt_tdr_core and streamlit_views are importable
_app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _app_root not in sys.path:
    sys.path.insert(0, _app_root)

st.set_page_config(page_title="Production – LVT TDR Delivery", page_icon="📋", layout="centered", initial_sidebar_state="expanded")

# Back to Portal
st.sidebar.markdown("[← Back to TDR Portal](/)")

# Single source of truth: use the same Production UI as app.py (radio + full bulk / TDR-only)
from streamlit_views.production import render_production
render_production()
