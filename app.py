"""
TDR Portal – launcher. Shows "Loading..." then runs app_main so any error is caught and displayed.
"""
import sys
import traceback
import streamlit as st

try:
    st.set_page_config(page_title="TDR Portal", page_icon="📋", layout="centered", initial_sidebar_state="expanded")
except Exception:
    pass

# Show something immediately so we have output before importing/running the rest
st.write("Loading…")

try:
    from app_main import run
    run()
except BaseException as e:
    traceback.print_exception(type(e), e, e.__traceback__, file=sys.stderr)
    st.error(f"**App error:** {e}")
    with st.expander("Technical details", expanded=True):
        st.code(traceback.format_exc(), language="text")
    st.info("Set **Python 3.12** in App settings → General. If this keeps happening, check **Logs** in the app menu.")
