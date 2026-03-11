"""
TDR Portal – launcher. Shows "Loading..." then runs app_main so any error is caught and displayed.
"""
import sys
import traceback
import streamlit as st


def _python_version():
    """Return short Python version string (e.g. 3.12.0) for debugging."""
    return f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"

try:
    from streamlit.runtime.scriptrunner_utils.exceptions import RerunException
except ImportError:
    RerunException = type("RerunException", (BaseException,), {})

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
    if isinstance(e, RerunException) or type(e).__name__ == "RerunException":
        raise  # st.rerun() uses this; let it propagate
    traceback.print_exception(type(e), e, e.__traceback__, file=sys.stderr)
    st.error(f"**App error:** {e}")
    with st.expander("Technical details", expanded=True):
        st.code(traceback.format_exc(), language="text")
    st.info(f"**Runtime:** Python {_python_version()}. In App settings → General, try changing **Python version** (e.g. to 3.12 or 3.13) and Save to force a fresh redeploy. Check **Logs** in the app menu for the full traceback.")
