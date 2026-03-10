"""
Minimal app for Streamlit Cloud debugging.
If this runs, deployment is OK; set main file back to app.py to use the full app.
"""
import streamlit as st

st.set_page_config(page_title="TDR check", page_icon="📋", layout="centered")
st.write("TDR Portal – deployment OK. Switch main file to **app.py** for the full app.")
