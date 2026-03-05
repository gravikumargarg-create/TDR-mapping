"""Synthetic data – TDR mapping sheet creation."""
import os
import sys
import tempfile
import zipfile
from datetime import datetime
from io import BytesIO

import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import tdr_core
import sharepoint_graph

st.set_page_config(page_title="Synthetic – TDR mapping", page_icon="📋", layout="centered")

st.markdown('[← Back to TDR Portal](/)')
st.markdown(
    """
    <style>
    .stApp { background: #f1f5f9 !important; }
    .block-container { padding: 1.75rem 1.5rem !important; max-width: 880px !important; }
    section[data-testid="stFileUploader"] { background: #fff !important; border-radius: 8px !important; border-top: 3px solid #0d9488 !important; }
    .stButton > button[kind="primary"] { background: #0d9488 !important; color: #fff !important; border-radius: 999px !important; }
    div[data-testid="stDownloadButton"] > button { border-radius: 8px !important; border: 1px solid #0d9488 !important; color: #0d9488 !important; }
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown(
    """
    <div style="background: linear-gradient(90deg, #0f766e 0%, #0d9488 50%, #14b8a6 100%); color: #fff; padding: 18px 20px; border-radius: 10px; margin-bottom: 12px; text-align: center;">
        <div style="font-size: 1.25rem; font-weight: 700;">TDR mapping sheet creation (Synthetic data)</div>
        <div style="font-size: 0.75rem; opacity: 0.95;">Upload TDR data + LVT report → mapping and TDR-wise reports.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

TDR_SHAREPOINT_URL = (
    "https://amdocs.sharepoint.com/:f:/r/sites/USCCTesting_Offshore/Shared%20Documents/"
    "Release%20%26%20OffCycle/UTMO%20-%20Migration%20(Data%20creation)/Data%20Creation/R2%20Data"
    "?csf=1&web=1&e=EZDfgA"
)
tdr_bytes = None
tdr_name = None
tdr_sheet = None
tdr_source = st.radio("TDR file from", options=["Local file", "SharePoint"], horizontal=True, key="tdr_source")
use_sharepoint_direct = tdr_source == "SharePoint" and sharepoint_graph.has_sharepoint_credentials()
if tdr_source == "SharePoint" and not use_sharepoint_direct:
    st.markdown(f'<a href="{TDR_SHAREPOINT_URL}" target="_blank" style="font-size: 0.85rem; color: #0d9488;">📂 Open TDR folder on SharePoint</a>', unsafe_allow_html=True)
    st.caption("Download the file from SharePoint, then upload it in the left box below.")

if use_sharepoint_direct:
    token = sharepoint_graph.get_token()
    if not token:
        st.warning("Could not get SharePoint token. Check AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET in secrets.")
    else:
        files = sharepoint_graph.list_tdr_excel_files(token)
        if not files:
            st.info("No Excel files found in the TDR folder.")
        else:
            selected_name = st.selectbox("TDR file (SharePoint)", options=[f["name"] for f in files], key="sp_tdr_file")
            selected = next((f for f in files if f["name"] == selected_name), None)
            if selected:
                cache_key = f"sp_tdr_{selected['id']}"
                if cache_key not in st.session_state:
                    with st.spinner("Loading file…"):
                        content = sharepoint_graph.download_file_content(token, selected["drive_id"], selected["id"])
                    if content:
                        st.session_state[cache_key] = content
                content = st.session_state.get(cache_key) if cache_key in st.session_state else None
                if content:
                    tdr_bytes = content
                    tdr_name = selected["name"]
                    try:
                        wb_tdr = tdr_core.load_workbook(BytesIO(content), read_only=True)
                        tdr_sheet_names = wb_tdr.sheetnames
                        wb_tdr.close()
                        if tdr_sheet_names:
                            tdr_sheet = st.selectbox("Sheet (TDR)", options=tdr_sheet_names, index=0, key="tdr_sheet")
                    except Exception as e:
                        st.warning(str(e))
    st.markdown("**Upload file(s) for LVT Report and/or Device Details**")
    upload_files = st.file_uploader("Excel file(s)", type=["xlsx", "xlsm"], accept_multiple_files=True, key="multi_upload")
else:
    st.markdown("**Upload file(s) – we'll detect TDR Data, LVT Report, and Device Details**")
    upload_files = st.file_uploader("Excel file(s)", type=["xlsx", "xlsm"], accept_multiple_files=True, key="multi_upload")

tdr_file = None
lvt_file = None
device_file = None
detection_cache_key = "tdr_detection_" + (str(len(upload_files)) if upload_files else "0") + "_" + "_".join((f.name + str(f.size) for f in upload_files)) if upload_files else ""
if upload_files:
    if detection_cache_key != st.session_state.get("_last_detection_key"):
        detected = {"tdr": None, "lvt": None, "device": None}
        for uf in upload_files:
            try:
                buf = BytesIO(uf.getvalue())
                wb = tdr_core.load_workbook(buf, read_only=True, data_only=True)
                roles = tdr_core.detect_excel_roles(wb)
                wb.close()
                if roles["tdr_sheets"] and detected["tdr"] is None:
                    detected["tdr"] = (uf, roles["tdr_sheets"])
                if roles["lvt_sheets"] and detected["lvt"] is None:
                    detected["lvt"] = (uf, roles["lvt_sheets"])
                if roles["device_sheets"] and detected["device"] is None:
                    detected["device"] = (uf, roles["device_sheets"])
            except Exception:
                pass
        st.session_state["_detected"] = detected
        st.session_state["_last_detection_key"] = detection_cache_key
    detected = st.session_state.get("_detected") or {"tdr": None, "lvt": None, "device": None}
    if not use_sharepoint_direct:
        tdr_file, tdr_sheet_list = detected["tdr"] or (None, [])
        if tdr_file and tdr_sheet_list:
            tdr_bytes = tdr_file.getvalue()
            tdr_name = tdr_file.name
            tdr_sheet = st.selectbox("**TDR Data** → " + tdr_file.name, options=tdr_sheet_list, index=0, key="tdr_sheet")
        elif upload_files:
            st.info("No TDR Data detected. Upload an Excel with TDR sections.")
    lvt_file, lvt_sheet_list = detected["lvt"] or (None, [])
    if lvt_file and lvt_sheet_list:
        default_idx = lvt_sheet_list.index(tdr_core.LVT_SHEET_NAME) if tdr_core.LVT_SHEET_NAME in lvt_sheet_list else 0
        lvt_sheet = st.selectbox("**LVT Report** → " + lvt_file.name, options=lvt_sheet_list, index=default_idx, key="lvt_sheet")
    else:
        lvt_sheet = None
        if upload_files and not detected["lvt"]:
            st.info("No LVT Report detected (no sheet named 'BAN Wise Result').")
    device_file, device_sheet_list = detected["device"] or (None, [])
    if device_file and device_sheet_list:
        st.caption("**Device Details** → " + device_file.name)
else:
    lvt_sheet = None

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
run = st.button("Run TDR", type="primary")

if run and not tdr_bytes:
    st.warning("Please provide **TDR Data** (upload file(s) or pick from SharePoint).")
elif run and (not lvt_file or lvt_file.size == 0):
    st.warning("Please upload file(s) that include **LVT Report** (sheet named BAN Wise Result).")
elif run and tdr_bytes and tdr_sheet and lvt_file and lvt_file.size > 0:
    tmpdir = tempfile.mkdtemp(prefix="tdr_streamlit_")
    try:
        os.environ["TDR_WEB_REPORT_FOLDER"] = tmpdir
        tdr_path = os.path.join(tmpdir, "tdr_input.xlsx")
        with open(tdr_path, "wb") as f:
            f.write(tdr_bytes)
        sheet_name = tdr_sheet
        if not sheet_name:
            st.error("Please select a TDR Data sheet.")
        else:
            lvt_path = os.path.join(tmpdir, "lvt_input.xlsx")
            with open(lvt_path, "wb") as f:
                f.write(lvt_file.getvalue())
            device_details_path = None
            if device_file and device_file.size > 0:
                device_details_path = os.path.join(tmpdir, "device_details.xlsx")
                with open(device_details_path, "wb") as f:
                    f.write(device_file.getvalue())
            sheet_to_use = (lvt_sheet or tdr_core.LVT_SHEET_NAME).strip() or tdr_core.LVT_SHEET_NAME
            out_path = os.path.join(tmpdir, "TDR_BAN_Report.xlsx")
            with st.spinner("Processing…"):
                result_path, summary = tdr_core.run_extraction_and_report(
                    [(tdr_path, [sheet_name])], output_excel=out_path,
                    lvt_report_path=lvt_path, lvt_sheet_name=sheet_to_use if lvt_path else None,
                    device_details_path=device_details_path,
                )
            if result_path and os.path.isfile(result_path):
                with open(result_path, "rb") as f:
                    report_bytes = f.read()
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                report_filename = f"TDR_BAN_Report_{ts}.xlsx"
                zip_bytes = None
                zip_filename = f"TDR_wise_report_{datetime.now().strftime('%Y%m%d')}.zip"
                per_tdr_folder = (summary or {}).get("per_tdr_folder")
                if per_tdr_folder and os.path.isdir(per_tdr_folder):
                    files = [n for n in os.listdir(per_tdr_folder) if n.endswith((".xlsx", ".xlsm"))]
                    if files:
                        buf = BytesIO()
                        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                            for n in files:
                                z.write(os.path.join(per_tdr_folder, n), n)
                        buf.seek(0)
                        zip_bytes = buf.getvalue()
                st.session_state["tdr_result"] = {
                    "report_bytes": report_bytes, "report_filename": report_filename,
                    "zip_bytes": zip_bytes, "zip_filename": zip_filename, "summary": summary, "lvt_used": True,
                }
            else:
                st.error("No TDR data found or report generation failed.")
    finally:
        if "TDR_WEB_REPORT_FOLDER" in os.environ:
            del os.environ["TDR_WEB_REPORT_FOLDER"]
        try:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass

if "tdr_result" in st.session_state:
    r = st.session_state["tdr_result"]
    st.success("Done — download below.")
    if r.get("zip_bytes"):
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("Download main report", data=r["report_bytes"], file_name=r["report_filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_main")
        with c2:
            st.download_button("Download per-TDR (ZIP)", data=r["zip_bytes"], file_name=r["zip_filename"], mime="application/zip", key="dl_zip")
    else:
        st.download_button("Download main report", data=r["report_bytes"], file_name=r["report_filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_main")
    if r.get("summary"):
        s = r["summary"]
        total, passed, failed, not_found = s.get("total", 0), s.get("passed", 0), s.get("failed", 0), s.get("not_found", 0)
        tdr_p, tdr_f, tdr_part = s.get("tdr_passed", 0), s.get("tdr_failed", 0), s.get("tdr_partial", 0)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total BAN", total)
        m2.metric("Passed", passed)
        m3.metric("Failed", failed)
        m4.metric("Not found", not_found)
