"""
TDR Data Excel – Streamlit app (free hosting on Streamlit Community Cloud).
Upload TDR Data + optional LVT report, run, download main report and per-TDR files.
Uses tdr_core (copy of local script); does not modify the original tdr_excel_script.py.
"""
import os
import sys
import tempfile
import zipfile
from datetime import datetime
from io import BytesIO

import streamlit as st

# Ensure we can import tdr_core and sharepoint_graph from this folder
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import tdr_core
import sharepoint_graph

st.set_page_config(page_title="TDR mapping sheet creation", page_icon="📋", layout="centered")

# New theme: teal/slate, compact spacing
st.markdown(
    """
    <style>
    .stApp { background: #f1f5f9 !important; }
    .block-container { padding: 1.75rem 1.5rem 1.5rem !important; max-width: 880px !important; }
    div[data-testid="stVerticalBlock"] > div { padding: 0.15rem 0 !important; }
    section[data-testid="stFileUploader"] {
        background: #fff !important; border-radius: 8px !important; padding: 14px !important;
        border: 1px solid #e2e8f0 !important; border-top: 3px solid #0d9488 !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06) !important;
    }
    .stSelectbox > div { background: #fff !important; border-radius: 8px !important; border: 1px solid #e2e8f0 !important; }
    .stButton > button[kind="primary"], .stButton > button:first-child {
        background: #0d9488 !important; color: #fff !important; border: none !important;
        border-radius: 999px !important; padding: 0.5rem 1.75rem !important; font-weight: 600 !important;
    }
    .stButton > button[kind="primary"]:hover, .stButton > button:first-child:hover {
        background: #0f766e !important; box-shadow: 0 2px 8px rgba(13, 148, 136, 0.4) !important;
    }
    div[data-testid="stDownloadButton"] > button { border-radius: 8px !important; border: 1px solid #0d9488 !important; color: #0d9488 !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# New header: teal bar, centered, extra padding so text is not clipped
st.markdown(
    """
    <div style="
        background: linear-gradient(90deg, #0f766e 0%, #0d9488 50%, #14b8a6 100%);
        color: #fff; padding: 18px 20px 16px; border-radius: 10px; margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(15, 118, 110, 0.3);
        text-align: center;
    ">
        <div style="font-size: 1.25rem; font-weight: 700; margin-bottom: 4px; line-height: 1.3;">TDR mapping sheet creation</div>
        <div style="font-size: 0.75rem; opacity: 0.95;">Upload TDR data + LVT report (both required) → get mapping and TDR-wise reports.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# SharePoint folder link for TDR (Option 2)
TDR_SHAREPOINT_URL = (
    "https://amdocs.sharepoint.com/:f:/r/sites/USCCTesting_Offshore/Shared%20Documents/"
    "Release%20%26%20OffCycle/UTMO%20-%20Migration%20(Data%20creation)/Data%20Creation/R2%20Data"
    "?csf=1&web=1&e=EZDfgA"
)

# Unified TDR input: bytes + name + sheet (from upload or from SharePoint direct)
tdr_bytes = None
tdr_name = None
tdr_sheet = None

# TDR source choice above both columns so upload row is perfectly aligned
tdr_source = st.radio(
    "TDR file from",
    options=["Local file", "SharePoint"],
    horizontal=True,
    key="tdr_source",
    help="Local upload or pick a file directly from SharePoint (if configured)",
)
use_sharepoint_direct = tdr_source == "SharePoint" and sharepoint_graph.has_sharepoint_credentials()
if tdr_source == "SharePoint" and not use_sharepoint_direct:
    st.markdown(
        f'<a href="{TDR_SHAREPOINT_URL}" target="_blank" rel="noopener" '
        'style="font-size: 0.85rem; color: #0d9488;">📂 Open TDR folder on SharePoint</a>',
        unsafe_allow_html=True,
    )
    st.caption("Download the file from SharePoint, then upload it in the left box below.")

# --- Single multi-file upload OR SharePoint for TDR ---
if use_sharepoint_direct:
    token = sharepoint_graph.get_token()
    if not token:
        st.warning("Could not get SharePoint token. Check AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET in secrets.")
    else:
        files = sharepoint_graph.list_tdr_excel_files(token)
        if not files:
            st.info("No Excel files found in the TDR folder.")
        else:
            selected_name = st.selectbox(
                "TDR file (SharePoint)",
                options=[f["name"] for f in files],
                key="sp_tdr_file",
                help="Pick a file from the R2 Data folder",
            )
            selected = next((f for f in files if f["name"] == selected_name), None)
            if selected:
                cache_key = f"sp_tdr_{selected['id']}"
                if cache_key not in st.session_state:
                    with st.spinner("Loading file…"):
                        content = sharepoint_graph.download_file_content(
                            token, selected["drive_id"], selected["id"]
                        )
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
                        else:
                            st.caption("No sheets.")
                    except Exception as e:
                        st.warning(str(e))
                else:
                    st.warning("Could not download file.")
    # When SharePoint TDR: still need LVT (and optional Device) from uploads
    st.markdown("**Upload file(s) for LVT Report and/or Device Details**")
    upload_files = st.file_uploader("Excel file(s)", type=["xlsx", "xlsm"], accept_multiple_files=True, key="multi_upload", help="Upload one or more Excel files. We'll detect LVT Report and Device Details.")
else:
    st.markdown("**Upload file(s) – we'll detect TDR Data, LVT Report, and Device Details**")
    upload_files = st.file_uploader("Excel file(s)", type=["xlsx", "xlsm"], accept_multiple_files=True, key="multi_upload", help="Select multiple files at once. The script will detect which is TDR Data, which is LVT Report, and which is Device Details.")

# Detect roles from uploaded files and show sheet selection by file name
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
        else:
            if upload_files:
                st.info("No TDR Data detected in uploaded files (no sheet with TDR sections). Upload an Excel that contains TDR-###### and BANs.")

    lvt_file, lvt_sheet_list = detected["lvt"] or (None, [])
    if lvt_file and lvt_sheet_list:
        default_idx = lvt_sheet_list.index(tdr_core.LVT_SHEET_NAME) if tdr_core.LVT_SHEET_NAME in lvt_sheet_list else 0
        lvt_sheet = st.selectbox("**LVT Report** → " + lvt_file.name, options=lvt_sheet_list, index=default_idx, key="lvt_sheet")
    else:
        lvt_sheet = None
        if upload_files and not detected["lvt"]:
            st.info("No LVT Report detected (no sheet named 'BAN Wise Result'). Add a file that has that sheet.")

    device_file, device_sheet_list = detected["device"] or (None, [])
    if device_file and device_sheet_list:
        st.caption("**Device Details** → " + device_file.name + " (sheet used: first with CUSTOMER_ID)")
    else:
        device_file = None
else:
    lvt_sheet = None
    if not use_sharepoint_direct:
        st.caption("Upload one or more Excel files to detect TDR Data, LVT Report, and Device Details.")

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
            st.error("Please select a TDR Data sheet (re-upload the file if the dropdown did not appear).")
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
                    [(tdr_path, [sheet_name])],
                    output_excel=out_path,
                    lvt_report_path=lvt_path,
                    lvt_sheet_name=sheet_to_use if lvt_path else None,
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
                lvt_used = bool(lvt_file and lvt_file.size > 0)
                st.session_state["tdr_result"] = {
                    "report_bytes": report_bytes,
                    "report_filename": report_filename,
                    "zip_bytes": zip_bytes,
                    "zip_filename": zip_filename,
                    "summary": summary,
                    "lvt_used": lvt_used,
                }
            else:
                st.error("No TDR data found in the sheet or report generation failed.")
    finally:
        if "TDR_WEB_REPORT_FOLDER" in os.environ:
            del os.environ["TDR_WEB_REPORT_FOLDER"]
        try:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass
# Results: compact, teal theme
if "tdr_result" in st.session_state:
    r = st.session_state["tdr_result"]
    st.markdown(
        '<div style="background: #ccfbf1; color: #0f766e; padding: 10px 14px; border-radius: 8px; margin-bottom: 12px; border-left: 4px solid #0d9488; font-weight: 600;">✓ Done — download below</div>',
        unsafe_allow_html=True,
    )
    show_zip = r.get("lvt_used", True) and r.get("zip_bytes")
    if show_zip:
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("Download main report", data=r["report_bytes"], file_name=r["report_filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_main_report")
        with c2:
            st.download_button("Download per-TDR (ZIP)", data=r["zip_bytes"], file_name=r["zip_filename"], mime="application/zip", key="download_per_tdr_zip")
    else:
        st.download_button("Download main report", data=r["report_bytes"], file_name=r["report_filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_main_report")
        if not r.get("lvt_used", True):
            st.caption("Upload LVT and run again for per-TDR ZIP.")
        else:
            st.caption("No per-TDR files.")
    if r.get("summary"):
        s = r["summary"]
        total, passed, failed, not_found = s.get("total", 0), s.get("passed", 0), s.get("failed", 0), s.get("not_found", 0)
        tdr_p, tdr_f, tdr_part = s.get("tdr_passed", 0), s.get("tdr_failed", 0), s.get("tdr_partial", 0)
        total_tdr = tdr_p + tdr_f + tdr_part
        per_tdr = s.get("per_tdr_count", 0)
        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
        # Row 1: BAN (4 cards)
        st.markdown(
            '<div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-bottom: 8px;">'
            f'<div style="background: #f8fafc; border-radius: 6px; padding: 10px 12px; border-left: 3px solid #64748b;"><div style="font-size: 0.7rem; color: #64748b; font-weight: 600;">TOTAL BAN</div><div style="font-size: 1.25rem; font-weight: 700; color: #1e293b;">{total}</div></div>'
            f'<div style="background: #ecfdf5; border-radius: 6px; padding: 10px 12px; border-left: 3px solid #10b981;"><div style="font-size: 0.7rem; color: #059669; font-weight: 600;">PASSED</div><div style="font-size: 1.25rem; font-weight: 700; color: #047857;">{passed}</div></div>'
            f'<div style="background: #fef2f2; border-radius: 6px; padding: 10px 12px; border-left: 3px solid #ef4444;"><div style="font-size: 0.7rem; color: #dc2626; font-weight: 600;">FAILED</div><div style="font-size: 1.25rem; font-weight: 700; color: #b91c1c;">{failed}</div></div>'
            f'<div style="background: #fffbeb; border-radius: 6px; padding: 10px 12px; border-left: 3px solid #f59e0b;"><div style="font-size: 0.7rem; color: #d97706; font-weight: 600;">NOT FOUND</div><div style="font-size: 1.25rem; font-weight: 700; color: #b45309;">{not_found}</div></div>'
            "</div>"
            # Row 2: TDR (4 cards, aligned with BAN row)
            '<div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-bottom: 8px;">'
            f'<div style="background: #f8fafc; border-radius: 6px; padding: 10px 12px; border-left: 3px solid #64748b;"><div style="font-size: 0.7rem; color: #64748b; font-weight: 600;">TOTAL TDR</div><div style="font-size: 1.25rem; font-weight: 700; color: #1e293b;">{total_tdr}</div></div>'
            f'<div style="background: #ecfdf5; border-radius: 6px; padding: 10px 12px; border-left: 3px solid #10b981;"><div style="font-size: 0.7rem; color: #059669; font-weight: 600;">PASSED</div><div style="font-size: 1.25rem; font-weight: 700; color: #047857;">{tdr_p}</div></div>'
            f'<div style="background: #fef2f2; border-radius: 6px; padding: 10px 12px; border-left: 3px solid #ef4444;"><div style="font-size: 0.7rem; color: #dc2626; font-weight: 600;">FAILED</div><div style="font-size: 1.25rem; font-weight: 700; color: #b91c1c;">{tdr_f}</div></div>'
            f'<div style="background: #fffbeb; border-radius: 6px; padding: 10px 12px; border-left: 3px solid #f59e0b;"><div style="font-size: 0.7rem; color: #d97706; font-weight: 600;">PARTIAL</div><div style="font-size: 1.25rem; font-weight: 700; color: #b45309;">{tdr_part}</div></div>'
            "</div>"
            f'<div style="background: #f0fdfa; border-radius: 6px; padding: 10px 14px; border-left: 3px solid #0d9488;">📁 Per-TDR files: <strong>{per_tdr}</strong></div>',
            unsafe_allow_html=True,
        )
