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

# Ensure we can import tdr_core from this folder
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import tdr_core

st.set_page_config(page_title="TDR mapping sheet creation", page_icon="📋", layout="centered")

# New theme: teal/slate, compact spacing
st.markdown(
    """
    <style>
    .stApp { background: #f1f5f9 !important; }
    .block-container { padding: 1rem 1.5rem 1.5rem !important; max-width: 880px !important; }
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

# New header: teal bar, centered - compact so fully visible (no cut-off on small screens)
st.markdown(
    """
    <div style="
        background: linear-gradient(90deg, #0f766e 0%, #0d9488 50%, #14b8a6 100%);
        color: #fff; padding: 12px 16px; border-radius: 10px; margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(15, 118, 110, 0.3);
        text-align: center;
    ">
        <div style="font-size: 1.25rem; font-weight: 700; margin-bottom: 2px;">TDR mapping sheet creation</div>
        <div style="font-size: 0.75rem; opacity: 0.95;">Upload TDR data + LVT report (both required) → get mapping and TDR-wise reports.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Two columns: TDR (left) and LVT (right)
col_tdr, col_lvt = st.columns(2)

with col_tdr:
    st.markdown("**TDR Data**")
    tdr_file = st.file_uploader("TDR Excel (required)", type=["xlsx", "xlsm"], help="TDR sections", key="tdr_upload")
    tdr_sheet = None
    if tdr_file and tdr_file.size > 0:
        try:
            wb_tdr = tdr_core.load_workbook(BytesIO(tdr_file.getvalue()), read_only=True)
            tdr_sheet_names = wb_tdr.sheetnames
            wb_tdr.close()
            if tdr_sheet_names:
                tdr_sheet = st.selectbox("Sheet", options=tdr_sheet_names, index=0, key="tdr_sheet")
            else:
                st.caption("No sheets.")
        except Exception as e:
            st.warning(str(e))
    else:
        st.caption("Upload file to pick sheet.")

with col_lvt:
    st.markdown("**LVT Report**")
    lvt_file = st.file_uploader("LVT Excel (required)", type=["xlsx", "xlsm"], help="BAN-wise list", key="lvt_upload")
    lvt_sheet = None
    if lvt_file and lvt_file.size > 0:
        try:
            wb_lvt = tdr_core.load_workbook(BytesIO(lvt_file.getvalue()), read_only=True)
            lvt_sheet_names = wb_lvt.sheetnames
            wb_lvt.close()
            if lvt_sheet_names:
                default_idx = lvt_sheet_names.index(tdr_core.LVT_SHEET_NAME) if tdr_core.LVT_SHEET_NAME in lvt_sheet_names else 0
                lvt_sheet = st.selectbox("Sheet", options=lvt_sheet_names, index=default_idx, key="lvt_sheet")
            else:
                st.caption("No sheets.")
        except Exception as e:
            st.warning(str(e))
            lvt_sheet = st.text_input("Sheet name", value=tdr_core.LVT_SHEET_NAME, key="lvt_fallback")
    else:
        st.caption("Upload file to pick sheet.")

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
run = st.button("Run TDR", type="primary")

if run and not tdr_file:
    st.warning("Please upload **TDR Data Excel** (required).")
elif run and (not lvt_file or lvt_file.size == 0):
    st.warning("Please upload **LVT Report Excel** (required).")
elif run and tdr_file and lvt_file and lvt_file.size > 0:
    tmpdir = tempfile.mkdtemp(prefix="tdr_streamlit_")
    try:
        os.environ["TDR_WEB_REPORT_FOLDER"] = tmpdir
        tdr_path = os.path.join(tmpdir, "tdr_input.xlsx")
        with open(tdr_path, "wb") as f:
            f.write(tdr_file.getvalue())

        # Use the sheet selected from dropdown (tdr_sheet set when TDR file was uploaded)
        sheet_name = tdr_sheet
        if not sheet_name:
            st.error("Please select a TDR Data sheet (re-upload the file if the dropdown did not appear).")
        else:
            lvt_path = os.path.join(tmpdir, "lvt_input.xlsx")
            with open(lvt_path, "wb") as f:
                f.write(lvt_file.getvalue())
            sheet_to_use = (lvt_sheet or tdr_core.LVT_SHEET_NAME).strip() or tdr_core.LVT_SHEET_NAME
            out_path = os.path.join(tmpdir, "TDR_BAN_Report.xlsx")

            with st.spinner("Processing…"):
                result_path, summary = tdr_core.run_extraction_and_report(
                    [(tdr_path, [sheet_name])],
                    output_excel=out_path,
                    lvt_report_path=lvt_path,
                    lvt_sheet_name=sheet_to_use if lvt_path else None,
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
