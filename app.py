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

st.set_page_config(page_title="TDR Data Excel", page_icon="📊", layout="centered")
st.title("📊 TDR Data Excel")
st.markdown("Upload your **TDR Data** sheet and optional **LVT report**. Get the main report and one Excel per TDR (same as the local script).")

tdr_file = st.file_uploader("TDR Data Excel (required)", type=["xlsx", "xlsm"], help="Excel file with TDR sections")
lvt_file = st.file_uploader("LVT Report Excel (optional)", type=["xlsx", "xlsm"], help="For BAN status column; if omitted, status will be 'Not found'")

# When TDR file is uploaded, read sheet names and show dropdown
tdr_sheet = None
if tdr_file and tdr_file.size > 0:
    try:
        wb_tdr = tdr_core.load_workbook(BytesIO(tdr_file.getvalue()), read_only=True)
        tdr_sheet_names = wb_tdr.sheetnames
        wb_tdr.close()
        if tdr_sheet_names:
            tdr_sheet = st.selectbox(
                "TDR Data sheet",
                options=tdr_sheet_names,
                index=0,
                key="tdr_sheet",
                help="Sheet in the TDR file that contains TDR sections (usually the first sheet)",
            )
        else:
            st.caption("TDR file has no sheets.")
    except Exception as e:
        st.warning(f"Could not read TDR file sheets: {e}.")
else:
    st.caption("Upload a TDR Data Excel file to choose which sheet to use.")

# When LVT is uploaded, read sheet names and show dropdown; otherwise no sheet choice needed
lvt_sheet = None
if lvt_file and lvt_file.size > 0:
    try:
        wb_lvt = tdr_core.load_workbook(BytesIO(lvt_file.getvalue()), read_only=True)
        lvt_sheet_names = wb_lvt.sheetnames
        wb_lvt.close()
        if lvt_sheet_names:
            default_idx = 0
            if tdr_core.LVT_SHEET_NAME in lvt_sheet_names:
                default_idx = lvt_sheet_names.index(tdr_core.LVT_SHEET_NAME)
            lvt_sheet = st.selectbox(
                "LVT sheet for BAN-wise list",
                options=lvt_sheet_names,
                index=default_idx,
                key="lvt_sheet",
                help="Sheet in the LVT report that contains the BAN-wise list (default: BAN Wise Result if present)",
            )
        else:
            st.caption("LVT file has no sheets.")
    except Exception as e:
        st.warning(f"Could not read LVT file sheets: {e}. Using default sheet name.")
        lvt_sheet = st.text_input("LVT sheet name (fallback)", value=tdr_core.LVT_SHEET_NAME)
else:
    st.caption("Upload an LVT report to choose which sheet to use for the BAN-wise list.")

run = st.button("Run TDR")

if run and tdr_file:
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
            lvt_path = None
            if lvt_file and lvt_file.size > 0:
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
                zip_filename = f"TDR_per_TDR_files_{ts}.zip"
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
                    "report_bytes": report_bytes,
                    "report_filename": report_filename,
                    "zip_bytes": zip_bytes,
                    "zip_filename": zip_filename,
                    "summary": summary,
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
elif run and not tdr_file:
    st.warning("Please upload a TDR Data Excel file.")

# Show download section from session state so both buttons stay visible after clicking either one
if "tdr_result" in st.session_state:
    r = st.session_state["tdr_result"]
    st.success("Done. Download your files below.")
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "Download main report",
            data=r["report_bytes"],
            file_name=r["report_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_main_report",
        )
    with col2:
        if r.get("zip_bytes"):
            st.download_button(
                "Download per-TDR files (ZIP)",
                data=r["zip_bytes"],
                file_name=r["zip_filename"],
                mime="application/zip",
                key="download_per_tdr_zip",
            )
        else:
            st.caption("No per-TDR files generated.")
    if r.get("summary"):
        st.json(r["summary"])
