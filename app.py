"""
TDR Data Excel – Streamlit app (free hosting on Streamlit Community Cloud).
Upload TDR Data + optional LVT report, run, download main report and per-TDR files.
Uses tdr_core (copy of local script); does not modify the original tdr_excel_script.py.
"""
import os
import sys
import tempfile
import zipfile
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
lvt_sheet = st.text_input("LVT sheet name for BAN-wise list (optional)", value="BAN Wise Result", help="Default: BAN Wise Result")

run = st.button("Run TDR")

if run and tdr_file:
    tmpdir = tempfile.mkdtemp(prefix="tdr_streamlit_")
    try:
        os.environ["TDR_WEB_REPORT_FOLDER"] = tmpdir
        tdr_path = os.path.join(tmpdir, "tdr_input.xlsx")
        with open(tdr_path, "wb") as f:
            f.write(tdr_file.getvalue())

        # First sheet from TDR file
        wb = tdr_core.load_workbook(tdr_path, read_only=True)
        sheet_name = wb.sheetnames[0] if wb.sheetnames else None
        wb.close()
        if not sheet_name:
            st.error("TDR file has no sheets.")
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
                st.success("Done. Download your files below.")
                with open(result_path, "rb") as f:
                    report_bytes = f.read()
                st.download_button("Download main report", data=report_bytes, file_name=os.path.basename(result_path), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                per_tdr_folder = (summary or {}).get("per_tdr_folder")
                if per_tdr_folder and os.path.isdir(per_tdr_folder):
                    files = [n for n in os.listdir(per_tdr_folder) if n.endswith((".xlsx", ".xlsm"))]
                    if files:
                        buf = BytesIO()
                        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                            for n in files:
                                z.write(os.path.join(per_tdr_folder, n), n)
                        buf.seek(0)
                        st.download_button("Download per-TDR files (ZIP)", data=buf.getvalue(), file_name="TDR_per_TDR_files.zip", mime="application/zip")
                if summary:
                    st.json(summary)
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
