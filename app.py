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

# Attractive header and description
st.markdown(
    """
    <div style="
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 50%, #e2e8f0 100%);
        border-radius: 12px;
        padding: 24px 28px;
        margin-bottom: 24px;
        border-left: 5px solid #2563eb;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    ">
        <p style="margin: 0 0 8px 0; font-size: 1.75rem; font-weight: 700; color: #1e293b; letter-spacing: -0.02em;">
            📋 TDR mapping sheet creation
        </p>
        <p style="margin: 0; font-size: 0.95rem; line-height: 1.5; color: #475569;">
            Upload your <strong>TDR data sheet</strong> and <strong>LVT report</strong> (input files). 
            The script will create a detailed mapping and TDR-wise report for further use.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

tdr_file = st.file_uploader("TDR Data Excel (required)", type=["xlsx", "xlsm"], help="Excel file with TDR sections")
# TDR sheet dropdown: right below TDR upload only
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

lvt_file = st.file_uploader("LVT Report Excel (optional)", type=["xlsx", "xlsm"], help="For BAN status column; if omitted, status will be 'Not found'")
# LVT sheet dropdown: right below LVT upload only
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
                zip_filename = "TDR_per_TDR_files.zip"
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
elif run and not tdr_file:
    st.warning("Please upload a TDR Data Excel file.")

# Show download section from session state so both buttons stay visible after clicking either one
if "tdr_result" in st.session_state:
    r = st.session_state["tdr_result"]
    st.success("Done. Download your files below.")
    show_zip = r.get("lvt_used", True) and r.get("zip_bytes")  # hide ZIP when run was without LVT
    if show_zip:
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
            st.download_button(
                "Download per-TDR files (ZIP)",
                data=r["zip_bytes"],
                file_name=r["zip_filename"],
                mime="application/zip",
                key="download_per_tdr_zip",
            )
    else:
        st.download_button(
            "Download main report",
            data=r["report_bytes"],
            file_name=r["report_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_main_report",
        )
        if not r.get("lvt_used", True):
            st.caption("Upload an LVT report and run again to get per-TDR files (ZIP).")
        else:
            st.caption("No per-TDR files generated.")
    # Summary: rich, colored dashboard-style layout
    if r.get("summary"):
        s = r["summary"]
        total = s.get("total", 0)
        passed = s.get("passed", 0)
        failed = s.get("failed", 0)
        not_found = s.get("not_found", 0)
        tdr_p = s.get("tdr_passed", 0)
        tdr_f = s.get("tdr_failed", 0)
        tdr_part = s.get("tdr_partial", 0)
        per_tdr = s.get("per_tdr_count", 0)
        st.markdown("---")
        st.markdown(
            '<p style="font-size: 1.35rem; font-weight: 700; color: #1e293b; margin-bottom: 0.5rem;">📋 Summary</p>',
            unsafe_allow_html=True,
        )
        # BAN counts: colored cards
        st.markdown(
            '<div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 16px;">'
            f'<div style="background: linear-gradient(145deg, #f1f5f9 0%, #e2e8f0 100%); border-radius: 10px; padding: 14px 16px; border-left: 4px solid #64748b; box-shadow: 0 1px 3px rgba(0,0,0,0.06);">'
            f'<div style="font-size: 0.8rem; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.02em;">Total BANs</div>'
            f'<div style="font-size: 1.75rem; font-weight: 700; color: #1e293b;">{total}</div></div>'
            f'<div style="background: linear-gradient(145deg, #dcfce7 0%, #bbf7d0 100%); border-radius: 10px; padding: 14px 16px; border-left: 4px solid #16a34a; box-shadow: 0 1px 3px rgba(0,0,0,0.06);">'
            f'<div style="font-size: 0.8rem; color: #15803d; font-weight: 600; text-transform: uppercase;">Passed</div>'
            f'<div style="font-size: 1.75rem; font-weight: 700; color: #166534;">{passed}</div></div>'
            f'<div style="background: linear-gradient(145deg, #fee2e2 0%, #fecaca 100%); border-radius: 10px; padding: 14px 16px; border-left: 4px solid #dc2626; box-shadow: 0 1px 3px rgba(0,0,0,0.06);">'
            f'<div style="font-size: 0.8rem; color: #b91c1c; font-weight: 600; text-transform: uppercase;">Failed</div>'
            f'<div style="font-size: 1.75rem; font-weight: 700; color: #991b1b;">{failed}</div></div>'
            f'<div style="background: linear-gradient(145deg, #fef3c7 0%, #fde68a 100%); border-radius: 10px; padding: 14px 16px; border-left: 4px solid #d97706; box-shadow: 0 1px 3px rgba(0,0,0,0.06);">'
            f'<div style="font-size: 0.8rem; color: #b45309; font-weight: 600; text-transform: uppercase;">Not found</div>'
            f'<div style="font-size: 1.75rem; font-weight: 700; color: #92400e;">{not_found}</div></div>'
            "</div>"
            "<p style='font-size: 0.9rem; font-weight: 600; color: #475569; margin: 4px 0 8px 0;'>TDR-wise</p>"
            '<div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 16px;">'
            f'<div style="background: linear-gradient(145deg, #dcfce7 0%, #bbf7d0 100%); border-radius: 10px; padding: 14px 16px; border-left: 4px solid #16a34a; box-shadow: 0 1px 3px rgba(0,0,0,0.06);">'
            f'<div style="font-size: 0.8rem; color: #15803d; font-weight: 600;">Passed</div><div style="font-size: 1.6rem; font-weight: 700; color: #166534;">{tdr_p}</div></div>'
            f'<div style="background: linear-gradient(145deg, #fee2e2 0%, #fecaca 100%); border-radius: 10px; padding: 14px 16px; border-left: 4px solid #dc2626; box-shadow: 0 1px 3px rgba(0,0,0,0.06);">'
            f'<div style="font-size: 0.8rem; color: #b91c1c; font-weight: 600;">Failed</div><div style="font-size: 1.6rem; font-weight: 700; color: #991b1b;">{tdr_f}</div></div>'
            f'<div style="background: linear-gradient(145deg, #fef3c7 0%, #fde68a 100%); border-radius: 10px; padding: 14px 16px; border-left: 4px solid #d97706; box-shadow: 0 1px 3px rgba(0,0,0,0.06);">'
            f'<div style="font-size: 0.8rem; color: #b45309; font-weight: 600;">Partial</div><div style="font-size: 1.6rem; font-weight: 700; color: #92400e;">{tdr_part}</div></div>'
            "</div>"
            f'<div style="background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%); border-radius: 10px; padding: 16px 20px; border-left: 5px solid #2563eb; box-shadow: 0 2px 6px rgba(37,99,235,0.15);">'
            f'<span style="font-size: 1.1rem; font-weight: 600; color: #1e40af;">📁 Per-TDR files</span> '
            f'<span style="font-size: 1.4rem; font-weight: 700; color: #1e3a8a;">{per_tdr}</span> '
            f'<span style="color: #3730a3;">Excel file(s) generated</span></div>',
            unsafe_allow_html=True,
        )
