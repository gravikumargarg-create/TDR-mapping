"""Synthetic TDR mapping view – run from app.py when portal_view == 'synthetic'."""
import os
import tempfile
import zipfile
from datetime import datetime
from io import BytesIO
import streamlit as st

import tdr_core
import sharepoint_graph


def _normalize_id(val):
    """Normalize BAN/CUSTOMER_ID for comparison."""
    if val is None:
        return None
    s = str(val).strip()
    return s if s else None


def render_synthetic():
    # TDR+BML merge section removed: use the four uploads above (Data details, LVT, Device details, BML) and Run TDR only.
    st.markdown(
        """
        <style>
        .stApp { background: #f1f5f9 !important; }
        .block-container { padding: 1.75rem 1.5rem 5rem !important; max-width: 880px !important; }
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
            <div style="font-size: 1.25rem; font-weight: 700;">TDR wise mapping</div>
            <div style="font-size: 0.75rem; opacity: 0.95;">TDR-wise mapping for synthetic data. Inputs needed: TDR data sheets, device details, and LVT report.</div>
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
    tdr_bytes = None
    tdr_name = None
    lvt_file = None
    lvt_sheet = None
    device_file = None
    bml_file = None
    data_details_files = []

    if tdr_source == "SharePoint" and not use_sharepoint_direct:
        st.markdown(f'<a href="{TDR_SHAREPOINT_URL}" target="_blank" style="font-size: 0.85rem; color: #0d9488;">📂 Open TDR folder on SharePoint</a>', unsafe_allow_html=True)
        st.caption("Download the file from SharePoint, then use the upload options below.")

    if use_sharepoint_direct:
        token = sharepoint_graph.get_token()
        if not token:
            st.warning("Could not get SharePoint token. Check AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET in secrets.")
        else:
            files = sharepoint_graph.list_tdr_excel_files(token)
            if not files:
                st.info("No Excel files found in the TDR folder.")
            else:
                selected_name = st.selectbox("1. TDR Data (SharePoint)", options=[f["name"] for f in files], key="sp_tdr_file")
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
                            wb_tdr = tdr_core.load_workbook(BytesIO(content), read_only=True, data_only=True)
                            roles = tdr_core.detect_excel_roles(wb_tdr)
                            all_sheets_sp = wb_tdr.sheetnames
                            wb_tdr.close()
                            tdr_sheet_list = roles.get("tdr_sheets") or all_sheets_sp
                            if tdr_sheet_list:
                                st.caption(f"TDR Data → {tdr_name} ({len(tdr_sheet_list)} sheet(s))")
                        except Exception as e:
                            st.warning(str(e))
        st.markdown("**2. LVT file** | **3. Device details** | **4. BML file**")
        _s1, _s2, _s3 = st.columns(3)
        with _s1:
            lvt_file = st.file_uploader("LVT file", type=["xlsx", "xlsm"], key="lvt_upload_sp", help="Excel with BAN Wise Result sheet.")
        with _s2:
            device_file = st.file_uploader("Device details", type=["xlsx", "xlsm"], key="device_upload_sp", help="Excel with CUSTOMER_ID column.")
        with _s3:
            bml_file = st.file_uploader("BML file", type=["xlsx", "xlsm"], key="bml_upload_sp", help="BML Excel (TDR + BML sheets).")
        if lvt_file and lvt_file.size > 0:
            lvt_sheet = tdr_core.LVT_SHEET_NAME  # Resolved when Run TDR is clicked
    else:
        tdr_sheet_list = []
        data_details_files = []
        st.markdown("**1. Data details input file(s)** | **2. LVT file**")
        _r1a, _r1b = st.columns(2)
        with _r1a:
            data_details_upload = st.file_uploader("Data details input file(s)", type=["xlsx", "xlsm"], accept_multiple_files=True, key="data_details_upload", help="One or more TDR data Excel files (all sheets will be processed).")
        with _r1b:
            lvt_file = st.file_uploader("LVT file", type=["xlsx", "xlsm"], key="lvt_upload", help="Excel with BAN Wise Result sheet.")
        st.markdown("**3. Device details** | **4. BML file**")
        _r2a, _r2b = st.columns(2)
        with _r2a:
            device_file = st.file_uploader("Device details", type=["xlsx", "xlsm"], key="device_upload", help="Excel with CUSTOMER_ID column. Added to ZIP when provided.")
        with _r2b:
            bml_file = st.file_uploader("BML file", type=["xlsx", "xlsm"], key="bml_upload", help="BML Excel. Added to ZIP when provided.")
        if data_details_upload:
            data_details_files = [f for f in data_details_upload if f and getattr(f, "size", 0) > 0]
        if data_details_files:
            tdr_bytes = data_details_files[0].getvalue()
            tdr_name = data_details_files[0].name if len(data_details_files) == 1 else f"{len(data_details_files)} file(s)"
            # Don't open workbooks here — only when Run TDR is clicked (avoids long "running" state on every upload)
            st.caption(f"Data details → {tdr_name}")
        if lvt_file and lvt_file.size > 0:
            # Defer LVT workbook read until Run TDR; use default sheet name for now
            lvt_sheet = tdr_core.LVT_SHEET_NAME

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    _run1, _run2, _run3 = st.columns([1, 2, 1])
    with _run2:
        has_tdr = bool(tdr_bytes)
        has_lvt = lvt_file and lvt_file.size > 0
        run = st.button("Run TDR", type="primary", use_container_width=True, disabled=not (has_tdr and has_lvt))
    st.markdown("<div style='height: 2rem;'></div>", unsafe_allow_html=True)

    if run and not tdr_bytes and not data_details_files:
        st.warning("Please provide **TDR Data** (upload file(s) or pick from SharePoint).")
    elif run and (not lvt_file or lvt_file.size == 0):
        st.warning("Please upload file(s) that include **LVT Report** (sheet named BAN Wise Result).")
    elif run and (tdr_bytes or data_details_files) and lvt_file and lvt_file.size > 0:
        tmpdir = tempfile.mkdtemp(prefix="tdr_streamlit_")
        try:
            all_sources = []
            source_display_names = []
            if data_details_files:
                for i, f in enumerate(data_details_files):
                    path = os.path.join(tmpdir, f"tdr_input_{i}.xlsx")
                    with open(path, "wb") as out:
                        out.write(f.getvalue())
                    wb_tdr = tdr_core.load_workbook(path, read_only=True, data_only=True)
                    roles = tdr_core.detect_excel_roles(wb_tdr)
                    sheet_list = roles.get("tdr_sheets") or wb_tdr.sheetnames
                    wb_tdr.close()
                    if sheet_list:
                        all_sources.append((path, sheet_list))
                        source_display_names.append(getattr(f, "name", f"Data_{i}.xlsx"))
            else:
                wb_tdr = tdr_core.load_workbook(BytesIO(tdr_bytes), read_only=True, data_only=True)
                roles = tdr_core.detect_excel_roles(wb_tdr)
                all_sheet_names = wb_tdr.sheetnames
                wb_tdr.close()
                sheet_list = roles.get("tdr_sheets") or all_sheet_names
                if sheet_list:
                    tdr_path = os.path.join(tmpdir, "tdr_input.xlsx")
                    with open(tdr_path, "wb") as f:
                        f.write(tdr_bytes)
                    all_sources.append((tdr_path, sheet_list))
                    source_display_names.append(tdr_name if tdr_name else "TDR Data.xlsx")
            if not all_sources:
                st.error("No TDR sheets found in the Data details Excel(s).")
            else:
                os.environ["TDR_WEB_REPORT_FOLDER"] = tmpdir
                lvt_path = os.path.join(tmpdir, "lvt_input.xlsx")
                with open(lvt_path, "wb") as f:
                    f.write(lvt_file.getvalue())
                # Resolve LVT sheet name only when running (we deferred workbook read above)
                try:
                    wb_lvt = tdr_core.load_workbook(lvt_path, read_only=True, data_only=True)
                    roles_lvt = tdr_core.detect_excel_roles(wb_lvt)
                    wb_lvt.close()
                    lvt_sheet_list = roles_lvt.get("lvt_sheets") or []
                    sheet_to_use = (tdr_core.LVT_SHEET_NAME if tdr_core.LVT_SHEET_NAME in (roles_lvt.get("lvt_sheets") or []) else (lvt_sheet_list[0] if lvt_sheet_list else tdr_core.LVT_SHEET_NAME)).strip() or tdr_core.LVT_SHEET_NAME
                except Exception:
                    sheet_to_use = tdr_core.LVT_SHEET_NAME
                device_details_path = None
                if device_file and device_file.size > 0:
                    device_details_path = os.path.join(tmpdir, "device_details.xlsx")
                    with open(device_details_path, "wb") as f:
                        f.write(device_file.getvalue())
                out_path = os.path.join(tmpdir, "TDR_BAN_Report.xlsx")
                bml_path = None
                if bml_file and bml_file.size > 0:
                    bml_path = os.path.join(tmpdir, "bml_input.xlsx")
                    with open(bml_path, "wb") as f:
                        f.write(bml_file.getvalue())
                with st.spinner("Running TDR extraction (this may take a minute for large files)…"):
                    result_path, summary = tdr_core.run_extraction_and_report(
                        all_sources, output_excel=out_path,
                        lvt_report_path=lvt_path, lvt_sheet_name=sheet_to_use if lvt_path else None,
                        device_details_path=device_details_path,
                        bml_path=bml_path,
                        source_display_names=source_display_names if source_display_names else None,
                    )
                if result_path and os.path.isfile(result_path):
                    with open(result_path, "rb") as f:
                        report_bytes = f.read()
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    report_filename = f"TDR_BAN_Report_{ts}.xlsx"
                    zip_bytes = None
                    zip_filename = f"TDR_wise_report_{datetime.now().strftime('%Y%m%d')}.zip"
                    per_tdr_folder = (summary or {}).get("per_tdr_folder")
                    per_tdr_file_names = (summary or {}).get("per_tdr_file_names") or []
                    if per_tdr_folder and os.path.isdir(per_tdr_folder):
                        files = per_tdr_file_names if per_tdr_file_names else [n for n in os.listdir(per_tdr_folder) if n.endswith((".xlsx", ".xlsm"))]
                        has_device = device_file and device_file.size > 0 and device_details_path and os.path.isfile(device_details_path)
                        has_bml = bml_file and bml_file.size > 0
                        if files or has_device or has_bml:
                            buf = BytesIO()
                            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                                for n in files:
                                    fp = os.path.join(per_tdr_folder, n)
                                    if os.path.isfile(fp):
                                        z.write(fp, n)
                                if has_device:
                                    z.write(device_details_path, "Pre-load device details.xlsx")
                                if has_bml and bml_path and os.path.isfile(bml_path):
                                    z.write(bml_path, "BML.xlsx")
                            buf.seek(0)
                            zip_bytes = buf.getvalue()
                    qe_mbl_bytes = None
                    qe_mbl_filename = f"QE_MBL_BAN_LIST_{datetime.now().strftime('%d%m%Y')}.xlsx"
                    delivery_status_rows = (summary or {}).get("delivery_status_rows")
                    if delivery_status_rows:
                        try:
                            qe_mbl_bytes = tdr_core.build_qe_mbl_ban_list_workbook(
                                bml_path, device_details_path,
                                delivery_status_rows,
                                device_details_sheet_name=None,
                            )
                        except Exception:
                            qe_mbl_bytes = None
                    st.session_state["tdr_result"] = {
                        "report_bytes": report_bytes, "report_filename": report_filename,
                        "zip_bytes": zip_bytes, "zip_filename": zip_filename,
                        "qe_mbl_bytes": qe_mbl_bytes, "qe_mbl_filename": qe_mbl_filename,
                        "summary": summary, "lvt_used": True,
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
        if r.get("summary"):
            if r["summary"].get("lvt_filter_applied"):
                st.info("Report shows only customers that appear in the LVT file (LVT filter applied).")
            else:
                st.warning("LVT filter could not be applied (no BANs found in LVT sheet). Report shows all rows from Data details. Ensure your LVT file has a sheet like **BAN Wise Result** with BAN and Status columns.")
        if r.get("zip_bytes") or r.get("qe_mbl_bytes"):
            n_cols = 3 if r.get("qe_mbl_bytes") else 2
            cols = st.columns(n_cols)
            with cols[0]:
                st.download_button("Download main report", data=r["report_bytes"], file_name=r["report_filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_main")
            with cols[1]:
                if r.get("zip_bytes"):
                    st.download_button("Download per-TDR (ZIP)", data=r["zip_bytes"], file_name=r["zip_filename"], mime="application/zip", key="dl_zip")
                else:
                    st.write("")
            if r.get("qe_mbl_bytes"):
                with cols[2]:
                    st.download_button("Download QE_MBL_BAN_LIST", data=r["qe_mbl_bytes"], file_name=r["qe_mbl_filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_qe_mbl")
        else:
            st.download_button("Download main report", data=r["report_bytes"], file_name=r["report_filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_main")
        if r.get("summary"):
            s = r["summary"]
            total = s.get("total", 0)
            passed, failed, not_found = s.get("passed", 0), s.get("failed", 0), s.get("not_found", 0)
            tdr_p, tdr_f, tdr_part = s.get("tdr_passed", 0), s.get("tdr_failed", 0), s.get("tdr_partial", 0)
            total_tdr = tdr_p + tdr_f + tdr_part

            st.markdown("---")
            st.markdown(
                """
                <div style="
                    background: linear-gradient(135deg, #f0fdfa 0%, #e0f2fe 100%);
                    border: 1px solid #0d9488;
                    border-radius: 10px;
                    padding: 1rem 1.25rem;
                    margin: 0.5rem 0 1rem 0;
                    box-shadow: 0 1px 3px rgba(13, 148, 136, 0.15);
                ">
                    <div style="font-size: 1.1rem; font-weight: 700; color: #0f766e; margin-bottom: 1rem;">High-level summary</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # BAN wise summary — only Total, Passed, Failed (no Not found)
            ban_html = f"""
            <div style="
                border: 1px solid #0d9488;
                border-radius: 8px;
                padding: 1rem 1.25rem;
                margin-bottom: 1rem;
                background: #fff;
            ">
                <div style="font-weight: 700; color: #0f766e; margin-bottom: 0.75rem; font-size: 1rem;">BAN wise summary</div>
                <div style="display: flex; flex-wrap: wrap; gap: 1rem;">
                    <div style="border: 1px solid #94a3b8; border-radius: 6px; padding: 0.5rem 1rem; min-width: 90px; background: #f8fafc;">
                        <div style="font-size: 0.8rem; color: #64748b;">Total BAN</div>
                        <div style="font-size: 1.25rem; font-weight: 700;">{total}</div>
                    </div>
                    <div style="border: 1px solid #22c55e; border-radius: 6px; padding: 0.5rem 1rem; min-width: 90px; background: #dcfce7;">
                        <div style="font-size: 0.8rem; color: #166534;">Passed</div>
                        <div style="font-size: 1.25rem; font-weight: 700; color: #15803d;">{passed}</div>
                    </div>
                    <div style="border: 1px solid #ef4444; border-radius: 6px; padding: 0.5rem 1rem; min-width: 90px; background: #fee2e2;">
                        <div style="font-size: 0.8rem; color: #991b1b;">Failed</div>
                        <div style="font-size: 1.25rem; font-weight: 700; color: #b91c1c;">{failed}</div>
                    </div>
                </div>
            </div>
            """
            st.markdown(ban_html, unsafe_allow_html=True)

            # TDR wise summary — bordered section, total first then breakdown with colors
            tdr_html = f"""
            <div style="
                border: 1px solid #0d9488;
                border-radius: 8px;
                padding: 1rem 1.25rem;
                margin-bottom: 0.5rem;
                background: #fff;
            ">
                <div style="font-weight: 700; color: #0f766e; margin-bottom: 0.75rem; font-size: 1rem;">TDR wise summary</div>
                <div style="display: flex; flex-wrap: wrap; gap: 1rem;">
                    <div style="border: 1px solid #94a3b8; border-radius: 6px; padding: 0.5rem 1rem; min-width: 90px; background: #f8fafc;">
                        <div style="font-size: 0.8rem; color: #64748b;">Total TDR</div>
                        <div style="font-size: 1.25rem; font-weight: 700;">{total_tdr}</div>
                    </div>
                    <div style="border: 1px solid #22c55e; border-radius: 6px; padding: 0.5rem 1rem; min-width: 90px; background: #dcfce7;">
                        <div style="font-size: 0.8rem; color: #166534;">TDR Passed</div>
                        <div style="font-size: 1.25rem; font-weight: 700; color: #15803d;">{tdr_p}</div>
                    </div>
                    <div style="border: 1px solid #ef4444; border-radius: 6px; padding: 0.5rem 1rem; min-width: 90px; background: #fee2e2;">
                        <div style="font-size: 0.8rem; color: #991b1b;">TDR Failed</div>
                        <div style="font-size: 1.25rem; font-weight: 700; color: #b91c1c;">{tdr_f}</div>
                    </div>
                    <div style="border: 1px solid #eab308; border-radius: 6px; padding: 0.5rem 1rem; min-width: 90px; background: #fef9c3;">
                        <div style="font-size: 0.8rem; color: #854d0e;">TDR Partial</div>
                        <div style="font-size: 1.25rem; font-weight: 700; color: #a16207;">{tdr_part}</div>
                    </div>
                </div>
            </div>
            """
            st.markdown(tdr_html, unsafe_allow_html=True)
