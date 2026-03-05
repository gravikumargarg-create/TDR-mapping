"""Production LVT TDR Delivery view – run from app.py when portal_view == 'production'."""
import tempfile
from pathlib import Path

import streamlit as st

try:
    from lvt_tdr_core import run_lvt_tdr_from_paths
except ImportError:
    run_lvt_tdr_from_paths = None


def render_production():
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
            <div style="font-size: 1.25rem; font-weight: 700;">LVT TDR Delivery (Production data)</div>
            <div style="font-size: 0.75rem; opacity: 0.95;">Upload LVT report + data Excel files → report + INSERT SQL (no DB connection).</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("**1. LVT report** (Excel with BAN/customer IDs and Pass/Fail status)")
    lvt_file = st.file_uploader("LVT Excel", type=["xlsx", "xlsm"], key="lvt_prod")
    lvt_sheet = st.text_input("LVT sheet name", value="BAN Wise Result", key="lvt_sheet_prod")

    st.markdown("**2. Data Excel files** (all non-LVT Excel files; TDR data, Rate Plan, etc.)")
    data_files = st.file_uploader("Data Excel files (multiple)", type=["xlsx", "xlsm"], accept_multiple_files=True, key="data_prod")

    st.markdown("**3. Optional – for INSERT SQL**")
    owner = st.text_input("OWNER (for SQL)", value="", key="owner_prod")
    requestor = st.text_input("REQUESTOR (for SQL)", value="", key="requestor_prod")
    default_tdr = st.text_input("Default TDR for 'Found but no TDR' rows", value="", key="default_tdr_prod")

    run = st.button("Run LVT TDR", type="primary")

    if run and run_lvt_tdr_from_paths is None:
        st.error("Could not load LVT TDR module. Ensure lvt_tdr_core.py is in the app folder.")
    elif run:
        if not lvt_file or lvt_file.size == 0:
            st.warning("Please upload the **LVT Excel** file.")
        elif not data_files or all(f.size == 0 for f in data_files):
            st.warning("Please upload at least one **data Excel** file.")
        else:
            tmpdir = Path(tempfile.mkdtemp(prefix="lvt_tdr_streamlit_"))
            try:
                lvt_path = tmpdir / "lvt_input.xlsx"
                lvt_path.write_bytes(lvt_file.getvalue())
                data_paths = []
                for i, uf in enumerate(data_files):
                    if uf.size == 0:
                        continue
                    name = uf.name or f"data_{i}.xlsx"
                    data_paths.append(tmpdir / name)
                    data_paths[-1].write_bytes(uf.getvalue())
                out_dir = tmpdir / "report"
                out_dir.mkdir(parents=True, exist_ok=True)
                log_lines = []
                def log_fn(msg):
                    log_lines.append(msg)
                with st.spinner("Processing…"):
                    report_path, sql_path = run_lvt_tdr_from_paths(
                        lvt_path, data_paths, out_dir,
                        lvt_sheet_name=lvt_sheet or "BAN Wise Result",
                        owner=owner or None, requestor=requestor or None, default_tdr_id=default_tdr or None,
                        log_fn=log_fn,
                    )
                for line in log_lines:
                    st.text(line)
                if report_path and report_path.is_file():
                    report_bytes = report_path.read_bytes()
                    sql_bytes = sql_path.read_bytes() if sql_path and sql_path.is_file() else None
                    st.session_state["lvt_result"] = {
                        "report_bytes": report_bytes,
                        "report_name": report_path.name,
                        "sql_bytes": sql_bytes,
                        "sql_name": sql_path.name if sql_path else None,
                    }
                else:
                    st.error("Report generation failed.")
            except Exception as e:
                st.exception(e)
            finally:
                try:
                    import shutil
                    shutil.rmtree(tmpdir, ignore_errors=True)
                except Exception:
                    pass

    if "lvt_result" in st.session_state:
        r = st.session_state["lvt_result"]
        st.success("Done — download report and INSERT SQL below.")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("Download report (Excel)", data=r["report_bytes"], file_name=r["report_name"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_report")
        with c2:
            if r.get("sql_bytes"):
                st.download_button("Download INSERT SQL", data=r["sql_bytes"], file_name=r["sql_name"], mime="text/plain", key="dl_sql")
            else:
                st.info("No INSERT SQL (no eligible rows: LVT Passed + Found/Found but no TDR).")
