# TDR Data Excel – Streamlit app (free)

Same TDR tool as the local script, running in the browser. **Free** on [Streamlit Community Cloud](https://streamlit.io/cloud) – no paid cloud needed.

## Run locally

```bash
cd streamlit_tdr
pip install -r requirements.txt
streamlit run app.py
```

Open http://localhost:8501 in your browser.

## Deploy to Streamlit Community Cloud (free)

1. Push this repo to **GitHub** (or use an existing repo that contains the `streamlit_tdr` folder).
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **New app**.
4. Set:
   - **Repository**: your repo (e.g. `yourusername/voxelize`)
   - **Branch**: `main` (or your default)
   - **Main file path**: `streamlit_tdr/app.py`
   - **Advanced** → **Python version**: 3.11 (or 3.10)
5. If the app root is the repo root, Streamlit will look for `streamlit_tdr/app.py` and `streamlit_tdr/tdr_core.py`. If your repo root is already `streamlit_tdr`, use `app.py` as main file path.
6. Click **Deploy**. Wait a few minutes.
7. You’ll get a URL like `https://your-app-name.streamlit.app`.

## Link from UTMO QE Hub

On your dashboard (e.g. QE Tools), add a link that opens the Streamlit app in a new tab:

- **Label:** e.g. "TDR Data Excel (web)"
- **URL:** your Streamlit app URL from step 7

Example: in the TDR Data Excel tab you can add:  
*"Or use the free web version: [Open TDR tool](https://your-app-name.streamlit.app)"*

## Runbook (for team sharing)

- **RUNBOOK.md** – Full runbook (purpose, access, inputs, steps, outputs, troubleshooting). Open in any editor or in Word (File → Open → select RUNBOOK.md, then Save As → .docx if needed).
- **Word file:** To generate **RUNBOOK.docx** for sharing:  
  `pip install python-docx`  
  `python generate_runbook.py`  
  This creates **RUNBOOK.docx** in the repo folder.

## Revert to state before direct SharePoint

If direct SharePoint integration causes issues, revert to link + upload only:

```bash
git checkout before-sharepoint-direct -- app.py
git commit -m "Revert to link + upload only"
git push origin main
```

## Notes

- Your **local** `tdr_excel_script.py` (in the project root) is **not** changed. This app uses the **copy** in `streamlit_tdr/tdr_core.py`.
- No billing: Streamlit Community Cloud is free for public apps.
- You don’t need to keep your computer on; the app runs on Streamlit’s servers.
