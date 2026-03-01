# TDR Mapping Sheet Creation – Runbook

**Version:** 1.0  
**Tool:** TDR Streamlit App (TDR mapping sheet creation)  
**Audience:** Team members who run the tool or need to support it.

---

## 1. Purpose

The TDR Mapping Sheet Creation tool:

- Reads **TDR Data** (Excel with TDR-###### sections and BANs), **LVT Report** (Excel with BAN-wise status, e.g. BAN Wise Result), and optionally **Device Details** (Excel with CUSTOMER_ID and device columns).
- Produces:
  - A **main report** Excel with **TDR Info** (TDR, BAN, Status, Failure Description, Check ID, Comments) and **TDR Summary**.
  - **Per-TDR Excel files** (one per TDR) in a ZIP, each with the TDR section and optionally a **Device Details** sheet (when Device Details Excel is provided).
- Optionally adds **Device Details** to each TDR-wise file by matching BAN/Customer ID from an uploaded Device Details Excel.

---

## 2. How to Access

The tool runs on **Streamlit Cloud**. No local installation required.

**App URL:**  
**https://tdr-streamlit-sharepoint.streamlit.app**

*(Use the URL above, or the URL shared by your team if different.)*

- Open the URL in a browser to use the app.

---

## 3. Prerequisites

- **TDR Data** – Excel (.xlsx/.xlsm) with at least one sheet containing TDR section headers (e.g. TDR-202958) and 9-digit BAN IDs.
- **LVT Report** – Excel with a sheet named **BAN Wise Result** (or similar; spaces/underscores allowed) containing BAN and Status (e.g. Passed/Failed).
- **Device Details (optional)** – Excel with a sheet whose first row has **CUSTOMER_ID** (or BAN/Customer ID) and device columns (MSISDN, IMEI, ESN, EID, DEVICE_MODEL, UICCID, UIMSI, TIMSI, DEVICE_LOCK_STATUS).

---

## 4. Input Options

### 4.1 TDR file from

- **Local file:** TDR Data will be chosen from the files you upload (see below).
- **SharePoint:** If configured, you can pick the TDR file directly from the SharePoint folder (no download-then-upload). Otherwise, use the link to open the folder, download the file, and upload it with the other files.

### 4.2 Upload file(s)

- **One file upload accepts multiple files.**  
  Upload one or more Excel files in a single step.
- The app **automatically detects**:
  - **TDR Data** – first file that has a sheet with TDR sections (TDR-###### and BANs).
  - **LVT Report** – first file that has a sheet named **BAN Wise Result**.
  - **Device Details** – first file that has a sheet with a **CUSTOMER_ID**-style column.
- After detection you get:
  - **TDR Data** → dropdown to choose **which sheet** to use (only sheets with TDR content).
  - **LVT Report** → dropdown to choose **which sheet** (default: BAN Wise Result).
  - **Device Details** → used automatically (first sheet with CUSTOMER_ID); no sheet pick needed.

You can upload 1, 2, or 3 files; the same file can be used for more than one role if it contains the right sheets.

---

## 5. Steps to Run

1. Choose **TDR file from:** Local file or SharePoint (if available).
2. **Upload** one or more Excel files (TDR Data, LVT Report, and/or Device Details).
3. Confirm the detected **TDR Data** and **LVT Report** files and select the **sheets** from the dropdowns.
4. If Device Details was detected, it will be used for the “Device Details” sheet in each TDR-wise file.
5. Click **Run TDR**.
6. When processing finishes:
   - **Download main report** – Excel with TDR Info and TDR Summary.
   - **Download per-TDR (ZIP)** – ZIP containing one Excel per TDR (and Device Details sheet when applicable).

---

## 6. Outputs

### 6.1 Main report Excel

- **TDR Info** sheet:  
  Columns: **TDR**, **BAN**, **Status**, **Failure Description**, **Check ID**, **Comments**.
  - Status comes from LVT (BAN Wise Result). Passed/Failed/Not found.
  - Failure Description and Check ID come from the **BAN Wise Failures** sheet in the LVT file (if present). For **Passed** rows these show **N/A**.
  - Comments: filled for certain Check IDs (e.g. 3106) as configured.
- **TDR Summary** sheet: Per-TDR counts (total, passed, failed, not found) and overall summary.

### 6.2 Per-TDR Excel files (in ZIP)

- One file per TDR (e.g. `TDR-202958.xlsx`).
- First sheet: TDR section data from the TDR Data file.
- **Device Details** sheet (if Device Details Excel was uploaded): Rows from that Excel where **CUSTOMER_ID** matches the BANs for that TDR, **sorted by CUSTOMER_ID ascending**. Long numeric columns (e.g. IMEI, ESN, EID, UICCID) are stored as text to avoid truncation.

---

## 7. Device Details

- **Source:** Optional Excel with **CUSTOMER_ID** (or equivalent) and device columns.
- **Usage:** For each TDR, the tool finds all rows in the Device Details file whose CUSTOMER_ID is in that TDR’s BAN list, copies them into the **Device Details** sheet of that TDR’s Excel, **sorted by CUSTOMER_ID ascending**.
- **Formatting:** Long IDs (CUSTOMER_ID, MSISDN, IMEI, ESN, EID, UICCID, UIMSI, TIMSI) are written as text so Excel does not change or truncate them.

---

## 8. SharePoint (optional)

- If your admin has granted **Sites.Read.All** for the app, you can choose **TDR file from → SharePoint** and pick the TDR file from a dropdown (no manual download/upload of that file).
- LVT and Device Details are still provided via the **Upload file(s)** option.
- Secrets (e.g. AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET) must be set in the app settings (e.g. Streamlit Cloud → Settings → Secrets). Do not commit secrets to code or Git.

---

## 9. Troubleshooting

| Issue | What to do |
|-------|------------|
| “No TDR Data detected” | Ensure at least one uploaded file has a sheet containing TDR-###### headers and 9-digit BANs. |
| “No LVT Report detected” | Ensure at least one file has a sheet named **BAN Wise Result** (or equivalent). |
| Status always “Not found” | Check that the LVT sheet has BAN and Status columns and that BAN values match the 9-digit BANs in TDR Data. |
| Device Details sheet empty | Confirm the Device Details Excel has a column like CUSTOMER_ID and that values match the BANs in the TDR. |
| Long numbers look wrong in Excel | Device Details sheet formats CUSTOMER_ID, IMEI, ESN, EID, UICCID, etc. as text; re-run with the latest version if you still see issues. |
| SharePoint dropdown not shown | Set the three Azure secrets in the app and ensure admin has granted **Sites.Read.All** for the app. |

---

## 10. Support and Repo

- **Repository:** TDR-mapping (e.g. GitHub repo used for deployment).
- **Main files:** `app.py` (Streamlit UI), `tdr_core.py` (extraction and report logic), `sharepoint_graph.py` (SharePoint/Graph API).
- For runbook updates or process changes, update this runbook and share the new version (e.g. Word export) with the team.

---

*End of runbook.*
