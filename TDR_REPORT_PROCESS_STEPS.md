# TDR Report – Step-by-step process

This document describes how the TDR report is built. **Current behaviour:** only customers that appear in the LVT report are kept; the main Excel and per-TDR Excels contain only those LVT-matched rows.

---

## 1. Inputs (UI)

| Step | What you provide | Used for |
|------|------------------|----------|
| **1. Data details** | One or more Excel file(s) (TDR data) | Source of **TDR → BAN** pairs (all rows in the report). |
| **2. LVT file** | One Excel with BAN Wise Result (or similar) | **Status** (Passed/Failed/Not found) per BAN; optional failure details. |
| **3. Device details** (optional) | Excel with CUSTOMER_ID | Added to per-TDR Excel files in the ZIP. |
| **4. BML file** (optional) | BML Excel | Included in the ZIP. |

**Run TDR** requires at least **Data details** and **LVT file**.

---

## 2. Building the list of (TDR, BAN) rows from Data details

**Where in code:** `tdr_core.run_extraction_and_report()` → loop over `all_sources` → `extract_tdr_ban_mapping(ws)`.

- **all_sources**  
  For each Data details file you uploaded, we get a file path and a list of **sheets** to process.

- **For each file and each sheet:**
  1. **Find TDR sections**  
     Sheet is scanned for rows that contain a TDR id (e.g. `TDR-201878`). Each such row starts a “section” up to the next TDR row (or end of sheet).
  2. **Collect BANs per section**  
     In each section, every cell is scanned for **9-digit BAN** numbers. All such BANs are collected for that TDR.
  3. **Append one row per (TDR, BAN)**  
     `all_rows.append((tdr_id, ban))` for each TDR and each BAN. We also store **tdr_section_ranges** (tdr_id, excel_path, sheet_name, start_row, end_row) for later per-TDR Excel creation.

- **Result:**  
  `all_rows` = all **(TDR, BAN)** pairs from Data details. No per-TDR files are written yet.

---

## 3. Loading LVT and keeping only LVT-matched customers

**Where in code:** Right after the Data details loop; `_build_ban_to_status_from_sheet(lvt_ws)`.

- LVT workbook is opened (sheet = chosen LVT sheet, usually “BAN Wise Result”).
- We build **ban_to_status**: normalized BAN (9-digit) → status (e.g. "Passed", "Failed").
- Optional: **ban_to_failures** from a “Failures” sheet for failure details.

**Filter:**  
- **rows_in_lvt** = only those (TDR, BAN) from `all_rows` whose BAN is in **ban_to_status**.  
- If LVT was not provided or could not be loaded, **rows_in_lvt = all_rows** (no filter).

**Result:**  
All downstream steps (main Excel, per-TDR Excels, summary) use **rows_in_lvt** only. Customers not in the LVT report are excluded from the output.

---

## 4. Per-TDR Excels (only for LVT-matched TDRs; 3 sheets each)

**Where in code:** After computing **rows_in_lvt**; only TDRs that have at least one row in **rows_in_lvt** get a file.

For each such **tdr_id**:

1. **Sheet 1 – Data**  
   Data from Data details for that TDR: we open the source Excel and copy the TDR section (same range we detected earlier) into a new workbook as sheet **"Data"**.

2. **Sheet 2 – Device Details**  
   If the user provided a Device details file, we add a sheet **"Device Details"** with rows where CUSTOMER_ID is in the set of BANs for this TDR (from **rows_in_lvt**).

3. **Sheet 3 – BML**  
   If the user provided a BML file, we add a sheet **"BML"** by copying the first sheet of the BML workbook.

Each file is saved as `{tdr_excel_folder}/{TDR-xxxxxx}.xlsx` and included in the per-TDR ZIP.

## 5. Main report Excel (LVT-matched rows only)

**Where in code:** When `output_excel` is set and **rows_in_lvt** is not empty.

### 5.1 TDR Info sheet

- **Data rows:** One row per entry in **rows_in_lvt** (TDR, BAN, Status). So the main Excel contains **only customers that exist in the LVT report**.
- **Status** is filled from **ban_to_status**; **Failure details** from **ban_to_failures** if present.

### 5.2 TDR Summary sheet

- **Input:** **rows_in_lvt** + **ban_to_status**.
- One row per TDR that appears in **rows_in_lvt**: TDR | Total BANs | Passed | Failed | Not found | TDR Status.

### 5.3 Dashboard summary (BAN wise)

- **Total BAN** = number of **distinct** BANs in **rows_in_lvt**.
- **Passed / Failed / Not found** = counts of those distinct BANs by status.

---

---

## 6. Quick check you can do

1. **Data details:**  
   Count how many **sheets** you selected (or how many files × sheets are used).  
   In each sheet, count how many **TDR sections** there are and how many **BANs** per section.  
   Rough total rows ≈ sum over (all TDRs × BANs in that TDR) across all sheets → should be in the same ballpark as 600+.

2. **LVT:**  
   Number of rows in “BAN Wise Result” (or the chosen LVT sheet) ≈ number of **distinct** BANs we can mark as Passed/Failed.  
   “Total BAN” on the dashboard is **distinct** BANs from the **Data details** side; “Passed + Failed + Not found” in the summary are those distinct BANs classified by LVT.

3. **TDR Summary sheet:**  
   “Total BANs” per TDR = row count for that TDR in **all_rows**.  
   Sum of “Total BANs” over all TDRs can be **greater** than the dashboard “Total BAN” because the same BAN can appear in multiple TDRs.

---

## 7. Code reference (tdr_core.py)

| Step | Function / location |
|------|----------------------|
| Build `all_rows` and `tdr_section_ranges` | `run_extraction_and_report()`: loop over `all_sources`, `extract_tdr_ban_mapping(ws)`, append to `all_rows` and `tdr_section_ranges`. |
| LVT and filter | Load LVT → `ban_to_status`; `rows_in_lvt = [(t,b) for (t,b) in all_rows if _normalize_ban(b) in ban_to_status]`. |
| Per-TDR workbooks | For each TDR in `rows_in_lvt`: copy Data sheet from source, `_add_device_details_sheet_to_workbook`, `_add_bml_sheet_to_workbook`; save. |
| Main Excel | Write **rows_in_lvt** to TDR Info; `_fill_tdr_info_status_column`, `_fill_tdr_info_failure_columns`; `_build_tdr_summary(rows_in_lvt, ban_to_status)`; `_add_tdr_summary_sheet`. |
| BML sheet | `_add_bml_sheet_to_workbook(wb, bml_path)` copies the first sheet of the BML workbook into `wb` as sheet "BML". |
