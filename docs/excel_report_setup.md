# Report Setup: Google Sheets and Excel

`scripts/build_report.py` produces **`ClaimsKPI_Report.xlsx`**:

| Sheet | Contents |
|---|---|
| `Summary` | Headline KPIs, three chart tables (formulas), three charts |
| `v_kpi_monthly`, `v_kpi_weekly`, `v_kpi_annual_pmpm`, `v_kpi_inpatient_quarterly`, `v_inpatient_stays`, `v_paid_by_diagnosis` | One sheet per SQL view, same columns in the same order |

The Summary tables are `SUMIFS` / `AVERAGEIFS` formulas with whole-column references (for example `v_kpi_monthly!$E:$E`). They don't care where the data came from, so the same Summary sheet works as a static snapshot (Google Sheets) and with live SQL Server queries (Excel).

---

## Option A: Google Sheets (static snapshot)

1. Go to <https://drive.google.com> → **New** → **File upload** → choose `ClaimsKPI_Report.xlsx`.
2. Right-click the uploaded file → **Open with** → **Google Sheets**. Then **File** → **Save as Google Sheets** to get a native copy.
3. Check the **Summary** sheet: the five headline KPIs and the three tables should show numbers. The charts are converted from Excel's format. If one looks off, select its table (for example `B12:E48` for monthly paid) → **Insert** → **Chart** and pick the chart type.

**Limitations to mention in an interview:**
- Google Sheets runs in Google's cloud and can't reach `localhost\SQLEXPRESS`, so the data is a **snapshot**. To refresh, re-run `py -3.9 scripts/build_report.py` and upload again.
- Google Sheets doesn't run VBA, so `RefreshReport.bas` only applies to the Excel version (the equivalent there would be Apps Script).

---

## Option B: Excel with live SQL Server queries and the VBA macro (`.xlsm`)

These steps need desktop Excel for Windows (Microsoft 365 or 2019+), on the same PC as SQL Server.

### 1. Connect each view sheet to SQL Server

Open `ClaimsKPI_Report.xlsx`, then repeat these steps for **each of the 6 `v_...` sheets**:

1. Go to the sheet (for example `v_kpi_monthly`). Click the corner square above row 1 to select everything, then press **Delete**. This clears the snapshot so the query has room. Don't delete the sheet itself.
2. **Data** tab → **Get Data** → **From Database** → **From SQL Server Database**.
3. **Server:** `localhost\SQLEXPRESS`  **Database:** `ClaimsKPI` → **OK**.
4. On the credentials screen, pick **Windows** → **Use my current credentials** → **Connect**.
   If you get an *encryption support* message, click **OK** to connect without encryption. This is a local server with a self-signed certificate.
5. In the **Navigator**, click the view with the **same name as the sheet** (for example `dbo.v_kpi_monthly`).
6. Click the small arrow next to **Load** → **Load To...** → choose **Table**, **Existing worksheet**, and type `=$A$1` → **OK**.
   (You've already selected the right sheet, so `$A$1` is that sheet's A1.)

Each view keeps the same column order as the snapshot, so the Summary formulas keep working.

### 2. Turn off background refresh (so the macro waits for the data)

1. **Data** → **Queries & Connections**. A pane opens on the right.
2. For **each** of the 6 queries: right-click it → **Properties...** → **untick "Enable background refresh"** → **OK**.

**Why:** with background refresh on, `RefreshAll` returns immediately while the queries are still running. The macro would then stamp the time and export the PDF using **old** data.

### 3. Save as macro-enabled

**File** → **Save As** → **Browse** → *Save as type:* **Excel Macro-Enabled Workbook (\*.xlsm)** → name it `ClaimsKPI_Report.xlsm` → **Save**.

### 4. Import the macro

1. If you don't see a **Developer** tab: **File** → **Options** → **Customize Ribbon** → tick **Developer** → **OK**.
2. Press **Alt+F11** to open the VBA editor.
3. **File** → **Import File...** → choose `vba\RefreshReport.bas` → **Open**. A module named `RefreshReport` appears under *Modules*.
4. Close the VBA editor.

### 5. Add the button on Summary

1. Go to the **Summary** sheet.
2. **Developer** → **Insert** → under *Form Controls*, click **Button** (the first icon).
3. Drag to draw the button around cells **D1:F2**. Leave **B1** clear, because the macro writes "Last refreshed: ..." there.
4. In the *Assign Macro* box, choose **RefreshAndExport** → **OK**.
5. Right-click the button → **Edit Text** → type `Refresh & Export PDF` → click any cell.
6. Press **Ctrl+S**.

### 6. Test it

1. Click the button. Excel refreshes all 6 queries. The large `v_inpatient_stays` takes a few seconds.
2. You should see a message *"Report refreshed and saved to: ...\KPI_Summary_YYYY-MM-DD.pdf"*, and **B1** should show the current time.
3. Open the PDF to check it shows the Summary sheet.
4. Take screenshots of the Summary sheet and the PDF and save them to `screenshots/`.

**If macros are blocked:** click **Enable Content** on the yellow bar when opening the file. If there's a red *"Microsoft has blocked macros"* bar instead, close Excel, right-click the `.xlsm` → **Properties** → tick **Unblock** → **OK**, then reopen it.

### What the macro does (for the interview)

`RefreshAndExport` in `vba/RefreshReport.bas`:
1. `ThisWorkbook.RefreshAll` re-runs every SQL query.
2. `Application.CalculateUntilAsyncQueriesDone` waits for any queries that are still running and for the formulas to recalculate.
3. Writes a "Last refreshed" timestamp to `Summary!B1`.
4. Exports the Summary sheet as a dated PDF next to the workbook. This is the file you'd email to stakeholders.
5. An error handler restores screen updating and shows a readable message instead of a VBA debug dialog.
