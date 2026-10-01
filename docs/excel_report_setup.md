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

This needs desktop Excel for Windows, on the same PC as SQL Server.

### 1. Build the workbook (automated)

```powershell
py -3.9 scripts\build_report.py
powershell -ExecutionPolicy Bypass -File scripts\build_excel_report.ps1
```

`build_excel_report.ps1` drives Excel through its COM automation interface:

1. Opens `ClaimsKPI_Report.xlsx` and, on each of the 6 `v_...` sheets, replaces the snapshot with an **Excel table backed by a live ODBC query**: `SELECT * FROM dbo.<view>` against `localhost\SQLEXPRESS`, using Windows authentication, so there's no password to store.
2. Sets **`BackgroundQuery = False`** on every query. Otherwise `RefreshAll` would return immediately while the queries are still running, and the macro would stamp the time and export the PDF using **old** data.
3. Adds the **Refresh & Export PDF** button at the top right of Summary (over `O1:Q2`, clear of the title and of `B1`, where the macro writes its timestamp), wired to `RefreshAndExport`. The button is set not to print.
4. Saves as **`ClaimsKPI_Report.xlsm`** (macro-enabled).

To see the queries in Excel: **Data** → **Queries & Connections** → **Connections** tab. Right-click a query → **Properties** to see "Enable background refresh" unticked and, under **Definition**, the SQL.

### 2. Import the macro (manual, once)

The script doesn't import the VBA, because that would mean turning on Excel's *"Trust access to the VBA project object model"* security setting, which is best left off.

1. Open `ClaimsKPI_Report.xlsm`.
2. Press **Alt+F11** to open the VBA editor.
3. **File** → **Import File...** → choose `vba\RefreshReport.bas` → **Open**. A module `RefreshReport` appears under *Modules*.
4. Close the VBA editor and press **Ctrl+S**.

### 3. Test it

1. On **Summary**, click **Refresh & Export PDF**.
2. Excel re-runs all 6 queries. The 66,705-row `v_inpatient_stays` takes a few seconds.
3. You should see *"Report refreshed and saved to: ...\KPI_Summary_YYYY-MM-DD.pdf"*, and **B1** should show the current time.
4. Open the PDF to check it's the Summary sheet.

**If macros are blocked:** click **Enable Content** on the yellow bar. If there's a red *"Microsoft has blocked macros"* bar instead, close Excel, right-click the `.xlsm` → **Properties** → tick **Unblock** → **OK**, then reopen it.

**If the refresh fails with a connection error:** check the SQL Server service is running (`Get-Service 'MSSQL$SQLEXPRESS'`) and that ODBC Driver 18 for SQL Server is installed.

### What the macro does (for the interview)

`RefreshAndExport` in `vba/RefreshReport.bas`:
1. `ThisWorkbook.RefreshAll` re-runs every SQL query.
2. `Application.CalculateUntilAsyncQueriesDone` waits for any queries still running and for the formulas to recalculate.
3. Writes a "Last refreshed" timestamp to `Summary!B1`.
4. Exports the Summary sheet as a dated PDF next to the workbook: **one landscape page** with the headline KPIs, the data note and the three charts. The chart tables further down the sheet are outside the print area. This is the file you'd email to stakeholders.
5. An error handler restores screen updating and shows a readable message instead of a VBA debug dialog.

### Doing it by hand instead (no script)

For each `v_...` sheet: clear the sheet → **Data** → **Get Data** → **From Database** → **From SQL Server Database** → server `localhost\SQLEXPRESS`, database `ClaimsKPI` → **Windows** / *Use my current credentials* → pick the view with the same name → **Load To...** → **Table**, existing worksheet `=$A$1`. Then, in **Queries & Connections**, untick **Enable background refresh** under each query's **Properties**. Save as `.xlsm`, import the macro (step 2), and add a button with **Developer** → **Insert** → **Button (Form Control)** → assign `RefreshAndExport`.
