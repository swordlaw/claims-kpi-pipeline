<#
Build ClaimsKPI_Report.xlsm: the Excel version of the report with live SQL Server queries.

Starts from ClaimsKPI_Report.xlsx (made by build_report.py) and, on each view sheet,
replaces the snapshot with a live ODBC query to the matching SQL view. Background
refresh is turned off on every query so RefreshAndExport waits for the data. Adds the
"Refresh & Export PDF" button on Summary and saves as macro-enabled .xlsm.

The VBA module itself is imported by hand (Alt+F11 > File > Import File >
vba\RefreshReport.bas). Importing it from a script would need Excel's "Trust access
to the VBA project object model" security setting, which is best left off.

Usage:  powershell -ExecutionPolicy Bypass -File scripts\build_excel_report.ps1
Needs:  Excel (desktop), ODBC Driver 18 for SQL Server, ClaimsKPI database
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$src  = Join-Path $root 'ClaimsKPI_Report.xlsx'
$out  = Join-Path $root 'ClaimsKPI_Report.xlsm'

$conn = 'ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=localhost\SQLEXPRESS;' +
        'DATABASE=ClaimsKPI;Trusted_Connection=yes;TrustServerCertificate=yes'

# sheet / view name -> ORDER BY (column order must stay the same as the view)
$views = [ordered]@{
    'v_kpi_monthly'             = 'month_start, claim_type'
    'v_kpi_weekly'              = 'week_start, claim_type'
    'v_kpi_annual_pmpm'         = 'yr, claim_type'
    'v_kpi_inpatient_quarterly' = 'yr, qtr'
    'v_inpatient_stays'         = 'member_id, admit_dt'
    'v_paid_by_diagnosis'       = 'total_paid DESC'
}
$formats = @{
    month_start = 'yyyy-mm-dd'; week_start = 'yyyy-mm-dd'; admit_dt = 'yyyy-mm-dd'; discharge_dt = 'yyyy-mm-dd'
    total_paid = '"$"#,##0'; avg_paid_per_claim = '"$"#,##0'; paid_amount = '"$"#,##0'; pmpm = '"$"#,##0.00'
    readmit_30_rate = '0.0%'
}

if (Test-Path $out) { Remove-Item $out }
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($src)

    foreach ($name in $views.Keys) {
        $ws = $wb.Worksheets.Item($name)
        [void]$ws.Cells.Clear()
        # 0 = xlSrcExternal: an Excel table backed by an external query
        $lo = $ws.ListObjects.Add(0, $conn, $null, 1, $ws.Range('A1'))
        $lo.Name = "tbl_$name"
        $qt = $lo.QueryTable
        $qt.CommandText = "SELECT * FROM dbo.$name ORDER BY $($views[$name])"
        $qt.BackgroundQuery = $false      # the macro waits for this query to finish
        $qt.RefreshOnFileOpen = $false
        $qt.PreserveColumnInfo = $true
        $qt.WorkbookConnection.Name = "SQL $name"
        [void]$qt.Refresh($false)

        foreach ($col in $lo.ListColumns) {
            if ($formats.ContainsKey($col.Name)) { $col.DataBodyRange.NumberFormat = $formats[$col.Name] }
        }
        [void]$ws.Range('A1').EntireRow.EntireColumn.AutoFit()
        Write-Output ("{0,-28} {1,7:N0} rows (live query, background refresh off)" -f $name, $lo.ListRows.Count)
    }

    # Button on Summary, top right over O1:Q2, clear of the title and of B1 (where the macro writes "Last refreshed")
    $sum = $wb.Worksheets.Item('Summary')
    $r = $sum.Range('O1:Q2')
    $btn = $sum.Buttons().Add($r.Left, $r.Top, $r.Width, $r.Height)
    $btn.Name = 'btnRefreshAndExport'
    $btn.OnAction = 'RefreshAndExport'
    $btn.Characters().Text = 'Refresh & Export PDF'
    $btn.PrintObject = $false            # keep the button out of the PDF
    $sum.Range('B1').Value2 = "Last refreshed: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $xl.CalculateFull()
    $sum.Activate()

    $wb.SaveAs($out, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    $wb.Close($false)
    Write-Output "Saved $out"
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
}
