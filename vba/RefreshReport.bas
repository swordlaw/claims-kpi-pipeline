Attribute VB_Name = "RefreshReport"
' One-click refresh of all SQL Server queries, timestamp, and PDF export of the Summary sheet.
' Import via the VBA editor (Alt+F11 > File > Import File), save the workbook as .xlsm,
' and assign RefreshAndExport to a button on the Summary sheet.
' For each query: Data > Queries & Connections > right-click > Properties >
' uncheck "Enable background refresh" so the refresh finishes before export.

Option Explicit

Public Sub RefreshAndExport()
    Dim ws As Worksheet
    Dim pdfPath As String

    On Error GoTo Fail
    Application.ScreenUpdating = False

    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone

    Set ws = ThisWorkbook.Worksheets("Summary")
    ws.Range("B1").Value = "Last refreshed: " & Format(Now, "yyyy-mm-dd hh:nn")

    pdfPath = ThisWorkbook.Path & Application.PathSeparator & _
              "KPI_Summary_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath

    Application.ScreenUpdating = True
    MsgBox "Report refreshed and saved to:" & vbCrLf & pdfPath, vbInformation
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Refresh failed: " & Err.Description, vbExclamation
End Sub
