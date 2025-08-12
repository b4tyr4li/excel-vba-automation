Attribute VB_Name = "Module1"
Option Explicit

Sub RefreshReport()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 1) Recalculate workbook
    Application.CalculateFull

    ' 2) Refresh pivots, queries, connections, charts
    ThisWorkbook.RefreshAll

    ' 3) Export Summary as PDF
    On Error Resume Next
    Sheets("Summary").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:="C:\Users\batyr\Desktop\Git\excel-vba-automation\reports\Sales_Report.pdf"
    On Error GoTo 0

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Refreshed.", vbInformation
End Sub


