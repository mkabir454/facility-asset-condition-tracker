Attribute VB_Name = "Module1"
Sub ExportDashboardPDF()
    Dim f As String
    Application.CalculateFull
    f = ThisWorkbook.Path & Application.PathSeparator & "Asset_Dashboard.pdf"
    ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat Type:=xlTypePDF, FileName:=f
    MsgBox "Saved: " & f
End Sub
