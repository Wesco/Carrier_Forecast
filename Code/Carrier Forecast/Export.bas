Attribute VB_Name = "Export"
Option Explicit

Sub ExportSlink(SourceSheet As Worksheet)
    Dim PrevDispAlert As Boolean
    Dim Path As String
    Dim Name As String

    PrevDispAlert = Application.DisplayAlerts
    Path = "\\br3615gaps\gaps\Carrier\" & Format(Date, "yyyy") & "Slink\"
    Name = SourceSheet.Name & " " & Format(Date, "yyyy-mm-dd")

    SourceSheet.Copy
    ActiveSheet.Name = Name
    If Not FolderExists(Path) Then RecMkDir (Path)

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Path & Name & ".xlsx", xlOpenXMLWorkbook
    Application.DisplayAlerts = PrevDispAlert
    ActiveWorkbook.Close
End Sub

Sub ExportForecast()
    Dim sPath As String
    Dim TotalRows As Long

    sPath = "\\br3615gaps\gaps\Carrier\" & Format(Date, "yyyy") & " Alerts\"

    If Not FileExists(sPath) Then RecMkDir sPath
    Sheets("Forecast").Copy
    Sheets.Add After:=Sheets(Sheets.Count), Count:=2
    Sheets("Sheet2").Name = "Order"
    Sheets("Sheet3").Name = "Expedite"

    Sheets("Forecast").Select
    TotalRows = Columns(4).Rows(Rows.Count).End(xlUp).Row
    ActiveSheet.UsedRange.Sort Key1:=Range("M1:M" & TotalRows), Order1:=xlDescending, Header:=xlYes

    On Error Resume Next
    ActiveWorkbook.SaveAs sPath & "Slink Alert " & Format(Date, "M-dd-yy") & ".xlsx", xlOpenXMLWorkbook
    On Error GoTo 0

    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
End Sub
