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
    sPath = "\\br3615gaps\gaps\Carrier\" & Format(Date, "yyyy") & " Alerts\"

    If FileExists(sPath) = False Then RecMkDir sPath
    Worksheets("Forecast").Copy
    Sheets.Add After:=Sheets(Sheets.Count), Count:=2

    With ActiveWorkbook.Worksheets("Forecast").ListObjects("Table1").Sort.SortFields
        .Clear
        .Add Key:=Range("Table1[LT/Days]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    End With
    With ActiveWorkbook.Worksheets("Forecast").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Worksheets("Sheet2").Name = "Order"
    Worksheets("Sheet3").Name = "Expedite"
    Worksheets(1).Select

    On Error Resume Next
    ActiveWorkbook.SaveAs sPath & "Slink Alert " & Format(Date, "M-dd-yy") & ".xlsx", xlOpenXMLWorkbook
    On Error GoTo 0

    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
End Sub
