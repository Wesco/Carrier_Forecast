Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Application.ScreenUpdating = False

    On Error GoTo MAIN_ERR
    ImportMaster
    ImportGaps SimsAsText:=False
    ImportForecast "Demand (*.xlsx), *.xlsx", "Demand", Sheets("Demand").Range("A1")
    ImportForecast "Weekly (*.xlsx), *.xlsx", "Weekly", Sheets("Weekly").Range("A1")

    ExportSlink Sheets("Demand")
    ExportSlink Sheets("Weekly")

    CombineForecasts
    MergeParts
    ExportCombined
    CreateOrderReport
    AddNotes
    ExportForecast
    On Error GoTo 0

    Application.ScreenUpdating = True
    ThisWorkbook.Saved = True
    MsgBox ("Complete!")
    Email SendTo:="JBarnhill@wesco.com", _
          CC:="ACoffey@wesco.com", _
          Subject:="Carrier Forecast", _
          Body:="""\\br3615gaps\gaps\Carrier\" & Format(Date, "yyyy") & " Alerts\Slink Alert " & Format(Date, "M-dd-yy") & ".xlsx"""
    Exit Sub

MAIN_ERR:
    MsgBox Prompt:="Error " & Err.Number & " (" & Err.Description & ") occurred in " & Err.Source & ".", _
           Title:="Oops!"
End Sub

Sub Clean()
    Dim PrevDispAlerts As Boolean
    Dim s As Worksheet

    PrevDispAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ThisWorkbook.Activate
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Columns.Hidden = False
            s.Rows.Hidden = False
            s.Cells.Delete
            s.Range("A1").Select
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select

    Application.DisplayAlerts = PrevDispAlerts
End Sub
