Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Application.ScreenUpdating = False
    CleanUp
    ImportMaster
    ImportGaps
    ImportDemandForecast    'Unmodified copy is saved during import
    ImportWeeklyForecast    'Unmodified copy is saved during import
    CombineForecasts
    MergeParts
    ExportCombined
    CreateOrderReport
    AddNotes
    ExportForecast
    Application.ScreenUpdating = True
    ThisWorkbook.Saved = True
    MsgBox ("Complete!")
    Email SendTo:="JBarnhill@wesco.com", _
          Subject:="Carrier Forecast", _
          Body:="""\\br3615gaps\gaps\Carrier\" & Format(Date, "yyyy") & " Alerts\Slink Alert " & Format(Date, "M-dd-yy") & ".xlsx"""
End Sub


