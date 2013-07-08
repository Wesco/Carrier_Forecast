Attribute VB_Name = "Import"
Option Explicit

Sub ImportDemandForecast()
    Dim fPath As String    'Stores the demand forecast file location
    Dim sName As String    'New name for unmodified forecast
    Const sPath As String = "\\br3615gaps\GAPS\Carrier\2012 Slink\"

    Application.DisplayAlerts = False
    fPath = Application.GetOpenFilename("demand (*.xlsx), *.xlsx", Title:="Select demand forecast")
    sName = "Demand " & Format(Date, "m-dd-yy") & ".xlsx"

    If FileExists(sPath) <> True Then RecMkDir sPath
    
    On Error GoTo OPEN_ERROR    'Displays 'No File Selected'
    Workbooks.Open FileName:=fPath
    On Error GoTo 0

    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Demand").Range("A1")


    On Error GoTo SAVE_ERROR    'sets sName = ActiveWorkbook.Name
    Application.DisplayAlerts = True    'Show "Overwrite existing file?" prompt if file already exists
    ActiveWorkbook.SaveAs FileName:=sPath & sName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = False
    On Error GoTo 0

    ActiveWorkbook.Close
    ThisWorkbook.Activate
    Worksheets("Demand").Select

    'Cleanup formatting and remove rows/columns that are not needed
    'After cleanup is complete Col 1 = Part Numbers, Col 2+ = Dates
    With Range(Rows(1), Rows(10))
        .UnMerge
        .Range(.Rows(1), .Rows(8)).Delete Shift:=xlShiftUp
    End With
    With Range("A1:C2")
        .Rows(2).Value = .Rows(1).Value
    End With
    Rows(1).Delete Shift:=xlShiftUp
    Range("B:F").Delete Shift:=xlShiftToLeft
    Range("B:D").Delete

    With Rows("1:1")
        'Empty character = alt + 255
        Range(Cells(1, 2), Cells(1, .CurrentRegion.Columns.Count)).Replace What:=" ", Replacement:=""
        Range(Cells(1, 2), Cells(1, .CurrentRegion.Columns.Count)).NumberFormat = "mm/dd"
    End With

    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Range("A1").Select

    Delete fPath
    Application.DisplayAlerts = True
    Exit Sub

SAVE_ERROR:
    sName = ActiveWorkbook.Name
    Resume Next

OPEN_ERROR:
    Err.Raise Number:=75, Description:="No File Selected."
    Exit Sub
End Sub

Sub ImportWeeklyForecast()
    Dim fPath As String    'Stores the weekly forecast file location
    Dim sName As String    'new name for unmodified forecast
    Const sPath As String = "\\br3615gaps\GAPS\Carrier\2012 Slink\"

    Application.DisplayAlerts = False
    fPath = Application.GetOpenFilename("Weekly Forecast (*.xlsx), *.xlsx", Title:="Select weekly forecast")
    sName = "Weekly " & Format(Date, "m-dd-yy") & ".xlsx"
    
    If FileExists(sPath) <> True Then RecMkDir sPath
    
    On Error GoTo OPEN_ERROR
    Workbooks.Open FileName:=fPath
    On Error GoTo 0

    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Weekly").Range("A1")

    'DisplayAlerts = True so that if
    'the file already exists the user
    'is prompted to overwrite
    On Error GoTo SAVE_ERROR
    Application.DisplayAlerts = True
    ActiveWorkbook.SaveAs FileName:=sPath & sName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = False
    On Error GoTo 0

    ActiveWorkbook.Close
    ThisWorkbook.Activate
    Worksheets("Weekly").Select

    'Cleanup formatting and remove rows/columns that are not needed
    'After cleanup is complete Col 1 = Part Numbers, Col 2+ = Dates
    With Range(Rows(1), Rows(10))
        .UnMerge
        .Range(.Rows(1), .Rows(8)).Delete Shift:=xlShiftUp
    End With
    With Range("A1:C2")
        .Rows(2).Value = .Rows(1).Value
    End With
    Rows(1).Delete Shift:=xlShiftUp
    Range("B:F").Delete Shift:=xlShiftToLeft

    With ActiveSheet.UsedRange
        'Column J to Last Column with data
        Range(Cells(1, 10), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)).Delete Shift:=xlShiftToLeft
        'Empty character = alt + 255
        Range(Cells(1, 2), Cells(1, .CurrentRegion.Columns.Count)).Replace What:=" ", Replacement:=""
        Range(Cells(1, 2), Cells(1, .CurrentRegion.Columns.Count)).NumberFormat = "mm/dd/yyyy"
    End With

    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Range("A1").Select

    Delete fPath
    Application.DisplayAlerts = True
    Exit Sub

SAVE_ERROR:
    sName = ActiveWorkbook.Name
    Resume Next

OPEN_ERROR:
    Err.Raise Number:=75, Description:="No File Selected."
    Exit Sub
End Sub

Sub ImportGaps()
    Dim fileFound As Boolean
    Dim sName As String    'Gaps Filename
    Dim i As Integer       'Counter
    Dim sPath As String
    
    sPath = "\\BR3615GAPS\GAPS\3615 GAPS DOWNLOAD\" & Format(Date, "yyyy") & "\"
    
    Worksheets("Gaps").Select
    
    For i = 0 To 10
        sName = "3615 " & Format(Date - i, "yyyy-mm-dd") & ".xlsx"
        If FileExists(sPath & sName) Then
            fileFound = True
            Exit For
        End If
    Next

    If fileFound = True Then
        Workbooks.Open (sPath & sName)
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Gaps").Range("A1")
        ActiveWorkbook.Close
    Else
        Err.Raise Number:=53, _
                  Description:="Gaps file not found." _
                               & vbCrLf & vbCrLf & _
                               "Please make sure you are connected to the network " & _
                               "and that gaps has been run in the past 10 days"
    End If
    Columns("D:D").ClearContents
    Range("D1").Value = "SIM"
    Range("D2").Formula = "=B2&C2"
    Range("D2").AutoFill Destination:=Range(Cells(2, 4), Cells(ActiveSheet.UsedRange.Rows.Count, 4))
    Range("D:D").Value = Range("D:D").Value
    Range("D:D").EntireColumn.AutoFit
End Sub

Sub ImportMaster()
    Dim sPath As String
    Dim Wkbk As Workbook

    sPath = "\\br3615gaps\gaps\Billy Mac-Master Lists\Carrier Master List " & Format(Date, "yyyy") & ".xls"
    ThisWorkbook.Sheets("Master").Cells.Delete
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Workbooks.Open FileName:=sPath
    ActiveSheet.AutoFilterMode = False
    Set Wkbk = ActiveWorkbook
    ActiveSheet.UsedRange.Copy
    ThisWorkbook.Activate
    Sheets("Master").Range("A1").PasteSpecial Paste:=xlPasteValues, _
                                              Operation:=xlNone, _
                                              SkipBlanks:=False, _
                                              Transpose:=False
    Application.CutCopyMode = False
    Wkbk.Close
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Sheets("Macro").Select
End Sub

