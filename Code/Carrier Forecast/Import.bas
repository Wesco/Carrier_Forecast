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

'---------------------------------------------------------------------------------------
' Proc  : Sub ImportGaps
' Date  : 12/12/2012
' Desc  : Imports gaps to the workbook containing this macro.
' Ex    : ImportGaps
'---------------------------------------------------------------------------------------
Sub ImportGaps(Optional Destination As Range, Optional SimsAsText As Boolean = True)
    Dim Path As String      'Gaps file path
    Dim Name As String      'Gaps Sheet Name
    Dim i As Long           'Counter to decrement the date
    Dim dt As Date          'Date for gaps file name and path
    Dim TotalRows As Long   'Total number of rows
    Dim Result As VbMsgBoxResult    'Yes/No to proceed with old gaps file if current one isn't found


    'This error is bypassed so you can determine whether or not the sheet exists
    On Error GoTo CREATE_GAPS
    If TypeName(Destination) = "Nothing" Then
        Set Destination = ThisWorkbook.Sheets("Gaps").Range("A1")
    End If
    On Error GoTo 0

    Application.DisplayAlerts = False

    'Try to find Gaps
    For i = 0 To 15
        dt = Date - i
        Path = "\\br3615gaps\gaps\3615 Gaps Download\" & Format(dt, "yyyy") & "\"
        Name = "3615 " & Format(dt, "yyyy-mm-dd") & ".csv"
        If Exists(Path & Name) Then
            Exit For
        End If
    Next

    'Make sure Gaps file was found
    If Exists(Path & Name) Then
        If dt <> Date Then
            Result = MsgBox( _
                     Prompt:="Gaps from " & Format(dt, "mmm dd, yyyy") & " was found." & vbCrLf & "Would you like to continue?", _
                     Buttons:=vbYesNo, _
                     Title:="Gaps not up to date")
        End If

        If Result <> vbNo Then
            ThisWorkbook.Activate
            Sheets(Destination.Parent.Name).Select

            'If there is data on the destination sheet delete it
            If Range("A1").Value <> "" Then
                Cells.Delete
            End If

            ImportCsvAsText Path, Name, Sheets("Gaps").Range("A1")
            TotalRows = ActiveSheet.UsedRange.Rows.Count
            Range("D1:D" & TotalRows).ClearContents
            Range("D1").Value = "SIM"

            'SIMs are 11 digits and can have leading 0's
            If SimsAsText = True Then
                Range("D2:D" & TotalRows).Formula = "=""=""&""""""""&RIGHT(""000000"" & B2, 6)&RIGHT(""00000"" & C2, 5)&"""""""""
            Else
                Range("D2:D" & TotalRows).Formula = "=B2&RIGHT(""00000"" & C2, 5)"
            End If

            Range("D2:D" & TotalRows).Value = Range("D2:D" & TotalRows).Value
        Else
            Err.Raise 18, "ImportGaps", "Import canceled"
        End If
    Else
        Err.Raise 53, "ImportGaps", "Gaps could not be found."
    End If

    Application.DisplayAlerts = True
    Exit Sub

CREATE_GAPS:
    ThisWorkbook.Sheets.Add After:=Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = "Gaps"
    Resume

End Sub

Sub ImportMaster()
    Dim sPath As String
    Dim Wkbk As Workbook

    sPath = "\\br3615gaps\gaps\Billy Mac-Master Lists\Carrier Master List " & Format(Date, "yyyy") & ".xls"
    ThisWorkbook.Sheets("Master").Cells.Delete
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Workbooks.Open FileName:=sPath
    On Error Resume Next
    Sheets("ACTIVE").Select
    On Error GoTo 0
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

'---------------------------------------------------------------------------------------
' Proc  : Function Exists
' Date  : 6/24/14
' Type  : Boolean
' Desc  : Checks if a file exists and can be read
' Ex    : FileExists "C:\autoexec.bat"
'---------------------------------------------------------------------------------------
Private Function Exists(ByVal FilePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Remove trailing backslash
    If InStr(Len(FilePath), FilePath, "\") > 0 Then
        FilePath = Left(FilePath, Len(FilePath) - 1)
    End If

    'Check to see if the file exists and has read access
    On Error GoTo File_Error
    If fso.FileExists(FilePath) Then
        fso.OpenTextFile(FilePath, 1).Read 0
        Exists = True
    Else
        Exists = False
    End If
    On Error GoTo 0

    Exit Function

File_Error:
    Exists = False
End Function
