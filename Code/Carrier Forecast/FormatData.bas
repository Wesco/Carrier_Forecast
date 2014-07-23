Attribute VB_Name = "FormatData"
Option Explicit

Sub FormatFcst(Sheet As Worksheet)
    Dim TotalCols As Integer

    Sheet.Select

    'Unmerge column headers
    Rows("9:10").UnMerge

    'Delete report header
    Rows("1:8").Delete Shift:=xlShiftUp

    'Fix column headers
    Range("A2:C2").Value = Range("A1:C1").Value
    Rows(1).Delete Shift:=xlShiftUp

    'Remove unused columns
    Range("B:I").Delete
    TotalCols = ActiveSheet.UsedRange.Rows.Count

    'Empty character = alt + 255
    Range(Cells(1, 2), Cells(1, TotalCols)).Replace What:=" ", Replacement:=""
    Range(Cells(1, 2), Cells(1, TotalCols)).NumberFormat = "mm/dd"

    Columns.AutoFit
    Rows.AutoFit
    Range("A1").Select

    Application.DisplayAlerts = True
End Sub
