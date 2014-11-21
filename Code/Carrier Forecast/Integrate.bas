Attribute VB_Name = "Integrate"
Option Explicit

Sub CombineForecasts()
    Dim i As Integer
    Dim DemandRows As Long
    Dim DemandCols As Integer
    Dim ColHeaders As Variant
    Dim WeeklyEndDt As Date
    Dim WeeklyRows As Long
    Dim WeeklyCols As Integer


    Sheets("Weekly").Select

    'Remove dates from the weekly forecast if they have already passed
    Do While CDate(Range("B1").Value) < Date
        Columns(2).Delete Shift:=xlToLeft
    Loop

    WeeklyRows = ActiveSheet.UsedRange.Rows.Count
    WeeklyCols = ActiveSheet.UsedRange.Columns.Count
    WeeklyEndDt = Cells(1, WeeklyCols).Value
    ColHeaders = Range(Cells(1, 2), Cells(1, WeeklyCols)).Value

    Sheets("Demand").Select
    DemandCols = ActiveSheet.UsedRange.Columns.Count
    DemandRows = ActiveSheet.UsedRange.Rows.Count

    'Remove intersection of weekly / demand columns
    For i = 2 To DemandCols
        If CDate(Range("B1").Value) <= WeeklyEndDt Then
            Columns(2).Delete
        Else
            Exit For
        End If
    Next

    'Insert columns for weekly data
    Range(Cells(1, 2), Cells(1, WeeklyCols)).EntireColumn.Insert
    Range(Cells(1, 2), Cells(1, WeeklyCols)).Value = ColHeaders
    Range(Cells(1, 2), Cells(1, WeeklyCols)).NumberFormat = "mm/dd/yyyy"
    
    'Copy weekly data below demand data
    Sheets("Weekly").Select
    Range(Cells(2, 1), Cells(WeeklyRows, WeeklyCols)).Copy Destination:=Sheets("Demand").Range("A" & DemandRows + 1)
    
    'Set all blanks to 0's
    Sheets("Demand").Select
    ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks).Value = 0
End Sub

Sub MergeParts()
    Dim i As Integer
    Dim PivDest As Range
    Dim PivSource As Range
    Dim PivName As String

    PivName = "PivotTable1"
    Set PivDest = Sheets("Combined").Range("A1")
    Set PivSource = Sheets("Demand").UsedRange

    ActiveWorkbook.PivotCaches.Create(xlDatabase, PivSource).CreatePivotTable PivDest, PivName
    
    Sheets("Combined").Select
    With ActiveSheet.PivotTables(PivName)
        .PivotFields(PivSource(1, 1).Text).Orientation = xlRowField
        .PivotFields(PivSource(1, 1).Text).Position = 1
        For i = 2 To PivSource.Columns.Count
            .AddDataField .PivotFields(PivSource(1, i).Text), "Sum of " & PivSource(1, i).Text, xlSum
        Next
    End With

    Cells.Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Rows("1:1").Delete Shift:=xlUp
    Range("A1").Value = "Part Number"
    Rows(ActiveSheet.UsedRange.Rows.Count).Delete Shift:=xlUp
    For i = 1 To ActiveSheet.Columns.Count - 1
        Cells(1, i).Value = Replace(Cells(1, i).Text, "Sum of ", "")
    Next
    Range(Cells(1, 2), Cells(1, ActiveSheet.Columns.Count)).NumberFormat = "mm/dd/yyyy"
End Sub
