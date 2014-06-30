Attribute VB_Name = "Integrate"
Option Explicit

Sub CombineForecasts()
    Dim rTemp As Range
    Dim vCell As Integer
    Dim i As Integer
    Dim demandRows As Long
    Dim weeklyRows As Long
    Dim weeklyCols As Long

    demandRows = Worksheets("Demand").UsedRange.Rows.Count
    weeklyRows = Worksheets("Weekly").UsedRange.Rows.Count

    Worksheets("Weekly").Select
    Do While CInt(Format(Range("B1").Value, "m")) < CInt(Format(Date, "m"))
        Columns(2).Delete Shift:=xlToLeft
    Loop

    weeklyCols = ActiveSheet.UsedRange.Columns.Count
    vCell = CInt(Format(Cells(1, weeklyCols).Value, "mmdd"))
    Set rTemp = Range(Cells(1, 2), Cells(1, weeklyCols))

    Worksheets("Demand").Select

    Do While vCell <> Format(Range("B1").Text, "mdd")
        Columns(2).Delete Shift:=xlToLeft
    Loop

    If vCell = Format(Range("B1").Text, "mdd") Then
        Columns(2).Delete Shift:=xlToLeft
    End If

    Worksheets("Demand").Range(Cells(1, 2), Cells(1, weeklyCols)).EntireColumn.Insert Shift:=xlToRight
    For i = 1 To weeklyCols - 1
        Worksheets("Demand").Cells(1, i + 1) = rTemp(i)
    Next

    Worksheets("Demand").Range(Cells(1, 2), Cells(1, i)).NumberFormat = "mm/dd"
    Worksheets("Weekly").Select
    Range(Cells(2, 1), Cells(weeklyRows, weeklyCols)).Copy Destination:=Worksheets("Demand").Cells(demandRows + 1, 1)
    Worksheets("Demand").Select
    ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks).Value = 0

End Sub

Sub MergeParts()
    Dim i As Integer
    Dim rPivDest As Range
    Dim rPivSource As Range
    Const sPivName As String = "PivotTable1"

    Worksheets("Combined").Select
    Set rPivDest = Worksheets("Combined").Range("A1")

    Worksheets("Demand").Select
    Set rPivSource = Worksheets("Demand").UsedRange

    ActiveWorkbook.PivotCaches.Create(xlDatabase, rPivSource).CreatePivotTable rPivDest, sPivName
    Worksheets("Combined").Select

    With Worksheets("Combined").PivotTables(sPivName)
        .PivotFields(rPivSource(1, 1).Text).Orientation = xlRowField
        .PivotFields(rPivSource(1, 1).Text).Position = 1
        For i = 2 To rPivSource.Columns.Count
            .AddDataField .PivotFields(rPivSource(1, i).Text), "Sum of " & rPivSource(1, i).Text, xlSum
        Next
    End With

    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Rows("1:1").Delete Shift:=xlUp
    Range("A1").Value = "Part Number"
    Rows(ActiveSheet.UsedRange.Rows.Count).Delete Shift:=xlUp
    For i = 1 To ActiveSheet.Columns.Count - 1
        Cells(1, i).Value = Replace(Cells(1, i).Text, "Sum of ", "")
    Next
    Range(Cells(1, 2), Cells(1, ActiveSheet.Columns.Count)).NumberFormat = "mm/dd"
End Sub
