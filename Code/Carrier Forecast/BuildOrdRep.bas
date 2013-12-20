Attribute VB_Name = "BuildOrdRep"
Option Explicit

Sub CreateOrderReport()
    Dim i As Long
    Dim iDest As Long
    Dim iRows As Long
    Dim iCols As Long
    Dim TotalRows As Long

    iDest = 16    'Column offset for vlookups
    iRows = Worksheets("Combined").UsedRange.Rows.Count
    iCols = Worksheets("Combined").UsedRange.Columns.Count

    Application.ScreenUpdating = False
    Worksheets("Combined").Range("A:A").SpecialCells(xlCellTypeConstants).Copy _
            Destination:=Worksheets("Forecast").Range("A1")

    Worksheets("Forecast").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    Range("B1").Value = "SIM"
    Range("B2:B" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(A2,Master!A:B,2,FALSE)=0,""-"",VLOOKUP(A2,Master!A:B,2,FALSE)),""-"")"
    Range("B2:B" & TotalRows).Value = Range("B2:B" & TotalRows).Value

    Range("C1").Value = "Description"
    Range("C2:C" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(A2,Master!A:D,4,FALSE)=0,"""",VLOOKUP(A2,Master!A:D,4,FALSE)),"""")"
    Range("C2:C" & TotalRows).Value = Range("C2:C" & TotalRows).Value

    Range("D1").Value = "On Hand"
    Range("D2:D" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!D:F,3,FALSE),""-"")"
    Range("D2:D" & TotalRows).Value = Range("D2:D" & TotalRows).Value
    
    Range("E1").Value = "Reserve"
    Range("E2:E" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!D:G,4,FALSE),""-"")"
    Range("E2:E" & TotalRows).Value = Range("E2:E" & TotalRows).Value

    Range("F1").Value = "OO"
    Range("F2:F" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!D:I,6,FALSE),""-"")"
    Range("F2:F" & TotalRows).Value = Range("F2:F" & TotalRows).Value

    Range("G1").Value = "BO"
    Range("G2:G" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!D:H,5,FALSE),""-"")"
    Range("G2:G" & TotalRows).Value = Range("G2:G" & TotalRows).Value

    Range("H1").Value = "WDC"
    Range("H2").Formula = "=IFERROR(VLOOKUP(B2,Gaps!D:AJ,33,FALSE),""-"")"
    Range("H2").AutoFill Destination:=Range(Cells(2, 8), Cells(ActiveSheet.UsedRange.Rows.Count, 8))

    Range("I1").Value = "Last Cost"
    Range("I2").Formula = "=IFERROR(VLOOKUP(B2,Gaps!D:AE,28,FALSE),""-"")"
    Range("I2").AutoFill Destination:=Range(Cells(2, 9), Cells(ActiveSheet.UsedRange.Rows.Count, 9))

    Range("J1").Value = "UOM"
    Range("J2").Formula = "=IFERROR(VLOOKUP(B2,Gaps!D:AI,32,FALSE),""-"")"
    Range("J2").AutoFill Destination:=Range(Cells(2, 10), Cells(ActiveSheet.UsedRange.Rows.Count, 10))

    Range("K1").Value = "Min/Mult"
    Range("K2").Formula = "=IFERROR(VLOOKUP(A2,Master!A:N,14,FALSE),""-"")"
    Range("K2").AutoFill Destination:=Range(Cells(2, 11), Cells(ActiveSheet.UsedRange.Rows.Count, 11))

    Range("L1").Value = "Supplier"
    Range("L2").Formula = "=IFERROR(VLOOKUP(B2,Gaps!D:AL,35,FALSE),""-"")"
    Range("L2").AutoFill Destination:=Range(Cells(2, 12), Cells(ActiveSheet.UsedRange.Rows.Count, 12))

    Range("M1").Value = "LT/Days"
    Range("M2").Formula = "=IFERROR(VLOOKUP(A2,Master!A:O,15,FALSE),""-"")"
    Range("M2").AutoFill Destination:=Range(Cells(2, 13), Cells(ActiveSheet.UsedRange.Rows.Count, 13))

    Range("N1").Value = "LT/Weeks"
    Range("N2").Formula = "=M2/7"
    Range("N2").AutoFill Destination:=Range(Cells(2, 14), Cells(ActiveSheet.UsedRange.Rows.Count, 14))

    Range("O1").Value = "Stock Visualization"

    For i = 2 To iCols
        Cells(1, iDest) = Worksheets("Combined").Cells(1, i)
        Cells(1, iDest).NumberFormat = "mm/dd"
        If iDest = 16 Then
            Cells(2, iDest).Formula = "=IFERROR(D2-VLOOKUP(A2,Combined!" & _
                                      Range(Cells(1, 1), Cells(iRows, iCols)).Address(False, False) & ",2,FALSE),0)"
            Cells(2, iDest).AutoFill Destination:=Range(Cells(2, iDest), Cells(ActiveSheet.UsedRange.Rows.Count, iDest))
        Else
            Cells(2, iDest).Formula = "=IFERROR(" & _
                                      Cells(2, iDest - 1).Address(False, False) & _
                                      "-VLOOKUP(A2,Combined!" & _
                                      Range(Cells(1, 1), Cells(iRows, iCols)).Address(False, False) & _
                                      "," & i & ",FALSE),0)"
            Cells(2, iDest).AutoFill _
                    Destination:=Range(Cells(2, iDest), Cells(ActiveSheet.UsedRange.Rows.Count, iDest))
        End If
        iDest = iDest + 1    'iDest starts at 15 so that the columns line up properly for vlookups
    Next

    Cells(1, ActiveSheet.UsedRange.Columns.Count + 1).Value = "Notes"
    Cells(2, ActiveSheet.UsedRange.Columns.Count).Formula = _
    "=IFERROR(IF(VLOOKUP(A2,Master!A:E,5,FALSE)=0,"""",VLOOKUP(A2,Master!A:E,5,FALSE)),"""")"
    Cells(2, ActiveSheet.UsedRange.Columns.Count).AutoFill _
            Destination:=Range( _
                         Cells(2, ActiveSheet.UsedRange.Columns.Count), _
                         Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count))
    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

    Range("O2").Select
    Range("O2").SparklineGroups.Add _
            Type:=xlSparkColumn, _
            SourceData:=Range(Cells(2, 16), Cells(2, ActiveSheet.UsedRange.Columns.Count - 1)) _
                        .Address(False, False)

    Selection.SparklineGroups.Item(1).Points.Negative.Visible = True
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 3289650
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0

    With Range("O:O")
        Range("O2").AutoFill Destination:=Range(Cells(2, 15), Cells(.CurrentRegion.Rows.Count, 15))
    End With

    Range(Cells(2, 15), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count - 1)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority

    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Range("O1").Select
    With Range("A:A")
        ActiveSheet.ListObjects.Add( _
                xlSrcRange, Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)), , _
                xlYes).Name = "Table1"
    End With
    Range(Cells(2, 4), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count)).HorizontalAlignment = xlCenter
    Range(Cells(2, 2), Cells(ActiveSheet.UsedRange.Rows.Count, 2)).HorizontalAlignment = xlCenter
    ActiveSheet.UsedRange.Columns.AutoFit
    
    Columns("O:O").ColumnWidth = 22.29
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Proc : AddNotes
' Date : 1/17/2013
' Desc : Add previous weeks expedite notes to the forecast
'---------------------------------------------------------------------------------------
Sub AddNotes()
    Dim sPath As String
    Dim sWkBk As String
    Dim sYear As String
    Dim iRows As Long
    Dim iCols As Integer
    Dim i As Integer

    Sheets("Temp").Cells.Delete

    For i = 1 To 30
        sYear = Date - i
        sWkBk = "Slink Alert " & Format(sYear, "m-dd-yy") & ".xlsx"
        sPath = "\\br3615gaps\gaps\Carrier\" & Format(sYear, "yyyy") & " Alerts\"

        If FileExists(sPath & sWkBk) = True Then
            Workbooks.Open sPath & sWkBk

            Sheets("Expedite").Select
            iRows = ActiveSheet.UsedRange.Rows.Count
            iCols = ActiveSheet.UsedRange.Columns.Count

            Range(Cells(1, 1), Cells(iRows, 1)).Copy Destination:=ThisWorkbook.Sheets("Temp").Range("A1")
            Range(Cells(1, iCols), Cells(iRows, iCols)).Copy Destination:=ThisWorkbook.Sheets("Temp").Range("B1")
            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = True

            Sheets("Forecast").Select
            iRows = ActiveSheet.UsedRange.Rows.Count
            iCols = ActiveSheet.UsedRange.Columns.Count + 1

            Cells(1, iCols).Value = "Expedite Notes"
            Cells(2, iCols).Formula = "=IFERROR(IF(VLOOKUP(A2,Temp!A:B,2,FALSE)=0,"""",VLOOKUP(A2,Temp!A:B,2,FALSE)),"""")"
            Cells(2, iCols).AutoFill Destination:=Range(Cells(2, iCols), Cells(iRows, iCols))
            Range(Cells(2, iCols), Cells(iRows, iCols)).Value = Range(Cells(2, iCols), Cells(iRows, iCols)).Value
            Columns(iCols).EntireColumn.AutoFit
            Exit For
        End If
    Next
End Sub


