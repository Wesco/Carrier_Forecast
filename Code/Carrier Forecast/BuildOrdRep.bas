Attribute VB_Name = "BuildOrdRep"
Option Explicit

Sub CreateOrderReport()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim DestCol As Integer
    Dim CombinedCols As Long
    Dim i As Long

    Sheets("Combined").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Range("A1:A" & TotalRows).Copy Destination:=Sheets("Forecast").Range("A1")

    Sheets("Forecast").Select

    Range("B1").Value = "SIM"
    Range("B2:B" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(A2,Master!A:B,2,FALSE)=0,"""",VLOOKUP(A2,Master!A:B,2,FALSE)),"""")"
    Range("B2:B" & TotalRows).Value = Range("B2:B" & TotalRows).Value

    Range("C1").Value = "Description"
    Range("C2:C" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(A2,Master!A:D,4,FALSE)=0,"""",VLOOKUP(A2,Master!A:D,4,FALSE)),"""")"
    Range("C2:C" & TotalRows).Value = Range("C2:C" & TotalRows).Value

    Range("D1").Value = "On Hand"
    Range("D2:D" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!A:G,7,FALSE),0)"
    Range("D2:D" & TotalRows).Value = Range("D2:D" & TotalRows).Value

    Range("E1").Value = "Reserve"
    Range("E2:E" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!A:H,8,FALSE),0)"
    Range("E2:E" & TotalRows).Value = Range("E2:E" & TotalRows).Value

    Range("F1").Value = "OO"
    Range("F2:F" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!A:J,10,FALSE),0)"
    Range("F2:F" & TotalRows).Value = Range("F2:F" & TotalRows).Value

    Range("G1").Value = "BO"
    Range("G2:G" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!A:I,9,FALSE),0)"
    Range("G2:G" & TotalRows).Value = Range("G2:G" & TotalRows).Value

    Range("H1").Value = "WDC"
    Range("H2:H" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!A:AK,37,FALSE),0)"
    Range("H2:H" & TotalRows).Value = Range("H2:H" & TotalRows).Value

    Range("I1").Value = "Last Cost"
    Range("I2:I" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!A:AF,32,FALSE),0)"
    Range("I2:I" & TotalRows).Value = Range("I2:I" & TotalRows).Value

    Range("J1").Value = "UOM"
    Range("J2:J" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!A:AJ,36,FALSE),"""")"
    Range("J2:J" & TotalRows).Value = Range("J2:J" & TotalRows).Value

    Range("K1").Value = "Min/Mult"
    Range("K2:K" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2,Master!A:N,14,FALSE),"""")"
    Range("K2:K" & TotalRows).Value = Range("K2:K" & TotalRows).Value

    Range("L1").Value = "Supplier"
    Range("L2:L" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!A:AM,39,FALSE),"""")"
    Range("L2:L" & TotalRows).Value = Range("L2:L" & TotalRows).Value

    Range("M1").Value = "LT/Days"
    Range("M2:M" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2,Master!A:O,15,FALSE),0)"
    Range("M2:M" & TotalRows).Value = Range("M2:M" & TotalRows).Value

    Range("N1").Value = "LT/Weeks"
    Range("N2:N" & TotalRows).Formula = "=IFERROR(ROUNDUP(M2/7,0),0)"
    Range("N2:N" & TotalRows).Value = Range("N2:N" & TotalRows).Value

    Range("O1").Value = "Stock Visualization"

    'Add forecast month data
    Range("P1:AL1").Formula = "=Combined!B1"
    Range("P1:AL1").Value = Range("P1:AL1").Value
    Range("P1:AL1").NumberFormat = "mm/dd"
    Range("P2:P" & TotalRows).Formula = "=D2-VLOOKUP(A2,Combined!A:B,2,FALSE)"
    Range("P2:P" & TotalRows).Value = Range("P2:P" & TotalRows).Value

    CombinedCols = Sheets("Combined").UsedRange.Columns.Count

    For i = 3 To 24
        'Set column data
        Range(Cells(2, i + 14), Cells(TotalRows, i + 14)).Formula = "=IFERROR(" & Cells(2, i + 14 - 1).Address(False, False) & "-VLOOKUP(A2,Combined!A:X," & i & ",FALSE),0)"
        Range(Cells(2, i + 14), Cells(TotalRows, i + 14)).Value = Range(Cells(2, i + 14), Cells(TotalRows, i + 14)).Value
    Next

    TotalCols = Columns(Columns.Count).End(xlToLeft).Column + 1
    Cells(1, TotalCols).Value = "Notes"
    Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols)).Formula = "=IFERROR(IF(VLOOKUP(A2,Master!A:E,5,FALSE)=0,"""",VLOOKUP(A2,Master!A:E,5,FALSE)),"""")"
    Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols)).Value = Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols)).Value

    Range("O2:O" & TotalRows).SparklineGroups.Add Type:=xlSparkColumn, SourceData:=Range(Cells(2, 16), Cells(TotalRows, TotalCols - 1)).Address(False, False)
    With Range("O2:O" & TotalRows).SparklineGroups.Item(1)
        .Points.Negative.Visible = True
        .SeriesColor.Color = 3289650
        .SeriesColor.TintAndShade = 0
        .Points.Negative.Color.Color = 208
        .Points.Negative.Color.TintAndShade = 0
        .Points.Markers.Color.Color = 208
        .Points.Markers.Color.TintAndShade = 0
        .Points.Highpoint.Color.Color = 208
        .Points.Highpoint.Color.TintAndShade = 0
        .Points.Lowpoint.Color.Color = 208
        .Points.Lowpoint.Color.TintAndShade = 0
        .Points.Firstpoint.Color.Color = 208
        .Points.Firstpoint.Color.TintAndShade = 0
        .Points.Lastpoint.Color.Color = 208
        .Points.Lastpoint.Color.TintAndShade = 0
    End With

    Range("P2:AL" & TotalRows).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Range("P2:AL" & TotalRows).FormatConditions(Range("P2:AL" & TotalRows).FormatConditions.Count).SetFirstPriority

    With Range("P2:AL" & TotalRows).FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Range("P2:AL" & TotalRows).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Range("P2:AL" & TotalRows).FormatConditions(1).StopIfTrue = False

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


    For i = 1 To 30
        sYear = Date - i
        sWkBk = "Slink Alert " & Format(sYear, "m-dd-yy") & ".xlsx"
        sPath = "\\br3615gaps\gaps\Carrier\" & Format(sYear, "yyyy") & " Alerts\"

        If FileExists(sPath & sWkBk) = True Then
            Workbooks.Open sPath & sWkBk

            Sheets("Expedite").Select
            iRows = ActiveSheet.UsedRange.Rows.Count
            iCols = ActiveSheet.UsedRange.Columns.Count

            Range(Cells(1, 1), Cells(iRows, 1)).Copy Destination:=ThisWorkbook.Sheets("Expedite").Range("A1")
            Range(Cells(1, iCols), Cells(iRows, iCols)).Copy Destination:=ThisWorkbook.Sheets("Expedite").Range("B1")
            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = True

            Sheets("Forecast").Select
            iRows = ActiveSheet.UsedRange.Rows.Count
            iCols = ActiveSheet.UsedRange.Columns.Count + 1

            Cells(1, iCols).Value = "Expedite Notes"
            Cells(2, iCols).Formula = "=IFERROR(IF(VLOOKUP(A2,Expedite!A:B,2,FALSE)=0,"""",VLOOKUP(A2,Expedite!A:B,2,FALSE)),"""")"
            Cells(2, iCols).AutoFill Destination:=Range(Cells(2, iCols), Cells(iRows, iCols))
            Range(Cells(2, iCols), Cells(iRows, iCols)).Value = Range(Cells(2, iCols), Cells(iRows, iCols)).Value
            Columns(iCols).EntireColumn.AutoFit
            Exit For
        End If
    Next
End Sub
