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
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Set column headers
    Range("B1:O1").Value = Array("SIM", _
                                 "Description", _
                                 "On Hand", _
                                 "Reserve", _
                                 "OO", _
                                 "BO", _
                                 "WDC", _
                                 "Last Cost", _
                                 "UOM", _
                                 "Min/Mult", _
                                 "Supplier", _
                                 "LT/Days", _
                                 "LT/Weeks", _
                                 "Stock Visualization")

    'Add column data
    Range("B2:N" & TotalRows).Formula = Array("=IFERROR(IF(VLOOKUP(A2,Master!A:B,2,FALSE)=0,"""",VLOOKUP(A2,Master!A:B,2,FALSE)),"""")", _
                                              "=IFERROR(IF(VLOOKUP(A2,Master!A:D,4,FALSE)=0,"""",VLOOKUP(A2,Master!A:D,4,FALSE)),"""")", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:G,7,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:H,8,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:J,10,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:I,9,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:AK,37,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:AF,32,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:AJ,36,FALSE),"""")", _
                                              "=IFERROR(VLOOKUP(A2,Master!A:N,14,FALSE),"""")", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:AM,39,FALSE),"""")", _
                                              "=IFERROR(VLOOKUP(A2,Master!A:O,15,FALSE),0)", _
                                              "=IFERROR(ROUNDUP(M2/7,0),0)")

    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

    'Add forecast month data
    Range("P1:AL1").Formula = "=Combined!B1"
    Range("P1:AL1").Value = Range("P1:AL1").Value
    Range("P1:AL1").NumberFormat = "mm/dd"
    Range("P2:P" & TotalRows).Formula = "=D2-VLOOKUP(A2,Combined!A:B,2,FALSE)"
    Range("P2:P" & TotalRows).Value = Range("P2:P" & TotalRows).Value

    Sheets("Combined").Select
    CombinedCols = ActiveSheet.UsedRange.Columns.Count

    Sheets("Forecast").Select
    'Set column data
    For i = 3 To 24
        Range(Cells(2, i + 14), Cells(TotalRows, i + 14)).Formula = "=IFERROR(" & Cells(2, i + 14 - 1).Address(False, False) & "-VLOOKUP(A2,Combined!A:X," & i & ",FALSE),0)"
        Range(Cells(2, i + 14), Cells(TotalRows, i + 14)).Value = Range(Cells(2, i + 14), Cells(TotalRows, i + 14)).Value
    Next

    'Add notes from master
    Range("AM1").Value = "Notes"
    Range("AM2:AM" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(A2,Master!A:E,5,FALSE)=0,"""",VLOOKUP(A2,Master!A:E,5,FALSE)),"""")"
    Range("AM2:AM" & TotalRows).Value = Range("AM2:AM" & TotalRows).Value

    'Add sparklines
    With Range("O2:O" & TotalRows).SparklineGroups
        .Add Type:=xlSparkColumn, SourceData:=Range("P2:AL" & TotalRows).Address(False, False)
        With .Item(1)
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
    End With

    'Add conditional formatting
    Range("P2:AL" & TotalRows).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    With Range("P2:AL" & TotalRows).FormatConditions(1)
        .Font.Color = -16383844
        .Font.TintAndShade = 0
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 13551615
        .Interior.TintAndShade = 0
        .StopIfTrue = False
    End With

    'Add alternating line colors
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:AM" & TotalRows), , xlYes).Name = "Table1"
    Sheet1.ListObjects(1).Unlist

    'Fix column alignment
    Range(Cells(2, 4), Cells(TotalRows, TotalCols)).HorizontalAlignment = xlCenter
    Range("B2:B" & TotalRows).HorizontalAlignment = xlCenter

    'Fix column width
    ActiveSheet.UsedRange.Columns.AutoFit
    Columns("O:O").ColumnWidth = 22.29
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
