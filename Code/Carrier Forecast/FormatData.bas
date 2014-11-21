Attribute VB_Name = "FormatData"
Option Explicit

Sub FormatFcst(Report As Fcst)
    Dim TotalCols As Integer
    Dim Yr As Integer
    Dim i As Long

    If Report = Demand Then
        Sheets("Demand").Select
    Else
        Sheets("Weekly").Select
    End If

    'Unmerge column headers
    Rows("9:10").UnMerge

    'Delete report header
    Rows("1:8").Delete Shift:=xlShiftUp

    'Fix column headers
    Range("A2:C2").Value = Range("A1:C1").Value
    Rows(1).Delete Shift:=xlShiftUp
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Remove unused columns
    If Report = Demand Then
        Range("B:I").Delete
    Else
        Range("B:F").Delete
        Range(Cells(1, 10), Cells(Rows.Count, TotalCols)).Columns.Delete
    End If

    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Empty character = alt + 255
    Range(Cells(1, 2), Cells(1, TotalCols)).Replace What:=" ", Replacement:=""

    'Add the year to column headers
    If Format(Date, "mmm") = "Jan" Then
        Yr = Year(Date) - 1
    Else
        Yr = Year(Date)
    End If

    For i = 2 To TotalCols
        If Format(Cells(1, i).Value, "mmm") = "Jan" And Format(Cells(1, i - 1).Value, "mmm") <> "Jan" Then
            Yr = Yr + 1
        End If

        Cells(1, i).Value = Format(Cells(1, i).Value, "mm/dd") & "/" & Yr
    Next

    Range(Cells(1, 2), Cells(1, TotalCols)).NumberFormat = "mm/dd/yyyy"
    Columns.AutoFit
    Rows.AutoFit
    Range("A1").Select
End Sub
