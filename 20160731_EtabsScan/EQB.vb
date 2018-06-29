Sub AssembledSub()

    Sheets("Assembled Point Masses").Select
    Lastrow = Sheets("Assembled Point Masses").UsedRange.Rows.Count '------------抓取最後一行

    ' Cells.Select
    ' With ActiveWorkbook.Worksheets("Assembled Point Masses").Sort
    '     .SortFields.Clear
    '     .SortFields.Add Key:=Range(Cells(2, 2), Cells(Lastrow, 2)), _
    '      SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    '     .SetRange Range(Cells(1, 1), Cells(Lastrow, 11))
    '     .Header = xlYes
    '     .MatchCase = False
    '     .Orientation = xlTopToBottom
    '     .SortMethod = xlPinYin
    '     .Apply
    ' End With


    ' If ActiveSheet.FilterMode Then
    '     ActiveSheet.ShowAllData
    '     Cells.Select
    '     ActiveSheet.Range("$A$1:$K$198").AutoFilter Field:=2, Criteria1:="All"

    ' Else
    Cells.Select
    Selection.AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(Lastrow, 11)).AutoFilter Field:=2, Criteria1:="All"

    ' End If

End Sub






