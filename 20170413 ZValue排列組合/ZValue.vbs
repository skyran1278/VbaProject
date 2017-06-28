Function ()

    Worksheets("Z").Activate

    ' 排序
    Worksheets("Z").Range(Cells(7, 3), Cells(zRowUsed, 10)).Sort _
        Key1:=Range(Cells(8, 10), Cells(zRowUsed, 10)), Order1:=xlAscending, _
        Key2:=Range(Cells(8, 8), Cells(zRowUsed, 8)), Order2:=xlDescending, Header:=xlYes

    rowStart = 1
    columnStart = 8
    rowEnd = Cells(Rows.Count, 8).End(xlUp).Row
    columnEnd = 8

    zValue = Range(Cells(rowStart, columnStart), Cells(rowEnd, columnEnd)).Value

    ' rowStart = 1
    columnStart = 10
    ' rowEnd = Cells(Rows.Count, 8).End(xlUp).Row
    columnEnd = 10

    group = Range(Cells(rowStart, columnStart), Cells(rowEnd, columnEnd)).Value

    For i = 8 To rowEnd
        If group(i) <> group(i + 1) Then

        End If
    Next

End Function

Function Controller()

End Function