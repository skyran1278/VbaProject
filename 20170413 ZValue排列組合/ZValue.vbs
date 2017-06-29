' DATA 資料命名
Const STORY = 4
Const LABEL = 5
Const MAX_M = 6
Const FY = 7
Const Z = 8
Const LENGTH = 9
Const GROUP = 10
Const SELECT_NUMBER = 11
Const REPLACE_NUMBER = 12
Const DIFF = 13


Function ExecutionTime(Time0)

    If Timer - Time0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - Time0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - Time0) / 60, 2) & " Min", vbOKOnly
    End If

End Function


Sub Controller()

    Time0 = Timer

    Dim zValue, sumArray(), arr()

    Worksheets("Z").Activate

    rowStart = 1
    rowEnd = Cells(Rows.Count, 8).End(xlUp).Row

    ' 排序
    Worksheets("Z").Range(Cells(7, 3), Cells(rowEnd, 10)).Sort _
        Key1:=Range(Cells(8, Z), Cells(rowEnd, Z)), Order1:=xlDescending, Header:=xlYes

    zValue = Range(Cells(rowStart, Z), Cells(rowEnd, Z)).Value

    rowSelect = 5
    columnSelect = 2
    selectNumber = Cells(rowSelect, columnSelect)

    rowStart = 8

    ReDim sumArray(rowStart To rowEnd)
    ReDim arr(selectCount)

    Cells(rowStart, SELECT_NUMBER) = "*"

    selectCount = selectNumber - 1

    Do While selectCount > 0

        For i = rowStart To rowEnd

            If Cells(i, SELECT_NUMBER) = "" Then

                Cells(i, SELECT_NUMBER) = "*"
                sumArray(i) = Application.Sum(Range(Cells(rowStart, REPLACE_NUMBER), Cells(rowEnd, REPLACE_NUMBER)))
                Cells(i, SELECT_NUMBER) = ""

            End If

        Next

        Cells(Application.Match(Application.Min(sumArray), sumArray, 0) + rowStart - 1, SELECT_NUMBER) = "*"

        selectCount = selectCount - 1

    Loop

    Do While selectCount > 0

        For i = rowStart + 1 To rowEnd

            If Cells(i, SELECT_NUMBER) = "*" Then

                Cells(i, SELECT_NUMBER) = ""

                For j = rowStart To rowEnd

                    If Cells(i, SELECT_NUMBER) = "" Then

                        Cells(i, SELECT_NUMBER) = "*"
                        sumArray(i) = Application.Sum(Range(Cells(rowStart, REPLACE_NUMBER), Cells(rowEnd, REPLACE_NUMBER)))
                        Cells(i, SELECT_NUMBER) = ""

                    End If

                Next

                Cells(Application.Match(Application.Min(sumArray), sumArray, 0) + rowStart - 1, SELECT_NUMBER) = "*"

                sumArray(i) = Application.Sum(Range(Cells(rowStart, REPLACE_NUMBER), Cells(rowEnd, REPLACE_NUMBER)))
                Cells(i, SELECT_NUMBER) = ""

            End If

        Next

        Cells(Application.Match(Application.Min(sumArray), sumArray, 0) + rowStart - 1, SELECT_NUMBER) = "*"

        selectCount = selectCount - 1

    Loop

    Call ExecutionTime(Time0)

End Sub


