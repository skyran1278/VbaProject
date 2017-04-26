Sub ZValue()
'
' 排序
' 讀進陣列
' 判斷是不是同一個Group
' 同個Group中取40個或30個或20個做排列組合
' 找出SUM最小的
' 就是答案

    Dim Value()
    Dim GroupValue()
    Dim Combo()

    ' 讀取第8欄之列數
    ZRowUsed = Cells(Rows.Count, 8).End(xlUp).Row

    ReDim Value(ZRowUsed - 9 , 1)


    ' 排序
    Worksheets("Z").Range(Cells(7, 3), Cells(ZRowUsed, 10)).Sort _
        Key1:=Range(Cells(8, 10), Cells(ZRowUsed, 10)), Order1:=xlAscending, _
        Key2:=Range(Cells(8, 8), Cells(ZRowUsed, 8)), Order2:=xlDescending, Header:=xlYes

    ' 取值進陣列
    For ZRowNumber = 8 To ZRowUsed

        'Group
        Value(ValueRowNumber, 0) = Cells(ZRowNumber, 10)

        'ZValue
        Value(ValueRowNumber, 1) = Cells(ZRowNumber, 8)

        '陣列計數
        ValueRowNumber = ValueRowNumber + 1

    Next

    For ValueRowNumber = 0 To ZRowUsed - 9

    	If Value(ValueRowNumber, 0) <> Value(ValueRowNumber + 1, 0) Then

    		EndNumber = ValueRowNumber

    		ZValueNumber = EndNumber - StartNumber

    		If ZValueNumber > Cells(4, 4) Then

    			Interval = FIX(ZValueNumber / Cells(4, 4))

    			For IntervalRowNumber = StartNumber To EndNumber step Interval

	    			GroupValue(GroupValueRowNumber) = Value(IntervalRowNumber, 1)

	    			GroupValueRowNumber = GroupValueRowNumber + 1

    			Next

                Cells(4, 7) = Value(StartNumber, 1)

                For j = 0 To UBound(GroupValue)
                    Combo(0, j) = GroupValue(0)
                Next

                For j = 1 To 4
                    For k = 1 To 4
                        Combo(j, k, 1) = GroupValue(j)
                        Combo(j, k, 2) = GroupValue(k)
                    Next
                Next

                For i = 1 To 5
                    For j = 1 To 4
                        For k = 1 To 3
                            Combo(k, j) = GroupValue(j)
                        Next
                    Next
                Next






    		End If


    	End If

    	StartNumber = ValueRowNumber + 1
    Next





End Sub

Sub kk()
    Dim n As Integer
    Dim r As Integer
    Dim a() As Integer

    n = 5
    r = 3
    ReDim a(r)
    a(1) = 1
    a(2) = 2
    a(3) = 3
    Call Combo(a, 1, 1, n, r)

End Sub

Sub Combo(a() As Integer, digit As Integer, lower As Integer, ByVal n As Integer, ByVal r As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim Temp As String

    For i = lower To n - r + digit
        ' a(digit) = i
        If digit <> r Then
            Call Combo(a, digit + 1, i + 1, n, r)
        Else
            Temp = ""
            For j = 1 To r
                Temp = Temp & Format(a(j), "#")
            Next j
            Debug.Print Temp
        End If
    Next i
End Sub

Sub Combo()

End Sub


Sub Permutations()
	dim LimitValue(3), CountValue(3)
	For i = 0 To 3
		CountValue(i) = 7 - i
		LimitValue(i) = 7 - i

	Next

End Sub

Sub 组合()
    Const n As Integer = 5, m As Integer = 3
    Dim i As Integer, a(1 To m), b(1 To m) As Integer
    For i = 1 To m '生成组合和上界值
        a(i) = i
        b(i) = n - m + i
    Next
    i = m
    Do
        If i = m Then Debug.Print Join(a, " ") '输出组合序列
        If a(i) < b(i) Then '从最后递增1，产生新序列
            a(i) = a(i) + 1
            If i < m Then
                For i = i To m - 1
                    a(i + 1) = a(i) + 1
                Next
            End If
        Else
            i = i - 1
        End If
    Loop Until i = 0
End Sub

Function SelectionSeltor(ComboPMM)

    ' 從第1筆資料Loop到最後一筆
    For RowNumber = 0 To UBound(ComboPMM) - 1

        ' 看看他與下一筆資料相不相同，如果相同就是一組。
        If ComboPMM(RowNumber, 0) <> ComboPMM(RowNumber + 1, 0) Then

            EndNumber = RowNumber

            ' 每一個Column（包含很多個Combo）重新初始化
            FinalSelectionNumber = 0
            FinalRatio = 0

            ' 相同的一組
            For ColumnNumber = StartNumber To EndNumber

                ' 每一個Combo重新初始化
                SelectionNumber = 0
                Ratio = 0

                For SelectionNumber = 1 To PMMNumber

                    PMM = PMMArray(SelectionNumber)

                    ' 19條線
                    For LineNumber = 1 To 19

                        ' PMM的資料格式：
                        ' M P Angle b c
                        ' ComboPMM的資料格式：
                        ' Name M P Angle
                        If Newton(ComboPMM(ColumnNumber, 1), PMM(LineNumber, 3), ComboPMM(ColumnNumber, 2), PMM(LineNumber, 4), PMM(LineNumber - 1, 2), PMM(LineNumber, 2), ComboPMM(ColumnNumber, 3)) Then
                            Ratio = CaculateRatio(ComboPMM(ColumnNumber, 1), ComboPMM(ColumnNumber, 2), PMM(LineNumber, 3), PMM(LineNumber, 4))
                            GoTo NextCombo
                        End If

                    Next
                Next



NextCombo:
                ' Combo Loop 結束
                ' 超出所有PMMCurve，例外處理
                If SelectionNumber = 0 Then
                    SelectionNumber = PMMNumber + 1
                    SelectionSection(ColumnNumber, 4) = PMMNumber + 1
                Else
                    SelectionSection(ColumnNumber, 4) = SelectionNumber
                End If



                ' 判斷有沒有大於FinalSelectionNumber，有的話才寫入
                If FinalSelectionNumber < SelectionNumber Then
                    FinalSelectionNumber = SelectionNumber
                    FinalRatio = Ratio
                End If

                ' 判斷有沒有大於Ratio，有的話才寫入
                If FinalRatio < Ratio And FinalSelectionNumber <= SelectionNumber Then
                    FinalRatio = Ratio
                End If

            Next


            ' 斷面的Loop 結束
            ' 寫入斷面資料
            SelectionSection(SelectionSectionNumber, 0) = ComboPMM(RowNumber, 0)
            SelectionSection(SelectionSectionNumber, 1) = FinalSelectionNumber
            SelectionSection(SelectionSectionNumber, 2) = PMMCurveName(FinalSelectionNumber)
            SelectionSection(SelectionSectionNumber, 3) = FinalRatio

            ' 下一組的開始編號
            StartNumber = RowNumber + 1

            ' 下一組
            SelectionSectionNumber = SelectionSectionNumber + 1

        End If

    Next

    SelectionSeltor = SelectionSection()

End Function

Sub ColorSort()
   'Set up your variables and turn off screen updating.
   Dim iCounter As Integer
   Application.ScreenUpdating = False

   'For each cell in column A, go through and place the color index value of the cell in column C.
   For iCounter = 2 To 55
      Cells(iCounter, 3) = _
         Cells(iCounter, 1).Interior.ColorIndex
   Next iCounter

   'Sort the rows based on the data in column C
   Range("C1") = "Index"
   Columns("A:C").Sort key1:=Range("C2"), _
      order1:=xlAscending, header:=xlYes

   'Clear out the temporary sorting value in column C, and turn screen updating back on.
   Columns(3).ClearContents
   Application.ScreenUpdating = True
End Sub

Sub Ex()
    Dim Ar, Rng As Range
    Ar = Array("SD", 100, "SA", 50, 777, "AAA", 5)
    With ActiveSheet
        Set Rng = .[a1].Resize(UBound(Ar) + 1)
        Rng.Value = Application.Transpose(Ar)
        Rng.Sort Key1:=Rng(1), Order1:=xlAscending, Header:=xlNo
        Ar = Application.Transpose(Rng)
    End With
End Sub

Sub 巨集1()
'
' 巨集1 巨集
'

'
    Cells.Select
    ActiveWorkbook.Worksheets("Z").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Z").Sort.SortFields.Add Key:=Range("J7:J1420"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("Z").Sort.SortFields.Add Key:=Range("H7:H1420"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Z").Sort
        .SetRange Range("A7:J1420")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub