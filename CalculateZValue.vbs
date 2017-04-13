Sub ZValue()
'
'排序
'讀進陣列
'判斷是不是同一個Group
'同個Group中取40個或30個或20個做排列組合
'找出SUM最小的
'就是答案

    Dim Value() As String
    Dim CopyValue() As String

    '讀取第8欄之列數
    ZRowUsed = Cells(Rows.Count, 8).End(xlUp).Row

    ReDim Value(ZRowUsed - 9 , 1)
    


    Worksheets("Z").Range(Cells(7, 3), Cells(ZRowUsed, 10)).Sort _
        Key1:=Range(Cells(8, 10), Cells(ZRowUsed, 10)), Order1:=xlAscending, _
        Key2:=Range(Cells(8, 8), Cells(ZRowUsed, 8)), Order2:=xlDescending, Header:=xlYes

    '陣列計數
    ValueRowNumber = -1

    For ZRowNumber = 8 To ZRowUsed

        '陣列計數
        ValueRowNumber = ValueRowNumber + 1

        'Group
        Value(ValueRowNumber, 0) = Cells(ZRowNumber, 10)

        'ZValue
        Value(ValueRowNumber, 1) = Cells(ZRowNumber, 8)

    Next

    For ValueRowNumber = 0 To ZRowUsed - 9

    	StartNumber = 0

    	If Value(ValueRowNumber, 0) <> Value(ValueRowNumber + 1, 0) Then
    		
    		EndNumber = ValueRowNumber

    		ZValueNumber = EndNumber - StartNumber

    		If ZValueNumber > Cells(4, 4) Then

    			CopyValue() = Value()

    			
    			Interval = FIX(ZValueNumber / Cells(4, 4))

    			For CopyValueRowNumber = StartNumber To EndNumber step Interval

    				Value(CopyValueRowNumber, 0)

	    			For IntervalNumber = 0 To Interval - 1

	    				CopyValue(CopyValueRowNumber + IntervalNumber, 0) = Value(CopyValueRowNumber, 0)

	    			Next    				

    			Next

    			

    		End If

    		
    	End If

    	StartNumber = ValueRowNumber + 1
    Next





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