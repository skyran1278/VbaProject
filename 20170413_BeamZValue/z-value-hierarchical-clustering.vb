Private WS_Z
Private APP
Private Const INFINITY = 1.00E+307


Function ClearBeforeOutputData()
'
' 清空前次輸出的資料
'

    With WS_Z
        rowEnd = .Cells(Rows.Count, 18).End(xlUp).Row
        If rowEnd > 18 Then
            .Range(.Cells(18, 1), .Cells(rowEnd, 2)).ClearContents
        End If
        .Range(.Columns(12), .Columns(100)).ClearContents
    End With

End Function


Sub SortZValue()
'
' 排序
'

    colZValue = 8



    With WS_Z
        rowStart = 7
        colStart = 3
        rowEnd = .Cells(Rows.Count, colZValue).End(xlUp).Row
        colEnd = 10
        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)).Sort _
        Key1:=.Range(.Cells(rowStart, colZValue), .Cells(rowEnd, colZValue)), Order1:=xlDescending, Header:=xlYes
    End With

End Sub


Function HierarchicalClustering(ByVal arrZValue, selectValue)
'
' 階層式分群法
'
' @param
' @returns

    Dim arrDistance()

    uBoundZValue = UBound(arrZValue)

    ReDim arrDistance(2 To uBoundZValue)

    For i = 2 To uBoundZValue
        arrDistance(i) = arrZValue(i - 1) - arrZValue(i)
    Next i

    countOfValue = uBoundZValue

    While countOfValue > selectValue

        minPointValue = MinPoint(arrDistance)

        If minPointValue <> 2 And minPointValue <> UBound(arrDistance) Then
            arrDistance(minPointValue - 1) = arrDistance(minPointValue - 1) + arrDistance(minPointValue) / 2
            arrDistance(minPointValue + 1) = arrDistance(minPointValue + 1) + arrDistance(minPointValue) / 2

        ElseIf minPointValue = 2 Then
            arrDistance(minPointValue + 1) = arrDistance(minPointValue + 1) + arrDistance(minPointValue) / 2

        ElseIf minPointValue = UBound(arrDistance) Then
            arrDistance(minPointValue - 1) = arrDistance(minPointValue - 1) + arrDistance(minPointValue) / 2

        End If

        Call DeleteElementAt(minPointValue, arrDistance)
        Call DeleteElementAt(minPointValue, arrZValue)

        countOfValue = countOfValue - 1

    Wend

    HierarchicalClustering = arrZValue

End Function


Sub DeleteElementAt(Byval index As Integer, Byref prLst as Variant)
'
'
'
' @param
' @returns

    ' Move all element back one position
    For i = index + 1 To UBound(prLst)
        prLst(i - 1) = prLst(i)
    Next

    ' Shrink the array by one, removing the last one
    ReDim Preserve prLst(LBound(prLst) To UBound(prLst) - 1)

End Sub


Function MinPoint(arrValue)
'
' 回傳陣列最小值所在位置
'
' @param
' @returns

   lBoundValue = LBound(arrValue)
   uBoundValue = UBound(arrValue)

   minValue = arrValue(lBoundValue)

   For i = lBoundValue To uBoundValue
       If arrValue(i) <= minValue Then
           minValue = arrValue(i)
           minPointValue = i
       End If
   Next i

   MinPoint = minPointValue

End Function


Sub Main()
'
' @purpose:
' 找出最佳化數值
'
'
' @algorithm:
' 階層式演算法
'
'
' @test:
'
'
'
    Dim arrOutputSelectValue()

    time0 = Timer

    Set APP = Application.WorksheetFunction

    Set WS_Z = Worksheets("Z-階層式分群法")

    Call ClearBeforeOutputData

    Call SortZValue

    arrZValue = APP.Transpose(GetArray(WS_Z, 8, 8, 8, 8))

    arrSelectValue = Split(WS_Z.Cells(5, 2), ",")
    uBoundSelectValue = UBound(arrSelectValue)

    ReDim arrOutputSelectValue(1 To 1000, 1 To uBoundSelectValue + 1)

    For i = 0 To uBoundSelectValue
        selectValue = Int(arrSelectValue(i))
        arrSelectZValue = HierarchicalClustering(arrZValue, selectValue)

        With WS_Z
            .Range(.Cells(LBound(arrSelectZValue) + 7, 12 + i), .Cells(UBound(arrSelectZValue) + 7, 12 + i)) = APP.Transpose(arrSelectZValue)
        End With
    Next i

End Sub
