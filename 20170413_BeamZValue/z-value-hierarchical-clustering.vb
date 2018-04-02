Private WS_Z
Private APP


Function ClearBeforeOutputData()
'
' 清空前次輸出的資料
'

    With WS_Z
        rowEnd = .Cells(Rows.Count, 1).End(xlUp).Row
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

    colZValue = 9



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

    ReDim arrDistance(2 To uBoundZValue, 1 To 3)

    Const distance = 1
    Const count = 2
    Const sum = 3


    For i = 2 To uBoundZValue
        arrDistance(i, distance) = arrZValue(i - 1, 1) - arrZValue(i, 1)
        arrDistance(i, count) = 1
        arrDistance(i, sum) = arrDistance(i, distance) * arrDistance(i, count)
    Next i

    countOfValue = uBoundZValue

    While countOfValue > selectValue

        minPointValue = MinPoint(arrDistance, sum)

        If minPointValue <> 2 And minPointValue <> UBound(arrDistance) Then

            ' 上面會吃掉下面，所以群集數增加
            arrDistance(minPointValue - 1, count) = arrDistance(minPointValue - 1, count) + arrDistance(minPointValue, count)

            ' 下面與上面的距離差變大
            arrDistance(minPointValue + 1, distance) = arrDistance(minPointValue + 1, distance) + arrDistance(minPointValue, distance)

            ' 重新計算變化的總和 = 群集數 * 距離差
            arrDistance(minPointValue - 1, sum) = arrDistance(minPointValue - 1, count) * arrDistance(minPointValue - 1, distance)
            arrDistance(minPointValue + 1, sum) = arrDistance(minPointValue + 1, count) * arrDistance(minPointValue + 1, distance)

        ElseIf minPointValue = 2 Then

            ' 下面與上面的距離差變大
            arrDistance(minPointValue + 1, distance) = arrDistance(minPointValue + 1, distance) + arrDistance(minPointValue, distance)
            arrDistance(minPointValue + 1, sum) = arrDistance(minPointValue + 1, count) * arrDistance(minPointValue + 1, distance)

        ElseIf minPointValue = UBound(arrDistance) Then
            ' 上面會吃掉下面，所以群集數增加
            arrDistance(minPointValue - 1, count) = arrDistance(minPointValue - 1, count) + arrDistance(minPointValue, count)
            arrDistance(minPointValue - 1, sum) = arrDistance(minPointValue - 1, count) * arrDistance(minPointValue - 1, distance)

        End If

        arrDistance = DeleteElementAt(minPointValue, arrDistance)
        arrZValue = DeleteElementAt(minPointValue, arrZValue)

        countOfValue = countOfValue - 1

    Wend

    HierarchicalClustering = arrZValue

End Function


Function DeleteElementAt(Byval index As Integer, Byval arrOld as Variant)
'
'
'
' @param
' @returns

    Dim arrNew()

    rowStart = LBound(arrOld)
    rowEnd = UBound(arrOld)
    colStart = LBound(arrOld, 2)
    colEnd = UBound(arrOld, 2)

    ReDim arrNew(rowStart To rowEnd - 1, colStart To colEnd)

    ' Move all element back one position
    For i = rowStart To index - 1
        For j = colStart To colEnd
            arrNew(i, j) = arrOld(i, j)
        Next j
    Next

    For i = index + 1 To rowEnd
        For j = colStart To colEnd
            arrNew(i - 1, j) = arrOld(i, j)
        Next j
    Next

    DeleteElementAt = arrNew

End Function


' Sub DeleteElementAt(Byval index As Integer, Byref prLst as Variant)
' '
' '
' '
' ' @param
' ' @returns

'     rowStart = LBound(prLst)
'     rowEnd = UBound(prLst)
'     colStart = LBound(prLst, 2)
'     colEnd = UBound(prLst, 2)

'     ' Move all element back one position
'     For i = index + 1 To rowEnd
'         For j = colStart To colEnd
'             prLst(i - 1, j) = prLst(i, j)
'         Next j
'     Next

'     ' Shrink the array by one, removing the last one
'     ReDim Preserve prLst(rowStart To rowEnd - 1, colStart To colEnd)

' End Sub


Function MinPoint(arrValue, colCompare)
'
' 回傳陣列最小值所在位置
'
' @param
' @returns

   lBoundValue = LBound(arrValue)
   uBoundValue = UBound(arrValue)

   minValue = arrValue(lBoundValue, colCompare)

   For i = lBoundValue To uBoundValue
       If arrValue(i, colCompare) <= minValue Then
           minValue = arrValue(i, colCompare)
           minPointValue = i
       End If
   Next i

   MinPoint = minPointValue

End Function


Function SumValue()
'
'
'
' @param
' @returns



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

    Set WS_Z = Worksheets("z-value-hierarchical-clustering")

    Call ClearBeforeOutputData

    Call SortZValue

    arrZValue = GetArray(WS_Z, 9, 9, 9, 9)
    dblOriginalSum = APP.Sum(arrZValue)

    arrSelectValue = Split(WS_Z.Cells(5, 2), ",")
    uBoundSelectValue = UBound(arrSelectValue)

    ReDim arrOutputSelectValue(1 To 1000, 1 To uBoundSelectValue + 1)

    For i = 0 To uBoundSelectValue
        selectValue = Int(arrSelectValue(i))
        arrSelectZValue = HierarchicalClustering(arrZValue, selectValue)

        sum = 0
        For j = UBound(arrZValue) To LBound(arrZValue) Step -1
            Do
                For k = UBound(arrSelectZValue) To LBound(arrSelectZValue) Step -1
                    If arrSelectZValue(k, 1) >= arrZValue(j, 1) Then
                        sum = sum + arrSelectZValue(k, 1)
                        Exit for
                    End If
                Next k
            Loop Until True
        Next j

        With WS_Z
            .Range(.Cells(LBound(arrSelectZValue) + 7, 12 + i), .Cells(UBound(arrSelectZValue) + 7, 12 + i)) = arrSelectZValue
            .Cells(18 + i, 1) = selectValue
            .Cells(6, 12 + i) = selectValue
            .Cells(18 + i, 2) = sum / dblOriginalSum
        End With
    Next i

End Sub
