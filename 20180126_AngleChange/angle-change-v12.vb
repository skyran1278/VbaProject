Private WS_LINE, WS_DISP, WS_RESULT


Function GetLines()
'
' 取得 safe beam length 資料
'
' @returns GetLines(Array)

    With WS_LINE
        rowStart = 4
        colStart = 1
        rowEnd = .Cells(Rows.Count, 1).End(xlUp).Row
        colEnd = 6

        GetLines = .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd))
    End With

End Function


Function GetDisp()
'
' 取得 safe node displacement 資料
'
' @returns GetDisp(Array)

    With WS_DISP
        rowStart = 4
        colStart = 1
        rowEnd = .Cells(Rows.Count, 2).End(xlUp).Row
        colEnd = 7

        GetDisp = .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd))
    End With

End Function


Function CombinedDispArray(dispArray)
'
' 合併 point 和 ASD，使之成為唯一ID。
' 取出 Z displacement
'
' @param dispArray(Array)
' @returns CombinedDispArray(Array)

    Dim combinedArray()

    dispLBound = LBound(dispArray, 1)
    dispUBound = UBound(dispArray, 1)

    ReDim combinedArray(dispLBound To dispUBound, 1 To 2)

    ID = 2
    ASD = 3
    zDisp = 7

    For dispRow = dispLBound To dispUBound
        combinedArray(dispRow, 1) = dispArray(dispRow, ID) & dispArray(dispRow, ASD)
        combinedArray(dispRow, 2) = dispArray(dispRow, zDisp)
    Next dispRow

    CombinedDispArray = combinedArray

End Function


Function SetDictionary(idAndLoadArray)
'
' 取代內建的 VLookup，效能大幅提升。
'
' @param idAndLoadArray(Array)
' @returns SetDictionary(Object)

    Set dictionary = CreateObject("Scripting.Dictionary")

    idAndLoadLBound = LBound(idAndLoadArray, 1)
    idAndLoadUBound = UBound(idAndLoadArray, 1)

    For idAndLoadRow = idAndLoadLBound To idAndLoadUBound
        If Not dictionary.Exists(idAndLoadArray(idAndLoadRow, 1)) Then
            Call dictionary.Add(idAndLoadArray(idAndLoadRow, 1), idAndLoadArray(idAndLoadRow, 2))
        End If
    Next idAndLoadRow

    Set SetDictionary = dictionary

End Function


Function Max(values)
   Dim maxValue, Value

   maxValue = values(1, 1)

   For Each Value In values
       If Value > maxValue Then maxValue = Value
   Next

   Max = maxValue

End Function


Function GetMax10List(max10list, cur)
'
' 取得最大的 10 個
'
' @param max10list(List)
' @param cur(Array)
' @returns GetMax10List(List)


    While max10list.Count < 10
        max10list.Add Array(Null, Null, Null, Null, Null, Null, 0)
    Wend

    For i = 0 To max10list.Count - 1
        If cur(6) > max10list.Item(i)(6) Then
            max10list.Insert i, cur
            max10list.RemoveAt 10
            Exit For
        End If
    Next i

    Set GetMax10List = max10list

End Function


Function printResult(result)
'
' 列出結果
'
' @param result(Array)

    rowStart = 6
    rowEnd = rowStart + UBound(result) - 1
    colStart = 17
    colEnd = colStart + 108 - 1

    With WS_RESULT
        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = result
    End With

End Function


Function PrintMaxResult(maxResult)
'
' 列出最大值結果
'
' @param maxResult(Array)

    rowStart = 6
    rowEnd = rowStart + 36 - 1
    colStart = 5
    colEnd = colStart

    With WS_RESULT
        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = maxResult
    End With

End Function

Function Print10MaxResult(result)
'
' 列出前 10 個最大值的結果
'
' @param result(Array)

    Dim printResult(1 To 10, 1 To 7)

    rowStart = 6
    rowEnd = rowStart + UBound(result) - 1
    colStart = 8
    colEnd = colStart + 7 - 1

    rowLBound = LBound(printResult, 1)
    rowUBound = UBound(printResult, 1)

    colLBound = LBound(printResult, 2)
    colUBound = UBound(printResult, 2)

    For i = rowLBound To rowUBound
        For j = colLBound To colUBound
            printResult(i, j) = result(i - 1)(j - 1)
        Next j
    Next i

    With WS_RESULT
        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = printResult
    End With

End Function


Sub Main()
'
' @purpose:
' check 角變量 是否符合規範
'
'
' @algorithm:
' 桿件兩點的沈陷量除以桿件長度
'
'
' @test:
'
'
'
    Dim time0 As Double
    Dim result()
    Dim maxResult(1 To 36, 1 To 1)
    Set maxLineResult = CreateObject("System.Collections.ArrayList")

    time0 = Timer

    Call PerformanceVBA(True)

    Set WS_LINE = Worksheets("Lines-v12")
    Set WS_DISP = Worksheets("Nodal Displacements-v12")
    Set WS_RESULT = Worksheets("Result-v12")

    lineArray = GetLines()
    dispArray = GetDisp()
    Set dictionary = SetDictionary(CombinedDispArray(dispArray))

    lineLBound = LBound(lineArray, 1)
    lineUBound = UBound(lineArray, 1)

    ReDim result(lineLBound To lineUBound, 1 To 108)

    lineName = 1
    lineID1 = 2
    lineID2 = 3
    lineLength = 6

    For ASD = 1 To 36
        loadCombo = "ASD" & Format(ASD, "00")
        id1 = (ASD - 1) * 3 + 1
        id2 = (ASD - 1) * 3 + 2
        angleChange = (ASD - 1) * 3 + 3
        For lineRow = lineLBound To lineUBound
            id1AndLoad = lineArray(lineRow, lineID1) & loadCombo
            id2AndLoad = lineArray(lineRow, lineID2) & loadCombo
            result(lineRow, id1) = dictionary.Item(id1AndLoad)
            result(lineRow, id2) = dictionary.Item(id2AndLoad)
            result(lineRow, angleChange) = Abs(result(lineRow, id1) - result(lineRow, id2)) / lineArray(lineRow, lineLength)
            Set maxLineResult = GetMax10List(maxLineResult, Array(lineArray(lineRow, lineName), lineArray(lineRow, lineID1), lineArray(lineRow, lineID2), loadCombo, result(lineRow, id1), result(lineRow, id2), result(lineRow, angleChange)))
        Next lineRow

        If Max(Application.Index(result, 0, angleChange)) = 0 Then
            maxResult(ASD, 1) = "NaN"
        Else
            maxResult(ASD, 1) = 1 / Max(Application.Index(result, 0, angleChange))
        End If

    Next ASD

    Call printResult(result)
    Call PrintMaxResult(maxResult)
    Call Print10MaxResult(maxLineResult.ToArray)

    Call FontSetting(WS_RESULT)
    Call PerformanceVBA(False)
    Call ExecutionTimeVBA(time0)

End Sub
