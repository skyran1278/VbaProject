Private WS_BEAM, WS_RESULT


Function ClearBeforeOutputData()
'
' 清空前次輸出的資料
'

    WS_RESULT.Cells.Clear

End Function


Function GetRawData()
'
' 取得 beam rebar 資料
'
' @returns GetRawData(Array)

    With WS_BEAM
        rowStart = 3
        colStart = 1
        rowEnd = .Cells(Rows.Count, 5).End(xlUp).Row
        colEnd = 16

        GetRawData = .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd))
    End With

End Function


Function CalRebarNumber(arrRawData)
'
' 計算上下排總支數
'
' @param
' @returns

    Dim arrRebarNumber()
    Redim arrRebarNumber(1 To UBound(arrRawData), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrRawData)
    colStart = 6
    colEnd = 8

    ' 一二排相加
    For i = rowStart To rowEnd Step 2
        For j = colStart To colEnd
            arrRebarNumber(i, j - 5) = Int(Split(arrRawData(i, j), "-")(0)) + Int(Split(arrRawData(i + 1, j), "-")(0))
        Next
    Next

    CalRebarNumber = arrRebarNumber

End Function


Function PrintResult(arrResult)
'
' 列印出最佳化結果
'
' @param arrResult(Array)

    With WS_RESULT
        rowStart = 3
        colStart = 3
        rowEnd = rowStart + UBound(arrResult, 1) - 1
        colEnd = colStart + UBound(arrResult, 2)

        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = arrResult
    End With

    ' 格式化條件
    For i = rowStart To rowEnd
        With WS_RESULT.Range(WS_RESULT.Cells(i, colStart), WS_RESULT.Cells(i, colEnd))
            .FormatConditions.AddColorScale ColorScaleType:=3
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = 8109667

            .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
            .FormatConditions(1).ColorScaleCriteria(2).Value = 50
            .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = 8711167

            .FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
            .FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = 7039480
        End With
    Next i

End Function


Sub Main()
'
' @purpose:
' reduce 鋼筋量
'
'
' @algorithm:
' 上層筋由耐震控制
' 下層筋由重力與耐震共同控制
'
' @test:
'
'
'

    Dim arrMultiBreakRebar
    Dim time0 As Double

    time0 = Timer

    Call PerformanceVBA(True)

    Set WS_BEAM = Worksheets("小梁配筋")
    Set WS_RESULT = Worksheets("最佳化斷筋點")

    Call ClearBeforeOutputData
    arrBeam = GetRawData()
    arrRebarNumber = CalRebarNumber(arrBeam)

    ubRebarNumber = UBound(arrRebarNumber)

    ReDim arrMultiBreakRebar(1 To ubRebarNumber, 21)

    ubMultiBreakRebar = UBound(arrMultiBreakRebar)

    varleft = 1
    varmid = 2
    varright = 3

    For i = 1 To ubMultiBreakRebar Step 4

        arrMultiBreakRebar(i, 0) = "上層"

        ' 左端到中央
        ratio = 1
        For j = 1 To 11
            arrMultiBreakRebar(i, j) = RoundUp(Max(ratio * arrRebarNumber(i, varleft), 2))
            ratio = ratio - 0.1
        Next j

        ' 中央到右端
        ratio = 0.1
        For j = 12 To 21
            arrMultiBreakRebar(i, j) = RoundUp(Max(ratio * arrRebarNumber(i, varright), 2))
            ratio = ratio + 0.1
        Next j

    Next i

    For i = 3 To ubMultiBreakRebar Step 4

        arrMultiBreakRebar(i, 0) = "下層"

        ' 左端到中央
        ratio = 1
        For j = 1 To 11
            arrMultiBreakRebar(i, j) = RoundUp(Max(ratio * arrRebarNumber(i, varleft), (1 - ratio ^ 2) * arrRebarNumber(i, varmid), 2))
            ratio = ratio - 0.1
        Next j

        ' 中央到右端
        ratio = 0.1
        For j = 12 To 21
            arrMultiBreakRebar(i, j) = RoundUp(Max(ratio * arrRebarNumber(i, varright), (1 - ratio ^ 2) * arrRebarNumber(i, varmid), 2))
            ratio = ratio + 0.1
        Next j

    Next i

    Call PrintResult(arrMultiBreakRebar)

    Call FontSetting(WS_RESULT)
    Call PerformanceVBA(False)
    Call ExecutionTimeVBA(time0)

End Sub
