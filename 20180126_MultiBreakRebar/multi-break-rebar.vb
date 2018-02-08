Private WS_BEAM


Function ClearBeforeOutputData()
'
' 清空前次輸出的資料
'

    With WS_BEAM
        rowStart = 3
        colStart = 17
        rowEnd = .Cells(Rows.Count, 17).End(xlUp).Row
        colEnd = 38

        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)).Clear
    End With

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
            arrRebarNumber(i, j - 5) = Split(arrRawData(i, j), "-") + Split(arrRawData(i + 1, j), "-")
        Next
    Next

    CalRebarNumber = arrRebarNumber

End Function


Function PrintResult(arrResult)
'
' 列印出最佳化結果
'
' @param arrResult(Array)

    With WS_BEAM
        rowStart = 3
        colStart = 18
        rowEnd = rowStart + UBound(arrResult, 1) - 1
        colEnd = colStart + UBound(arrResult, 2) - 1

        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = arrResult
    End With

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
        rate = 1
        For j = 1 To 11
            arrMultiBreakRebar(i, j) = Fix(rate * arrRebarNumber(i, varleft)) + 1
            rate = rate - 0.1
        Next j

        ' 中央到右端
        rate = 0.1
        For j = 12 To 21
            arrMultiBreakRebar(i, j) = Fix(rate * arrRebarNumber(i, varright)) + 1
            rate = rate + 0.1
        Next j

    Next i

    For i = 3 To ubMultiBreakRebar Step 4

        arrMultiBreakRebar(i, 0) = "下層"

        ' 左端到中央
        rate = 1
        For j = 1 To 11
            arrMultiBreakRebar(i, j) = Fix(Max(rate * arrRebarNumber(i, varleft), (1 - rate ^ 2) * arrRebarNumber(i, varmid))) + 1
            rate = rate - 0.1
        Next j

        ' 中央到右端
        rate = 0.1
        For j = 12 To 21
            arrMultiBreakRebar(i, j) = Fix(Max(rate * arrRebarNumber(i, varright), (1 - rate ^ 2) * arrRebarNumber(i, varmid))) + 1
            rate = rate + 0.1
        Next j

    Next i

    Call PrintResult(arrMultiBreakRebar)

    Call FontSetting(WS_BEAM)
    Call PerformanceVBA(False)
    Call ExecutionTimeVBA(time0)

End Sub
