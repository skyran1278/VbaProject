Function SetTestGlobalVar()
'
' set global variable.
'

    ' global var
    Set wsBeam = Worksheets("大梁配筋 TEST")
    Set wsResult = Worksheets("最佳化斷筋點 TEST")
    Set wsErr = Worksheets("ERROR")

    ' 從第二列開始
    varErrNum = 2

    arrRebarSize = ran.GetRangeToArray(Worksheets("Rebar Size"), 1, 1, 1, 10)

    ' #3 => 0.9525cm
    Set objRebarSizeToDb = ran.CreateDictionary(arrRebarSize, 1, 7)

    ' #3 => 0.71cm^2
    Set objRebarSizeToArea = ran.CreateDictionary(arrRebarSize, 1, 10)

    ' 第一列也抓進來，方便秀出錯誤訊息。
    arrInfo = ran.GetRangeToArray(Worksheets("General Information"), 1, 4, 4, 12)

    lbRowInfo = LBound(arrInfo, 1)
    ubRowInfo = UBound(arrInfo, 1)
    lbColInfo = LBound(arrInfo, 2)
    ubColInfo = UBound(arrInfo, 2)

    ' 掃描是否有沒輸入的數值
    For i = lbRowInfo To ubRowInfo
        For j = lbColInfo To ubColInfo

            If arrInfo(i, j) = "" Then
                PrintErr "General Information " & arrInfo(i, 1) & " " & arrInfo(1, j) & " 是否空白？"
            End If

        Next j
    Next i

    Set objStoryToFy = ran.CreateDictionary(arrInfo, 1, 2)
    Set objStoryToFyt = ran.CreateDictionary(arrInfo, 1, 3)
    Set objStoryToFc = ran.CreateDictionary(arrInfo, 1, 4)
    Set objStoryToSDL = ran.CreateDictionary(arrInfo, 1, 5)
    Set objStoryToLL = ran.CreateDictionary(arrInfo, 1, 6)
    Set objStoryToBand = ran.CreateDictionary(arrInfo, 1, 7)
    Set objStoryToSlab = ran.CreateDictionary(arrInfo, 1, 8)
    Set objStoryToCover = ran.CreateDictionary(arrInfo, 1, 9)

End Function


Function ClearPrevOutputData()
'
' 清空前次輸出的資料.
'
    With wsResult
        .Range(.Cells(3, 6), .Cells(.Cells(Rows.Count, 6).End(xlUp).Row, 26)).ClearContents
    End With

End Function

Function PrintResult(ByVal arrResult, ByVal rowStart)
'
' 列印出最佳化結果
'
' @param {Array} [arrResult] 需要 print 出的陣列.
' @return {Number} [rowStartNext] 回傳下一次從第幾列 print.
'

    colStart = 6

    rowEnd = rowStart + UBound(arrResult, 1) - LBound(arrResult, 1)
    colEnd = colStart + UBound(arrResult, 2) - LBound(arrResult, 2)

    With wsResult
        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = arrResult
    End With

    rowStartNext = rowStart + 4
    PrintResult = rowStartNext

End Function


Sub Test()


    Set ran = New UTILS_CLASS
    Set APP = Application.WorksheetFunction

    Call ran.ExecutionTime(True)
    Call ran.PerformanceVBA(True)

    Call SetTestGlobalVar

    Call ClearPrevOutputData

    ' 不包含標題
    arrBeam = ran.GetRangeToArray(wsBeam, 3, 1, 5, 16)

    arrRebar1stNum = GetRebar1stNum(arrBeam)

    arrRebarTotalNum = GetRebarTotalNum(arrBeam)

    arrRebarTotalArea = GetRebarTotalArea(arrBeam)

    arrNormalSplice = CalNormalSplice(arrRebarTotalNum)

    arrGravity = CalGravityDemand(arrBeam)

    arrMultiRebar = OptimizeMultiRebar(arrBeam, arrRebarTotalArea, arrGravity, arrNormalSplice)

    arrLapLength = CalLapLength(arrBeam, arrRebar1stNum, arrMultiRebar)

    arrSmartSplice = CalSmartSplice(arrMultiRebar, arrLapLength)

    arrSmartSpliceModify = CalOptimizeNoMoreThanNormal(arrSmartSplice, arrNormalSplice)

    arrThreePoints = ThreePoints(arrBeam, arrSmartSpliceModify)

    arrMultiThreePoints = ConvertThreePoints(arrThreePoints)

    rowStartNext = PrintResult(arrRebar1stNum, 3)
    rowStartNext = PrintResult(arrRebarTotalNum, rowStartNext)
    rowStartNext = PrintResult(arrRebarTotalArea, rowStartNext)
    rowStartNext = PrintResult(arrNormalSplice, rowStartNext)
    rowStartNext = PrintResult(arrGravity, rowStartNext)
    rowStartNext = PrintResult(arrMultiRebar, rowStartNext)
    rowStartNext = PrintResult(arrLapLength, rowStartNext)
    rowStartNext = PrintResult(arrSmartSplice, rowStartNext)
    rowStartNext = PrintResult(arrSmartSpliceModify, rowStartNext)
    rowStartNext = PrintResult(arrThreePoints, rowStartNext)
    rowStartNext = PrintResult(arrMultiThreePoints, rowStartNext)

    wsResult.Activate

    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTime(False)

End Sub
