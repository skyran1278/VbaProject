Function SetTestGlobalVar()
'
' set global variable.
'

    ' global var
    Set wsBeam = Worksheets("大梁配筋 TEST")
    Set wsResult = Worksheets("最佳化斷筋點 TEST")

    arrRebarSize = ran.GetRangeToArray(Worksheets("Rebar Size"), 1, 1, 1, 10)

    ' #3 => 0.9525cm
    Set objRebarSizeToDb = ran.CreateDictionary(arrRebarSize, 1, 7)

    ' #3 => 0.71cm^2
    Set objRebarSizeToArea = ran.CreateDictionary(arrRebarSize, 1, 10)

    arrInfo = ran.GetRangeToArray(Worksheets("General Information TEST"), 2, 4, 4, 12)

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
        .Range(.Cells(3, 28), .Cells(.Cells(Rows.Count, 29).End(xlUp).Row, 49)).ClearContents
    End With

End Function

Function PrintResult(ByVal arrResult, ByVal rowStart, ByVal colStart)
'
' 列印出最佳化結果
'
' @param {Array} [arrResult] 需要 print 出的陣列.
' @param {Array} [colStart] 從哪一列開始.
' @return {Number} [rowStartNext] 回傳下一次從第幾列 print.
'

    With wsResult
        rowEnd = rowStart + UBound(arrResult, 1) - LBound(arrResult, 1)
        colEnd = colStart + UBound(arrResult, 2) - LBound(arrResult, 2)

        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = arrResult
    End With

    rowStartNext = rowStart + 4
    PrintResult = rowStartNext

End Function


Sub Test()

    Dim time0 As Double

    time0 = Timer

    Set ran = New UTILS_CLASS
    Set APP = Application.WorksheetFunction

    Call ran.PerformanceVBA(True)

    Call SetTestGlobalVar

    Call ClearPrevOutputData

    ' 不包含標題
    arrBeam = ran.GetRangeToArray(wsBeam, 3, 1, 5, 16)

    arrRebarTotalNum = CalRebarTotalNum(arrBeam)

    arrNormalSplice = CalNormalGirderMultiRebar(arrRebarTotalNum)

    arrRebarTotalArea = CalRebarTotalArea(arrBeam)

    arrGirderMultiRebar = OptimizeGirderMultiRebar(arrBeam, arrRebarTotalArea)

    arrLapLengthRatio = CalLapLengthRatio(arrBeam)
    arrMultiLapLength = CalMultiLapLength(arrLapLengthRatio)

    arrSmartSplice = CalSplice(arrGirderMultiRebar, arrMultiLapLength)

    arrSmartSpliceModify = CalOptimizeNoMoreThanNormal(arrSmartSplice, arrNormalSplice)

    rowStartNext = PrintResult(arrRebarTotalNum, 3, 29)
    rowStartNext = PrintResult(arrRebarTotalArea, rowStartNext, 29)
    rowStartNext = PrintResult(arrNormalSplice, rowStartNext, 28)
    rowStartNext = PrintResult(arrGirderMultiRebar, rowStartNext, 28)
    rowStartNext = PrintResult(arrLapLengthRatio, rowStartNext, 29)
    rowStartNext = PrintResult(arrMultiLapLength, rowStartNext, 28)
    rowStartNext = PrintResult(arrSmartSplice, rowStartNext, 28)
    rowStartNext = PrintResult(arrSmartSpliceModify, rowStartNext, 28)

    Call ran.FontSetting(wsResult)
    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTimeVBA(time0)

End Sub
