Function SetTestGlobalVar()
'
' set global variable.
'

    ' global var
    Set wsBeam = Worksheets("大梁配筋 TEST")
    Set wsResult = Worksheets("最佳化斷筋點 TEST")

    ' #3 => 0.9525cm
    Set objRebarSizeToDb = ran.CreateDictionary(ran.GetRangeToArray(Worksheets("Rebar Size"), 1, 1, 1, 10), 1, 7)

    arrInfo = ran.GetRangeToArray(Worksheets("General Information TEST"), 2, 4, 4, 13)

    Set objStoryToFy = ran.CreateDictionary(arrInfo, 1, 2)
    Set objStoryToFyt = ran.CreateDictionary(arrInfo, 1, 3)
    Set objStoryToFc = ran.CreateDictionary(arrInfo, 1, 4)
    Set objStoryToCover = ran.CreateDictionary(arrInfo, 1, 10)

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
'

    With wsResult
        rowEnd = rowStart + UBound(arrResult, 1) - 1
        colEnd = colStart + UBound(arrResult, 2)

        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = arrResult
    End With

End Function


Private Function CalTotalRebarTest(ByVal arrTotalRebar)
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'

    Expect arrTotalRebar(1, 1) = 9
    Expect arrTotalRebar(1, 2) = 4
    Expect arrTotalRebar(1, 3) = 10

    Expect arrTotalRebar(2, 1) = 0
    Expect arrTotalRebar(2, 2) = 0
    Expect arrTotalRebar(2, 3) = 0

    Expect arrTotalRebar(3, 1) = 10
    Expect arrTotalRebar(3, 2) = 5
    Expect arrTotalRebar(3, 3) = 9

    Expect arrTotalRebar(4, 1) = 0
    Expect arrTotalRebar(4, 2) = 0
    Expect arrTotalRebar(4, 3) = 0

End Function

Private Function Expect(ByVal bol, Optional ByVal title = "Title")
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'

    If Not bol Then
        MsgBox actual & " <> " & expected, vbOKOnly, title
    End If

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

    arrTotalRebar = CalTotalRebar(arrBeam)

    Call CalTotalRebarTest(arrTotalRebar)

    arrGirderMultiRebar = OptimizeGirderMultiRebar(arrTotalRebar)
    arrNormalGirderMultiRebar = CalNormalGirderMultiRebar(arrTotalRebar)

    Call PrintResult(arrGirderMultiRebar, 3, 28)
    Call PrintResult(arrNormalGirderMultiRebar, 7, 28)

    arrGirderMultiRebar = CalOptimizeNoMoreThanNormal(arrGirderMultiRebar, arrNormalGirderMultiRebar)

    arrLapLengthRatio = CalLapLengthRatio(arrBeam)
    arrMultiLapLength = CalMultiLapLength(arrLapLengthRatio)

    arrSmartSplice = CalSplice(arrGirderMultiRebar, arrMultiLapLength)
    arrNormalSplice = CalSplice(arrNormalGirderMultiRebar, arrMultiLapLength)

    ' arrSmartSplice = OptimizeGirderMultiRebar(arrTotalRebar)
    ' arrNormalSplice = CalNormalGirderMultiRebar(arrTotalRebar)

    ' arrOptimizeResult = CalOptimizeResult(arrSmartSplice, arrNormalSplice)

    ' Call PrintResult(arrSmartSplice, 3)
    ' Call PrintResult(arrNormalSplice, varSpliceNum + 3 + 1)
    ' Call PrintResult(arrOptimizeResult, 2 * varSpliceNum + 3 + 2)
    ' wsResult.Cells(2, 2) = APP.Average(arrOptimizeResult)

    Call ran.FontSetting(wsResult)
    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTimeVBA(time0)

End Sub
