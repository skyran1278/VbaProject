Private Const varSpliceNum = 10

Private ran As UTILS_CLASS
Private APP
Private wsBeam As Worksheet
Private wsResult As Worksheet
Private objRebarSizeToDb As Object
Private objStoryToFy As Object
Private objStoryToFyt As Object
Private objStoryToFc As Object
Private objStoryToCover As Object

' TODO: 需要重新書裡一次邏輯
' 大梁考慮耐震
' 小梁考慮負彎矩

Private Function SetGlobalVar()
'
' set global variable.
'

    ' global var
    Set wsBeam = Worksheets("大梁配筋")
    Set wsResult = Worksheets("最佳化斷筋點")

    ' #3 => 0.9525cm
    Set objRebarSizeToDb = ran.CreateDictionary(ran.GetRangeToArray(Worksheets("Rebar Size"), 1, 1, 1, 10), 1, 7)

    arrInfo = ran.GetRangeToArray(Worksheets("General Information"), 2, 4, 8, 8)

    Set objStoryToFy = ran.CreateDictionary(arrInfo, 1, 2)
    Set objStoryToFyt = ran.CreateDictionary(arrInfo, 1, 3)
    Set objStoryToFc = ran.CreateDictionary(arrInfo, 1, 4)
    Set objStoryToCover = ran.CreateDictionary(arrInfo, 1, 5)

End Function


Function ClearPrevOutputData()
'
' 清空前次輸出的資料.
'

    wsResult.Cells.Clear

End Function


Function CalTotalRebar(ByVal arrBeam)
'
' 計算上下排總支數
'
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrTotalRebar] 總支數，列數與 arrBeam 對齊，行數分左中右.
'

    Dim arrTotalRebar()
    ReDim arrTotalRebar(1 To UBound(arrBeam), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrBeam)
    colStart = 6
    colEnd = 8

    ' 一二排相加
    For i = rowStart To rowEnd Step 2
        For j = colStart To colEnd

            colTotalRebar = j - 5

            ' 計算上下排總支數
            arrTotalRebar(i, colTotalRebar) = Int(Split(arrBeam(i, j), "-")(0)) + Int(Split(arrBeam(i + 1, j), "-")(0))

        Next
    Next

    CalTotalRebar = arrTotalRebar

End Function


Function OptimizeGirderRebar(ByVal arrTotalRebar)
'
' 上層筋由耐震控制.
' 下層筋由重力與耐震共同控制.
'
' @param {Array} [arrTotalRebar] 總支數，列數與 arrBeam 對齊，行數分左中右.
' @return {Array} [arrGirderRebar] descrip.
'

    Dim arrGirderRebar

    ubTotalRebar = UBound(arrTotalRebar)

    ReDim arrGirderRebar(1 To ubTotalRebar, varSpliceNum)

    ubGirderRebar = UBound(arrGirderRebar)

    varLeft = 1
    varMid = 2
    varRight = 3

    ' 一半的地方
    varHalfOfSpliceNum = APP.RoundUp(varSpliceNum / 2, 0)

    ' 遞減的斜率
    slope = 1 / varHalfOfSpliceNum

    ' 上層筋由耐震控制.
    For i = 1 To ubGirderRebar Step 4

        arrGirderRebar(i, 0) = "上層"

        ' 左端到中央
        ratio = 1
        For j = 1 To varHalfOfSpliceNum
            ' 耐震和 2 支取大值
            arrGirderRebar(i, j) = APP.RoundUp(ran.Max(ratio * arrTotalRebar(i, varLeft), 2), 0)
            ratio = ratio - slope
        Next j

        ' 中央到右端
        ratio = slope
        For j = varHalfOfSpliceNum + 1 To varSpliceNum
            ' 耐震和 2 支取大值
            arrGirderRebar(i, j) = APP.RoundUp(ran.Max(ratio * arrTotalRebar(i, varRight), 2), 0)
            ratio = ratio + slope
        Next j

    Next i

    ' 下層筋由重力與耐震共同控制.
    For i = 3 To ubGirderRebar Step 4

        arrGirderRebar(i, 0) = "下層"

        ' 左端到中央
        ratio = 1
        For j = 1 To varHalfOfSpliceNum
            arrGirderRebar(i, j) = APP.RoundUp(ran.Max(ratio * arrTotalRebar(i, varLeft), (1 - ratio ^ 2) * arrTotalRebar(i, varMid), 2), 0)
            ratio = ratio - slope
        Next j

        ' 中央到右端
        ratio = slope
        For j = varHalfOfSpliceNum + 1 To varSpliceNum
            arrGirderRebar(i, j) = APP.RoundUp(ran.Max(ratio * arrTotalRebar(i, varRight), (1 - ratio ^ 2) * arrTotalRebar(i, varMid), 2), 0)
            ratio = ratio + slope
        Next j

    Next i

    OptimizeGirderRebar = arrGirderRebar


End Function


Function PrintResult(ByVal arrResult, ByVal colStart)
'
' 列印出最佳化結果
'
' @param arrResult(Array)

    With wsResult
        rowStart = 3
        rowEnd = rowStart + UBound(arrResult, 1) - 1
        colEnd = colStart + UBound(arrResult, 2)

        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = arrResult
    End With

    ' 格式化條件
    For i = rowStart To rowEnd
        With wsResult.Range(wsResult.Cells(i, colStart), wsResult.Cells(i, colEnd))
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


Private Function CalLapLength(ByVal arrBeam)
'
' TODO: 可以做優化，如果算過了就不用再算一次.
' 計算不同主筋的搭接長度.
' 回傳比例 除以梁長.
'
' @since 1.0.0
' @param {array} [arrBeam] RCAD 匯出的配筋.
' @return {array} [CalLapLength] 精算法的搭接長度 格式與 arrRebarTotalNumber 對齊.
'

    Dim arrLapLength

    pi_ = 3.1415926

    ' 鋼筋塗布修正因數
    ' 未塗布鋼筋
    psiE = 1

    ' 混凝土單位重之修正因數
    ' 於常重混凝土內之鋼筋
    lambda = 1

    ubBeam = UBound(arrBeam)

    ReDim arrLapLength(1 To ubBeam, 1 To 3)

    ubLapLength = UBound(arrLapLength)

    ' loop 全部
    For i = 1 To ubLapLength Step 4

        story = arrBeam(i, 1)

        fy_ = objStoryToFy.Item(story)
        fyt_ = objStoryToFyt.Item(story)
        fc_ = objStoryToFc.Item(story)
        cover = objStoryToCover.Item(story)

        width_ = arrBeam(i, 3)
        length_ = arrBeam(i, 13)

        ' loop 左中右
        For j = 6 To 8

            colBar = j
            colStirrup = j + 4
            colLapLength = j - 5

            tmp = Split(arrBeam(i, colStirrup), "@")
            stirrupSize = tmp(0)
            stirrupSpace = Int(tmp(1))

            fytDb = objRebarSizeToDb.Item(stirrupSize)

            ' loop 上下排
            For k = 0 To 3

                tmp = Split(arrBeam(i + k, colBar), "-")
                fyNum = Int(tmp(0))

                If fyNum = 0 Then
                    arrLapLength(i + k, colLapLength) = 0
                Else
                    barSize = tmp(1)
                    fyDb = objRebarSizeToDb.Item(barSize)

                    If k < 2 Then
                    ' 上層筋
                        psiT = 1.3
                    Else
                    ' 下層筋
                        psiT = 1
                    End If

                    ' 由於詳細計算法沒有收入簡算法可以修正的條件，所以到最後會比簡算法長，所以用簡算法來訂定上限。
                    ' 鋼筋尺寸修正因數
                    If fyDb >= 2 Then
                        simpleLd = 0.19 * fy_ * psiT * psiE * lambda / Sqr(fc_) * fyDb
                        psiS = 1
                    Else
                        simpleLd = 0.15 * fy_ * psiT * psiE * lambda / Sqr(fc_) * fyDb
                        psiS = 0.8
                    End If



                    ' 有加主筋之半
                    cc_ = cover + fytDb + fyDb / 2

                    ' 有加主筋之半
                    cs_ = ((width_ - fyDb * fyNum - fytDb * 2 - cover * 2) / (fyNum - 1) + fyDb) / 2

                    If cs_ <= cc_ Then
                    ' 水平劈裂

                        cb_ = cs_
                        atr_ = 2 * pi_ * fytDb ^ 2 / 4
                        ktr_ = atr_ * fyt_ / 105 / stirrupSpace / fyNum

                    Else
                    ' 垂直劈裂

                        cb_ = cc_
                        atr_ = pi_ * fytDb ^ 2 / 4
                        ktr_ = atr_ * fyt_ / 105 / stirrupSpace

                    End If

                    ldb_ = 0.28 * fy_ / Sqr(fc_) * fyDb

                    factor = psiT * psiE * psiS * lambda / ran.Min((cb_ + ktr_) / fyDb, 2.5)

                    ld_ = factor * ldb_

                    ' 乙級搭接 * 1.3
                    arrLapLength(i + k, colLapLength) = APP.RoundUp(1.3 * ran.Min(ld_, simpleLd), 0)

                    ' 換算成比例
                    arrLapLength(i + k, colLapLength) = arrLapLength(i + k, colLapLength) / length_

                End If

            Next k

        Next j

    Next i

    CalLapLength = arrLapLength

End Function


Function CalMultiLapLength(ByVal arrLapLengthRatio)
'
' 有計算 1 2 排最大值
'
' @param
' @returns

    Dim arrMultiLapLength

    ubLapLengthRatio = UBound(arrLapLengthRatio)

    ReDim arrMultiLapLength(1 To ubLapLengthRatio, varSpliceNum)

    ubMultiLapLength = UBound(arrMultiLapLength)

    varLeft = 1
    varMid = 2
    varRight = 3

    varOneThreeSpliceNum = APP.RoundUp(varSpliceNum / 3, 0)
    varTwoThreeSpliceNum = APP.RoundUp(2 * varSpliceNum / 3, 0)

    For i = 1 To ubMultiLapLength

        For j = varLeft To varRight

            ' 轉換成格數
            arrLapLengthRatio(i, j) = APP.RoundUp(arrLapLengthRatio(i, j) / (1 / varSpliceNum), 0)

        Next j

    Next i

    For i = 1 To ubMultiLapLength Step 2

        ' 這裡有一個 bug 就是要先抽離變數，否則進去 Max 型態會改變造成錯誤.
        ' 左端
        For j = 1 To varOneThreeSpliceNum
            row1 = arrLapLengthRatio(i, varLeft)
            row2 = arrLapLengthRatio(i + 1, varLeft)
            arrMultiLapLength(i, j) = APP.RoundUp(ran.Max(row1, row2), 0)
        Next j

        ' 中央
        For j = varOneThreeSpliceNum + 1 To varTwoThreeSpliceNum
            row1 = arrLapLengthRatio(i, varMid)
            row2 = arrLapLengthRatio(i + 1, varMid)
            arrMultiLapLength(i, j) = APP.RoundUp(ran.Max(row1, row2), 0)
        Next j

        ' 右端
        For j = varTwoThreeSpliceNum + 1 To varSpliceNum
            row1 = arrLapLengthRatio(i, varRight)
            row2 = arrLapLengthRatio(i + 1, varRight)
            arrMultiLapLength(i, j) = APP.RoundUp(ran.Max(row1, row2), 0)
        Next j

    Next i

    CalMultiLapLength = arrMultiLapLength


End Function


Private Function CalSplice(ByVal arrGirderRebar, ByVal arrMultiLapLength)
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'

    Dim arrSplice

    arrSplice = arrGirderRebar

    ubSmartSplice = UBound(arrSplice)

    For i = 1 To ubSmartSplice

        For j = 1 To varSpliceNum

            ' 輸出要延伸幾格
            lapLength = arrMultiLapLength(i, j)

            For k = 1 To lapLength

                If j + k <= varSpliceNum Then

                    prevBar = arrSplice(i, j + k)
                    lapBar = arrGirderRebar(i, j)

                    arrSplice(i, j + k) = ran.Max(prevBar, lapBar)

                End If

            Next k

        Next j

        For j = varSpliceNum To 1 Step -1

            ' 輸出要延伸幾格
            lapLength = arrMultiLapLength(i, j)

            For k = 1 To lapLength

                If j - k >= 1 Then

                    prevBar = arrSplice(i, j - k)
                    lapBar = arrGirderRebar(i, j)

                    arrSplice(i, j - k) = ran.Max(prevBar, lapBar)

                End If

            Next k

        Next j

    Next i

    CalSplice = arrSplice

End Function


Function CalMultiBreakRebarNormal(arrRebarTotalNumber)
'
'
' @param
' @returns

    Dim arrGirderRebarNormal

    ubRebarNumber = UBound(arrRebarTotalNumber)

    ReDim arrGirderRebarNormal(1 To ubRebarNumber, varSpliceNum)

    ubGirderRebar = UBound(arrGirderRebarNormal)

    varLeft = 1
    varMid = 2
    varRight = 3

    varOneThreeSpliceNum = APP.RoundUp(varSpliceNum / 3, 0)
    varTwoThreeSpliceNum = APP.RoundUp(2 * varSpliceNum / 3, 0)

    For i = 1 To ubGirderRebar Step 2

        ' 左端
        For j = 1 To varOneThreeSpliceNum
            arrGirderRebarNormal(i, j) = arrRebarTotalNumber(i, varLeft)
        Next j

        ' 中央
        For j = varOneThreeSpliceNum + 1 To varTwoThreeSpliceNum
            arrGirderRebarNormal(i, j) = arrRebarTotalNumber(i, varMid)
        Next j

        ' 右端
        For j = varTwoThreeSpliceNum + 1 To varSpliceNum
            arrGirderRebarNormal(i, j) = arrRebarTotalNumber(i, varRight)
        Next j

        ' 在四捨五入處取大值
        If arrRebarTotalNumber(i, varMid) > arrGirderRebarNormal(i, varOneThreeSpliceNum) Then
            arrGirderRebarNormal(i, varOneThreeSpliceNum) = arrRebarTotalNumber(i, varMid)
        End If

        If arrRebarTotalNumber(i, varRight) > arrGirderRebarNormal(i, varTwoThreeSpliceNum) Then
            arrGirderRebarNormal(i, varTwoThreeSpliceNum) = arrRebarTotalNumber(i, varRight)
        End If

    Next i

    CalMultiBreakRebarNormal = arrGirderRebarNormal


End Function

Private Function CalOptimizeResult(ByVal arrOptimized, ByVal arrInitial)
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'

    Dim arrOptimizeResult

    ubOptimized = UBound(arrOptimized)

    ReDim arrOptimizeResult(1 To ubOptimized, varSpliceNum)

    For i = 1 To ubOptimized Step 2

        For j = 1 To varSpliceNum

            arrOptimizeResult(i, j) = arrOptimized(i, j) / arrInitial(i, j)

        Next j

    Next i

    CalOptimizeResult = arrOptimizeResult

End Function


Private Function CalOptimizeNoMoreThanNormal(ByVal arrGirderRebar, ByVal arrGirderRebarNormal)
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'

    ubGirderRebar = UBound(arrGirderRebar)

    For i = 1 To ubGirderRebar Step 2

        For j = 1 To varSpliceNum

            If arrGirderRebar(i, j) > arrGirderRebarNormal(i, j) Then

                arrGirderRebar(i, j) = arrGirderRebarNormal(i, j)

            End If

        Next j

    Next i

    CalOptimizeNoMoreThanNormal = arrGirderRebar

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
' 計算完成後加上延伸長度
'
' @test:
'
'
'

    Dim time0 As Double
    Dim objRebarSizeToDb As Object

    Set ran = New UTILS_CLASS
    Set APP = Application.WorksheetFunction

    time0 = Timer

    Call ran.PerformanceVBA(True)

    Call SetGlobalVar

    Call ClearPrevOutputData

    ' 不包含標題
    arrBeam = ran.GetRangeToArray(wsBeam, 3, 1, 5, 16)

    arrTotalRebar = CalTotalRebar(arrBeam)

    arrGirderRebar = OptimizeGirderRebar(arrTotalRebar)
    arrGirderRebarNormal = CalMultiBreakRebarNormal(arrTotalRebar)

    arrGirderRebar = CalOptimizeNoMoreThanNormal(arrGirderRebar, arrGirderRebarNormal)

    arrLapLengthRatio = CalLapLength(arrBeam)
    arrMultiLapLength = CalMultiLapLength(arrLapLengthRatio)

    arrSmartSplice = CalSplice(arrGirderRebar, arrMultiLapLength)
    arrNormalSplice = CalSplice(arrGirderRebarNormal, arrMultiLapLength)

    ' arrSmartSplice = OptimizeGirderRebar(arrTotalRebar)
    ' arrNormalSplice = CalMultiBreakRebarNormal(arrTotalRebar)

    arrOptimizeResult =  CalOptimizeResult(arrSmartSplice, arrNormalSplice)

    Call PrintResult(arrSmartSplice, 3)
    Call PrintResult(arrNormalSplice, varSpliceNum + 3 + 1)
    Call PrintResult(arrOptimizeResult, 2 * varSpliceNum + 3 + 2)
    wsResult.Cells(2, 2) = APP.Average(arrOptimizeResult)

    Call ran.FontSetting(wsResult)
    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTimeVBA(time0)

End Sub
