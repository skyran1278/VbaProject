Public Const varSpliceNum = 21

Public ran As UTILS_CLASS
Public APP
Public wsBeam As Worksheet
Public wsResult As Worksheet
Public objRebarSizeToDb As Object
Public objRebarSizeToArea As Object
Public objStoryToFy As Object
Public objStoryToFyt As Object
Public objStoryToFc As Object
Public objStoryToSDL As Object
Public objStoryToLL As Object
Public objStoryToBand As Object
Public objStoryToSlab As Object
Public objStoryToCover As Object

' 大梁考慮耐震
' 小梁考慮負彎矩

Function SetGlobalVar()
'
' set global variable.
'

    ' global var
    Set wsBeam = Worksheets("大梁配筋")
    Set wsResult = Worksheets("最佳化斷筋點")

    arrRebarSize = ran.GetRangeToArray(Worksheets("Rebar Size"), 1, 1, 1, 10)

    ' #3 => 0.9525cm
    Set objRebarSizeToDb = ran.CreateDictionary(arrRebarSize, 1, 7)

    ' #3 => 0.71cm^2
    Set objRebarSizeToArea = ran.CreateDictionary(arrRebarSize, 1, 10)

    arrInfo = ran.GetRangeToArray(Worksheets("General Information"), 2, 4, 4, 12)

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

    wsResult.Cells.Clear

End Function


Function CalRebarTotalArea(ByVal arrBeam)
'
' 計算上下排總鋼筋量
'
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrRebarTotalArea] 總鋼筋量，列數與 arrBeam 對齊，行數分左中右.
'

    Dim arrRebarTotalArea() As Double
    ReDim arrRebarTotalArea(1 To UBound(arrBeam), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrBeam)
    colStart = 6
    colEnd = 8

    ' 一二排相加
    For i = rowStart To rowEnd Step 2
        For j = colStart To colEnd

            colRebarTotalArea = j - 5

            rebar1st = Split(arrBeam(i, j), "-")
            rebar1stNum = Int(rebar1st(0))
            rebar1stSize = rebar1st(1)

            rebar2nd = Split(arrBeam(i + 1, j), "-")
            rebar2ndNum = Int(rebar2nd(0))
            rebar2ndSize = rebar2nd(1)

            If rebar1stSize = rebar2ndSize Then

                ' 計算鋼筋量
                arrRebarTotalArea(i, colRebarTotalArea) = (rebar1stNum + rebar2ndNum) * objRebarSizeToArea.Item(rebar1stSize)

            Else

                MsgBox "第" & i & " 列鋼筋第一排與第二排號數不相等", vbOKOnly, "Error"
            End If

        Next
    Next

    CalRebarTotalArea = arrRebarTotalArea

End Function


Function CalRebarTotalNum(ByVal arrBeam)
'
' 計算上下排總支數
'
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrRebarTotalNum] 總支數，列數與 arrBeam 對齊，行數分左中右.
'

    Dim arrRebarTotalNum() As Integer
    ReDim arrRebarTotalNum(1 To UBound(arrBeam), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrBeam)
    colStart = 6
    colEnd = 8

    ' 一二排相加
    For i = rowStart To rowEnd Step 2
        For j = colStart To colEnd

            colRebarTotalNum = j - 5

            ' 計算上下排總支數
            arrRebarTotalNum(i, colRebarTotalNum) = Int(Split(arrBeam(i, j), "-")(0)) + Int(Split(arrBeam(i + 1, j), "-")(0))

        Next
    Next

    CalRebarTotalNum = arrRebarTotalNum

End Function


Function CalGravityDemand(ByVal arrBeam)
'
' 計算重力所需要的鋼筋量
'
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrRebarTotalNum] 總支數，列數與 arrBeam 對齊，行數分左中右.
'

    Dim arrGravity() As Integer
    ReDim arrGravity(1 To UBound(arrBeam), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrBeam)

    For i = rowStart To rowEnd Step 4

        top = i
        bot = i + 2

        story = arrBeam(i, 1)
        span = arrBeam(i, 13)

        fy_ = objStoryToFy.Item(story)
        fyt_ = objStoryToFyt.Item(story)
        fc_ = objStoryToFc.Item(story)
        SDL = objStoryToSDL.Item(story)
        LL = objStoryToLL.Item(story)
        band = objStoryToBand.Item(story)
        slab = objStoryToSlab.Item(story)
        cover = objStoryToCover.Item(story)

        mn = 1 / 24 * (0.9 * ((SDL + 2.4 * slab) * band)) * (span ^ 2)




    Next

End Function


Function OptimizeGirderMultiRebar(ByVal arrBeam, ByVal arrRebarTotalNum)
'
' 上層筋由耐震控制.
' 下層筋由重力與耐震共同控制.
' FIXME: 隱含了鋼筋必須相同的限制，思考要不要轉換成鋼筋量
' FIXME: 演算法具有問題
'
' @param {Array} [arrRebarTotalNum] 總支數，列數與 arrBeam 對齊，行數分左中右.
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrGirderMultiRebar] 依據演算法的配筋，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
'

    Dim arrGirderMultiRebar

    ubRebarTotalNum = UBound(arrRebarTotalNum)

    ReDim arrGirderMultiRebar(1 To ubRebarTotalNum, varSpliceNum)

    ubGirderMultiRebar = UBound(arrGirderMultiRebar)

    varLeft = 1
    varMid = 2
    varRight = 3

    ' 一半的地方
    varHalfOfSpliceNum = APP.RoundUp(varSpliceNum / 2, 0)

    ' 遞減的斜率
    slope_ = 1 / varHalfOfSpliceNum

    ' 上層筋由耐震控制.
    For i = 1 To ubGirderMultiRebar Step 4

        arrGirderMultiRebar(i, 0) = "上層"

        ' 左端到中央
        ratio = 1
        For j = 1 To varHalfOfSpliceNum
            ' 耐震和 2 支取大值
            arrGirderMultiRebar(i, j) = APP.RoundUp(ran.Max(ratio * arrRebarTotalNum(i, varLeft), 2), 0)
            ratio = ratio - slope_
        Next j

        ' 中央到右端
        ratio = slope_
        For j = varHalfOfSpliceNum + 1 To varSpliceNum
            ' 耐震和 2 支取大值
            arrGirderMultiRebar(i, j) = APP.RoundUp(ran.Max(ratio * arrRebarTotalNum(i, varRight), 2), 0)
            ratio = ratio + slope_
        Next j

    Next i

    ' 下層筋由重力與耐震共同控制.
    For i = 3 To ubGirderMultiRebar Step 4

        arrGirderMultiRebar(i, 0) = "下層"

        ' 左端到中央
        ratio = 1
        For j = 1 To varHalfOfSpliceNum
            ' 耐震、重力、2 支取大值
            arrGirderMultiRebar(i, j) = APP.RoundUp(ran.Max(ratio * arrRebarTotalNum(i, varLeft), (1 - ratio ^ 2) * arrRebarTotalNum(i, varMid), 2), 0)
            ' arrGirderMultiRebar(i, j) = APP.RoundUp(ran.Max(ratio * (arrRebarTotalNum(i, varLeft) - arrRebarTotalNum(i, varMid)) + arrRebarTotalNum(i, varMid), 2), 0)
            ratio = ratio - slope_
        Next j

        ' 中央到右端
        ratio = slope_
        For j = varHalfOfSpliceNum + 1 To varSpliceNum
            ' 耐震、重力、2 支取大值
            arrGirderMultiRebar(i, j) = APP.RoundUp(ran.Max(ratio * arrRebarTotalNum(i, varRight), (1 - ratio ^ 2) * arrRebarTotalNum(i, varMid), 2), 0)
            ' arrGirderMultiRebar(i, j) = APP.RoundUp(ran.Max(ratio * (arrRebarTotalNum(i, varRight) - arrRebarTotalNum(i, varMid)) + arrRebarTotalNum(i, varMid), 2), 0)
            ratio = ratio + slope_
        Next j

    Next i

    OptimizeGirderMultiRebar = arrGirderMultiRebar


End Function


Function PrintResult(ByVal arrResult, ByVal colStart)
'
' 列印出最佳化結果
'
' @param {Array} [arrResult] 需要 print 出的陣列.
' @param {Array} [colStart] 從哪一列開始.
'

    With wsResult
        rowStart = 3
        rowEnd = rowStart + UBound(arrResult, 1) - 1
        colEnd = colStart + UBound(arrResult, 2)

        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = arrResult
    End With

    ' 格式化條件
    ' For i = rowStart To rowEnd
    '     With wsResult.Range(wsResult.Cells(i, colStart), wsResult.Cells(i, colEnd))
    '         .FormatConditions.AddColorScale ColorScaleType:=3
    '         .FormatConditions(.FormatConditions.Count).SetFirstPriority
    '         .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    '         .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = 8109667

    '         .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
    '         .FormatConditions(1).ColorScaleCriteria(2).Value = 50
    '         .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = 8711167

    '         .FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    '         .FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = 7039480
    '     End With
    ' Next i

End Function


Function CalLapLengthRatio(ByVal arrBeam)
'
' TODO: 可以做優化，如果算過了就不用再算一次.
' 搭接長度還是延伸長度，是延伸長度。
' 計算不同主筋的搭接長度.
'
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrLapLengthRatio] 回傳精算法的搭接長度比例，列數與 arrBeam 對齊，行數分左中右.
'

    Dim arrLapLengthRatio

    pi_ = 3.1415926

    ' 鋼筋塗布修正因數
    ' 未塗布鋼筋
    psiE = 1

    ' 混凝土單位重之修正因數
    ' 於常重混凝土內之鋼筋
    lambda = 1

    ubBeam = UBound(arrBeam)

    ReDim arrLapLengthRatio(1 To ubBeam, 1 To 3)

    ubLapLengthRatio = UBound(arrLapLengthRatio)

    ' loop 全部
    For i = 1 To ubLapLengthRatio Step 4

        story = arrBeam(i, 1)

        fy_ = objStoryToFy.Item(story)
        fyt_ = objStoryToFyt.Item(story)
        fc_ = objStoryToFc.Item(story)
        cover = objStoryToCover.Item(story)

        If IsEmpty(fy_) Then
            MsgBox "請確認 " & story & " 是否存在於 General Information", vbOKOnly, "Error"
        End If

        width_ = arrBeam(i, 3)
        length_ = arrBeam(i, 13)

        ' loop 左中右
        For j = 6 To 8

            colBar = j
            colStirrup = j + 4

            ' arrLapLengthRatio 的行數
            colLapLengthRatio = j - 5

            ' 箍筋
            ' 避免有雙箍的狀況
            tmp = Split(arrBeam(i, colStirrup), "@")
            stirrupSpace = Int(tmp(1))
            tmp = Split(tmp(0), "#")
            stirrupSize = "#" & tmp(1)

            fytDb = objRebarSizeToDb.Item(stirrupSize)

            ' loop 上下排
            ' k = 0 => 上層第一排
            For k = 0 To 3

                tmp = Split(arrBeam(i + k, colBar), "-")
                fyNum = Int(tmp(0))

                ' 看主筋支數有幾根
                ' 0 的話代表沒有配筋，所以搭接長度也為 0
                If fyNum = 0 Then
                    arrLapLengthRatio(i + k, colLapLengthRatio) = 0

                Else
                    ' 之所以這裡才取 tmp(1)，是因為如果 fyNum = 0，會沒有 tmp(1)
                    barSize = tmp(1)
                    fyDb = objRebarSizeToDb.Item(barSize)

                    ' 鋼筋位置修正因數
                    If k < 2 Then
                        ' 水平鋼筋其下混凝土一次澆置厚度大於30 cm者
                        ' 上層筋 1.3 倍
                        psiT = 1.3
                    Else
                        ' 下層筋
                        psiT = 1
                    End If

                    ' 由於詳細計算法沒有收入簡算法可以修正的條件，所以到最後會比簡算法長，所以用簡算法來訂定上限。
                    ' 鋼筋尺寸修正因數
                    If fyDb <= 2 Then
                        ' D19或較小之鋼筋及麻面鋼線
                        simpleLd = 0.15 * fy_ * psiT * psiE * lambda / Sqr(fc_) * fyDb
                        psiS = 0.8
                    Else
                        ' D22或較大之鋼筋
                        simpleLd = 0.19 * fy_ * psiT * psiE * lambda / Sqr(fc_) * fyDb
                        psiS = 1
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

                    ' 不是搭接長度，是延伸長度
                    arrLapLengthRatio(i + k, colLapLengthRatio) = APP.RoundUp(ran.Min(ld_, simpleLd), 0)

                    ' 換算成比例
                    ' 搭接長度 / 梁長
                    arrLapLengthRatio(i + k, colLapLengthRatio) = arrLapLengthRatio(i + k, colLapLengthRatio) / length_

                End If

            Next k

        Next j

    Next i

    CalLapLengthRatio = arrLapLengthRatio

End Function


Function CalMultiLapLength(ByVal arrLapLengthRatio)
'
' ratio => 格數
' 左中右 => multi
' 1 2 排取大值
'
' @param {Array} [arrLapLength] 精算法的搭接長度比例，列數與 arrBeam 對齊，行數分左中右.
' @return {Array} [arrMultiLapLength] 回傳精算法的搭接長度格數，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.

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
            arrLapLengthRatio(i, j) = APP.RoundUp(arrLapLengthRatio(i, j) * varSpliceNum, 0)

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


Function CalSplice(ByVal arrGirderMultiRebar, ByVal arrMultiLapLength)
'
' 斷筋點 + 延伸長度.
'
' @param {Array} [arrGirderMultiRebar] 依據演算法的配筋，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
' @param {Array} [arrMultiLapLength] 精算法的搭接長度格數，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
' @return {Array} [arrSplice] 回傳加上延伸長度的斷筋點，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
'

    Dim arrSplice

    arrSplice = arrGirderMultiRebar

    ubSmartSplice = UBound(arrSplice)

    For i = 1 To ubSmartSplice Step 2

        ' 從左至右
        For j = 1 To varSpliceNum

            ' 要延伸幾格
            lapLength = arrMultiLapLength(i, j)

            For k = 1 To lapLength

                If j + k <= varSpliceNum Then

                    prevBar = arrSplice(i, j + k)
                    lapBar = arrGirderMultiRebar(i, j)

                    arrSplice(i, j + k) = ran.Max(prevBar, lapBar)

                End If

            Next k

        Next j

        ' 從右至左
        For j = varSpliceNum To 1 Step -1

            ' 輸出要延伸幾格
            lapLength = arrMultiLapLength(i, j)

            For k = 1 To lapLength

                If j - k >= 1 Then

                    prevBar = arrSplice(i, j - k)
                    lapBar = arrGirderMultiRebar(i, j)

                    arrSplice(i, j - k) = ran.Max(prevBar, lapBar)

                End If

            Next k

        Next j

    Next i

    CalSplice = arrSplice

End Function


Function CalNormalGirderMultiRebar(ByVal arrRebarTotalNumber)
'
' 原始配筋
' 分成 1/3 1/3 1/3
'
' @param {Array} [arrRebarTotalNum] 總支數，列數與 arrBeam 對齊，行數分左中右.
' @return {Array} [arrNormalGirderMultiRebar] 依據 arrRebarTotalNum，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
'

    Dim arrNormalGirderMultiRebar

    ubRebarNumber = UBound(arrRebarTotalNumber)

    ReDim arrNormalGirderMultiRebar(1 To ubRebarNumber, varSpliceNum)

    ubGirderRebar = UBound(arrNormalGirderMultiRebar)

    varLeft = 1
    varMid = 2
    varRight = 3

    varOneThreeSpliceNum = APP.RoundUp(varSpliceNum / 3, 0)
    varTwoThreeSpliceNum = APP.RoundUp(2 * varSpliceNum / 3, 0)

    varOneFiveSpliceNum = APP.RoundUp(varSpliceNum / 5, 0)
    varFourFiveSpliceNum = APP.RoundUp(4 * varSpliceNum / 5, 0)

    ' 上層
    For i = 1 To ubGirderRebar Step 4

        ' 左端
        For j = 1 To varOneThreeSpliceNum
            arrNormalGirderMultiRebar(i, j) = arrRebarTotalNumber(i, varLeft)
        Next j

        ' 中央
        ' 中央通常會比較少，取保守由兩端佔滿 1/3 2/3 處
        For j = varOneThreeSpliceNum + 1 To varTwoThreeSpliceNum - 1
            arrNormalGirderMultiRebar(i, j) = arrRebarTotalNumber(i, varMid)
        Next j

        ' 右端
        For j = varTwoThreeSpliceNum To varSpliceNum
            arrNormalGirderMultiRebar(i, j) = arrRebarTotalNumber(i, varRight)
        Next j

    Next i

    ' 下層
    For i = 3 To ubGirderRebar Step 4

        ' 中央
        ' 先填滿全部都是中央
        For j = 1 To varSpliceNum
            arrNormalGirderMultiRebar(i, j) = arrRebarTotalNumber(i, varMid)
        Next j

        ' 左端
        If arrRebarTotalNumber(i, varLeft) < arrRebarTotalNumber(i, varMid) Then

            ' 左端比較少
            For j = 1 To varOneFiveSpliceNum - 1
                arrNormalGirderMultiRebar(i, j) = arrRebarTotalNumber(i, varLeft)
            Next j

        Else

            ' 左端比較多
            For j = 1 To varOneThreeSpliceNum
                arrNormalGirderMultiRebar(i, j) = arrRebarTotalNumber(i, varLeft)
            Next j

        End If

        ' 右端
        If arrRebarTotalNumber(i, varRight) < arrRebarTotalNumber(i, varMid) Then

            ' 右端比較少
            For j = varFourFiveSpliceNum + 1 To varSpliceNum
                arrNormalGirderMultiRebar(i, j) = arrRebarTotalNumber(i, varRight)
            Next j

        Else

            ' 右端比較多
            For j = varTwoThreeSpliceNum To varSpliceNum
                arrNormalGirderMultiRebar(i, j) = arrRebarTotalNumber(i, varRight)
            Next j

        End If

    Next i

    CalNormalGirderMultiRebar = arrNormalGirderMultiRebar


End Function

Function CalOptimizeResult(ByVal arrOptimized, ByVal arrInitial)
'
' 回傳最佳化結果.
' arrOptimized / arrInitial
'
' @param {Array} [arrOptimized] 最佳化過後的配筋.
' @param {Array} [arrInitial] 原始配筋.
' @return {Array} [varOptimized / varInitial] 回傳最佳化結果.
'

    ubOptimized = UBound(arrOptimized)

    varOptimized = 0
    varInitial = 0

    For i = 1 To ubOptimized Step 2

        For j = 1 To varSpliceNum

            varInitial = varInitial + arrInitial(i, j)
            varOptimized = varOptimized + arrOptimized(i, j)

        Next j

    Next i

    CalOptimizeResult = varOptimized / varInitial

End Function


Function CalOptimizeNoMoreThanNormal(ByVal arrGirderMultiRebar, ByVal arrNormalGirderMultiRebar)
'
' 最佳化的結果不應該超過初始的.
' 如果大於初始 => 最佳化 = 初始
'
' @param {Array} [arrGirderMultiRebar] 依據演算法的配筋，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
' @param {Array} [arrNormalGirderMultiRebar] 依據 arrRebarTotalNum，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
' @return {Array} [arrGirderMultiRebar] 回傳不大於初始的最佳化配筋.
'

    ubGirderMultiRebar = UBound(arrGirderMultiRebar)

    For i = 1 To ubGirderMultiRebar Step 2

        For j = 1 To varSpliceNum

            If arrGirderMultiRebar(i, j) > arrNormalGirderMultiRebar(i, j) Then

                arrGirderMultiRebar(i, j) = arrNormalGirderMultiRebar(i, j)

            End If

        Next j

    Next i

    CalOptimizeNoMoreThanNormal = arrGirderMultiRebar

End Function





Sub Main()

    Dim time0 As Double

    time0 = Timer

    Set ran = New UTILS_CLASS
    Set APP = Application.WorksheetFunction

    Call ran.PerformanceVBA(True)

    Call SetGlobalVar

    Call ClearPrevOutputData

    ' 不包含標題
    arrBeam = ran.GetRangeToArray(wsBeam, 3, 1, 5, 16)

    arrRebarTotalNum = CalRebarTotalNum(arrBeam)

    arrNormalSplice = CalNormalGirderMultiRebar(arrRebarTotalNum)

    ' arrRebarTotalArea = CalRebarTotalArea(arrBeam)

    arrGirderMultiRebar = OptimizeGirderMultiRebar(arrBeam, arrRebarTotalNum)

    arrLapLengthRatio = CalLapLengthRatio(arrBeam)
    arrMultiLapLength = CalMultiLapLength(arrLapLengthRatio)

    arrSmartSplice = CalSplice(arrGirderMultiRebar, arrMultiLapLength)

    arrSmartSplice = CalOptimizeNoMoreThanNormal(arrSmartSplice, arrNormalSplice)

    ' arrSmartSplice = OptimizeGirderMultiRebar(arrRebarTotalNum)

    varOptimizeResult = CalOptimizeResult(arrSmartSplice, arrNormalSplice)

    Call PrintResult(arrSmartSplice, 3)
    Call PrintResult(arrNormalSplice, varSpliceNum + 3 + 1)
    wsResult.Cells(2, 2) = varOptimizeResult

    Call ran.FontSetting(wsResult)
    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTimeVBA(time0)

End Sub
