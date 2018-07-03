Private ran As UTILS_CLASS
Private wsBeam As Worksheet
Private wsResult As Worksheet
Private objRebarSizeToDb As Object
Private objStoryToFy As Object
Private objStoryToFyt As Object
Private objStoryToFc As Object
Private objStoryToCover As Object

Private Function SetGlobalVar()
'
' set global variable.
'
' @since 1.0.0
'

    ' global var
    Set wsBeam = Worksheets("小梁配筋")
    Set wsResult = Worksheets("最佳化斷筋點")

    Set objRebarSizeToDb = ran.CreateDictionary(ran.GetRangeToArray(Worksheets("Rebar Size"), 1, 1, 1, 10), 1, 7)

    arrInfo = ran.GetRangeToArray(Worksheets("General Information"), 2, 4, 8, 8)

    Set objStoryToFy = ran.CreateDictionary(arrInfo, 1, 2)
    Set objStoryToFyt = ran.CreateDictionary(arrInfo, 1, 3)
    Set objStoryToFc = ran.CreateDictionary(arrInfo, 1, 4)
    Set objStoryToCover = ran.CreateDictionary(arrInfo, 1, 5)

End Function


Function ClearPrevOutputData()
'
' 清空前次輸出的資料
'

    wsResult.Cells.Clear

End Function


Function CalRebarTotalNumber(arrRawData)
'
' 計算上下排總支數
'
' @param
' @returns

    Dim arrRebarNumber()
    ReDim arrRebarNumber(1 To UBound(arrRawData), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrRawData)
    colStart = 6
    colEnd = 8

    ' 一二排相加
    For i = rowStart To rowEnd Step 2
        For j = colStart To colEnd

            ' 計算上下排總支數
            arrRebarNumber(i, j - 5) = Int(Split(arrRawData(i, j), "-")(0)) + Int(Split(arrRawData(i + 1, j), "-")(0))

        Next
    Next

    CalRebarTotalNumber = arrRebarNumber

End Function


Function CalRebarMaxNumber(arrRawData)
'
' 計算單排最大支數
'
' @param
' @returns

    Dim arrRebarNumber()
    ReDim arrRebarNumber(1 To UBound(arrRawData), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrRawData)
    colStart = 6
    colEnd = 8

    ' 一二排相加
    For i = rowStart To rowEnd Step 2
        For j = colStart To colEnd

            ' 計算單排最大支數
            arrRebarNumber(i, j - 5) = ran.Max(Int(Split(arrRawData(i, j), "-")(0)), Int(Split(arrRawData(i + 1, j), "-")(0)))

        Next
    Next

    CalRebarMaxNumber = arrRebarNumber

End Function


Function CalMultiBreakPoint(arrRebarNumber)
'
'
' @param
' @returns

    Dim arrMultiBreakRebar

    ubRebarNumber = UBound(arrRebarNumber)

    ReDim arrMultiBreakRebar(1 To ubRebarNumber, 21)

    ubMultiBreakRebar = UBound(arrMultiBreakRebar)

    varLeft = 1
    varMid = 2
    varRight = 3

    For i = 1 To ubMultiBreakRebar Step 4

        arrMultiBreakRebar(i, 0) = "上層"

        ' 左端到中央
        ratio = 1
        For j = 1 To 11
            arrMultiBreakRebar(i, j) = ran.RoundUp(ran.Max(ratio * arrRebarNumber(i, varLeft), 2))
            ratio = ratio - 0.1
        Next j

        ' 中央到右端
        ratio = 0.1
        For j = 12 To 21
            arrMultiBreakRebar(i, j) = ran.RoundUp(ran.Max(ratio * arrRebarNumber(i, varRight), 2))
            ratio = ratio + 0.1
        Next j

    Next i

    For i = 3 To ubMultiBreakRebar Step 4

        arrMultiBreakRebar(i, 0) = "下層"

        ' 左端到中央
        ratio = 1
        For j = 1 To 11
            arrMultiBreakRebar(i, j) = ran.RoundUp(ran.Max(ratio * arrRebarNumber(i, varLeft), (1 - ratio ^ 2) * arrRebarNumber(i, varMid), 2))
            ratio = ratio - 0.1
        Next j

        ' 中央到右端
        ratio = 0.1
        For j = 12 To 21
            arrMultiBreakRebar(i, j) = ran.RoundUp(ran.Max(ratio * arrRebarNumber(i, varRight), (1 - ratio ^ 2) * arrRebarNumber(i, varMid), 2))
            ratio = ratio + 0.1
        Next j

    Next i

    CalMultiBreakPoint = arrMultiBreakRebar


End Function


Function PrintResult(arrResult)
'
' 列印出最佳化結果
'
' @param arrResult(Array)

    With wsResult
        rowStart = 3
        colStart = 3
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


Private Function CalLapLength(arrBeam)
'
' TODO: 可以做優化，如果算過了就不用再算一次.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
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
                    arrLapLength(i + k, colLapLength) = Fix(1.3 * ran.Min(ld_, simpleLd)) + 1

                End If

            Next k

        Next j

    Next i

    CalLapLength = arrLapLength

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

    time0 = Timer

    Call ran.PerformanceVBA(True)

    Call SetGlobalVar

    Call ClearPrevOutputData

    ' 不包含標題
    arrBeam = ran.GetRangeToArray(wsBeam, 3, 1, 5, 16)

    arrRebarTotalNumber = CalRebarTotalNumber(arrBeam)
    arrRebarMaxNumber = CalRebarMaxNumber(arrBeam)

    arrLapLength = CalLapLength(arrBeam)

    arrMultiBreakRebar = CalMultiBreakPoint(arrRebarTotalNumber)

    Call PrintResult(arrMultiBreakRebar)

    Call ran.FontSetting(wsResult)
    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTimeVBA(time0)

End Sub
