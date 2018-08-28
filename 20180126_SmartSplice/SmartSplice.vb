Public Const varSpliceNum = 21

Public ran As UTILS_CLASS
Public APP
Public wsBeam As Worksheet
Public wsResult As Worksheet
Public wsErr As Worksheet
Public varErrNum As Integer
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
' Scan Info 是否為空
'

    ' global var
    Set wsBeam = Worksheets("大梁配筋")
    Set wsResult = Worksheets("最佳化斷筋點")
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


Function PrintErr(ByVal msg)
'
' 列印出錯誤.
'
' @param {String} [msg] 錯誤訊息.
'

    With wsErr
        .Cells(varErrNum, 2) = varErrNum - 1
        .Cells(varErrNum, 2) = msg
        varErrNum = varErrNum + 1
    End With

End Function


Function ClearPrevOutputData()
'
' 清空前次輸出的資料.
'

    wsResult.Cells.Clear

End Function


Function GetRebar1stNum(ByVal arrBeam)
'
' 計算上下排總支數
'
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrRebar1stNum] 總支數，列數與 arrBeam 對齊，行數分左中右.
'

    Dim arrRebar1stNum() As Integer
    ReDim arrRebar1stNum(1 To UBound(arrBeam), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrBeam)
    colStart = 6
    colEnd = 8

    ' 一二排相加
    For i = rowStart To rowEnd Step 4
        For j = colStart To colEnd

            colRebarTotalNum = j - 5

            ' 計算上下排總支數
            arrRebar1stNum(i, colRebarTotalNum) = Int(Split(arrBeam(i, j), "-")(0))
            arrRebar1stNum(i + 2, colRebarTotalNum) = Int(Split(arrBeam(i + 3, j), "-")(0))

        Next
    Next

    GetRebar1stNum = arrRebar1stNum

End Function


Function GetRebarTotalNum(ByVal arrBeam)
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

    GetRebarTotalNum = arrRebarTotalNum

End Function


Function GetRebarTotalArea(ByVal arrBeam)
'
' Scan 左中右號數相等
' Scan 上下排號數相等
' 計算上下排總鋼筋量
'
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrRebarTotalArea] 總鋼筋量，列數與 arrBeam 對齊，行數分左中右.
'

    Dim arrRebarTotalArea() As Double
    ReDim arrRebarTotalArea(1 To UBound(arrBeam), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrBeam)
    colLeft = 6
    colRight = 8

    ' 確認鋼筋號數左中右相等
    ' 上排
    For i = 1 To rowEnd Step 4

        rebarLeft = Split(arrBeam(i, 6), "-")(1)
        rebarMid = Split(arrBeam(i, 7), "-")(1)
        rebarRight = Split(arrBeam(i, 8), "-")(1)

        If rebarLeft <> rebarMid Or rebarLeft <> rebarRight Then
            MsgBox "第" & i & " 列鋼筋左中右號數不相等", vbOKOnly, "Error"
        End If

    Next i

    ' 確認鋼筋號數左中右相等
    ' 下排
    For i = 4 To rowEnd Step 4

        rebarLeft = Split(arrBeam(i, 6), "-")(1)
        rebarMid = Split(arrBeam(i, 7), "-")(1)
        rebarRight = Split(arrBeam(i, 8), "-")(1)

        If rebarLeft <> rebarMid Or rebarLeft <> rebarRight Then
            MsgBox "第" & i & " 列鋼筋左中右號數不相等", vbOKOnly, "Error"
        End If

    Next i

    ' 一二排相加
    For i = rowStart To rowEnd Step 4
        For j = colLeft To colRight

            colRebarTotalArea = j - 5

            top_ = i
            bot_ = i + 2

            rebarTop1st = Split(arrBeam(top_, j), "-")
            rebarTop2nd = Split(arrBeam(top_ + 1, j), "-")
            rebarBot2nd = Split(arrBeam(bot_, j), "-")
            rebarBot1st = Split(arrBeam(bot_ + 1, j), "-")

            rebarTop1stNum = Int(rebarTop1st(0))
            rebarTop2ndNum = Int(rebarTop2nd(0))
            rebarBot2ndNum = Int(rebarBot2nd(0))
            rebarBot1stNum = Int(rebarBot1st(0))

            rebarTop1stSize = rebarTop1st(1)
            rebarBot1stSize = rebarBot1st(1)

            ' 由於可能會沒有第二排的鋼筋號數，所以在 IF 裡面才取
            ' rebarTop2ndSize = rebarTop2nd(1)
            ' rebarBot2ndSize = rebarBot2nd(1)

            ' 判斷第二排是否有鋼筋
            If rebarTop2ndNum = 0 Then

                ' 第一排鋼筋量
                arrRebarTotalArea(top_, colRebarTotalArea) = rebarTop1stNum * objRebarSizeToArea.Item(rebarTop1stSize)

            ' 第二排有鋼筋的話，確定第一排與第二排的號數相同
            ElseIf rebarTop1stSize = rebarTop2nd(1) Then

                ' 第一排加第二排鋼筋量
                arrRebarTotalArea(top_, colRebarTotalArea) = (rebarTop1stNum + rebarTop2ndNum) * objRebarSizeToArea.Item(rebarTop1stSize)

            ' 有鋼筋，但號數不同，則 ERROR
            Else

                MsgBox "第" & top_ & " 列鋼筋第一排與第二排號數不相等", vbOKOnly, "Error"

            End If

            ' 判斷第二排是否有鋼筋
            If rebarBot2ndNum = 0 Then

                ' 第一排鋼筋量
                arrRebarTotalArea(bot_, colRebarTotalArea) = rebarBot1stNum * objRebarSizeToArea.Item(rebarBot1stSize)

            ' 第二排有鋼筋的話，確定第一排與第二排的號數相同
            ElseIf rebarBot1stSize = rebarBot2nd(1) Then

                ' 第一排加第二排鋼筋量
                arrRebarTotalArea(bot_, colRebarTotalArea) = (rebarBot1stNum + rebarBot2ndNum) * objRebarSizeToArea.Item(rebarBot1stSize)

            ' 有鋼筋，但號數不同，則 ERROR
            Else

                MsgBox "第" & bot_ & " 列鋼筋第一排與第二排號數不相等", vbOKOnly, "Error"

            End If



        Next
    Next

    GetRebarTotalArea = arrRebarTotalArea

End Function


Function CalGravityDemand(ByVal arrBeam)
'
' 計算重力所需要的鋼筋量
' Scan 箍筋號數左中右相等
'
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrRebarTotalNum] 總支數，列數與 arrBeam 對齊，行數分左中右.
'

    Dim arrGravity() As Double
    ReDim arrGravity(1 To UBound(arrBeam), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrBeam)

    For i = rowStart To rowEnd Step 4

        top_ = i
        bot_ = i + 2

        storey = arrBeam(i, 1)
        h = arrBeam(i, 4) ' cm
        span = arrBeam(i, 13) ' cm

        stirrupLeft = "#" & Split(Split(arrBeam(i, 10), "@")(0), "#")(1)
        stirrupMid = "#" & Split(Split(arrBeam(i, 11), "@")(0), "#")(1)
        stirrupRight = "#" & Split(Split(arrBeam(i, 12), "@")(0), "#")(1)

        If stirrupLeft <> stirrupMid Or stirrupLeft <> stirrupRight Then
            PrintErr "第" & i & " 列箍筋左中右號數不相等"
        End If

        stirrupSize = stirrupLeft
        fytDb = objRebarSizeToDb.Item(stirrupSize) ' cm

        barSizeTop = Split(arrBeam(i, 6), "-")(1)
        barSizeBot = Split(arrBeam(i + 3, 6), "-")(1)
        fyDbTop = objRebarSizeToDb.Item(barSizeTop) ' cm
        fyDbBot = objRebarSizeToDb.Item(barSizeBot) ' cm

        fy_ = objStoryToFy.Item(storey) ' kgf/cm^2
        fyt_ = objStoryToFyt.Item(storey) ' kgf/cm^2
        fc_ = objStoryToFc.Item(storey) ' kgf/cm^2
        SDL = objStoryToSDL.Item(storey) * 1000 / 10000 ' kgf/cm^2
        LL = objStoryToLL.Item(storey) * 1000 / 10000 ' kgf/cm^2
        band = objStoryToBand.Item(storey) * 100 ' cm
        slab = objStoryToSlab.Item(storey) * 100 ' cm
        cover = objStoryToCover.Item(storey) ' cm

        ' 鋼筋混凝土單位重 2.4 tf/m^3 = 2.4 * 1000 kgf / 1000000 cm^3
        ' kgf * cm
        mn_top = 1 / 24 * (0.9 * ((SDL + (2.4 * 0.001) * slab) * band)) * (span ^ 2)
        mn_bot = 1 / 8 * (1.2 * ((SDL + (2.4 * 0.001) * slab) * band) + 1.6 * (LL * band)) * (span ^ 2)

        ' 轉換成鋼筋量
        ' cm^2
        as_top = mn_top / (fy_ * 0.9 * (h - cover - fytDb - 1.5 * fyDbTop))
        as_bot = mn_bot / (fy_ * 0.9 * (h - cover - fytDb - 1.5 * fyDbBot))

        arrGravity(i, 1) = as_top

        ' 上層中央
        arrGravity(i, 2) = as_top * 2

        arrGravity(i, 3) = as_top

        arrGravity(i + 2, 2) = as_bot

    Next

    CalGravityDemand = arrGravity

End Function


Function OptimizeMultiRebar(ByVal arrBeam, ByVal arrRebarTotalArea, ByVal arrGravity, ByVal arrNormalSplice)
'
' 上層筋由耐震控制.
' 下層筋由重力與耐震共同控制.
' FIXME: 隱含了鋼筋必須相同的限制，思考要不要轉換成鋼筋量
' FIXME: 演算法具有問題
' TODO: 可能要做很多個判斷式了
'
' @param {Array} [arrRebarTotalArea] 總鋼筋量，列數與 arrBeam 對齊，行數分左中右.
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrGirderMultiRebar] 依據演算法的配筋，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
'

    Dim arrGirderMultiRebar() As Double

    ubRebarTotalNum = UBound(arrRebarTotalArea)

    ReDim arrGirderMultiRebar(1 To ubRebarTotalNum, 1 To varSpliceNum)

    ubGirderMultiRebar = UBound(arrGirderMultiRebar)

    varLeft = 1
    varMid = 2
    varRight = 3

    ' 一半的地方
    varHalfOfSpliceNum = (varSpliceNum / 2)

    ' 遞減的斜率
    slope_ = 1 / varHalfOfSpliceNum

    ' 上層筋由耐震控制.
    For i = 1 To ubGirderMultiRebar Step 4

        If arrRebarTotalArea(i, varLeft) > arrRebarTotalArea(i, varMid) And arrRebarTotalArea(i, varRight) > arrRebarTotalArea(i, varMid) Then

            rebar1stSize = Split(arrBeam(i, 6), "-")(1)
            area_ = objRebarSizeToArea.Item(rebar1stSize)

            ' 地震力需求
            ' 總鋼筋量 - 重力
            EQLeft = arrRebarTotalArea(i, varLeft) - arrGravity(i, varLeft)

            ratio = 1

            ' 左端到中央
            For j = 1 To varHalfOfSpliceNum

                ' 重力需求
                gravityLeft = - (1 - ratio ^ 2) * (arrGravity(i, varLeft) + arrGravity(i, varMid)) + arrGravity(i, varLeft)

                ' 當重力 > 0 才加
                ' 和 2 支取大值
                If gravityLeft > 0 Then
                    arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max((ratio * EQLeft + gravityLeft) / area_, 2))
                Else
                    arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max((ratio * EQLeft) / area_, 2))
                End If

                ratio = ratio - slope_

            Next j

            ' 地震力需求
            ' 總鋼筋量 - 重力
            EQRight = arrRebarTotalArea(i, varRight) - arrGravity(i, varRight)

            ratio = 1

            ' 右端到中央
            For j = varSpliceNum To Fix(varHalfOfSpliceNum) + 1 Step -1

                ' 重力需求
                gravityRight = - (1 - ratio ^ 2) * (arrGravity(i, varRight) + arrGravity(i, varMid)) + arrGravity(i, varRight)

                ' 當重力 > 0 才加
                ' 和 2 支取大值
                If gravityRight > 0 Then
                    arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max((ratio * EQRight + gravityRight) / area_, 2))
                Else
                    arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max((ratio * EQRight) / area_, 2))
                End If

                ratio = ratio - slope_

            Next j

        Else

            ' Hack: 直接還原成原本的配筋，不做更動，雖然接下來會經過延伸長度，但還會減少下來。
            For j = 1 To varSpliceNum
                arrGirderMultiRebar(i, j) = arrNormalSplice(i, j)
            Next j

        End If

    Next i

    ' 下層筋由重力與耐震共同控制.
    For i = 3 To ubGirderMultiRebar Step 4

        ' 地震力需求
        ' 總鋼筋量
        EQLeft = arrRebarTotalArea(i, varLeft)
        EQRight = arrRebarTotalArea(i, varRight)

        If EQLeft > arrRebarTotalArea(i, varMid) And EQRight > arrRebarTotalArea(i, varMid) Then

            ' 下層筋的鋼筋面積
            rebar1stSize = Split(arrBeam(i + 1, 6), "-")(1)
            area_ = objRebarSizeToArea.Item(rebar1stSize)

            If arrRebarTotalArea(i, varMid) > arrGravity(i, varMid) Then

                ' 左端到中央遞減
                ratio = 1
                For j = 1 To varHalfOfSpliceNum

                    ' 重力需求
                    gravityLeft = (1 - ratio ^ 2) * arrGravity(i, varMid)

                    arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max((ratio * EQLeft + gravityLeft) / area_, 2))
                    ' arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max(ratio * (EQLeft / area_ - arrRebarTotalArea(i, varMid) / area_) + arrRebarTotalArea(i, varMid) / area_, 2))
                    ratio = ratio - slope_
                Next j

                ' 右端到中央遞減
                ratio = 1
                For j = varSpliceNum To Fix(varHalfOfSpliceNum) + 1 Step -1

                    ' 重力需求
                    gravityRight = (1 - ratio ^ 2) * arrGravity(i, varMid)

                    arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max((ratio * EQRight + gravityRight) / area_, 2))
                    ' arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max(ratio * (EQRight / area_ - arrRebarTotalArea(i, varMid) / area_) + arrRebarTotalArea(i, varMid) / area_, 2))
                    ratio = ratio - slope_
                Next j

            Else

                ' 左端到中央遞減
                ratio = 1
                For j = 1 To varHalfOfSpliceNum

                    ' 耐震、重力、2 支取大值
                            ' arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max(ratio * EQLeft / area_, (1 - ratio ^ 2) * arrRebarTotalArea(i, varMid) / area_, 2))
                            arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max((ratio * (EQLeft - arrRebarTotalArea(i, varMid)) + arrRebarTotalArea(i, varMid)) / area_, 2))

                    ratio = ratio - slope_

                Next j

                        ' 右端到中央遞減
                ratio = 1
                For j = varSpliceNum To Fix(varHalfOfSpliceNum) + 1 Step -1

                    ' 耐震、重力、2 支取大值
                            ' arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max(ratio * EQRight / area_, (1 - ratio ^ 2) * arrRebarTotalArea(i, varMid) / area_, 2))
                            arrGirderMultiRebar(i, j) = ran.RoundUp(ran.Max((ratio * (EQRight - arrRebarTotalArea(i, varMid)) + arrRebarTotalArea(i, varMid)) / area_, 2))

                    ratio = ratio - slope_

                        Next j

            End If

        Else

            ' Hack: 直接還原成原本的配筋，不做更動，雖然接下來會經過延伸長度，但還會減少下來。
            For j = 1 To varSpliceNum
                arrGirderMultiRebar(i, j) = arrNormalSplice(i, j)
            Next j

        End If

    Next i

    OptimizeMultiRebar = arrGirderMultiRebar


End Function


Function CalLapLength(ByVal arrBeam, ByVal arrRebar1stNum, ByVal arrMultiRebar)
'
' TODO: 可以做優化，如果算過了就不用再算一次.
' 搭接長度還是延伸長度，是延伸長度。
' 計算不同主筋的搭接長度.
'
' @param {Array} [arrBeam] RCAD 輸出資料.
' @return {Array} [arrLapLength] 回傳精算法的搭接長度比例，列數與 arrBeam 對齊，行數分左中右.
'

    Dim arrLapLength() As Double

    pi_ = 3.1415926

    ' 鋼筋塗布修正因數
    ' 未塗布鋼筋
    psiE = 1

    ' 混凝土單位重之修正因數
    ' 於常重混凝土內之鋼筋
    lambda = 1

    ubBeam = UBound(arrBeam)

    ReDim arrLapLength(1 To ubBeam, 1 To varSpliceNum)

    ubLapLength = UBound(arrLapLength)

    ' loop 全部
    For i = 1 To ubLapLength Step 4

        storey = arrBeam(i, 1)
        width_ = arrBeam(i, 3)
        length_ = arrBeam(i, 13)

        fy_ = objStoryToFy.Item(storey)
        fyt_ = objStoryToFyt.Item(storey)
        fc_ = objStoryToFc.Item(storey)
        cover = objStoryToCover.Item(storey)

        ' 抽取一個出來看看是否有該 storey
        If IsEmpty(fy_) Then
            MsgBox "請確認 " & storey & " 是否存在於 General Information", vbOKOnly, "Error"
        End If

        ' loop 上下排
        ' j = 0 => 上層第一排
        ' 這裡不另外做一個 function 的原因是，傳太多參數感覺也會亂，所以用兩次 for loop 來解決上下排的問題
        For j = 1 To 2

            If j = 1 Then
                ' 上排
                row_ = i
                rowBeam = i
            Else
                ' 下排
                row_ = i + 2
                rowBeam = i + 3
            End If

            maxRebarCapacity = APP.Max(arrRebar1stNum(row_, 1), arrRebar1stNum(row_, 2), arrRebar1stNum(row_, 3))

            For col_ = 1 To varSpliceNum

                If col_ <= varSpliceNum / 4 Then
                    ' 左
                    colStirrup = 10
                ElseIf col_ >= 3 * varSpliceNum / 4 Then
                    ' 右
                    colStirrup = 12
                Else
                    ' 中
                    colStirrup = 11
                End If

                ' 箍筋
                ' 避免有雙箍的狀況
                tmp = Split(arrBeam(i, colStirrup), "@")
                stirrupSpace = Int(tmp(1))
                tmp = Split(tmp(0), "#")
                stirrupSize = "#" & tmp(1)

                fytDb = objRebarSizeToDb.Item(stirrupSize)



                tmp = Split(arrBeam(rowBeam, 6), "-")
                ' fyNum = Int(tmp(0))

                ' 只比可容納多出一支，由於第二排不能只排一支，所以要扣 2 支
                If arrMultiRebar(row_, col_) - maxRebarCapacity = 1 Then
                    fyNum = arrMultiRebar(row_, col_) - 2
                Else
                    fyNum = APP.Min(arrMultiRebar(row_, col_), maxRebarCapacity)
                End If

                ' 看主筋支數有幾根
                ' 0 的話代表沒有配筋，所以搭接長度也為 0
                ' If fyNum = 0 Then
                '     arrLapLength(i + j, colLapLength) = 0

                ' Else
                ' 之所以這裡才取 tmp(1)，是因為如果 fyNum = 0，會沒有 tmp(1)
                barSize = tmp(1)
                fyDb = objRebarSizeToDb.Item(barSize)

                ' 鋼筋位置修正因數
                If j = 1 Then
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
                ' 搭接長度 / 梁長 * 格數
                arrLapLength(row_, col_) = ran.RoundUp(ran.Min(ld_, simpleLd) / length_ * varSpliceNum)

                ' End If

            Next col_

        Next j

    Next i

    CalLapLength = arrLapLength

End Function


' Function CalMultiLapLength(ByVal arrLapLength)
' '
' ' ratio => 格數
' ' 左中右 => multi
' ' 1 2 排取大值
' '
' ' @param {Array} [arrLapLength] 精算法的搭接長度比例，列數與 arrBeam 對齊，行數分左中右.
' ' @return {Array} [arrMultiLapLength] 回傳精算法的搭接長度格數，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.

'     Dim arrMultiLapLength() As Double

'     ubLapLength = UBound(arrLapLength)

'     ReDim arrMultiLapLength(1 To ubLapLength, 1 To varSpliceNum)

'     ubMultiLapLength = UBound(arrMultiLapLength)

'     varLeft = 1
'     varMid = 2
'     varRight = 3

'     varOneThreeSpliceNum = ran.RoundUp(varSpliceNum / 3)
'     varTwoThreeSpliceNum = ran.RoundUp(2 * varSpliceNum / 3)

'     For i = 1 To ubMultiLapLength

'         For j = varLeft To varRight

'             ' 轉換成格數
'             arrLapLength(i, j) = ran.RoundUp(arrLapLength(i, j) * varSpliceNum)

'         Next j

'     Next i

'     For i = 1 To ubMultiLapLength Step 2

'         ' 這裡有一個 bug 就是要先抽離變數，否則進去 Max 型態會改變造成錯誤.
'         ' 左端
'         For j = 1 To varOneThreeSpliceNum
'             row1 = arrLapLength(i, varLeft)
'             row2 = arrLapLength(i + 1, varLeft)
'             arrMultiLapLength(i, j) = ran.RoundUp(ran.Max(row1, row2))
'         Next j

'         ' 中央
'         For j = varOneThreeSpliceNum + 1 To varTwoThreeSpliceNum
'             row1 = arrLapLength(i, varMid)
'             row2 = arrLapLength(i + 1, varMid)
'             arrMultiLapLength(i, j) = ran.RoundUp(ran.Max(row1, row2))
'         Next j

'         ' 右端
'         For j = varTwoThreeSpliceNum + 1 To varSpliceNum
'             row1 = arrLapLength(i, varRight)
'             row2 = arrLapLength(i + 1, varRight)
'             arrMultiLapLength(i, j) = ran.RoundUp(ran.Max(row1, row2))
'         Next j

'     Next i

'     CalMultiLapLength = arrMultiLapLength


' End Function


Function CalSmartSplice(ByVal arrMultiRebar, ByVal arrLapLength)
'
' 斷筋點 + 延伸長度.
'
' @param {Array} [arrMultiRebar] 依據演算法的配筋，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
' @param {Array} [arrLapLength] 精算法的搭接長度格數，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
' @return {Array} [arrSplice] 回傳加上延伸長度的斷筋點，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
'

    Dim arrSplice() As Double

    arrSplice = arrMultiRebar

    ubSmartSplice = UBound(arrSplice)

    For i = 1 To ubSmartSplice Step 2

        ' 從左至右
        For j = 1 To varSpliceNum

            ' 要延伸幾格
            lapLength = arrLapLength(i, j)

            ' 包含自己原本的長度，因為需求是從端點開始的
            For k = 1 To lapLength - 1

                If j + k <= varSpliceNum Then

                    prevBar = arrSplice(i, j + k)
                    lapBar = arrMultiRebar(i, j)

                    arrSplice(i, j + k) = ran.Max(prevBar, lapBar)

                End If

            Next k

        Next j

        ' 從右至左
        For j = varSpliceNum To 1 Step -1

            ' 輸出要延伸幾格
            lapLength = arrLapLength(i, j)

            ' 包含自己原本的長度，因為需求是從端點開始的
            For k = 1 To lapLength - 1

                If j - k >= 1 Then

                    prevBar = arrSplice(i, j - k)
                    lapBar = arrMultiRebar(i, j)

                    arrSplice(i, j - k) = ran.Max(prevBar, lapBar)

                End If

            Next k

        Next j

    Next i

    CalSmartSplice = arrSplice

End Function


Function CalNormalSplice(ByVal arrRebarTotalNum)
'
' 原始配筋
' 分成 1/3 1/3 1/3
'
' @param {Array} [arrRebarTotalNum] 總支數，列數與 arrBeam 對齊，行數分左中右.
' @return {Array} [arrNormalSplice] 依據 arrRebarTotalNum，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
'

    Dim arrNormalSplice() As Double

    ubRebarNumber = UBound(arrRebarTotalNum)

    ReDim arrNormalSplice(1 To ubRebarNumber, 1 To varSpliceNum)

    ubGirderRebar = UBound(arrNormalSplice)

    varLeft = 1
    varMid = 2
    varRight = 3

    varOneThreeSpliceNum = (varSpliceNum / 3)
    varTwoThreeSpliceNum = (2 * varSpliceNum / 3)

    varOneFiveSpliceNum = (varSpliceNum / 5)
    varFourFiveSpliceNum = (4 * varSpliceNum / 5)

    ' 上層
    For i = 1 To ubGirderRebar Step 4

        ' 左端
        For j = 1 To varOneThreeSpliceNum
            arrNormalSplice(i, j) = arrRebarTotalNum(i, varLeft)
        Next j

        ' 中央
        ' 中央通常會比較少，取保守由兩端佔滿 1/3 2/3 處
        For j = Fix(varOneThreeSpliceNum) + 1 To varTwoThreeSpliceNum
            arrNormalSplice(i, j) = arrRebarTotalNum(i, varMid)
        Next j

        ' 右端
        For j = Fix(varTwoThreeSpliceNum) + 1 To varSpliceNum
            arrNormalSplice(i, j) = arrRebarTotalNum(i, varRight)
        Next j

    Next i

    ' 下層
    For i = 3 To ubGirderRebar Step 4

        ' 中央
        ' 先填滿全部都是中央
        For j = 1 To varSpliceNum
            arrNormalSplice(i, j) = arrRebarTotalNum(i, varMid)
        Next j

        ' 左端
        If arrRebarTotalNum(i, varLeft) < arrRebarTotalNum(i, varMid) Then

            ' 左端比較少
            For j = 1 To varOneFiveSpliceNum
                arrNormalSplice(i, j) = arrRebarTotalNum(i, varLeft)
            Next j

        Else

            ' 左端比較多
            For j = 1 To ran.RoundUp(varOneThreeSpliceNum)
                arrNormalSplice(i, j) = arrRebarTotalNum(i, varLeft)
            Next j

        End If

        ' 右端
        If arrRebarTotalNum(i, varRight) < arrRebarTotalNum(i, varMid) Then

            ' 右端比較少
            For j = varSpliceNum To Fix(varFourFiveSpliceNum) + 1 Step -1
                arrNormalSplice(i, j) = arrRebarTotalNum(i, varRight)
            Next j

        Else

            ' 右端比較多
            For j = varSpliceNum To Fix(varTwoThreeSpliceNum) + 1 Step -1
                arrNormalSplice(i, j) = arrRebarTotalNum(i, varRight)
            Next j

        End If

    Next i

    CalNormalSplice = arrNormalSplice


End Function

Function CalOptimizeResult(ByVal arrOptimized, ByVal arrInitial) As Double
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


Function CalOptimizeNoMoreThanNormal(ByVal arrSmartSplice, ByVal arrNormalSplice)
'
' 最佳化的結果不應該超過初始的.
' 如果大於初始 => 最佳化 = 初始
'
' @param {Array} [arrSmartSplice] 依據演算法的配筋，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
' @param {Array} [arrNormalSplice] 依據 arrRebarTotalNum，列數與 arrBeam 對齊，行數由 varSpliceNum 控制.
' @return {Array} [arrSmartSplice] 回傳不大於初始的最佳化配筋.
'

    ubGirderMultiRebar = UBound(arrSmartSplice)

    For i = 1 To ubGirderMultiRebar Step 2

        For j = 1 To varSpliceNum

            If arrSmartSplice(i, j) > arrNormalSplice(i, j) Then

                arrSmartSplice(i, j) = arrNormalSplice(i, j)

            End If

        Next j

    Next i

    CalOptimizeNoMoreThanNormal = arrSmartSplice

End Function


Function ThreePoints(ByVal arrBeam, ByVal arrSmartSplice)
'
' 從無限多點限縮到三個點.
'
' @since 1.0.0
' @param {Array} [arrSmartSplice] 最佳化配筋.
' @return {type} [name] descrip.
'

    ubSmartSplice = UBound(arrSmartSplice)

    ubLeft = ran.RoundUp(0.15 * varSpliceNum)
    lbMid = Fix(0.4 * varSpliceNum)
    ubMid = ran.RoundUp(0.6 * varSpliceNum)
    lbRight = Fix(0.85 * varSpliceNum)

    combo = (lbMid - ubLeft + 1) * (lbRight - ubMid + 1)

    Dim arrThreePoints() As Double
    Dim arrComboUsage
    Dim arrCombo
    ReDim arrThreePoints(1 To ubSmartSplice, 1 To 6)
    ReDim arrCombo(1 To combo, 1 To 6)
    ReDim arrComboUsage(1 To combo)

    For row_ = 1 To ubSmartSplice Step 4

        span = arrBeam(row_, 13)

        ' 每四個迴圈中，迴圈其中兩排，來方便取得梁長
        For rowRebar = row_ To row_ + 2 Step 2

            i = 1

            ' 左邊先決定
            For colLeft = ubLeft To lbMid

                ' 右邊再決定
                For colRight = ubMid To lbRight

                    midMaxRebar = 0

                    ' 中間取最大值
                    For colMid = colLeft + 1 To colRight - 1

                        midMaxRebar = ran.Max(midMaxRebar, arrSmartSplice(rowRebar, colMid) * 1)

                    Next colMid

                    arrCombo(i, 1) = arrSmartSplice(rowRebar, 1)
                    arrCombo(i, 4) = colLeft / varSpliceNum * span

                    arrCombo(i, 2) = midMaxRebar
                    arrCombo(i, 5) = (colRight - colLeft - 1) / varSpliceNum * span

                    arrCombo(i, 3) = arrSmartSplice(rowRebar, varSpliceNum)
                    arrCombo(i, 6) = (varSpliceNum - colRight + 1) / varSpliceNum * span

                    leftUsage = arrCombo(i, 1) * arrCombo(i, 4)

                    midUsage = arrCombo(i, 2) * arrCombo(i, 5)

                    rightUsage = arrCombo(i, 3) * arrCombo(i, 6)

                    arrComboUsage(i) = leftUsage + midUsage + rightUsage

                    i = i + 1

                Next colRight

            Next colLeft

            arrComboUsageMin = APP.Min(arrComboUsage)

            comboUsageMinIndex = APP.Match(arrComboUsageMin, arrComboUsage, 0)

            For i = 1 To 6
                arrThreePoints(rowRebar, i) = arrCombo(comboUsageMinIndex, i)
            Next i

        Next rowRebar

    Next row_

    ThreePoints = arrThreePoints

End Function


Function ConvertThreePoints(ByVal arrThreePoints)
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'
    ubThreePoints = UBound(arrThreePoints)

    Dim arrMultiThreePoints() As Double
    ReDim arrMultiThreePoints(1 To ubThreePoints, 1 To varSpliceNum)

    For i = 1 To ubThreePoints Step 2

        span = arrThreePoints(i, 4) + arrThreePoints(i, 5) + arrThreePoints(i, 6)

        leftLength = arrThreePoints(i, 4) / span * varSpliceNum
        midLength = arrThreePoints(i, 5) / span * varSpliceNum
        rightLength = arrThreePoints(i, 6) / span * varSpliceNum

        For j = 1 To varSpliceNum

            If j <= leftLength Then

                ' 左
                arrMultiThreePoints(i, j) = arrThreePoints(i, 1)

            ElseIf j > leftLength + midLength Then

                ' 右
                arrMultiThreePoints(i, j) = arrThreePoints(i, 3)

            Else

                ' 中
                arrMultiThreePoints(i, j) = arrThreePoints(i, 2)

            End If

        Next j

    Next i

    ConvertThreePoints = arrMultiThreePoints


End Function


Function PrintResult(ByVal arrResult, ByVal colStart, ByVal strTitle)
'
' 列印出最佳化結果
' 隱含著從 0 開始
'
' @param {Array} [arrResult] 需要 print 出的陣列.
' @param {Array} [colStart] 從哪一列開始.
'

    rowStart = 3

    rowEnd = rowStart + UBound(arrResult, 1) - 1
    colEnd = colStart + UBound(arrResult, 2) - 1

    With wsResult
        .Cells(2, colStart) = strTitle
        .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd)) = arrResult
    End With

    ' 格式化條件
    ' For i = rowStart To rowEnd Step 2
    '     With wsResult.Range(wsResult.Cells(i, colStart), wsResult.Cells(i, colEnd))
    '         .FormatConditions.AddColorScale ColorScaleType:=3
    '         .FormatConditions(.FormatConditions.Count).SetFirstPriority
    '         .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    '         .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = 8109667

    '         .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
    '         .FormatConditions(1).ColorScaleCriteria(2).value = 50
    '         .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = 8711167

    '         .FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    '         .FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = 7039480
    '     End With
    ' Next i

    colStartNext = colEnd + 2
    PrintResult = colStartNext

End Function


Private Function PrintRebarTable(ByVal arrThreePoints)
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'

    arrBeamIncludeTitle = ran.GetRangeToArray(wsBeam, 1, 1, 5, 15)

    ubBeamIncludeTitle = UBound(arrBeamIncludeTitle)

    Dim arrRebarTable
    ReDim arrRebarTable(1 To ubBeamIncludeTitle, 1 To 18)

    For i = 1 To ubBeamIncludeTitle

        For j = 1 To 5
            arrRebarTable(i, j) = arrBeamIncludeTitle(i, j)
        Next j

        For j = 9 To 15
            arrRebarTable(i, j + 3) = arrBeamIncludeTitle(i, j)
        Next j

    Next i



End Function


Sub Main()

    Set ran = New UTILS_CLASS
    Set APP = Application.WorksheetFunction

    Call ran.ExecutionTime(True)
    Call ran.PerformanceVBA(True)

    Call SetGlobalVar

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
    ' arrLapLength = CalMultiLapLength(arrLapLength)

    arrSmartSplice = CalSmartSplice(arrMultiRebar, arrLapLength)

    arrSmartSpliceModify = CalOptimizeNoMoreThanNormal(arrSmartSplice, arrNormalSplice)

    arrThreePoints = ThreePoints(arrBeam, arrSmartSpliceModify)

    arrMultiThreePoints = ConvertThreePoints(arrThreePoints)
    ' arrSmartSplice = OptimizeMultiRebar(arrRebarTotalNum)

    varOptimizeResult = CalOptimizeResult(arrSmartSpliceModify, arrNormalSplice)

    colNext = PrintResult(arrRebar1stNum, 3, "第一排支數")
    colNext = PrintResult(arrRebarTotalNum, colNext, "總支數")
    colNext = PrintResult(arrRebarTotalArea, colNext, "鋼筋量")
    colNext = PrintResult(arrNormalSplice, colNext, "初始斷筋")
    colNext = PrintResult(arrGravity, colNext, "重力曲線")
    colNext = PrintResult(arrMultiRebar, colNext, "多點斷筋")
    colNext = PrintResult(arrLapLength, colNext, "延伸長度格數")
    colNext = PrintResult(arrSmartSplice, colNext, "多點斷筋 + 延伸長度")
    colNext = PrintResult(arrSmartSpliceModify, colNext, "多點斷筋 + 延伸長度 修正")
    ' colNext = PrintResult(arrThreePoints, colNext)

    wsResult.Cells(2, 2) = varOptimizeResult

    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTime(False)

End Sub

