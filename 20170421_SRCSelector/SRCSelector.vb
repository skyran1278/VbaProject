Dim CURVES_NAME
Dim CONTROL_COMBO


Function AutoFill()
' 公式自動填滿

    ' Worksheets("EtabsPMMCombo").Activate
    ' comboRowUsed = Cells(Rows.Count, 1).End(xlUp).row
    comboRowUsed = Worksheets("EtabsPMMCombo").Cells(Rows.Count, 1).End(xlUp).row

    ' Worksheets("PMM").Activate
    ' Range(Cells(3, 1), Cells(3, Cells(Rows.Count, 1).End(xlUp).row)).ClearContents
    ' Range(Cells(2, 1), Cells(2, 6)).AutoFill Destination:=Range(Cells(2, 1), Cells(comboRowUsed, 6))

    With Worksheets("PMM")

        .Range(.Cells(3, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).row, 7)).ClearContents
        .Range(.Cells(2, 1), .Cells(2, 6)).AutoFill Destination:=.Range(.Cells(2, 1), .Cells(comboRowUsed, 6))

    End With

End Function


Function ReadCombo()
' 讀取每個 Combo
' 回傳：combo array
' 資料格式：
' Name P M2 M3
' combo(2 To comboRowUsed, 1 To 4)

    Worksheets("PMM").Activate
    Dim combo()
    comboRowUsed = Cells(Rows.Count, 1).End(xlUp).row
    ReDim combo(2 To comboRowUsed, 1 To 4)

    ' 讀取所有的PMM
    For row = 2 To comboRowUsed

        ' Name
        combo(row, 1) = Cells(row, 1)

        ' P
        combo(row, 2) = Cells(row, 4)

        ' M2
        combo(row, 3) = Cells(row, 5)

        ' M3
        combo(row, 4) = Cells(row, 6)

    Next

    ' 沒有辦法指定 2 ~ comboRowUsed，1 ~ comboRowUsed，1 不確定會不會有問題，所以先以 for loop 迴圈為主
    ' combo = range(cells(1, 1), cells(comboRowUsed, 4))

    ReadCombo = combo()

End Function


Function ReadCurves()
' 讀取所有 curves
' 並修改全域變數 CURVES_NAME(1 To curveNumber + 1)
' 回傳：curves array
' curves(1 To curveNumber)

    Dim curves()

    ' 定義數值意義
    nameColumn = 2

    ' 讀取PMMCurve最後一列
    Worksheets("PMMCurve").Activate
    curveRowUsed = Cells(Rows.Count, 4).End(xlUp).row

    ' 統計有幾個非空白儲存格
    curveNumber = Application.WorksheetFunction.CountA(Range(Cells(2, nameColumn), Cells(curveRowUsed, nameColumn)))

    ReDim curves(1 To curveNumber)
    ReDim CURVES_NAME(1 To curveNumber + 1)

    For row = 2 To curveRowUsed

        If Cells(row, nameColumn) <> "" Then

            Index = Index + 1

            CURVES_NAME(Index) = Cells(row, nameColumn)

            curves(Index) = ReadCurve(row)

        End If


    Next

    CURVES_NAME(curveNumber + 1) = "超過所有斷面，請選擇更大的斷面！"

    ReadCurves = curves()


End Function


Function ReadCurve(row)
' 讀取單個 curve
' 排序後內插求值

' 參數：當前讀取的 curve row
' 回傳：curve array
' 資料格式：
' P M0 M45 M90
' curve(1 To 60, 3)

    Dim curve(1 To 60, 3)

    ' 讀取
    For Column = 1 To 3

        loading = Column * 4 + 1
        moment = Column * 4 + 2

        For Point = 1 To 20

            pointCumulativeNumber = pointCumulativeNumber + 1

            ' P
            curve(pointCumulativeNumber, 0) = Cells(row + Point, loading)

            ' M
            curve(pointCumulativeNumber, Column) = Cells(row + Point, moment)

        Next

    Next

    ' 以 load 排序
    Call QuickSortArray(curve, , , 0)

    ' 內插
    For moment = 1 To 3

        ' 設定第一個 0 為最小值
        For Point = 1 To 60

            ' 初始化，第一次讀到的非空值起始
            If Not IsEmpty(curve(Point, moment)) Then

                pointMin = Point
                loadMin = curve(Point, 0)
                momentMin = curve(Point, moment)

                Exit For

            End If

        Next

        ' 開始內插
        For Point = 1 To 60

            ' 非空元素
            If Not IsEmpty(curve(Point, moment)) Then

                pointMax = Point
                loadMax = curve(Point, 0)
                momentMax = curve(Point, moment)

                For pointMid = pointMin + 1 To pointMax - 1

                    ' 內插公式
                    curve(pointMid, moment) = Interpolate(loadMax, curve(pointMid, 0), loadMin, momentMax, momentMin)

                Next

                pointMin = pointMax
                loadMin = loadMax
                momentMin = momentMax

            End If

        Next

    Next

    ReadCurve = curve()

End Function


Public Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortArray(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortArray(SortArray, i, lngMax, lngColumn)

End Sub


Function Interpolate(varMax, varMid, varMin, aimsMax, aimsMin)
' 內插法
' 參數：varMax, varMid, varMin, aimsMax, aimsMin
' 回傳：aimsMid

    If varMax > varMin Then

        Interpolate = (varMid - varMin) / (varMax - varMin) * (aimsMax - aimsMin) + aimsMin

    ' 抓住 varMax = varMin 造成除以 0 的錯誤
    Else

        Interpolate = aimsMin

    End If

End Function


Function SectionSelector(combo, curves)

    Dim section()

    ' 定義
    ' combo
    comboName = 1
    loading = 2
    m2 = 3
    m3 = 4

    ' 定義
    ' curve
    p = 0
    m0 = 1
    m45 = 2
    m90 = 3

    ' 取出上限
    comboUBound = UBound(combo)
    curvesBound = UBound(curves)

    ' 計算 combo 數
    For row = 2 To comboUBound

        ' 計算 combo 數
        If combo(row, 1) <> combo(row + 1, 1) Then
            comboNumber = row - 2 + 1
            Exit For
        End If

    Next

    ' CONTROL_COMBO 大小和 combo 相同
    ReDim CONTROL_COMBO(2 To comboUBound, 1 To 1)

    ' section 為 combo 除以載重組合數
    ReDim section(2 To (comboUBound - 1) / comboNumber + 1, 3)

    sectionNumber = 2

    For row = 2 To comboUBound Step comboNumber

        ' 每一個Column（包含很多個Combo）重新初始化
        comboSelectNumber = 0
        comboRatio = 0

        ' 相同的一組
        For comboRow = row To row + comboNumber - 1

            ' 每一個 Combo 重新初始化
            ratio = 0
            curvesNumber = 1
            loadMid = combo(comboRow, loading)

            ' 循環 curves
            ' 為了跳出特定迴圈，使用 do loop
            Do While curvesNumber <= curvesBound

                curve = curves(curvesNumber)

                ' 至少要比最小的還大
                If loadMid > curve(1, p) Then

                    For Point = 1 To 60

                        ' 小於哪個 load
                        If loadMid < curve(Point, p) Then

                            loadMax = curve(Point, p)
                            loadMin = curve(Point - 1, p)

                            ' 內插
                            interM45 = Interpolate(loadMax, loadMid, loadMin, curve(Point, m45), curve(Point - 1, m45))

                            If combo(comboRow, m2) > combo(comboRow, m3) Then

                                ' 內插
                                interM0 = Interpolate(loadMax, loadMid, loadMin, curve(Point, m0), curve(Point - 1, m0))

                                If Newton(interM0, 0, interM45 / Sqr(2), interM45 / Sqr(2), combo(comboRow, m2), combo(comboRow, m3)) Then

                                    ratio = calRatio(interM0, 0, interM45 / Sqr(2), interM45 / Sqr(2), combo(comboRow, m2), combo(comboRow, m3))

                                    Exit Do

                                End If

                            Else

                                ' 內插
                                interM90 = Interpolate(loadMax, loadMid, loadMin, curve(Point, m90), curve(Point - 1, m90))

                                If Newton(0, interM90, interM45 / Sqr(2), interM45 / Sqr(2), combo(comboRow, m2), combo(comboRow, m3)) Then

                                    ratio = calRatio(0, interM90, interM45 / Sqr(2), interM45 / Sqr(2), combo(comboRow, m2), combo(comboRow, m3))

                                    Exit Do

                                End If

                            End If

                            Exit For

                        End If

                    Next Point

                End If

                curvesNumber = curvesNumber + 1

            Loop

            CONTROL_COMBO(comboRow, 1) = curvesNumber

            ' 判斷有沒有大於comboSelectNumber，有的話才寫入
            If curvesNumber > comboSelectNumber Then

                comboSelectNumber = curvesNumber
                comboRatio = ratio

            ' 如果相等的話，判斷有沒有大於Ratio，有的話才寫入
            ElseIf comboSelectNumber = curvesNumber And ratio > comboRatio Then

                comboRatio = ratio

            End If

        Next

        ' 寫入斷面資料
        section(sectionNumber, 0) = combo(row, comboName)
        section(sectionNumber, 1) = comboSelectNumber
        section(sectionNumber, 2) = CURVES_NAME(comboSelectNumber)
        section(sectionNumber, 3) = comboRatio

        ' 下一組
        sectionNumber = sectionNumber + 1

    Next

    SectionSelector = section()

End Function


Function Newton(x0, y0, x1, y1, x2, y2)
' 牛頓法
' 參數：x0, y0, x1, y1, x2, y2
' 回傳：boolean
' 判斷是否與(0, 0)同側

    m = (y1 - y0) / (x1 - x0)

    Newton = ((y2 - y0) - m * (x2 - x0)) * (-y0 + m * x0) > 0

End Function


Function calRatio(x0, y0, x1, y1, x2, y2)
' 參數：x0, y0, x1, y1, x2, y2
' 回傳：Ratio
' Ratio = 點到 (0, 0) 距離 / ( 點到直線距離 + 點到 (0, 0) 距離 )

    m = (y1 - y0) / (x1 - x0)

    calRatio = Sqr(x2 ^ 2 + y2 ^ 2) / ((Abs((y2 - y0) - m * (x2 - x0)) / Sqr(1 + m ^ 2)) + Sqr(x2 ^ 2 + y2 ^ 2))

End Function


Function PrintSection(section)
' 輸出資料：
' PMM Control Section
' SectionSelector

    ' 寫入資料在 PMM
    Worksheets("PMM").Activate
    Columns(7).ClearContents
    Range(Cells(2, 7), Cells(UBound(CONTROL_COMBO), 7)) = CONTROL_COMBO
    Cells(1, 7) = "Control Section"

    ' 寫入資料在 SectionSelector
    Worksheets("SectionSelector").Activate
    Range(Columns(11), Columns(14)).ClearContents
    Range(Cells(2, 11), Cells(UBound(section), 14)) = section
    Cells(1, 11) = "Column"
    Cells(1, 12) = "NO."
    Cells(1, 13) = "SectionName"
    Cells(1, 14) = "Ratio"

End Function


Function ExecutionTime(time0)

    If Timer - time0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - time0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - time0) / 60, 2) & " Min", vbOKOnly
    End If

End Function


Sub SRCSelector()
'
' 目的：
' 由於在ETABS不會 Design SRC 斷面，所以由 ETABS 輸出 PMM。
' 以 SectionBuilder 建立 SRC 斷面，產生包絡線，檢測 ETABS PMM 有沒有在包絡線裡面。
'
'
' 演算法：
' 1. PMM curve 取 0 45 90 度
' 2. 由於 P 不一定會相同，排序內差求值
' 3. 以 PMM 點求得該 P 下的 0 45 90 度的 M
' 4. 以 PMM 點 M2 M3 判斷要和哪一條線比較
' 5. 以牛頓法判斷是不是與 (0, 0) 同側
'
'
' 測試：
' 執行時間：
'

    time0 = Timer

    Call AutoFill

    combo = ReadCombo()

    ' 有副作用，會修改全域變數 CURVES_NAME
    curves = ReadCurves()

    ' 有副作用，會修改全域變數 CONTROL_COMBO
    ' 使用全域變數 CURVES_NAME
    section = SectionSelector(combo, curves)

    Call PrintSection(section)

    ExecutionTime (time0)

End Sub
