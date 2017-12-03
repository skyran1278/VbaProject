dim CURVES_NAME

Sub SRCSelector()
' 目的：
' 由於在ETABS不會 Design SRC 斷面，所以由 ETABS 輸出 PMM。
' 以 SectionBuilder 建立 SRC 斷面，產生包絡線，檢測 ETABS PMM 有沒有在包絡線裡面。


' 演算法：
' 1. PMM curve 取 0 45 90 度
' 2. 由於 P 不一定會相同，排序內差求值
' 3. 以 PMM 點求得該 P 下的 0 45 90 度的 M
' 4. 以 PMM 點 M2 M3 判斷要和哪一條線比較
' 5. 以牛頓法判斷是不是與 (0, 0) 同側


' 執行時間：
' 1.41s 7 萬資料量
' 6.9s 40 萬資料量
'
' 增加 Ratio 計算後的執行時間：
' 32.36s 40 萬資料量
' 重構程式碼後的執行時間：
' 21.61s 40 萬資料量
'

    time0 = Timer

    Call AutoFill

    combo = ReadCombo()

    ' 有副作用，會修改全域變數 CURVES_NAME
    curves = ReadCurves()

    SelectSection = SectionSelector(combo, curves, CURVES_NAME)

    ExecutionTime (time0)

End Sub


Function AutoFill()
' 公式自動填滿

    Worksheets("EtabsPMMCombo").Activate
    comboRowUsed = Cells(Rows.Count, 1).End(xlUp).row

    Worksheets("PMM").Activate
    Range(Cells(2, 1), Cells(2, 4)).AutoFill Destination:=Range(Cells(2, 1), Cells(comboRowUsed, 4))

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
        combo(row, 2) = Cells(row, 2)

        ' M2
        combo(row, 3) = Cells(row, 3)

        ' M3
        combo(row, 4) = Cells(row, 4)

    Next

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
    For Degree = 1 To 3

        loading = Degree * 4 + 1
        mement = Degree * 4 + 2

        For Point = 1 To 20

            pointCumulativeNumber = pointCumulativeNumber + 1

            ' P
            curve(pointCumulativeNumber, 0) = Cells(row + Point, loading)

            ' M
            curve(pointCumulativeNumber, Degree) = Cells(row + Point, mement)

        Next

    Next

    ' 以 load 排序
    Call QuickSortArray(curve, , , 0)

    ' 內插
    for moment = 1 to 3

        index = 0

        for point = 1 to 60

            ' 初始化，第一次讀到的非空值起始
            If Not IsEmpty(curve(point, moment)) and index = 0 Then

                pointMin = point
                loadMin = curve(point, 0)
                momentMin = curve(point, moment)

                index = 1

            ' 計算中間值
            elseif Not IsEmpty(curve(point, moment)) Then

                pointMax = point
                loadMax = curve(point, 0)
                momentMax = curve(point, moment)

                for pointMid = pointMin + 1 to pointMax - 1

                    ' 內插公式
                    curve(pointMid, moment) = ((curve(pointMid, 0) - loadMin) / (loadMax - loadMin)) * (momentMax - momentMin) + momentMin

                next

                pointMin = pointMax
                loadMin = loadMax
                momentMin = momentMax

            End If

        next

    next

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


Function SectionSelector(combo, curves, CURVES_NAME)

    ' 取出上限
    comboUBound = UBound(combo)
    curvesBound = UBound(curves)

    ' 使輸出結果陣列與 combo 相同
    Dim selectSection()
    ReDim selectSection(comboUBound, 4)

    for row = 2 To comboUBound

        ' 看看他與下一筆資料相不相同，如果相同就是一組。
        If combo(row, 1) <> combo(row + 1, 1) Then
            comboNumber = row - 2 + 1
            exit for
        End If

    next

    ' 從第1筆資料Loop到最後一筆
    For row = 2 To comboUBound step comboNumber

        ' 每一個Column（包含很多個Combo）重新初始化
        comboSelectNumber = 0
        comboRatio = 0

        ' 相同的一組
        For comboRow = row To row + comboNumber

            ' 每一個Combo重新初始化
            ratio = 0

            For curvesNumber = 1 To curvesBound

                curve = curves(curvesNumber)

                If combo(comboRow, 3) > combo(comboRow, 4) Then
                    ' PMM的資料格式：
                    ' M P Angle b c
                    ' combo的資料格式：
                    ' Name M P Angle
                    If Newton(combo(comboRow, 1), curve(LineNumber, 3), combo(comboRow, 2), curve(LineNumber, 4), curve(LineNumber - 1, 2), curve(LineNumber, 2), combo(comboRow, 3)) Then
                        ratio = CaculateRatio(combo(comboRow, 1), combo(comboRow, 2), curve(LineNumber, 3), curve(LineNumber, 4))
                        GoTo NextCombo
                    End If

                End If

            Next



NextCombo:
            ' Combo Loop 結束
            ' 超出所有PMMCurve，例外處理
            If curvesNumber = 0 Then
                curvesNumber = PMMNumber + 1
                selectSection(index, 4) = PMMNumber + 1
            Else
                selectSection(index, 4) = curvesNumber
            End If



            ' 判斷有沒有大於comboSelectNumber，有的話才寫入
            If comboSelectNumber < curvesNumber Then
                comboSelectNumber = curvesNumber
                comboRatio = ratio
            End If

            ' 判斷有沒有大於Ratio，有的話才寫入
            If comboRatio < ratio And comboSelectNumber <= curvesNumber Then
                comboRatio = ratio
            End If

        Next


        ' 斷面的Loop 結束
        ' 寫入斷面資料
        selectSection(selectSectionNumber, 0) = combo(row, 0)
        selectSection(selectSectionNumber, 1) = comboSelectNumber
        selectSection(selectSectionNumber, 2) = PMMCurveName(comboSelectNumber)
        selectSection(selectSectionNumber, 3) = comboRatio

        ' 下一組的開始編號
        StartNumber = row + 1

        ' 下一組
        selectSectionNumber = selectSectionNumber + 1

    Next

    SectionSelector = selectSection()

End Function


Function ExecutionTime(time0)

    If Timer - time0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - time0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - time0) / 60, 2) & " Min", vbOKOnly
    End If

End Function



