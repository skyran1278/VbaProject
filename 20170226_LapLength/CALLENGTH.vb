Private WS_LAP As Worksheet
Private WS_LENGTH As Worksheet


Sub GlobalVariable()
'
' 宣告全域變數：Worksheets
'
' @returns WS_LAP(Worksheet)
' @returns WS_LENGTH(Worksheet)

    ' worksheets
    Set WS_LAP = ThisWorkbook.Worksheets("搭接長度精細計算")
    ThisWorkbook.Worksheets.Add After:=WS_LAP
    Set WS_LENGTH = ThisWorkbook.ActiveSheet

End Sub


Function ReadCombo()
'
' 讀取所有 Combo 種類
'
' @returns comboTable(Array)

    Dim comboTable()

    name_ = 6
    rowFirstCombo = 21

    colFirstCombo = 8
    colLastCombo = 14

    rowLastCombo = WS_LAP.Cells(WS_LAP.Rows.Count, name_).End(xlUp).Row

    comboTable = WS_LAP.Range(WS_LAP.Cells(rowFirstCombo, colFirstCombo), WS_LAP.Cells(rowLastCombo, colLastCombo))

    ReadCombo = comboTable

End Function


Function ReadWidth()
'
' 讀取梁寬
'
' @returns widthTable(Array)

    Dim widthTable()

    rowFirstInput = 5
    rowLastInput = 19

    columnWidth_ = 7

    rowLastInput = WS_LAP.Cells(WS_LAP.Rows.Count, columnWidth_).End(xlUp).Row

    widthTable = WS_LAP.Range(WS_LAP.Cells(rowFirstInput, columnWidth_), WS_LAP.Cells(rowLastInput, columnWidth_))

    ReadWidth = widthTable

End Function

Function ReadName()
'
' 讀取梁名
'
' @returns lapName(String)

    columnName_ = 6
    rowFirstInput = 5

    lapName = WS_LAP.Cells(rowFirstInput, columnName_)

    ReadName = lapName

End Function

Function CalLength(comboTable, widthTable)
'
' 核心演算法：
' 計算 Cc、Cs => 可得知 Cb 和破壞模式（水平或是垂直）Ktr
' 得 Ktr => 得修正係數
' 修正係數 * ldb = ld(延伸長度)
' ld(延伸長度) * 1.3 = 搭接長度
'
' @param comboTable(Array)
' @param widthTable(Array)
' @returns lapTable(Array)

    Dim lapTable()

    colCover = 1
    colFy = 2
    colFyt = 3
    colFc = 4
    colFydb = 5
    colFytdb = 6
    colSpacing = 7

    rowTableSpace = 2
    rowTitleSpace = 4

    colTitleSpace = 2

    pi_ = Application.Pi()

    ' 修正因數
    psitTop_ = 1.3
    psitBot_ = 1
    psie_ = 1
    lamda_ = 1

    comboUBound = UBound(comboTable, 1)
    widthUBound = UBound(widthTable, 1)

    ' 最多的支數 + colTitleSpace - 1，之所以 - 1 是因為從 2 開始，會少一欄。
    colLapTableUBound = Fix((Application.Max(widthTable) - Application.Min(Application.Index(comboTable, 0, colCover)) * 2 - Application.Min(Application.Index(comboTable, 0, colFytdb)) / 10 * 2 - Application.Min(Application.Index(comboTable, 0, colFydb)) / 10) / (2 * Application.Min(Application.Index(comboTable, 0, colFydb)) / 10)) + 1 + colTitleSpace - 1

    ReDim lapTable(1 To (widthUBound * 2 + 6) * comboUBound - 2, 1 To colLapTableUBound)

    For rowCombo = 1 To comboUBound

        cover_ = comboTable(rowCombo, colCover)
        fy_ = comboTable(rowCombo, colFy)
        fyt_ = comboTable(rowCombo, colFyt)
        fc_ = comboTable(rowCombo, colFc)
        fydb_ = comboTable(rowCombo, colFydb) / 10
        fytdb_ = comboTable(rowCombo, colFytdb) / 10
        spacing_ = comboTable(rowCombo, colSpacing)

        ldb_ = 0.28 * fy_ / Sqr(fc_) * fydb_

        ' 由於詳細計算法沒有收入簡算法可以修正的條件，所以到最後會比簡算法長，所以用簡算法來訂定上限。
        ' 修正因數
        If fydb_ >= 2 Then
            psis_ = 1
            ldSimpleTop = 0.19 * fy_ * psitTop_ * psie_ * lamda_ / Sqr(fc_) * fydb_
            ldSimpleBot = 0.19 * fy_ * psitBot_ * psie_ * lamda_ / Sqr(fc_) * fydb_
        Else
            psis_ = 0.8
            ldSimpleTop = 0.15 * fy_ * psitTop_ * psie_ * lamda_ / Sqr(fc_) * fydb_
            ldSimpleBot = 0.15 * fy_ * psitBot_ * psie_ * lamda_ / Sqr(fc_) * fydb_
        End If

        For rowWidth = 1 To widthUBound

            width_ = widthTable(rowWidth, 1)
            maxFyNum = Fix((width_ - cover_ * 2 - fytdb_ * 2 - fydb_) / (2 * fydb_)) + 1

            ' 有加主筋之半
            cc_ = cover_ + fytdb_ + fydb_ / 2

            For fyNum = 2 To maxFyNum

                ' 有加主筋之半
                cs_ = ((width_ - fydb_ * fyNum - fytdb_ * 2 - cover_ * 2) / (fyNum - 1) + fydb_) / 2

                If cs_ <= cc_ Then

                    cb_ = cs_
                    atr_ = 2 * pi_ * fytdb_ ^ 2 / 4
                    ktr_ = atr_ * fyt_ / 105 / spacing_ / fyNum

                Else

                    cb_ = cc_
                    atr_ = pi_ * fytdb_ ^ 2 / 4
                    ktr_ = atr_ * fyt_ / 105 / spacing_

                End If

                botFactor = psitBot_ * psie_ * psis_ * lamda_ / Min((cb_ + ktr_) / fydb_, 2.5)
                topFactor = psitTop_ * botFactor

                ldBot_ = botFactor * ldb_
                ldTop_ = topFactor * ldb_

                lapTable((rowCombo - 1) * (widthUBound * 2 + rowTableSpace + rowTitleSpace) + (rowWidth - 1) * 2 + rowTitleSpace + 1, fyNum + colTitleSpace - 1) = Fix(1.3 * Min(ldTop_, ldSimpleTop)) + 1
                lapTable((rowCombo - 1) * (widthUBound * 2 + rowTableSpace + rowTitleSpace) + (rowWidth - 1) * 2 + rowTitleSpace + 2, fyNum + colTitleSpace - 1) = Fix(1.3 * Min(ldBot_, ldSimpleBot)) + 1

            Next fyNum

        Next rowWidth

    Next rowCombo

    CalLength = lapTable

End Function


Function Min(param1, param2)
'
' 最佳化取最小值速度
'
' @param param1(Double)
' @param param2(Double)
' @returns Min(Double)

    If param1 < param2 Then
        Min = param1
    Else
        Min = param2
    End If

End Function


Function AddText(lapTable, comboTable, widthTable, lapName)
'
' 增加文字
'
' @param lapTable(Array)
' @param comboTable(Array)
' @param widthTable(Array)
' @param lapName(String)
' @returns lapTable(Array)

    colCover = 1
    colFy = 2
    colFyt = 3
    colFc = 4
    colFydb = 5
    colFytdb = 6
    colSpacing = 7

    rowTableSpace = 2
    rowTitleSpace = 4
    colTitleSpace = 2

    comboUBound = UBound(comboTable, 1)
    widthUBound = UBound(widthTable, 1)
    lapRowUBound = UBound(lapTable, 1)
    lapColUBound = UBound(lapTable, 2)


    For rowCombo = 1 To comboUBound

        cover_ = comboTable(rowCombo, colCover)
        fy_ = comboTable(rowCombo, colFy)
        fyt_ = comboTable(rowCombo, colFyt)
        fc_ = comboTable(rowCombo, colFc)
        fydb_ = comboTable(rowCombo, colFydb)
        fytdb_ = comboTable(rowCombo, colFytdb)
        spacing_ = comboTable(rowCombo, colSpacing)

        rowComboFirst = (rowCombo - 1) * (widthUBound * 2 + rowTableSpace + rowTitleSpace)

        lapTable(rowComboFirst + 1, 1) = lapName
        lapTable(rowComboFirst + 1, 2) = "表 " & rowCombo & " 受拉竹節鋼筋搭接長度（乙級搭接）（單位：公分）"

        lapTable(rowComboFirst + 2, 2) = "適用條件"
        lapTable(rowComboFirst + 2, 3) = "保護層" & vbLf & "cm"
        lapTable(rowComboFirst + 2, 4) = "fy" & vbLf & "kgf/cm2"
        lapTable(rowComboFirst + 2, 5) = "fyt" & vbLf & "kgf/cm2"
        lapTable(rowComboFirst + 2, 6) = "fc'" & vbLf & "kgf/cm2"
        lapTable(rowComboFirst + 2, 7) = "主筋直徑" & vbLf & "mm"
        lapTable(rowComboFirst + 2, 8) = "箍筋直徑" & vbLf & "mm"
        lapTable(rowComboFirst + 2, 9) = "箍筋間距" & vbLf & "cm"

        lapTable(rowComboFirst + 3, 3) = cover_
        lapTable(rowComboFirst + 3, 4) = fy_
        lapTable(rowComboFirst + 3, 5) = fyt_
        lapTable(rowComboFirst + 3, 6) = fc_
        lapTable(rowComboFirst + 3, 7) = fydb_
        lapTable(rowComboFirst + 3, 8) = fytdb_
        lapTable(rowComboFirst + 3, 9) = spacing_

        lapTable(rowComboFirst + 4, 2) = "梁寬\主筋根數"

        fyNum = 2
        For col_ = colTitleSpace + 1 To lapColUBound
            lapTable(rowComboFirst + 4, col_) = fyNum
            fyNum = fyNum + 1
        Next col_

        For rowWidth = 1 To widthUBound
            lapTable(rowComboFirst + rowTitleSpace + (rowWidth - 1) * 2 + 1, 2) = widthTable(rowWidth, 1)
        Next rowWidth

    Next rowCombo

    AddText = lapTable

End Function


Sub Format(lapTable, comboTable, widthTable, lapName)
'
' 調整格式
' 由於框線會覆蓋，所以需要調整順序。
'
' @param lapTable(Array)
' @param comboTable(Array)
' @param widthTable(Array)
' @param lapName(String)

    rowTableSpace = 2
    rowTitleSpace = 4
    colTitleSpace = 2

    comboUBound = UBound(comboTable, 1)
    widthUBound = UBound(widthTable, 1)
    lapRowUBound = UBound(lapTable, 1)
    lapColUBound = UBound(lapTable, 2)

    WS_LENGTH.Range(WS_LENGTH.Cells(1, 1), WS_LENGTH.Cells(lapRowUBound, lapColUBound)) = lapTable


    For rowCombo = 1 To comboUBound

        rowComboFirst = (rowCombo - 1) * (widthUBound * 2 + rowTableSpace + rowTitleSpace)

        ' "適用條件" Merge
        With WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + 2, 2), WS_LENGTH.Cells(rowComboFirst + 3, 2))
            .Merge
            .BorderAround Weight:=xlThin
        End With

        ' 梁寬\主筋根數 ColumnWidth
        WS_LENGTH.Columns(2).ColumnWidth = 9.88

        ' "fy" & vbLf & "kgf/cm2" Superscript
        WS_LENGTH.Cells(rowComboFirst + 2, 4).Characters(Start:=10, Length:=1).Font.Superscript = True

        ' "fyt" & vbLf & "kgf/cm2" Superscript
        WS_LENGTH.Cells(rowComboFirst + 2, 5).Characters(Start:=11, Length:=1).Font.Superscript = True

        ' "fc'" & vbLf & "kgf/cm2" Superscript
        WS_LENGTH.Cells(rowComboFirst + 2, 6).Characters(Start:=11, Length:=1).Font.Superscript = True

        ' fydb_ """D""0"
        WS_LENGTH.Cells(rowComboFirst + 3, 7).NumberFormatLocal = """D""0"

        ' fytdb_ """D""0"
        WS_LENGTH.Cells(rowComboFirst + 3, 8).NumberFormatLocal = """D""0"

        ' input red
        WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + 3, colTitleSpace + 1), WS_LENGTH.Cells(rowComboFirst + 3, lapColUBound)).Font.Color = vbRed

        ' "梁寬\主筋根數" Subscript Superscript xlThin
        With WS_LENGTH.Cells(rowComboFirst + 4, 2)
            .Characters(Start:=1, Length:=2).Font.Subscript = True
            .Characters(Start:=4, Length:=4).Font.Superscript = True
            .BorderAround Weight:=xlThin
        End With

        ' 主筋根數 xlThin
        For col_ = colTitleSpace + 1 To lapColUBound
            WS_LENGTH.Cells(rowComboFirst + 4, col_).BorderAround Weight:=xlThin
        Next col_

        ' 寬度 Merge xlThin
        For rowWidth = 1 To widthUBound
            With WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + rowTitleSpace + (rowWidth - 1) * 2 + 1, 2), WS_LENGTH.Cells(rowComboFirst + rowTitleSpace + (rowWidth - 1) * 2 + 2, 2))
                .Merge
                .BorderAround Weight:=xlThin
            End With
        Next rowWidth

        ' 搭接長度 xlThin
        For row_ = 1 To widthUBound
            For col_ = colTitleSpace + 1 To lapColUBound
                WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + rowTitleSpace + (row_ - 1) * 2 + 1, col_), WS_LENGTH.Cells(rowComboFirst + rowTitleSpace + (row_ - 1) * 2 + 2, col_)).BorderAround Weight:=xlThin
            Next col_
        Next row_

        ' 搭接長度 格式化條件
        With WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + rowTitleSpace + 1, colTitleSpace + 1), WS_LENGTH.Cells(rowComboFirst + rowTitleSpace + widthUBound * 2, lapColUBound))
            .FormatConditions.AddColorScale ColorScaleType:=2
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = 16776444
            .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueHighestValue
            .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = 7039480
        End With

        ' lapName Merge 中等 xlMedium
        With WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + 1, 1), WS_LENGTH.Cells(rowComboFirst + widthUBound * 2 + rowTitleSpace, 1))
            .Merge
            .Style = "中等"
            .BorderAround Weight:=xlMedium
        End With

        ' "表 " & rowCombo & " 受拉竹節鋼筋搭接長度（乙級搭接）（單位：公分）" Merge 好 xlMedium
        With WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + 1, 2), WS_LENGTH.Cells(rowComboFirst + 1, lapColUBound))
            .Merge
            .Style = "好"
            .BorderAround Weight:=xlMedium
        End With

        ' 雙劃線 xlDouble
        WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + 2, 2), WS_LENGTH.Cells(rowComboFirst + 3, lapColUBound)).Borders(xlEdgeBottom).LineStyle = xlDouble

        ' 外圍粗外框線 xlThick
        WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + 1, 1), WS_LENGTH.Cells(rowComboFirst + widthUBound * 2 + rowTitleSpace, lapColUBound)).BorderAround Weight:=xlThick

    Next rowCombo

    With WS_LENGTH

        ' 移動到指定位置
        .Range(.Columns(1), .Columns(3)).Insert (xlToRight)
        .Range(.Rows(1), .Rows(4)).Insert (xlDown)

        .Cells(1, 1) = "UPDATE"
        .Cells(1, 2) = Date
        .Cells(2, 1) = "PROJECT"
        .Cells(2, 2) = "搭接長度精細計算"
        .Columns(2).ColumnWidth = 16.67
        .Cells(3, 1) = "SUBJECT"


        .Cells.Font.NAME = "微軟正黑體"
        .Cells.Font.NAME = "Calibri"

        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter

    End With

End Sub


Sub PrintInput(lapTable)
'
' 紀錄 input 的參數，避免二次輸入浪費時間。
'
' @param lapTable(Array)

    Dim inputTable()

    name_ = 6

    colSpace = 3 + 10

    columnFirstInput = 6
    columnLastInput = 14

    rowFirstInput = 5
    rowLastCombo = WS_LAP.Cells(WS_LAP.Rows.Count, name_).End(xlUp).Row

    lapColUBound = UBound(lapTable, 2)

    inputTable = WS_LAP.Range(WS_LAP.Cells(rowFirstInput, columnFirstInput), WS_LAP.Cells(rowLastCombo, columnLastInput))

    WS_LENGTH.Range(WS_LENGTH.Cells(rowFirstInput, colSpace + lapColUBound), WS_LENGTH.Cells(rowLastCombo, colSpace + lapColUBound + columnLastInput - columnFirstInput)) = inputTable

End Sub


Sub Main()
'
' @purpose:
' 計算每個 combo 的搭接長度
'
'
' @algorithm:
' 核心演算法：
' 計算 Cc、Cs => 可得知 Cb 和破壞模式（水平或是垂直）Ktr
' 得 Ktr => 得修正係數
' 修正係數 * ldb = ld(延伸長度)
' ld(延伸長度) * 1.3 = 搭接長度
'
'
'
' @test:
'
' [0.4] 執行時間： 308.10 sec
' [1.12] 執行時間： 58.12 sec
' [1.13] 執行時間： 33.35 sec
' [1.14] 執行時間： 7.98 sec 8.79 sec
'
' [1.12] vs [0.4]：搭接長度與之前差 1~2 公分
'
'
    Dim time0 As Double

    Dim comboTable()
    Dim widthTable()
    Dim lapTable()

    time0 = Timer

    Call PerformanceVBA(True)

    Call GlobalVariable

    comboTable = ReadCombo()
    widthTable = ReadWidth()
    lapName = ReadName()

    lapTable = CalLength(comboTable, widthTable)
    lapTable = AddText(lapTable, comboTable, widthTable, lapName)

    Call Format(lapTable, comboTable, widthTable, lapName)

    Call PrintInput(lapTable)

    Call PerformanceVBA(False)

    Call ExecutionTimeVBA(time0)

End Sub
