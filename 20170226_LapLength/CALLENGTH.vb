' Dim NAME As Integer
' Dim WIDTH As Integer
' Dim COVER As Integer
' Dim Fy As Integer
' Dim Fyt As Integer
' Dim FC As Integer
' Dim FY_DB As Integer
' Dim FYT_DB As Integer
' Dim SPACING As Integer

Dim WS_LAP As Worksheet
Dim WS_LENGTH As Worksheet

' Dim ROW_FIRST_INPUT As Integer
' Dim ROW_LAST_INPUT As Integer

' Dim COLUMN_FIRST_INPUT As Integer
' Dim COLUMN_LAST_INPUT As Integer

' Dim ROW_FIRST_COMBO As Integer

' Dim COLUMN_FIRST_COMBO As Integer
' Dim COLUMN_LAST_COMBO As Integer


Sub GlobalVariable()
'
' 宣告全域變數：Column 位置、Worksheets
'
' @returns NAME(Integer)
' @returns COVER(Integer)
' @returns WIDTH(Integer)
' @returns Fy(Integer)
' @returns Fyt(Integer)
' @returns FC(Integer)
' @returns FY_DB(Integer)
' @returns FYT_DB(Integer)
' @returns SPACING(Integer)
' @returns WS_LAP(Worksheet)
' @returns ROW_FIRST_INPUT(Integer)
' @returns ROW_LAST_INPUT(Integer)
' @returns COLUMN_FIRST_INPUT(Integer)
' @returns COLUMN_LAST_INPUT(Integer)
' @returns ROW_FIRST_COMBO(Integer)
' @returns COLUMN_FIRST_COMBO(Integer)
' @returns COLUMN_LAST_COMBO(Integer)

    ' ' Column 位置
    ' NAME = 6
    ' WIDTH = 7
    ' COVER = 8
    ' Fy = 9
    ' Fyt = 10
    ' FC = 11
    ' FY_DB = 12
    ' FYT_DB = 13
    ' SPACING = 14

    ' worksheets
    Set WS_LAP = ThisWorkbook.Worksheets("搭接長度精細計算")
    Set WS_LENGTH = ThisWorkbook.Worksheets("大梁")

    ' ' Input Variable
    ' ROW_FIRST_INPUT = 5
    ' ROW_LAST_INPUT = 19

    ' COLUMN_FIRST_INPUT = 6
    ' COLUMN_LAST_INPUT = 14

    ' ' Combo Variable
    ' ROW_FIRST_COMBO = 21

    ' COLUMN_FIRST_COMBO = 8
    ' COLUMN_LAST_COMBO = 14

End Sub


Function ReadCombo()
'
'
'
' @param
' @returns

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
'
'
' @param
' @returns

    Dim widthTable()

    rowFirstInput = 5
    rowLastInput = 19

    ColumnWidth = 7

    rowLastInput = WS_LAP.Cells(WS_LAP.Rows.Count, ColumnWidth).End(xlUp).Row

    widthTable = WS_LAP.Range(WS_LAP.Cells(rowFirstInput, ColumnWidth), WS_LAP.Cells(rowLastInput, ColumnWidth))

    ReadWidth = widthTable

End Function

Function ReadName()
'
'
'
' @param
' @returns

    ColumnName = 6
    rowFirstInput = 5

    lapName = WS_LAP.Cells(rowFirstInput, ColumnName)

    ReadName = lapName

End Function

Function CalLength(comboTable, widthTable)
'
'
'
' @param
' @returns

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

    ' maxLapLengthColumn = 50

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

        ' 修正因數
        If fydb_ >= 2 Then
            psis_ = 1
        Else
            psis_ = 0.8
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
                    atr_ = 2 * Application.Pi() * fytdb_ ^ 2 / 4
                    ktr_ = atr_ * fyt_ / 105 / spacing_ / fyNum

                Else

                    cb_ = cc_
                    atr_ = Application.Pi() * fytdb_ ^ 2 / 4
                    ktr_ = atr_ * fyt_ / 105 / spacing_

                End If

                botFactor = psitBot_ * psie_ * psis_ * lamda_ / Application.Min((cb_ + ktr_) / fydb_, 2.5)
                topFactor = psitTop_ * botFactor

                ldBot_ = botFactor * ldb_
                ldTop_ = topFactor * ldb_

                lapTable((rowCombo - 1) * (widthUBound * 2 + rowTableSpace + rowTitleSpace) + (rowWidth - 1) * 2 + rowTitleSpace + 1, fyNum + colTitleSpace - 1) = Fix(1.3 * ldTop_) + 1
                lapTable((rowCombo - 1) * (widthUBound * 2 + rowTableSpace + rowTitleSpace) + (rowWidth - 1) * 2 + rowTitleSpace + 2, fyNum + colTitleSpace - 1) = Fix(1.3 * ldBot_) + 1

            Next fyNum

        Next rowWidth

    Next rowCombo

    CalLength = lapTable

End Function


Function AddText(lapTable, comboTable, widthTable, lapName)
'
'
'
' @param
' @returns

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
        lapTable(rowComboFirst + 2, 3) = "保護層" & vbCrLf & "cm"
        lapTable(rowComboFirst + 2, 4) = "fy" & vbCrLf & "kgf/cm2"
        lapTable(rowComboFirst + 2, 5) = "fyt" & vbCrLf & "kgf/cm2"
        lapTable(rowComboFirst + 2, 6) = "fc'" & vbCrLf & "kgf/cm2"
        lapTable(rowComboFirst + 2, 7) = "主筋直徑" & vbCrLf & "mm"
        lapTable(rowComboFirst + 2, 8) = "箍筋直徑" & vbCrLf & "mm"
        lapTable(rowComboFirst + 2, 9) = "箍筋間距" & vbCrLf & "cm"

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
'
'
' @param
' @returns

    rowTableSpace = 2
    rowTitleSpace = 4
    colTitleSpace = 2

    comboUBound = UBound(comboTable, 1)
    widthUBound = UBound(widthTable, 1)
    lapRowUBound = UBound(lapTable, 1)
    lapColUBound = UBound(lapTable, 2)


    For rowCombo = 1 To comboUBound

        rowComboFirst = (rowCombo - 1) * (widthUBound * 2 + rowTableSpace + rowTitleSpace)

        ' lapTable(rowComboFirst + 1, 1) = lapName
        WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + 1, 1), WS_LENGTH.Cells(rowComboFirst + widthUBound * 2 + rowTitleSpace, 1)).Merge

        ' lapTable(rowComboFirst + 1, 2) = "表 " & rowCombo & " 受拉竹節鋼筋搭接長度（乙級搭接）（單位：公分）"
        WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + 1, 2), WS_LENGTH.Cells(rowComboFirst + 1, lapColUBound)).Merge

        ' lapTable(rowComboFirst + 2, 2) = "適用條件"
        WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + 2, 2), WS_LENGTH.Cells(rowComboFirst + 3, 2)).Merge

        ' lapTable(rowComboFirst + 2, 4) = "fy" & vbCrLf & "kgf/cm2"
        WS_LENGTH.Cells(rowComboFirst + 2, 4).Characters(Start:=9, Length:=1).Font.Subscript = True

        ' lapTable(rowComboFirst + 2, 5) = "fyt" & vbCrLf & "kgf/cm2"
        WS_LENGTH.Cells(rowComboFirst + 2, 5).Characters(Start:=9, Length:=1).Font.Subscript = True

        ' lapTable(rowComboFirst + 2, 6) = "fc'" & vbCrLf & "kgf/cm2"
        WS_LENGTH.Cells(rowComboFirst + 2, 6).Characters(Start:=9, Length:=1).Font.Subscript = True

        ' lapTable(rowComboFirst + 3, 7) = fydb_
        WS_LENGTH.Cells(rowComboFirst + 3, 7).NumberFormatLocal = """D""0"

        ' lapTable(rowComboFirst + 3, 8) = fytdb_
        WS_LENGTH.Cells(rowComboFirst + 3, 8).NumberFormatLocal = """D""0"

        ' lapTable(rowComboFirst + 4, 2) = "梁寬\主筋根數"
        WS_LENGTH.Cells(rowComboFirst + 4, 2).Characters(Start:=1, Length:=2).Font.Subscript = True
        WS_LENGTH.Cells(rowComboFirst + 4, 2).Characters(Start:=4, Length:=4).Font.Superscript = True

        For rowWidth = 1 To widthUBound
            ' lapTable(rowComboFirst + rowTitleSpace + (rowWidth - 1) * 2 + 1, 2) = widthTable(rowWidth, 1)
            WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + rowTitleSpace + (rowWidth - 1) * 2 + 1, 2), WS_LENGTH.Cells(rowComboFirst + rowTitleSpace + (rowWidth - 1) * 2 + 2, 2)).Merge
        Next rowWidth

        WS_LENGTH.Range(WS_LENGTH.Cells(rowComboFirst + 1, 1), WS_LENGTH.Cells(rowComboFirst + widthUBound * 2 + rowTitleSpace, lapColUBound)).BorderAround(Weight:=xlThick).Weight = xlMedium

    Next rowCombo

    WS_LENGTH.Cells.Font.NAME = "微軟正黑體"
    WS_LENGTH.Cells.Font.NAME = "Calibri"

    WS_LENGTH.Cells.HorizontalAlignment = xlCenter
    WS_LENGTH.Cells.VerticalAlignment = xlCenter

End Sub


Sub Main()
'
' @purpose:
' 計算每個 combo 的搭接長度
'
'
' @algorithm:
'
'
'
' @test:
'
'
'

    Dim comboTable()
    Dim widthTable()
    Dim lapTable()

    Call GlobalVariable
    comboTable = ReadCombo()
    widthTable = ReadWidth()
    lapName = ReadName()
    lapTable = CalLength(comboTable, widthTable)
    lapTable = AddText(lapTable, comboTable, widthTable, lapName)
    Call Format(lapTable, comboTable, widthTable, lapName)

    lapRowUBound = UBound(lapTable, 1)
    lapColUBound = UBound(lapTable, 2)
    WS_LENGTH.Range(WS_LENGTH.Cells(1, 1), WS_LENGTH.Cells(lapRowUBound, lapColUBound)) = lapTable

End Sub



