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
    Set WS_LAP = Worksheets("搭接長度精細計算")
    Set WS_LENGTH = Worksheets("大梁")

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

    columnFirstCombo = 8
    columnLastCombo = 14

    rowLastCombo = WS_LAP.Cells(WS_LAP.Rows.Count, name_).End(xlUp).Row

    comboTable = WS_LAP.Range(WS_LAP.Cells(rowFirstCombo, columnFirstCombo), WS_LAP.Cells(rowLastCombo, columnLastCombo))

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

    columnWidth = 7

    rowLastInput = WS_LAP.Cells(WS_LAP.Rows.Count, columnWidth).End(xlUp).Row

    widthTable = WS_LAP.Range(WS_LAP.Cells(rowFirstInput, columnWidth), WS_LAP.Cells(rowLastInput, columnWidth))

    ReadWidth = widthTable

End Function


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

    columnName = 6
    rowFirstInput = 5
    lapName = WS_LAP.Cells(rowFirstInput, columnName)

    comboCover = 1
    comboFy = 2
    comboFyt = 3
    comboFc = 4
    comboFydb = 5
    comboFytdb = 6
    comboSpacing = 7

    ' 修正因數
    psitTop_ = 1.3
    psitBot_ = 1
    psie_ = 1
    lamda_ = 1

    rowTableGap = 6
    rowTitleGap = 3
    columnGap = 1

    maxLapLengthColumn = 50

    comboUBound = UBound(comboTable)
    widthUBound = UBound(widthTable)

    redim lapTable(1 to (widthUBound * 2 + 6) * comboUBound,  1 to maxLapLengthColumn)

    For rowCombo = 1 To comboUBound
        cover_ = comboTable(rowCombo, comboCover)
        fy_ = comboTable(rowCombo, comboFy)
        fyt_ = comboTable(rowCombo, comboFyt)
        fc_ = comboTable(rowCombo, comboFc)
        fydb_ = comboTable(rowCombo, comboFydb) / 10
        fytdb_ = comboTable(rowCombo, comboFytdb) / 10
        spacing_ = comboTable(rowCombo, comboSpacing)

        ldb_ = 0.28 * fy_ / sqr(fc_) * fydb_

        ' 修正因數
        If fydb_ >= 2 Then
            psis_ = 1
        Else
            psis_ = 0.8
        End If



        For rowWidth = 1 To widthUBound
            width_ = widthTable(rowWidth, 1)
            maxFyNum = fix((width_ - cover_ * 2 - fytdb_ * 2 - fydb_) / (2 * fydb_)) + 1

            ' 有加主筋之半
            cc_ = cover_ + fytdb_ + fydb_ / 2

            For fyNum = 2 To maxFyNum

                ' 有加主筋之半
                cs_ = ((width_ - fydb_ * fyNum - fytdb_ * 2 - cover_ * 2) / (fyNum - 1) + fydb_) / 2

                If cs_ <= cc_ Then
                    cb_ = cs_
                    atr_ = 2 * application.pi() * fytdb_ ^ 2 / 4
                    ktr_ = atr_ * fyt_ / 105 / spacing_ / fyNum
                Else
                    cb_ = cc_
                    atr_ = application.pi() * fytdb_ ^ 2 / 4
                    ktr_ = atr_ * fyt_ / 105 / spacing_
                End If

                botFactor = psitBot_ * psie_ * psis_ * lamda_ / application.min((cb_ + ktr_) / fydb_, 2.5)
                topFactor = psitTop_ * botFactor

                ldBot_ = botFactor * ldb_
                ldTop_ = topFactor * ldb_

                lapTable((rowCombo - 1) * (widthUBound * 2 + rowTableGap) + rowWidth * 2 + rowTitleGap, fyNum + columnGap) = fix(1.3 * ldTop_) + 1
                lapTable((rowCombo - 1) * (widthUBound * 2 + rowTableGap) + rowWidth * 2 + rowTitleGap + 1, fyNum + columnGap) = fix(1.3 * ldBot_) + 1

            Next fyNum
        Next rowWidth
    Next rowCombo

    columnName = 1
    columnTitle = 2

    For rowCombo = 1 To comboUBound
        cover_ = comboTable(rowCombo, comboCover)
        fy_ = comboTable(rowCombo, comboFy)
        fyt_ = comboTable(rowCombo, comboFyt)
        fc_ = comboTable(rowCombo, comboFc)
        fydb_ = comboTable(rowCombo, comboFydb) / 10
        fytdb_ = comboTable(rowCombo, comboFytdb) / 10
        spacing_ = comboTable(rowCombo, comboSpacing)

        lapTable((rowCombo - 1) * (widthUBound * 2 + rowTableGap), columnName) = lapName
        lapTable((rowCombo - 1) * (widthUBound * 2 + rowTableGap), columnTitle) = "表" & rowCombo & "  受拉竹節鋼筋搭接長度（乙級搭接）"
        lapTable((rowCombo - 1) * (widthUBound * 2 + rowTableGap), columnTitle) = "表" & rowCombo & "  受拉竹節鋼筋搭接長度（乙級搭接）"

    Next rowCombo

    WS_LENGTH.Range(WS_LENGTH.Cells(5, 4), WS_LENGTH.Cells(UBound(lapTable), maxLapLengthColumn)) = lapTable

End Sub
