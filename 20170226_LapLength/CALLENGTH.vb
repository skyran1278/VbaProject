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

    comboUBound = UBound(comboTable)
    widthUBound = UBound(widthTable)
    For row = 1 To comboUBound
        cover_ = comboTable(row, comboCover)
        fy_ = comboTable(row, comboFy)
        fyt_ = comboTable(row, comboFyt)
        fc_ = comboTable(row, comboFc)
        fydb_ = comboTable(row, comboFydb)
        fytdb_ = comboTable(row, comboFytdb)
        spacing_ = comboTable(row, comboSpacing)

        For rowWidth = 1 To widthUBound
            width_ = widthTable(rowWidth, 1)
            ' maxFy = fix()
        Next rowWidth



    Next row

End Sub
