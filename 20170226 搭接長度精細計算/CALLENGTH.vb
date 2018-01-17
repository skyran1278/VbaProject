Dim NAME As Integer
Dim WIDTH As Integer
Dim COVER As Integer
Dim Fy As Integer
Dim Fyt As Integer
Dim FC As Integer
Dim FY_DB As Integer
Dim FYT_DB As Integer
Dim SPACING As Integer

Dim WS_LAP As Worksheet

Dim ROW_FIRST_INPUT As Integer
Dim ROW_LAST_INPUT As Integer

Dim COLUMN_FIRST_INPUT As Integer
Dim COLUMN_LAST_INPUT As Integer

Dim ROW_FIRST_COMBO As Integer

Dim COLUMN_FIRST_COMBO As Integer
Dim COLUMN_LAST_COMBO As Integer


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

    ' Column 位置
    NAME = 6
    WIDTH = 7
    COVER = 8
    Fy = 9
    Fyt = 10
    FC = 11
    FY_DB = 12
    FYT_DB = 13
    SPACING = 14

    ' worksheets
    Set WS_LAP = Worksheets("搭接長度精細計算")

    ' Input Variable
    ROW_FIRST_INPUT = 5
    ROW_LAST_INPUT = 19

    COLUMN_FIRST_INPUT = 6
    COLUMN_LAST_INPUT = 14

    ' Combo Variable
    ROW_FIRST_COMBO = 21

    COLUMN_FIRST_COMBO = 8
    COLUMN_LAST_COMBO = 14

End Sub


Function ReadCombo()
'
'
'
' @param
' @returns

    Dim comboTable()
    rowLastCombo = WS_LAP.Cells(WS_LAP.Rows.Count, NAME).End(xlUp).Row

    comboTable = WS_LAP.Range(WS_LAP.Cells(ROW_FIRST_COMBO, COLUMN_FIRST_COMBO), WS_LAP.Cells(rowLastCombo, COLUMN_LAST_COMBO))

    ReadCombo = comboTable

End Function


Function ReadWidth()
'
'
'
' @param
' @returns

    Dim widthArray()
    rowLastInput = WS_LAP.Cells(WS_LAP.Rows.Count, WIDTH).End(xlUp).Row

    widthArray = WS_LAP.Range(WS_LAP.Cells(ROW_FIRST_INPUT, WIDTH), WS_LAP.Cells(rowLastInput, WIDTH))

    ReadWidth = widthArray

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
    Dim widthArray()

    Call GlobalVariable
    comboTable = ReadCombo()
    widthArray = ReadWidth()
    lapName = WS_LAP.Cells(ROW_FIRST_INPUT, NAME)

End Sub
