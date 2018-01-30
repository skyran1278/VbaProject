Private NAME As Integer
Private WIDTH As Integer
Private COVER As Integer
Private Fy As Integer
Private Fyt As Integer
Private FC As Integer
Private FY_DB As Integer
Private FYT_DB As Integer
Private SPACING As Integer

Private WS_LAP As Worksheet

Private ROW_FIRST_INPUT As Integer
Private ROW_LAST_INPUT As Integer

Private COLUMN_FIRST_INPUT As Integer
Private COLUMN_LAST_INPUT As Integer

Private ROW_FIRST_COMBO As Integer

Private COLUMN_FIRST_COMBO As Integer
Private COLUMN_LAST_COMBO As Integer


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
    Set WS_LAP = ThisWorkbook.Worksheets("搭接長度精細計算")

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


Sub ClearCombo()
'
' 清除上一次產生的表格，以 NAME 那一 column 來取
'

    rowLastCombo = WS_LAP.Cells(WS_LAP.Rows.Count, NAME).End(xlUp).Row

    If rowLastCombo >= ROW_FIRST_COMBO Then
        WS_LAP.Range(WS_LAP.Rows(rowLastCombo), WS_LAP.Rows(ROW_FIRST_COMBO)).ClearContents
    End If


End Sub


Sub SortInput()
'
' Input 由小到大排列
'

    For col_ = COLUMN_FIRST_INPUT To COLUMN_LAST_INPUT
        WS_LAP.Range(WS_LAP.Cells(ROW_FIRST_INPUT, col_), WS_LAP.Cells(ROW_LAST_INPUT, col_)).Sort _
            Key1:=Range(WS_LAP.Cells(ROW_FIRST_INPUT, col_), WS_LAP.Cells(ROW_LAST_INPUT, col_)), _
            order1:=xlAscending
    Next

End Sub


Function ReadInput()
'
' 讀取需要排列組合的 Input 進 Array 操作
'
' @returns inputTable(Array)

    Dim inputTable()
    Dim doubleToArray(1 To 1, 1 To 1)
    ReDim inputTable(COLUMN_FIRST_COMBO To COLUMN_LAST_COMBO)

    For col_ = COLUMN_FIRST_COMBO To COLUMN_LAST_COMBO
        ROW_LAST_INPUT = WS_LAP.Cells(Rows.Count, col_).End(xlUp).Row
        inputTable(col_) = WS_LAP.Range(WS_LAP.Cells(ROW_FIRST_INPUT, col_), WS_LAP.Cells(ROW_LAST_INPUT, col_))

        ' 重要：處理回傳 double，重新 assign 一個 array
        If TypeName(inputTable(col_)) = "Double" Then
            doubleToArray(1, 1) = inputTable(col_)
            inputTable(col_) = doubleToArray
        End If
    Next

    ReadInput = inputTable

End Function


Function UboundInput(inputTable)
'
' 回傳每個 column 的上限
'
' @param inputTable(Array)
' @returns inputUbound(Array)

    Dim inputUbound()
    ReDim inputUbound(COLUMN_FIRST_COMBO To COLUMN_LAST_COMBO)

    For col_ = COLUMN_FIRST_COMBO To COLUMN_LAST_COMBO
        inputUbound(col_) = UBound(inputTable(col_))
    Next

    UboundInput = inputUbound

End Function


Function Combo(inputTable, inputUbound)
'
' 排列組合
'
' @param inputTable(Array)
' @param inputUbound(Array)
' @returns comboTable(Array)


    Dim comboTable()

    ' 計算總共有幾個 combo
    comboUbound = 1
    For col_ = COLUMN_FIRST_COMBO To COLUMN_LAST_COMBO
        comboUbound = comboUbound * inputUbound(col_)
    Next

    ReDim comboTable(1 To comboUbound, COLUMN_FIRST_INPUT To COLUMN_LAST_INPUT)


    count_ = 0

    For rowCover = 1 To inputUbound(COVER)
        cover_ = inputTable(COVER)(rowCover, 1)

        For rowFy = 1 To inputUbound(Fy)
            fy_ = inputTable(Fy)(rowFy, 1)

            For rowFyt = 1 To inputUbound(Fyt)
                fyt_ = inputTable(Fyt)(rowFyt, 1)

                For rowFc = 1 To inputUbound(FC)
                    fc_ = inputTable(FC)(rowFc, 1)

                    For rowFydb = 1 To inputUbound(FY_DB)
                        fydb_ = inputTable(FY_DB)(rowFydb, 1)

                        For rowFytdb = 1 To inputUbound(FYT_DB)
                            fytdb_ = inputTable(FYT_DB)(rowFytdb, 1)

                            For rowSpacing = 1 To inputUbound(SPACING)
                                spacing_ = inputTable(SPACING)(rowSpacing, 1)

                                count_ = count_ + 1

                                comboTable(count_, NAME) = count_
                                comboTable(count_, COVER) = cover_
                                comboTable(count_, Fy) = fy_
                                comboTable(count_, Fyt) = fyt_
                                comboTable(count_, FC) = fc_
                                comboTable(count_, FY_DB) = fydb_
                                comboTable(count_, FYT_DB) = fytdb_
                                comboTable(count_, SPACING) = spacing_

                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next

    Combo = comboTable

End Function

Sub PrintCombo(comboTable)
'
' 印出 Combo
'
' @param comboTable(Array)

    WS_LAP.Range(WS_LAP.Cells(ROW_FIRST_COMBO, COLUMN_FIRST_INPUT), WS_LAP.Cells(UBound(comboTable) + ROW_FIRST_COMBO - 1, COLUMN_LAST_COMBO)) = comboTable

End Sub


Sub Format()
'
' 格式化表格
'

    With WS_LAP
        .Cells.HorizontalAlignment = xlCenter
        .Cells.Font.NAME = "微軟正黑體"
        .Cells.Font.NAME = "Calibri"
        .Range(.Cells(ROW_FIRST_INPUT, COLUMN_FIRST_INPUT), .Cells(ROW_LAST_INPUT, COLUMN_LAST_INPUT)).Font.Color = vbRed

        .Columns(FY_DB).NumberFormatLocal = """D""0"
        .Columns(FYT_DB).NumberFormatLocal = """D""0"
    End With

End Sub


Sub Main()
'
' @purpose:
' 排列組合
'
'
' @algorithm:
' 排列組合
'
'
' @test:
' [0.4] 執行時間：154.00 sec 120.92 sec
' [1.11] 執行時間： 0.14 sec 0.26 sec
' [1.14] 執行時間： 0.07 sec 0.09 sec
' 多種狀況測試

    Dim time0 As Double
    Dim inputTable()
    Dim inputUbound()
    Dim comboTable()

    time0 = Timer
    Call PerformanceVBA(True)

    Call GlobalVariable

    Call ClearCombo

    Call SortInput
    inputTable = ReadInput()
    inputUbound = UboundInput(inputTable)

    comboTable = Combo(inputTable, inputUbound)
    Call PrintCombo(comboTable)

    Call Format

    Call PerformanceVBA(False)
    Call ExecutionTimeVBA(time0)

End Sub


