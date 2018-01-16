' FIXME: 開兩個工作簿有問題，需要繼續測試
Dim NAME As Integer
Dim WIDTH As Integer
Dim Fy As Integer
Dim Fyt As Integer
Dim FC As Integer
Dim FY_DB As Integer
Dim FYT_DB As Integer
Dim SPACING As Integer

Dim WS_LAP As Worksheet

Sub Pretreatment()
    Dim time0 As Double
    time0 = Timer

    ' =====================================
    ' 宣告全域變數名稱 Column 位置
    NAME = 6
    WIDTH = 7
    Fy = 8
    Fyt = 9
    FC = 10
    FY_DB = 11
    FYT_DB = 12
    SPACING = 13
    Set WS_LAP = Worksheets("搭接長度精細計算")

    ' =====================================
    ' 清除上一次產生的表格
    rowFirstCombo = 21

    rowLastCombo = WS_LAP.Cells(WS_LAP.Rows.Count, NAME).End(xlUp).Row
    If rowLastCombo > rowFirstCombo Then
        WS_LAP.Range(WS_LAP.Rows(rowLastCombo), WS_LAP.Rows(rowFirstCombo)).ClearContents
    End If

    ' =====================================
    ' 格式化表格

    rowFirstInput = 5
    rowLastInput = 19
    columnFirstInput = 6
    columnLastInput = 13

    WS_LAP.Cells.HorizontalAlignment = xlCenter
    WS_LAP.Cells.Font.NAME = "微軟正黑體"
    WS_LAP.Cells.Font.NAME = "Calibri"
    WS_LAP.Range(WS_LAP.Cells(rowFirstInput, columnFirstInput), WS_LAP.Cells(rowLastInput, columnLastInput)).Font.Color = vbRed

    WS_LAP.Columns(FY_DB).NumberFormatLocal = """D""0"
    WS_LAP.Columns(FYT_DB).NumberFormatLocal = """D""0"

    ' =====================================
    ' 由小到大排列
    rowFirstInput = 5
    rowLastInput = 19
    columnFirstInput = 6
    columnLastInput = 13

    For column = columnFirstInput To columnLastInput
        WS_LAP.Range(WS_LAP.Cells(rowFirstInput, column), WS_LAP.Cells(rowLastInput, column)).Sort _
            Key1:=Range(WS_LAP.Cells(rowFirstInput, column), WS_LAP.Cells(rowLastInput, column)), _
            order1:=xlAscending
    Next

    ' =====================================
    ' 排列組合

    ' 計入每個 column 有多少個 row

    columnFirstCombo = 8
    columnLastCombo = 13

    Dim rowUsed() As Integer
    ReDim rowUsed(columnFirstCombo To columnLastCombo)

    For column = columnFirstCombo To columnLastCombo
        rowUsed(column) = WS_LAP.Cells(Rows.Count, column).End(xlUp).Row
    Next

    rowFirstInput = 5
    rowFirstCombo = 21
    count_ = 0
    For rowFy = rowFirstInput To rowUsed(Fy)
        fy_ = WS_LAP.Cells(rowFy, Fy)

        For rowFyt = rowFirstInput To rowUsed(Fyt)
            fyt_ = WS_LAP.Cells(rowFyt, Fyt)

            For rowFc = rowFirstInput To rowUsed(FC)
                fc_ = WS_LAP.Cells(rowFc, FC)

                For rowFydb = rowFirstInput To rowUsed(FY_DB)
                    fydb_ = WS_LAP.Cells(rowFydb, FY_DB)

                    For rowFytdb = rowFirstInput To rowUsed(FYT_DB)
                        fytdb_ = WS_LAP.Cells(rowFytdb, FYT_DB)

                        count_ = count_ + 1
                        WS_LAP.Cells(rowFirstCombo + count_, NAME) = count_
                        WS_LAP.Cells(rowFirstCombo + count_, Fy) = fy_
                        WS_LAP.Cells(rowFirstCombo + count_, Fyt) = fyt_
                        WS_LAP.Cells(rowFirstCombo + count_, FC) = fc_
                        WS_LAP.Cells(rowFirstCombo + count_, FY_DB) = fydb_
                        WS_LAP.Cells(rowFirstCombo + count_, FYT_DB) = fytdb_

                    Next
                Next
            Next
        Next
    Next

    ExecutionTime (time0)

End Sub

Function ExecutionTime(time0)

    If Timer - time0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - time0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - time0) / 60, 2) & " Min", vbOKOnly
    End If

End Function

' Sub SortTable()


' End Sub

' Sub CountNumberTable(BeamPositionColumn)
'     Dim lastrow(4) As Integer, Diameter As Double, TieDiameter As Double, Fyt As Double
'     Dim Concrete As Double, TieSpacing As Double, CountNumber As Double
'     Dim Fy As Double, Cover As Double, RebarSpacing As Double
'     Dim I As Integer, a As Integer, b As Integer, c As Integer, d As Integer, e As Integer
'     Dim CountAllTable As Integer, CountFirstTable As Integer, CountLastTable As Integer

'     For CountFirstTable = 1 To 5
'         For I = 0 To 4
'             lastrow(I) = Cells(Rows.Count, I + BeamPositionColumn + 2).End(xlUp).Row - 5
'         Next
'         CountNumber = 0
'         For a = 1 To lastrow(0)
'             Diameter = Cells(a + 5, BeamPositionColumn + 2)
'             For b = 1 To lastrow(1)
'                 TieDiameter = Cells(b + 5, BeamPositionColumn + 3)
'                 For c = 1 To lastrow(2)
'                     TieSpacing = Cells(c + 5, BeamPositionColumn + 4)
'                     For d = 1 To lastrow(3)
'                         Fyt = Cells(d + 5, BeamPositionColumn + 5)
'                         For e = 1 To lastrow(4)
'                             Concrete = Cells(e + 5, BeamPositionColumn + 6)
'                             CountNumber = CountNumber + 1
'                             Cells(29 + CountNumber, BeamPositionColumn) = CountNumber
'                             Cells(29 + CountNumber, BeamPositionColumn + 2) = Diameter
'                             Cells(29 + CountNumber, BeamPositionColumn + 3) = TieDiameter
'                             Cells(29 + CountNumber, BeamPositionColumn + 4) = TieSpacing
'                             Cells(29 + CountNumber, BeamPositionColumn + 5) = Fyt
'                             Cells(29 + CountNumber, BeamPositionColumn + 6) = Concrete
'                         Next
'                     Next
'                 Next
'             Next
'         Next
'         BeamPositionColumn = BeamPositionColumn + 7
'     Next

'     For CountLastTable = 1 To 5
'         For I = 0 To 4
'             lastrow(I) = Cells(Rows.Count, I + BeamPositionColumn + 1).End(xlUp).Row - 5
'         Next
'         CountNumber = 0
'         For a = 1 To lastrow(0)
'             Diameter = Cells(a + 5, BeamPositionColumn + 1)
'             For b = 1 To lastrow(1)
'                 Cover = Cells(b + 5, BeamPositionColumn + 2)
'                 For c = 1 To lastrow(2)
'                     RebarSpacing = Cells(c + 5, BeamPositionColumn + 3)
'                     For d = 1 To lastrow(3)
'                         Fy = Cells(d + 5, BeamPositionColumn + 4)
'                         For e = 1 To lastrow(4)
'                             Concrete = Cells(e + 5, BeamPositionColumn + 5)
'                             CountNumber = CountNumber + 1
'                             Cells(29 + CountNumber, BeamPositionColumn) = CountNumber
'                             Cells(29 + CountNumber, BeamPositionColumn + 1) = Diameter
'                             Cells(29 + CountNumber, BeamPositionColumn + 2) = Cover
'                             Cells(29 + CountNumber, BeamPositionColumn + 3) = RebarSpacing
'                             Cells(29 + CountNumber, BeamPositionColumn + 4) = Fy
'                             Cells(29 + CountNumber, BeamPositionColumn + 5) = Concrete
'                         Next
'                     Next
'                 Next
'             Next
'         Next
'         BeamPositionColumn = BeamPositionColumn + 6
'     Next

' End Sub







