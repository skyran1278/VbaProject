Option Explicit
Dim BeamName As String

Sub Main()
    Dim Time0#, PlotNumber As Integer, CalculateNumber As Integer, BeamPositionColumn As Integer
    Time0 = Timer
    Application.ScreenUpdating = False

    'PlotNumber = 出圖表格位置
    'CalculateNumber = 計算表格位置
    'BeamPositionColumn = 輸入的梁欄位

    BeamName = InputBox("名稱", , "大梁")

    Call BEAM(2, 5, 4)
    Call BEAM(3, 6, 11)
    Call BEAM(4, 7, 18)

    Call FormatText(2, 5, 4)
    Call FormatText(3, 6, 11)
    Call FormatText(4, 7, 18)

    Application.ScreenUpdating = True
    MsgBox "執行時間 " & Timer - Time0 & " 秒", vbOKOnly '927
End Sub

Sub BEAM(PlotNumber, CalculateNumber, BeamPositionColumn)
    Dim BeamWidthNumber As Integer, BeamWidth(20) As Integer, LastRowNumber As Integer, CountColumnNumber As Integer, PlotRowNumber1 As Integer
    Dim PlotRowNumber2 As Integer, StirrupStrength As Integer, Concrete As Integer, StirrupSpace As Integer, CountRowNumber As Integer
    Dim MaximumNumber As Integer, PlotColumnNumber As Integer, j As Integer, k As Integer, I As Integer
    Dim MainDiameter As Double, StirrupDiameter As Double
    Dim Ws1 As Worksheet, WsPlot As Worksheet, WsCalculate As Worksheet
    Set Ws1 = Worksheets(1)
    Set WsPlot = Worksheets(PlotNumber)
    Set WsCalculate = Worksheets(CalculateNumber)
    WsPlot.Cells.Delete
    BeamWidthNumber = Ws1.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5
    LastRowNumber = Ws1.Cells(Rows.Count, BeamPositionColumn + 2).End(xlUp).Row
    CountColumnNumber = 1
    PlotRowNumber1 = 1
    PlotRowNumber2 = 2
    WsPlot.Cells.Clear
    For j = 30 To LastRowNumber
        MainDiameter = Ws1.Cells(j, BeamPositionColumn + 2).Value
        StirrupDiameter = Ws1.Cells(j, BeamPositionColumn + 3).Value
        StirrupStrength = Ws1.Cells(j, BeamPositionColumn + 4).Value
        Concrete = Ws1.Cells(j, BeamPositionColumn + 5).Value
        StirrupSpace = Ws1.Cells(j, BeamPositionColumn + 6).Value
        CountRowNumber = 10
        PlotRowNumber1 = PlotRowNumber1 + 5
        PlotRowNumber2 = PlotRowNumber2 + 5
        For I = 0 To BeamWidthNumber - 1
            BeamWidth(I) = Ws1.Cells(I + 6, BeamPositionColumn + 1).Value
            MaximumNumber = Fix((BeamWidth(I) - 4 * 2 - StirrupDiameter / 10 * 2 - MainDiameter / 10) / (2 * MainDiameter / 10)) + 1
            PlotRowNumber1 = PlotRowNumber1 + 2
            PlotRowNumber2 = PlotRowNumber2 + 2
            If MaximumNumber <= 10 Then
                For k = 2 To MaximumNumber
                    CountRowNumber = CountRowNumber + 1
                    PlotColumnNumber = k + 4
                    Application.Calculation = xlCalculationManual 'xlCalculationManual
                    With WsCalculate
                        .Cells(CountRowNumber, CountColumnNumber + 3).Value = BeamWidth(I)
                        .Cells(CountRowNumber, CountColumnNumber + 5).Value = MainDiameter
                        .Cells(CountRowNumber, CountColumnNumber + 6).Value = StirrupDiameter
                        .Cells(CountRowNumber, CountColumnNumber + 8).Value = Concrete
                        .Cells(CountRowNumber, CountColumnNumber + 10).Value = StirrupStrength
                        .Cells(CountRowNumber, CountColumnNumber + 11).Value = k
                        .Cells(CountRowNumber, CountColumnNumber + 12).Value = StirrupSpace
                    End With
                    Application.Calculation = xlCalculationAutomatic 'xlCalculationAutomatic
                    WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = WsCalculate.Cells(CountRowNumber, CountColumnNumber + 32).Value
                    WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = WsCalculate.Cells(CountRowNumber, CountColumnNumber + 33).Value
                Next
            Else
                For k = 2 To 10
                    CountRowNumber = CountRowNumber + 1
                    PlotColumnNumber = k + 4
                    Application.Calculation = xlCalculationManual 'xlCalculationManual
                    With WsCalculate
                        .Cells(CountRowNumber, CountColumnNumber + 3).Value = BeamWidth(I)
                        .Cells(CountRowNumber, CountColumnNumber + 5).Value = MainDiameter
                        .Cells(CountRowNumber, CountColumnNumber + 6).Value = StirrupDiameter
                        .Cells(CountRowNumber, CountColumnNumber + 8).Value = Concrete
                        .Cells(CountRowNumber, CountColumnNumber + 10).Value = StirrupStrength
                        .Cells(CountRowNumber, CountColumnNumber + 11).Value = k
                        .Cells(CountRowNumber, CountColumnNumber + 12).Value = StirrupSpace
                    End With
                    Application.Calculation = xlCalculationAutomatic 'xlCalculationAutomatic
                    WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = WsCalculate.Cells(CountRowNumber, CountColumnNumber + 32).Value
                    WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = WsCalculate.Cells(CountRowNumber, CountColumnNumber + 33).Value

                Next
                For k = 12 To MaximumNumber Step 2
                    CountRowNumber = CountRowNumber + 1
                    PlotColumnNumber = k / 2 + 9
                    Application.Calculation = xlCalculationManual 'xlCalculationManual
                    With WsCalculate
                        .Cells(CountRowNumber, CountColumnNumber + 3).Value = BeamWidth(I)
                        .Cells(CountRowNumber, CountColumnNumber + 5).Value = MainDiameter
                        .Cells(CountRowNumber, CountColumnNumber + 6).Value = StirrupDiameter
                        .Cells(CountRowNumber, CountColumnNumber + 8).Value = Concrete
                        .Cells(CountRowNumber, CountColumnNumber + 10).Value = StirrupStrength
                        .Cells(CountRowNumber, CountColumnNumber + 11).Value = k
                        .Cells(CountRowNumber, CountColumnNumber + 12).Value = StirrupSpace
                    End With
                    Application.Calculation = xlCalculationAutomatic 'xlCalculationAutomatic
                    WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = WsCalculate.Cells(CountRowNumber, CountColumnNumber + 32).Value
                    WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = WsCalculate.Cells(CountRowNumber, CountColumnNumber + 33).Value

                Next
            End If
        Next
        CountColumnNumber = CountColumnNumber + 35
    Next
End Sub

Sub FormatText(PlotNumber, CalculateNumber, BeamPositionColumn)
    Dim BeamWidthNumber2 As Integer, BeamWidth(20) As Integer, LastColumnNumber As Integer, BeamWidthNumber As Integer, TableNumber As Integer, CountRowNumber As Integer
    Dim I As Integer, j As Integer, CalculateCountColumnNumber As Integer, LimitColumnWidth As Double
    Dim Ws1 As Worksheet, WsPlot As Worksheet, WsCalculate As Worksheet
    Set Ws1 = Worksheets(1)
    Set WsPlot = Worksheets(PlotNumber)
    Set WsCalculate = Worksheets(CalculateNumber)

    WsPlot.Activate
    Cells.HorizontalAlignment = xlCenter
    Cells.Font.Name = "微軟正黑體"
    WsPlot.Columns(5).ColumnWidth = 10
    LastColumnNumber = WsPlot.UsedRange.Columns.Count + 5
    BeamWidthNumber = Ws1.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5
    BeamWidthNumber2 = (Ws1.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5) * 2
    TableNumber = Ws1.Cells(Rows.Count, BeamPositionColumn + 2).End(xlUp).Row - 29
    CountRowNumber = 5
    CalculateCountColumnNumber = 1


    For I = 1 To TableNumber

        '格式化條件
        Range(Cells(CountRowNumber + 3, 6), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Select
        Selection.FormatConditions.AddColorScale ColorScaleType:=2
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
            xlConditionValueLowestValue
        Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
            xlConditionValueLowestValue
        With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            .Color = 16776444
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
            xlConditionValueHighestValue
        With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .Color = 7039480
            .TintAndShade = 0
        End With


        '合併儲存格
        CenterText Range(Cells(CountRowNumber, 5), Cells(CountRowNumber, LastColumnNumber))
        CenterText Range(Cells(CountRowNumber + 1, 5), Cells(CountRowNumber + 1, LastColumnNumber))
        For j = 1 To BeamWidthNumber2 Step 2
            CenterText Range(Cells(CountRowNumber + 2 + j, 5), Cells(CountRowNumber + 3 + j, 5))
        Next
        CenterText Range(Cells(CountRowNumber, 4), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, 4))


        '數值
        Cells(CountRowNumber, 4) = BeamName
        Cells(CountRowNumber, 4).Style = "中等"
        Cells(CountRowNumber + 2, 5) = "梁寬\主筋根數"
        Cells(CountRowNumber + 2, 5).Characters(Start:=1, Length:=2).Font.Subscript = True
        Cells(CountRowNumber + 2, 5).Characters(Start:=4, Length:=4).Font.Superscript = True
        For j = 0 To BeamWidthNumber - 1
            WsPlot.Cells(CountRowNumber + 3 + j * 2, 5) = Ws1.Cells(j + 6, BeamPositionColumn + 1).Value
        Next
        If LastColumnNumber > 14 Then
            For j = 6 To 14
                WsPlot.Cells(CountRowNumber + 2, j) = j - 4
            Next

            For j = 15 To LastColumnNumber
                WsPlot.Cells(CountRowNumber + 2, j) = j * 2 - 18
            Next
        Else
            For j = 6 To LastColumnNumber
                WsPlot.Cells(CountRowNumber + 2, j) = j - 4
            Next
        End If

        Cells(CountRowNumber, 5) = "表" & I & "  受拉竹節鋼筋搭接長度（乙級搭接）"
        Cells(CountRowNumber + 1, 5) = WsCalculate.Cells(10, CalculateCountColumnNumber)


        '框線
        Range(Cells(CountRowNumber + 2, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlInsideVertical).Weight = xlThin
        Range(Cells(CountRowNumber + 2, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlInsideHorizontal).Weight = xlThin
        Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeLeft).Weight = xlMedium
        ' Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeTop).Weight = xlMedium
        ' Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeBottom).Weight = xlMedium
        ' Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeRight).Weight = xlMedium
        Range(Cells(CountRowNumber, 4), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeLeft).Weight = xlMedium
        Range(Cells(CountRowNumber, 4), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeTop).Weight = xlMedium
        Range(Cells(CountRowNumber, 4), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeBottom).Weight = xlMedium
        Range(Cells(CountRowNumber, 4), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeRight).Weight = xlMedium
        Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 1, LastColumnNumber)).Borders(xlEdgeBottom).LineStyle = xlDouble
        For j = 1 To BeamWidthNumber2 Step 2
            Range(Cells(CountRowNumber + 2 + j, 6), Cells(CountRowNumber + 3 + j, LastColumnNumber)).Borders(xlInsideHorizontal).LineStyle = xlNone
        Next


        CountRowNumber = CountRowNumber + BeamWidthNumber2 + 5
        CalculateCountColumnNumber = CalculateCountColumnNumber + 35
    Next

    '調整欄寬
    LimitColumnWidth = 10
    For I = 6 To LastColumnNumber
        LimitColumnWidth = WsPlot.Columns(I).ColumnWidth + LimitColumnWidth
    Next
    If LastColumnNumber <> 6 Then
        If LimitColumnWidth < 108.5 Then
            For I = 6 To LastColumnNumber
                WsPlot.Columns(I).ColumnWidth = (108.5 - 10) / (LastColumnNumber - 5)
            Next
        End If
    End If


End Sub
Function CenterText(Range)
    With Range
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Merge
    End With
End Function





