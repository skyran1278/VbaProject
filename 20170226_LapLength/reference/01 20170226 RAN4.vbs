Option Explicit

Sub PostProcessing()
    Dim Time0#, PlotNumber As Integer, CalculateNumber As Integer, BeamPositionColumn As Integer
    Time0 = Timer
    Application.ScreenUpdating = False

    'PlotNumber = 出圖表格位置
    'CalculateNumber = 計算表格位置
    'BeamPositionColumn = 輸入的梁欄位

    Call Addlength(2, 5, 4)
    Call Addlength(3, 6, 11)
    Call Addlength(4, 7, 18)
    
    Application.ScreenUpdating = True
    MsgBox "執行時間 " & Timer - Time0 & " 秒", vbOKOnly '927
End Sub

Sub Addlength(PlotNumber, CalculateNumber, BeamPositionColumn)
    Dim BeamWidthNumber2 As Integer, LastColumnNumber As Integer, BeamWidthNumber As Integer, _
    TableNumber As Integer, CountRowNumber As Integer
    Dim I As Integer, CalculateCountColumnNumber As Integer
    Dim Ws1 As Worksheet, WsPlot As Worksheet, WsCalculate As Worksheet, WsExplanation As Worksheet
    Dim NeedToAddLengthRange As Range, NeedToAddLengthCell As Range
    Set Ws1 = Worksheets(1)
    Set WsPlot = Worksheets(PlotNumber)
    'Set WsCalculate = Worksheets(CalculateNumber)
    Set WsExplanation = Worksheets(8)
    
    WsPlot.Activate
    LastColumnNumber = WsPlot.UsedRange.Columns.Count + 4
    BeamWidthNumber = Ws1.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5
    BeamWidthNumber2 = (Ws1.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5) * 2
    TableNumber = Ws1.Cells(Rows.Count, BeamPositionColumn + 2).End(xlUp).Row - 29
    CountRowNumber = 5
    CalculateCountColumnNumber = 1


    For I = 1 To TableNumber
        Set NeedToAddLengthRange = Range(Cells(CountRowNumber + 3, 6), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber))
        
        For Each NeedToAddLengthCell In NeedToAddLengthRange
            If NeedToAddLengthCell <> "" Then
                NeedToAddLengthCell.Value = NeedToAddLengthCell.Value + WsExplanation.Cells(23, 11).Value
            End If
        Next NeedToAddLengthCell

        CountRowNumber = CountRowNumber + BeamWidthNumber2 + 5
        CalculateCountColumnNumber = CalculateCountColumnNumber + 35
    Next
    
End Sub

