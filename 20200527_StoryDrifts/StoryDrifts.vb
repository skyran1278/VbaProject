Private RAN As UTILS_CLASS

Sub Main()
'
' @purpose:
'
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
    Dim storys

    Set RAN = New UTILS_CLASS
    Set wsStoryDrifts = Worksheets("Story Drifts")
    Set wsDrifts = Worksheets("Drifts")

    Call RAN.ExecutionTime(True)
    Call RAN.PerformanceVBA(False)

    arrStoryDrifts = RAN.GetRangeToArray(wsStoryDrifts, 2, 2, 2, 10)

    With wsStoryDrifts

        ' 消除之前的公式
        .Range(.Cells(1, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 1)).ClearContents

        ' 創建公式
        .Range("A2").Formula = "=B2&C2&D2"
        .Range("A2").AutoFill Destination:=.Range(.Cells(2, 1), .Cells(UBound(arrStoryDrifts), 1))

    End With


    ' MY Was Here
    With wsDrifts

        ' 初始化
        if .Cells(Rows.Count, 2).End(xlUp).Row > 7 then
            .Range(.Cells(7, 1), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 6)).ClearContents
        end if

        topStoryName = .Cells(1, 9)
        botStoryName = .Cells(2, 9)

        ' unique array
        allStorys = RAN.CreateDictionary(arrStoryDrifts, 1, False).keys()

        indextopStory = Application.WorksheetFunction.Match(topStoryName, allStorys, 0)
        indexbotStory = Application.WorksheetFunction.Match(botStoryName, allStorys, 0)

        ReDim storys(indexbotStory - indextopStory)

        For i = indextopStory To indexbotStory
            storys(i - indextopStory) = allStorys(i - 1)
        Next i


        botStoryRow = 7 + UBound(storys)

        .Cells(botStoryRow, 1) = 2
        .Cells(botStoryRow - 1, 1) = 3
        .Range(.Cells(botStoryRow, 1), .Cells(botStoryRow - 1, 1)).AutoFill Destination:=.Range(.Cells(7, 1), .Cells(botStoryRow, 1)), Type:=xlFillSeries

        topStory = .Cells(7, 1)

        ' 樓層
        .Range(.Cells(7, 2), .Cells(botStoryRow, 2)) = Application.WorksheetFunction.Transpose(storys)

        ' 層間位移 VLOOKUP
        .Range("C7").Formula = "=VLOOKUP($B7&C$2&C$1, 'Story Drifts'!$A:$J, 9, 0) * C$3"
        .Range("D7").Formula = "=VLOOKUP($B7&D$2&D$1, 'Story Drifts'!$A:$J, 9, 0) * D$3"
        .Range("E7").Formula = "=VLOOKUP($B7&E$2&E$1, 'Story Drifts'!$A:$J, 10, 0) * E$3"
        .Range("F7").Formula = "=VLOOKUP($B7&F$2&F$1, 'Story Drifts'!$A:$J, 10, 0) * F$3"
        .Range(.Cells(7, 3), .Cells(7, 7)).AutoFill Destination:=.Range(.Cells(7, 3), .Cells(botStoryRow, 7))

        ' 圖表
        Set chartX = .ChartObjects("X 向層間位移").Chart

        chartX.SeriesCollection(1).Name = "+X"
        chartX.SeriesCollection(1).XValues = .Range(.Cells(7, 3), .Cells(botStoryRow, 3))
        chartX.SeriesCollection(1).values = .Range(.Cells(7, 1), .Cells(botStoryRow, 1))

        chartX.SeriesCollection(2).Name = "-X"
        chartX.SeriesCollection(2).XValues = .Range(.Cells(7, 4), .Cells(botStoryRow, 4))
        chartX.SeriesCollection(2).values = .Range(.Cells(7, 1), .Cells(botStoryRow, 1))

        chartX.SeriesCollection(3).Name = "法規上限"
        chartX.SeriesCollection(3).XValues = Array(0.005, 0.005)
        chartX.SeriesCollection(3).values = Array(topStory, 2)

        chartX.Axes(xlValue).TickLabels.NumberFormatLocal = "[=" + CStr(topStory) + "] """ + topStoryName + """;0""F"""
        chartX.Axes(xlValue).MinimumScale = 2
        chartX.Axes(xlValue).MaximumScale = topStory

        Set chartY = .ChartObjects("Y 向層間位移").Chart

        chartY.SeriesCollection(1).Name = "+Y"
        chartY.SeriesCollection(1).XValues = .Range(.Cells(7, 5), .Cells(botStoryRow, 5))
        chartY.SeriesCollection(1).values = .Range(.Cells(7, 1), .Cells(botStoryRow, 1))

        chartY.SeriesCollection(2).Name = "-Y"
        chartY.SeriesCollection(2).XValues = .Range(.Cells(7, 6), .Cells(botStoryRow, 6))
        chartY.SeriesCollection(2).values = .Range(.Cells(7, 1), .Cells(botStoryRow, 1))

        chartY.SeriesCollection(3).Name = "法規上限"
        chartY.SeriesCollection(3).XValues = Array(0.005, 0.005)
        chartY.SeriesCollection(3).values = Array(topStory, 2)

        chartY.Axes(xlValue).TickLabels.NumberFormatLocal = "[=" + CStr(topStory) + "] """ + topStoryName + """;0""F"""
        chartY.Axes(xlValue).MinimumScale = 2
        chartY.Axes(xlValue).MaximumScale = topStory

    End With

    Call RAN.PerformanceVBA(False)
    Call RAN.ExecutionTime(False)

End Sub
