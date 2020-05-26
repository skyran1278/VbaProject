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
    Set RAN = New UTILS_CLASS
    Set wsStoryDrifts = Worksheets("Story Drifts")
    Set wsDrifts = Worksheets("Drifts")

    Call RAN.ExecutionTime(True)
    Call RAN.PerformanceVBA(False)

    arrStoryDrifts = RAN.GetRangeToArray(wsStoryDrifts, 2, 2, 2, 9)

    ' autofill
    storyDriftsUsed = UBound(arrStoryDrifts)

    With wsStoryDrifts

        .Range(.Cells(1, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 1)).ClearContents
        .Range("A2").Formula = "=B2&C2&D2"
        .Range("A2").AutoFill Destination:=.Range(.Cells(2, 1), .Cells(storyDriftsUsed, 1))

    End With

    ' unique array
   storys = RAN.CreateDictionary(arrStoryDrifts, 1, False).keys()
    ' MY Was Here
    With wsDrifts
        ' 初始化
        .Range(.Cells(7, 1), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 6)).ClearContents

        topStory = .Cells(1, 9)
        botStory = .Cells(2, 9)

        indexbotStory = Application.WorksheetFunction.Match(botStory, storys, 0)
        indextopStory = Application.WorksheetFunction.Match(topStory, storys, 0)

        .Cells(7 + indexbotStory - 1, 1) = 2
        .Range(.Cells(7 + indexbotStory - 1, 1), .Cells(7 + indexbotStory - 1, 1)).AutoFill Destination:=.Range(.Cells(7, 1), .Cells(7 + indexbotStory - 1, 1))

        .Range(.Cells(7, 2), .Cells(7 + UBound(storys) - 1, 2)) = Application.WorksheetFunction.Transpose(storys())
        .Range("C7").Formula = "=VLOOKUP($B7&C$2&C$1, 'Story Drifts'!$A:$J, 9, 0) * C$3"
        .Range("D7").Formula = "=VLOOKUP($B7&D$2&D$1, 'Story Drifts'!$A:$J, 9, 0) * D$3"
        .Range("E7").Formula = "=VLOOKUP($B7&E$2&E$1, 'Story Drifts'!$A:$J, 10, 0) * E$3"
        .Range("F7").Formula = "=VLOOKUP($B7&F$2&F$1, 'Story Drifts'!$A:$J, 10, 0) * F$3"
        .Range(.Cells(7, 3), .Cells(7, 7)).AutoFill Destination:=.Range(.Cells(7, 3), .Cells(2 + UBound(storys) - 1, 7))
    End With

    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).Name = "=""+X"""
    ActiveChart.FullSeriesCollection(1).XValues = "=Drifts!$C$7:$C$40"
    ActiveChart.FullSeriesCollection(1).values = "=Drifts!$B$7:$B$40"

    Call RAN.PerformanceVBA(False)
    Call RAN.ExecutionTime(False)

End Sub

