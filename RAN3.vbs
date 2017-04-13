'Option Explicit

Sub Main()
    Dim Time0#, PlotNumber As Integer, BeamPositionColumn As Integer
    Set WsInput = Worksheets("輸入")
    Time0 = Timer
    Application.ScreenUpdating = False

    'PlotNumber = 出圖表格位置
    'BeamPositionColumn = 輸入的梁欄位
    BeamPositionColumn = 4
    For PlotNumber =  3 To 4
        BeamPositionColumn = BeamPositionColumn + 7
        If WsInput.Cells(4, BeamPositionColumn).Interior.Color = RGB(198, 239, 206) Then
            Call BeamSeismicDesign(PlotNumber, BeamPositionColumn)
            Call BeamFormatText(PlotNumber, BeamPositionColumn)
        End If        
    Next
    For PlotNumber =  5 To 7
        BeamPositionColumn = BeamPositionColumn + 7
        If WsInput.Cells(4, BeamPositionColumn).Interior.Color = RGB(198, 239, 206) Then
            Call BeamDesign(PlotNumber, BeamPositionColumn)
            Call BeamFormatText(PlotNumber, BeamPositionColumn)
        End If        
    Next
    For PlotNumber =  8 To 11
        BeamPositionColumn = BeamPositionColumn + 6
        If WsInput.Cells(4, BeamPositionColumn).Interior.Color = RGB(198, 239, 206) Then
            Call Design(PlotNumber, BeamPositionColumn)
            Call FormatText(PlotNumber, BeamPositionColumn)
        End If
    Next
        PlotNumber = 12
        BeamPositionColumn = BeamPositionColumn + 6
        If WsInput.Cells(4, BeamPositionColumn).Interior.Color = RGB(198, 239, 206) Then
            Call SeismicDesign(PlotNumber, BeamPositionColumn)
            Call FormatText(PlotNumber, BeamPositionColumn)
        End If
    

    Application.ScreenUpdating = True
    MsgBox "執行時間 " & Timer - Time0 & " 秒", vbOKOnly '927
End Sub

Sub BeamSeismicDesign(PlotNumber, BeamPositionColumn)
    Dim BeamWidthNumber As Integer, BeamWidth(20) As Integer ', LastRowNumber As Integer, CountColumnNumber As Integer, PlotRowNumber1 As Integer
    'Dim PlotRowNumber2 As Integer, Fyt As Integer, Concrete As Integer, TieSpacing As Integer, CountRowNumber As Integer
    'Dim MaximumNumber As Integer, PlotColumnNumber As Integer, j As Integer, ReinforcementNumber As Integer, I As Integer
    'Dim Db As Double, TieDiameter As Double
    'Dim WsInput As Worksheet, WsPlot As Worksheet
    
    Set WsInput = Worksheets("輸入")
    Set WsPlot = Worksheets(PlotNumber)
    WsPlot.Cells.Delete
    BeamWidthNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5
    LastRowNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 2).End(xlUp).Row    
    PlotRowNumber1 = 1
    PlotRowNumber2 = 2
    
    For j = 30 To LastRowNumber
        Db = WsInput.Cells(j, BeamPositionColumn + 2).Value / 10
        TieDiameter = WsInput.Cells(j, BeamPositionColumn + 3).Value / 10
        TieSpacing = WsInput.Cells(j, BeamPositionColumn + 4).Value
        Fyt = WsInput.Cells(j, BeamPositionColumn + 5).Value
        Concrete = WsInput.Cells(j, BeamPositionColumn + 6).Value
        Cover = 4
        Fy = 4200            
        PlotRowNumber1 = PlotRowNumber1 + 5
        PlotRowNumber2 = PlotRowNumber2 + 5
        For I = 0 To BeamWidthNumber - 1
            BeamWidth(I) = WsInput.Cells(I + 6, BeamPositionColumn + 1).Value
            MaximumNumber = Fix((BeamWidth(I) - 4 * 2 - TieDiameter * 2 - Db) / (2 * Db)) + 1
            Cc = Cover + TieDiameter            
            PlotRowNumber1 = PlotRowNumber1 + 2
            PlotRowNumber2 = PlotRowNumber2 + 2
            If MaximumNumber <= 10 Then
                For ReinforcementNumber = 2 To MaximumNumber                        
                    PlotColumnNumber = ReinforcementNumber + 4
                    Cs = (BeamWidth(I) - Db * ReinforcementNumber - TieDiameter * 2 - Cover * 2) / 2 / (ReinforcementNumber - 1)
                    Cb = Db / 2 + Application.Min(Cc, Cs)
                    If Cs < Cc Then
                        Ktr = 2 * Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing / ReinforcementNumber
                    Else
                        Ktr = Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing
                    End If
                    CorrectionFactor = Application.Max(0.4, Db / (Cb + Ktr))
                    If Db < 2 Then
                        Ldb = 0.28 * 0.8 * Fy * Db / Application.Sqrt(Concrete)
                    Else
                        Ldb = 0.28 * Fy * Db / Sqr(Concrete)
                    End If
                    Ld = CorrectionFactor * Ldb
                    Ldh = 0.06 * Fy * Db / Sqr(Concrete)
                    ClassBSpliceTop = Application.RoundUp(Application.Max(1.3 * 1.3 * Ld, 3.25 * Ldh, 30), 0)
                    ClassBSpliceBot = Application.RoundUp(Application.Max(1.3 * Ld, 2.5 * Ldh, 30), 0)

                    WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = ClassBSpliceTop
                    WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = ClassBSpliceBot
                Next
            Else
                For ReinforcementNumber = 2 To 10
                    PlotColumnNumber = ReinforcementNumber + 4

                    Cs = (BeamWidth(I) - Db * ReinforcementNumber - TieDiameter * 2 - Cover * 2) / 2 / (ReinforcementNumber - 1)
                    Cb = Db / 2 + Application.Min(Cc, Cs)
                    If Cs < Cc Then
                        Ktr = 2 * Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing / ReinforcementNumber
                    Else
                        Ktr = Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing
                    End If
                    CorrectionFactor = Application.Max(0.4, Db / (Cb + Ktr))
                    If Db < 2 Then
                        Ldb = 0.28 * 0.8 * Fy * Db / Application.Sqrt(Concrete)
                    Else
                        Ldb = 0.28 * Fy * Db / Sqr(Concrete)
                    End If
                    Ld = CorrectionFactor * Ldb
                    Ldh = 0.06 * Fy * Db / Sqr(Concrete)
                    ClassBSpliceTop = Application.RoundUp(Application.Max(1.3 * 1.3 * Ld, 3.25 * Ldh, 30), 0)
                    ClassBSpliceBot = Application.RoundUp(Application.Max(1.3 * Ld, 2.5 * Ldh, 30), 0)

                    WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = ClassBSpliceTop
                    WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = ClassBSpliceBot
                Next
                For ReinforcementNumber = 12 To MaximumNumber Step 2
                    PlotColumnNumber = ReinforcementNumber / 2 + 9

                    Cs = (BeamWidth(I) - Db * ReinforcementNumber - TieDiameter * 2 - Cover * 2) / 2 / (ReinforcementNumber - 1)
                    Cb = Db / 2 + Application.Min(Cc, Cs)
                    If Cs < Cc Then
                        Ktr = 2 * Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing / ReinforcementNumber
                    Else
                        Ktr = Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing
                    End If
                    CorrectionFactor = Application.Max(0.4, Db / (Cb + Ktr))
                    If Db < 2 Then
                        Ldb = 0.28 * 0.8 * Fy * Db / Application.Sqrt(Concrete)
                    Else
                        Ldb = 0.28 * Fy * Db / Sqr(Concrete)
                    End If
                    Ld = CorrectionFactor * Ldb
                    Ldh = 0.06 * Fy * Db / Sqr(Concrete)
                    ClassBSpliceTop = Application.RoundUp(Application.Max(1.3 * 1.3 * Ld, 3.25 * Ldh, 30), 0)
                    ClassBSpliceBot = Application.RoundUp(Application.Max(1.3 * Ld, 2.5 * Ldh, 30), 0)

                    WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = ClassBSpliceTop
                    WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = ClassBSpliceBot
                Next
            End If
        Next
    Next                
End Sub

Sub BeamDesign(PlotNumber, BeamPositionColumn)
    Dim BeamWidthNumber As Integer, BeamWidth(20) As Integer ', LastRowNumber As Integer, CountColumnNumber As Integer, PlotRowNumber1 As Integer
    'Dim PlotRowNumber2 As Integer, Fyt As Integer, Concrete As Integer, TieSpacing As Integer, CountRowNumber As Integer
    'Dim MaximumNumber As Integer, PlotColumnNumber As Integer, j As Integer, ReinforcementNumber As Integer, I As Integer
    'Dim Db As Double, TieDiameter As Double
    'Dim WsInput As Worksheet, WsPlot As Worksheet
    
    Set WsInput = Worksheets("輸入")
    Set WsPlot = Worksheets(PlotNumber)
    WsPlot.Cells.Delete
    BeamWidthNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5
    LastRowNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 2).End(xlUp).Row    
    PlotRowNumber1 = 1
    PlotRowNumber2 = 2
    
    For j = 30 To LastRowNumber
        Db = WsInput.Cells(j, BeamPositionColumn + 2).Value / 10
        TieDiameter = WsInput.Cells(j, BeamPositionColumn + 3).Value / 10
        TieSpacing = WsInput.Cells(j, BeamPositionColumn + 4).Value
        Fyt = WsInput.Cells(j, BeamPositionColumn + 5).Value
        Concrete = WsInput.Cells(j, BeamPositionColumn + 6).Value
        Cover = 4
        Fy = 4200            
        PlotRowNumber1 = PlotRowNumber1 + 5
        PlotRowNumber2 = PlotRowNumber2 + 5
        For I = 0 To BeamWidthNumber - 1
            BeamWidth(I) = WsInput.Cells(I + 6, BeamPositionColumn + 1).Value
            MaximumNumber = Fix((BeamWidth(I) - 4 * 2 - TieDiameter * 2 - Db) / (2 * Db)) + 1
            Cc = Cover + TieDiameter            
            PlotRowNumber1 = PlotRowNumber1 + 2
            PlotRowNumber2 = PlotRowNumber2 + 2
            If MaximumNumber <= 10 Then
                For ReinforcementNumber = 2 To MaximumNumber                        
                    PlotColumnNumber = ReinforcementNumber + 4
                    Cs = (BeamWidth(I) - Db * ReinforcementNumber - TieDiameter * 2 - Cover * 2) / 2 / (ReinforcementNumber - 1)
                    Cb = Db / 2 + Application.Min(Cc, Cs)
                    If Cs < Cc Then
                        Ktr = 2 * Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing / ReinforcementNumber
                    Else
                        Ktr = Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing
                    End If
                    CorrectionFactor = Application.Max(0.4, Db / (Cb + Ktr))
                    If Db < 2 Then
                        Ldb = 0.28 * 0.8 * Fy * Db / Application.Sqrt(Concrete)
                    Else
                        Ldb = 0.28 * Fy * Db / Sqr(Concrete)
                    End If
                    Ld = CorrectionFactor * Ldb                    
                    ClassBSpliceTop = Application.RoundUp(Application.Max(1.3 * 1.3 * Ld, 30), 0)
                    ClassBSpliceBot = Application.RoundUp(Application.Max(1.3 * Ld, 30), 0)
                    WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = ClassBSpliceTop
                    WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = ClassBSpliceBot
                Next
            Else
                For ReinforcementNumber = 2 To 10
                    PlotColumnNumber = ReinforcementNumber + 4
                    Cs = (BeamWidth(I) - Db * ReinforcementNumber - TieDiameter * 2 - Cover * 2) / 2 / (ReinforcementNumber - 1)
                    Cb = Db / 2 + Application.Min(Cc, Cs)
                    If Cs < Cc Then
                        Ktr = 2 * Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing / ReinforcementNumber
                    Else
                        Ktr = Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing
                    End If
                    CorrectionFactor = Application.Max(0.4, Db / (Cb + Ktr))
                    If Db < 2 Then
                        Ldb = 0.28 * 0.8 * Fy * Db / Application.Sqrt(Concrete)
                    Else
                        Ldb = 0.28 * Fy * Db / Sqr(Concrete)
                    End If
                    Ld = CorrectionFactor * Ldb                    
                    ClassBSpliceTop = Application.RoundUp(Application.Max(1.3 * 1.3 * Ld, 30), 0)
                    ClassBSpliceBot = Application.RoundUp(Application.Max(1.3 * Ld, 30), 0)
                    WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = ClassBSpliceTop
                    WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = ClassBSpliceBot
                Next
                For ReinforcementNumber = 12 To MaximumNumber Step 2
                    PlotColumnNumber = ReinforcementNumber / 2 + 9
                    Cs = (BeamWidth(I) - Db * ReinforcementNumber - TieDiameter * 2 - Cover * 2) / 2 / (ReinforcementNumber - 1)
                    Cb = Db / 2 + Application.Min(Cc, Cs)
                    If Cs < Cc Then
                        Ktr = 2 * Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing / ReinforcementNumber
                    Else
                        Ktr = Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing
                    End If
                    CorrectionFactor = Application.Max(0.4, Db / (Cb + Ktr))
                    If Db < 2 Then
                        Ldb = 0.28 * 0.8 * Fy * Db / Application.Sqrt(Concrete)
                    Else
                        Ldb = 0.28 * Fy * Db / Sqr(Concrete)
                    End If
                    Ld = CorrectionFactor * Ldb                    
                    ClassBSpliceTop = Application.RoundUp(Application.Max(1.3 * 1.3 * Ld, 30), 0)
                    ClassBSpliceBot = Application.RoundUp(Application.Max(1.3 * Ld, 30), 0)
                    WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = ClassBSpliceTop
                    WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = ClassBSpliceBot
                Next
            End If
        Next
    Next                
End Sub

Sub WallSeismicDesign(PlotNumber, BeamPositionColumn)
   Dim BeamWidthNumber As Integer, BeamWidth(20) As Integer ', LastRowNumber As Integer, CountColumnNumber As Integer, PlotRowNumber1 As Integer
   Dim PlotRowNumber2 As Integer, Fyt As Integer, Concrete As Integer, TieSpacing As Integer, CountRowNumber As Integer
   Dim MaximumNumber As Integer, PlotColumnNumber As Integer, j As Integer, ReinforcementNumber As Integer, I As Integer
   Dim Db As Double, TieDiameter As Double
   Dim WsInput As Worksheet, WsPlot As Worksheet
    
    Set WsInput = Worksheets("輸入")
    Set WsPlot = Worksheets(PlotNumber)
    WsPlot.Cells.Delete
    'BeamWidthNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5
    LastRowNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 2).End(xlUp).Row
    
    PlotRowNumber1 = 1
    'PlotRowNumber2 = 2
    
    For j = 30 To LastRowNumber
        Db = WsInput.Cells(j, BeamPositionColumn + 1).Value / 10
        Cover = WsInput.Cells(j, BeamPositionColumn + 2).Value
        Spacing = WsInput.Cells(j, BeamPositionColumn + 3).Value
        Fy = WsInput.Cells(j, BeamPositionColumn + 4).Value
        Concrete = WsInput.Cells(j, BeamPositionColumn + 6).Value
        TieDiameter = 0
        TieSpacing = 0
        Fyt = 0

        'CountRowNumber = 10
        'PlotRowNumber1 = PlotRowNumber1 + 5
        'PlotRowNumber2 = PlotRowNumber2 + 5
        'For I = 0 To BeamWidthNumber - 1
        'BeamWidth(I) = WsInput.Cells(I + 6, BeamPositionColumn + 1).Value
        'MaximumNumber = Fix((BeamWidth(I) - 4 * 2 - TieDiameter * 2 - Db) / (2 * Db)) + 1
        Cc = Cover + TieDiameter
        
        PlotRowNumber1 = PlotRowNumber1 + 2
        'PlotRowNumber2 = PlotRowNumber2 + 2                
        'For ReinforcementNumber = 2 To MaximumNumber

        'CountRowNumber = CountRowNumber + 1
        PlotColumnNumber = 4

        Cs = (Spacing - Db) / 2 
        Cb = Db / 2 + Application.Min(Cc, Cs)
        'If Cs < Cc Then
        '    Ktr = 2 * Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing / ReinforcementNumber
        'Else
        '    Ktr = Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing
        'End If
        CorrectionFactor = Application.Max(0.4, Db / Cb )
        If Db < 2 Then
            Ldb = 0.28 * 0.8 * Fy * Db / Application.Sqrt(Concrete)
        Else
            Ldb = 0.28 * Fy * Db / Sqr(Concrete)
        End If
        Ld = CorrectionFactor * Ldb
        Ldh = 0.06 * Fy * Db / Sqr(Concrete)
        ClassBSpliceTop = Application.RoundUp(Application.Max(1.3 * 1.3 * Ld, 3.25 * Ldh, 30), 0)
        ClassBSpliceBot = Application.RoundUp(Application.Max(1.3 * Ld, 2.5 * Ldh, 30), 0)

        WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = ClassBSpliceTop
        WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = ClassBSpliceBot
        'Next                
        'Next
    Next  
End Sub

Sub WallDesign(PlotNumber, BeamPositionColumn)
   Dim BeamWidthNumber As Integer, BeamWidth(20) As Integer ', LastRowNumber As Integer, CountColumnNumber As Integer, PlotRowNumber1 As Integer
   Dim PlotRowNumber2 As Integer, Fyt As Integer, Concrete As Integer, TieSpacing As Integer, CountRowNumber As Integer
   Dim MaximumNumber As Integer, PlotColumnNumber As Integer, j As Integer, ReinforcementNumber As Integer, I As Integer
   Dim Db As Double, TieDiameter As Double
   Dim WsInput As Worksheet, WsPlot As Worksheet
    
    Set WsInput = Worksheets("輸入")
    Set WsPlot = Worksheets(PlotNumber)
    WsPlot.Cells.Delete
    'BeamWidthNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5
    LastRowNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 2).End(xlUp).Row
    
    PlotRowNumber1 = 1
    'PlotRowNumber2 = 2
    
    For j = 30 To LastRowNumber
        Db = WsInput.Cells(j, BeamPositionColumn + 1).Value / 10
        Cover = WsInput.Cells(j, BeamPositionColumn + 2).Value
        Spacing = WsInput.Cells(j, BeamPositionColumn + 3).Value
        Fy = WsInput.Cells(j, BeamPositionColumn + 4).Value
        Concrete = WsInput.Cells(j, BeamPositionColumn + 6).Value
        TieDiameter = 0
        TieSpacing = 0
        Fyt = 0

        'CountRowNumber = 10
        'PlotRowNumber1 = PlotRowNumber1 + 5
        'PlotRowNumber2 = PlotRowNumber2 + 5
        'For I = 0 To BeamWidthNumber - 1
        'BeamWidth(I) = WsInput.Cells(I + 6, BeamPositionColumn + 1).Value
        'MaximumNumber = Fix((BeamWidth(I) - 4 * 2 - TieDiameter * 2 - Db) / (2 * Db)) + 1
        Cc = Cover + TieDiameter
        
        PlotRowNumber1 = PlotRowNumber1 + 2
        'PlotRowNumber2 = PlotRowNumber2 + 2                
        'For ReinforcementNumber = 2 To MaximumNumber

        'CountRowNumber = CountRowNumber + 1
        PlotColumnNumber = 4

        Cs = (Spacing - Db) / 2 
        Cb = Db / 2 + Application.Min(Cc, Cs)
        'If Cs < Cc Then
        '    Ktr = 2 * Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing / ReinforcementNumber
        'Else
        '    Ktr = Application.Pi() * TieDiameter ^ 2 / 4 * Fyt / 105 / TieSpacing
        'End If
        CorrectionFactor = Application.Max(0.4, Db / Cb )
        If Db < 2 Then
            Ldb = 0.28 * 0.8 * Fy * Db / Application.Sqrt(Concrete)
        Else
            Ldb = 0.28 * Fy * Db / Sqr(Concrete)
        End If
        Ld = CorrectionFactor * Ldb
        Ldh = 0.06 * Fy * Db / Sqr(Concrete)
        ClassBSpliceTop = Application.RoundUp(Application.Max(1.3 * 1.3 * Ld, 3.25 * Ldh, 30), 0)
        ClassBSpliceBot = Application.RoundUp(Application.Max(1.3 * Ld, 2.5 * Ldh, 30), 0)

        WsPlot.Cells(PlotRowNumber1, PlotColumnNumber).Value = ClassBSpliceTop
        WsPlot.Cells(PlotRowNumber2, PlotColumnNumber).Value = ClassBSpliceBot
        'Next                
        'Next
    Next  
End Sub

Sub BeamFormatText(PlotNumber, BeamPositionColumn)
    Dim BeamWidthNumber2 As Integer, BeamWidth(20) As Integer, LastColumnNumber As Integer, BeamWidthNumber As Integer, TableNumber As Integer, CountRowNumber As Integer
    Dim I As Integer, j As Integer, CalculateCountColumnNumber As Integer, LimitColumnWidth As Double
    Dim WsInput As Worksheet, WsPlot As Worksheet, WsCalculate As Worksheet
    Set WsInput = Worksheets("輸入")
    Set WsPlot = Worksheets(PlotNumber)
    'Set WsCalculate = Worksheets(CalculateNumber)
    
    WsPlot.Activate
    Cells.HorizontalAlignment = xlCenter
    Cells.Font.Name = "微軟正黑體"
    WsPlot.Columns(5).ColumnWidth = 10
    LastColumnNumber = WsPlot.UsedRange.Columns.Count + 5
    BeamWidthNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5
    BeamWidthNumber2 = (WsInput.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5) * 2
    TableNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 2).End(xlUp).Row - 29
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


        '數值
        Cells(CountRowNumber + 2, 5) = "梁寬\主筋根數"
        Cells(CountRowNumber + 2, 5).Characters(Start:=1, Length:=2).Font.Subscript = True
        Cells(CountRowNumber + 2, 5).Characters(Start:=4, Length:=4).Font.Superscript = True
        For j = 0 To BeamWidthNumber - 1
            WsPlot.Cells(CountRowNumber + 3 + j * 2, 5) = WsInput.Cells(j + 6, BeamPositionColumn + 1).Value
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
        'Cells(CountRowNumber + 1, 5) = "受拉竹節鋼筋搭接長度（乙級搭接）"


        '框線
        Range(Cells(CountRowNumber + 2, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlInsideVertical).Weight = xlThin
        Range(Cells(CountRowNumber + 2, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlInsideHorizontal).Weight = xlThin
        Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeLeft).Weight = xlMedium
        Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeTop).Weight = xlMedium
        Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeBottom).Weight = xlMedium
        Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeRight).Weight = xlMedium
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
        If LimitColumnWidth < 115 Then
            For I = 6 To LastColumnNumber
                WsPlot.Columns(I).ColumnWidth = (115 - 10) / (LastColumnNumber - 5)
            Next
        End If
    End If
End Sub

Sub WallFormatText(PlotNumber, BeamPositionColumn)
    Dim BeamWidthNumber2 As Integer, BeamWidth(20) As Integer, LastColumnNumber As Integer, BeamWidthNumber As Integer, TableNumber As Integer, CountRowNumber As Integer
    Dim I As Integer, j As Integer, CalculateCountColumnNumber As Integer, LimitColumnWidth As Double
    Dim WsInput As Worksheet, WsPlot As Worksheet, WsCalculate As Worksheet
    Set WsInput = Worksheets("輸入")
    Set WsPlot = Worksheets(PlotNumber)
    'Set WsCalculate = Worksheets(CalculateNumber)
    
    WsPlot.Activate
    Cells.HorizontalAlignment = xlCenter
    Cells.Font.Name = "微軟正黑體"
    WsPlot.Columns(5).ColumnWidth = 10
    LastColumnNumber = WsPlot.UsedRange.Columns.Count + 5
    BeamWidthNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5
    BeamWidthNumber2 = (WsInput.Cells(Rows.Count, BeamPositionColumn + 1).End(xlUp).Row - 5) * 2
    TableNumber = WsInput.Cells(Rows.Count, BeamPositionColumn + 2).End(xlUp).Row - 29
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


        '數值
        Cells(CountRowNumber + 2, 5) = "梁寬\主筋根數"
        Cells(CountRowNumber + 2, 5).Characters(Start:=1, Length:=2).Font.Subscript = True
        Cells(CountRowNumber + 2, 5).Characters(Start:=4, Length:=4).Font.Superscript = True
        For j = 0 To BeamWidthNumber - 1
            WsPlot.Cells(CountRowNumber + 3 + j * 2, 5) = WsInput.Cells(j + 6, BeamPositionColumn + 1).Value
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
        Cells(CountRowNumber + 1, 5) = "受拉竹節鋼筋搭接長度（乙級搭接）"


        '框線
        Range(Cells(CountRowNumber + 2, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlInsideVertical).Weight = xlThin
        Range(Cells(CountRowNumber + 2, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlInsideHorizontal).Weight = xlThin
        Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeLeft).Weight = xlMedium
        Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeTop).Weight = xlMedium
        Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeBottom).Weight = xlMedium
        Range(Cells(CountRowNumber, 5), Cells(CountRowNumber + 3 + BeamWidthNumber2 - 1, LastColumnNumber)).Borders(xlEdgeRight).Weight = xlMedium
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
        If LimitColumnWidth < 115 Then
            For I = 6 To LastColumnNumber
                WsPlot.Columns(I).ColumnWidth = (115 - 10) / (LastColumnNumber - 5)
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




