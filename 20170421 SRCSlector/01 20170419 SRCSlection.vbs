Sub SRCSeltor()


' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！
' 最新程式碼，不要覆蓋！！！！！！！！！！！

' 需要增加一點註釋，不然絕對看不懂
' 目的：
' 由於在ETABS不會 Design SRC斷面，所以由ETABS輸出PMM。
' 以SectionBuilder建立SRC斷面，產生包絡線，看PMM有沒有在選取的斷面裡面。
'
' 演算法：由SectionBuilder的20個點，產生19條方程式，用牛頓法看有沒有和((0,0)或是(650,2000))在一起。
'
' 執行時間：
' 1.41s 7萬資料量
' 6.9s 40萬資料量
'
' 增加Ratio 計算後的執行時間：
' 32.36s 40萬資料量
'
'

    Time0 = Timer

    ' PMM1 = PMMCurve(6)
    ' PMM2 = PMMCurve(30)
    ' PMM3 = PMMCurve(54)
    ' PMM4 = PMMCurve(78)
    ' PMM5 = PMMCurve(102)
    ' PMM6 = PMMCurve(126)

    ComboPMM = Combo()

    ' SelectionSection = CreatFunction(PMM1, PMM2, PMM3, PMM4, PMM5, PMM6, ComboPMM)
    SelectionSection = CreatFunction(ComboPMM)

    Range(Columns(15), Columns(18)).ClearContents

    Range(Cells(2, 15), Cells(UBound(SelectionSection) + 1, 19)) = SelectionSection

    Worksheets("SectionSelector").Activate

    Range(Columns(11), Columns(14)).ClearContents

    Cells(1, 11) = "Column"
    Cells(1, 12) = "NO."
    Cells(1, 13) = "SectionName"
    Cells(1, 14) = "Ratio"

    Range(Cells(2, 11), Cells(UBound(SelectionSection), 14)) = SelectionSection

    MsgBox "執行時間 " & Timer - Time0 & " 秒", vbOKOnly

End Sub

Function Asin(X)

    If X = -1 Then
        Asin = -Pi / 2

    ElseIf X = 1 Then
        Asin = Pi / 2

    Else
        Asin = Atn(X / Sqr(1 - X ^ 2))

    End If

End Function

Function PMMCurve(RowNumber)

    Dim PMM(19, 4) As Double

    Worksheets("PMMCurve").Activate

    ' 讀取PMM
    For RowNumberCount = 0 To 19
        ' M
        PMM(RowNumberCount, 0) = Cells(RowNumber + RowNumberCount, 4)
        ' P
        PMM(RowNumberCount, 1) = Cells(RowNumber + RowNumberCount, 3)
        ' 角度
        PMM(RowNumberCount, 4) = Asin(PMM(RowNumberCount, 1) / Sqr(PMM(RowNumberCount, 0) ^ 2 + PMM(RowNumberCount, 1) ^ 2))
    Next

    ' x + b * y + c = 0
    ' M + b * P + c = 0
    ' b = PMM(RowNumberCount, 2)
    ' c = PMM(RowNumberCount, 3)
    For RowNumberCount = 1 To 19
        PMM(RowNumberCount, 2) = -(PMM(RowNumberCount, 0) - PMM(RowNumberCount - 1, 0)) / (PMM(RowNumberCount, 1) - PMM(RowNumberCount - 1, 1))
        PMM(RowNumberCount, 3) = -PMM(RowNumberCount - 1, 0) - PMM(RowNumberCount, 2) * PMM(RowNumberCount - 1, 1)
    Next

    ' 資料格式：
    ' M P b c
    PMMCurve = PMM()

End Function

Function Combo()

    Worksheets("EtabsPMMCombo").Activate
    Dim ComboPMM()
    ComboRowUsed = Cells(Rows.Count, 3).End(xlUp).Row
    ReDim ComboPMM(ComboRowUsed - 1, 3)

    ' 讀取所有的PMM
    For ComboRowNumber = 2 To ComboRowUsed
        ComboPMM(ComboRowNumber - 2, 0) = Cells(ComboRowNumber, 12)
        ComboPMM(ComboRowNumber - 2, 1) = Cells(ComboRowNumber, 13)
        ComboPMM(ComboRowNumber - 2, 2) = Cells(ComboRowNumber, 14)
        ComboPMM(ComboRowNumber - 2, 3) = Asin(ComboPMM(ComboRowNumber - 2, 2) / Sqr(ComboPMM(ComboRowNumber - 2, 1) ^ 2 + ComboPMM(ComboRowNumber - 2, 2) ^ 2))
    Next

    ' 給最後一個不一樣的值，為下一步的演算法做準備，免得無法比較出不同
    ComboPMM(ComboRowUsed - 1, 0) = 0

    ' 資料格式：
    ' Name M P
    Combo = ComboPMM()

End Function

Function CreatFunction(ComboPMM)

    StartNumber = 0
    SelectionSectionNumber = -1

    Dim SelectionSection()
    ReDim SelectionSection(UBound(ComboPMM), 4)

    PMM1 = PMMCurve(6)
    PMM2 = PMMCurve(30)
    PMM3 = PMMCurve(54)
    PMM4 = PMMCurve(78)
    PMM5 = PMMCurve(102)
    PMM6 = PMMCurve(126)

    ' For i = 6 To 126 step 24
    '     PMM = PMMCurve(i)
    ' Next

    For RowNumber = 0 To UBound(ComboPMM) - 1

        ' 看看他與下一筆資料相不相同，如果相同就是一組。
        If ComboPMM(RowNumber, 0) <> ComboPMM(RowNumber + 1, 0) Then

            SelectionSectionNumber = SelectionSectionNumber + 1
            EndNumber = RowNumber

            ' 每一個Column（包含很多個Combo）重新初始化
            FinalSelectionNumber = 0
            FinalRatio = 0

            ' 相同的一組
            For ColumnNumber = StartNumber To EndNumber

                ' 每一個Combo重新初始化
                SelectionNumber = 0
                MaxRatio = 0

                For i = 1 To 6

                    Select Case i
                        Case 1
                            PMM = PMM1
                        Case 2
                            PMM = PMM2
                        Case 3
                            PMM = PMM3
                        Case 4
                            PMM = PMM4
                        Case 5
                            PMM = PMM5
                        Case 6
                            PMM = PMM6
                    End Select

                    ' 19條線
                    For LineNumber = 1 To 19

                        ' (x + b * y + c) * c > 0 牛頓法
                        ' (M + b * P + c) * c > 0 牛頓法
                        ' 判斷兩個點是否在同一邊
                        '
                        ' PMM的資料格式：
                        ' M P b c
                        '
                        ' ComboPMM的資料格式：
                        ' Name M P
                        '

                        ' 判斷角度
                        If PMM(LineNumber - 1, 4) < ComboPMM(ColumnNumber, 3) And ComboPMM(ColumnNumber, 3) <= PMM(LineNumber, 4) And (ComboPMM(ColumnNumber, 1) + PMM(LineNumber, 2) * ComboPMM(ColumnNumber, 2) + PMM(LineNumber, 3)) * PMM(LineNumber, 3) > 0 Then

                            MaxRatio = Sqr(ComboPMM(ColumnNumber, 1) ^ 2 + ComboPMM(ColumnNumber, 2) ^ 2) / (Abs(ComboPMM(ColumnNumber, 1) + PMM(LineNumber, 2) * ComboPMM(ColumnNumber, 2) + PMM(LineNumber, 3)) / Sqr(1 + PMM(LineNumber, 2) ^ 2) + Sqr(ComboPMM(ColumnNumber, 1) ^ 2 + ComboPMM(ColumnNumber, 2) ^ 2))

                            ' 這一段很不和邏輯，一定要重構
                            Select Case i
                                Case 1
                                    SelectionNumber = 1
                                Case 2
                                    SelectionNumber = 2
                                Case 3
                                    SelectionNumber = 3
                                Case 4
                                    SelectionNumber = 4
                                Case 5
                                    SelectionNumber = 5
                                Case 6
                                    SelectionNumber = 6
                                Case Else
                                    SelectionNumber = 7
                            End Select

                            GoTo NextCombo

                        End If

                    Next

                Next

NextCombo:

                SelectionSection(ColumnNumber, 4) = SelectionNumber

                ' 判斷有沒有大於FinalSelectionNumber，有的話才寫入
                If FinalSelectionNumber < SelectionNumber Then
                    FinalSelectionNumber = SelectionNumber
                    FinalRatio = MaxRatio
                End If

                If FinalRatio < MaxRatio And FinalSelectionNumber <= SelectionNumber Then
                    FinalRatio = MaxRatio
                End If

            Next

            ' 下一組的開始編號
            StartNumber = RowNumber + 1

            ' 給編號命名，讓人更容易看懂
            Select Case FinalSelectionNumber

                Case 1
                    FinalSelection = "雙H800X150X12X20 12-#10"

                Case 2
                    FinalSelection = "雙H800X150X12X20 20-#10"

                Case 3
                    FinalSelection = "B600X600X20X20 12-#10"

                Case 4
                    FinalSelection = "B600X600X20X20 20-#10"

                Case 5
                    FinalSelection = "B800X800X50X50 20-#10"

                Case 6
                    FinalSelection = "B900X900X50X50 20-#10"

                Case Else
                    FinalSelection = "錯誤，超過所選斷面"
            End Select


            SelectionSection(SelectionSectionNumber, 0) = ComboPMM(RowNumber, 0)
            SelectionSection(SelectionSectionNumber, 1) = FinalSelectionNumber
            SelectionSection(SelectionSectionNumber, 2) = FinalSelection
            SelectionSection(SelectionSectionNumber, 3) = FinalRatio

        End If

    Next

    CreatFunction = SelectionSection()

End Function







