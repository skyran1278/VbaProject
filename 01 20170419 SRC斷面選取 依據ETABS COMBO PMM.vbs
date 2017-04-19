Sub SRCSeltor()

' 目的：由於在ETABS不會 Design SRC斷面，所以由ETABS輸出PMM。
'       以SectionBuilder建立SRC斷面，產生包絡線，看PMM有沒有在選取的斷面裡面。
'
' 演算法：由SectionBuilder的20個點，產生19條方程式，用牛頓法看有沒有和(0,0)在一起。
'
' 執行時間：1.41s 7萬資料量
'           6.9s 40萬資料量
'
    
    Time0 = Timer
    
    PMM1 = PMMCurve(6)
    PMM2 = PMMCurve(29)
    PMM3 = PMMCurve(52)
    PMM4 = PMMCurve(76)
    PMM5 = PMMCurve(100)
    PMM6 = PMMCurve(124)
    
    ComboPMM = Combo()

    SelectionSection = CreatFunction(PMM1, PMM2, PMM3, PMM4, PMM5, PMM6, ComboPMM)

    Range(Cells(1, 16), Cells(UBound(SelectionSection), 18)) = SelectionSection

    MsgBox "執行時間 " & Timer - Time0 & " 秒", vbOKOnly

End Sub

Function PMMCurve(RowNumber)

    Dim PMM(19, 3) As Double

    Worksheets("PMMCurve").Activate

    ' 讀取PMM
    For RowNumberCount = 0 To 19
        PMM(RowNumberCount, 0) = Cells(RowNumber + RowNumberCount, 4)
        PMM(RowNumberCount, 1) = Cells(RowNumber + RowNumberCount, 3)
    Next

    ' x + b * y + c = 0
    ' b = PMM(RowNumberCount, 2)
    ' c = PMM(RowNumberCount, 3)
    For RowNumberCount = 1 To 19
        PMM(RowNumberCount, 2) = -(PMM(RowNumberCount, 0) - PMM(RowNumberCount - 1, 0)) / (PMM(RowNumberCount, 1) - PMM(RowNumberCount - 1, 1))
        PMM(RowNumberCount, 3) = -PMM(RowNumberCount - 1, 0) - PMM(RowNumberCount, 0) * PMM(RowNumberCount - 1, 1)
    Next

    PMMCurve = PMM()

End Function

Function Combo()

    Worksheets("EtabsPMMCombo").Activate
    Dim ComboPMM()
    ComboRowUsed = Cells(Rows.Count, 3).End(xlUp).Row
    ReDim ComboPMM(ComboRowUsed - 1, 2)

    ' 讀取所有的PMM
    For ComboRowNumber = 2 To ComboRowUsed
        ComboPMM(ComboRowNumber - 2, 0) = Cells(ComboRowNumber, 12)
        ComboPMM(ComboRowNumber - 2, 1) = Cells(ComboRowNumber, 13)
        ComboPMM(ComboRowNumber - 2, 2) = Cells(ComboRowNumber, 14)
    Next

    ' 給最後一個不一樣的值，為下一步的演算法做準備，免得無法比較出不同
    ComboPMM(ComboRowUsed - 1, 0) = 0

    Combo = ComboPMM()
    
End Function

Function CreatFunction(PMM1, PMM2, PMM3, PMM4, PMM5, PMM6, ComboPMM)
    
    StartNumber = 0
    SelectionSectionNumber = -1
    Dim SelectionSection()
    ReDim SelectionSection(UBound(ComboPMM), 2)

    For RowNumber = 0 To UBound(ComboPMM) - 1

        ' 看看他與下一筆資料相不相同，如果相同就是一組。
        If ComboPMM(RowNumber, 0) <> ComboPMM(RowNumber + 1, 0) Then

            SelectionSectionNumber = SelectionSectionNumber + 1
            EndNumber = RowNumber
            FinalSelectionNumber = 0

            ' 相同的一組
            For ColumnNumber = StartNumber To EndNumber

                ' 19條線
                For LineNumber = 1 To 19

                    ' (x + b * y + c) * c < 0 牛頓法
                    If (ComboPMM(ColumnNumber, 1) + PMM1(LineNumber, 2) * ComboPMM(ColumnNumber, 2) + PMM1(LineNumber, 3)) * PMM1(LineNumber, 3) > 0 Then
                        If FinalSelectionNumber < 1 Then
                            FinalSelectionNumber = 1
                        End If
                    
                    ElseIf (ComboPMM(ColumnNumber, 1) + PMM2(LineNumber, 2) * ComboPMM(ColumnNumber, 2) + PMM2(LineNumber, 3)) * PMM2(LineNumber, 3) > 0 Then
                        If FinalSelectionNumber < 2 Then
                            FinalSelectionNumber = 2
                        End If
                        
                    ElseIf (ComboPMM(ColumnNumber, 1) + PMM3(LineNumber, 2) * ComboPMM(ColumnNumber, 2) + PMM3(LineNumber, 3)) * PMM3(LineNumber, 3) > 0 Then
                        If FinalSelectionNumber < 3 Then
                            FinalSelectionNumber = 3
                        End If
                            
                    ElseIf (ComboPMM(ColumnNumber, 1) + PMM4(LineNumber, 2) * ComboPMM(ColumnNumber, 2) + PMM4(LineNumber, 3)) * PMM4(LineNumber, 3) > 0 Then
                        If FinalSelectionNumber < 4 Then
                            FinalSelectionNumber = 4
                        End If
                                
                    ElseIf (ComboPMM(ColumnNumber, 1) + PMM5(LineNumber, 2) * ComboPMM(ColumnNumber, 2) + PMM5(LineNumber, 3)) * PMM5(LineNumber, 3) > 0 Then
                       If FinalSelectionNumber < 5 Then
                           FinalSelectionNumber = 5
                       End If
                                    
                    ElseIf (ComboPMM(ColumnNumber, 1) + PMM6(LineNumber, 2) * ComboPMM(ColumnNumber, 2) + PMM6(LineNumber, 3)) * PMM6(LineNumber, 3) > 0 Then
                        FinalSelectionNumber = 7

                    End If
                Next
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

        End If

    Next

    CreatFunction = SelectionSection()

End Function
