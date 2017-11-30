Sub SRCSelector()
' 目的：
' 由於在ETABS不會 Design SRC 斷面，所以由 ETABS 輸出 PMM。
' 以 SectionBuilder 建立 SRC 斷面，產生包絡線，檢測 ETABS PMM 有沒有在包絡線裡面。


' 演算法：
' 1. PMM curve 取 0 45 90 度
' 2. 由於 P 不一定會相同，排序內差求值
' 3. 以 PMM 點求得該 P 下的 0 45 90 度的 M
' 4. 以 PMM 點 M2 M3 判斷要和哪一條線比較
' 5. 以牛頓法判斷是不是與 (0, 0) 同側


' 執行時間：
' 1.41s 7 萬資料量
' 6.9s 40 萬資料量
'
' 增加 Ratio 計算後的執行時間：
' 32.36s 40 萬資料量
' 重構程式碼後的執行時間：
' 21.61s 40 萬資料量
'

    Time0 = Timer

    Call AutoFill

    ComboPMM = Combo()

    PMMCurve = ReadPMMCurve()

    SelectionSection = SelectionSelector(ComboPMM)



End Function


Function AutoFill()

' 公式自動填滿

    Worksheets("EtabsPMMCombo").Activate
    ComboRowUsed = Cells(Rows.Count, 1).End(xlUp).Row

    Worksheets("PMM").Activate
    Range(Cells(2, 1), Cells(2, 4)).AutoFill Destination := Range(Cells(2, 1), Cells(ComboRowUsed, 4))

End Function


Function Combo()

' 讀取每個 Combo
' 資料格式：
' Name P M2 M3

    Worksheets("PMM").Activate
    Dim ComboPMM()
    ComboRowUsed = Cells(Rows.Count, 1).End(xlUp).Row
    ReDim ComboPMM(2 to ComboRowUsed + 1, 1 to 4)

    ' 讀取所有的PMM
    For ComboRowNumber = 2 To ComboRowUsed

        ' Name
        ComboPMM(ComboRowNumber, 1) = Cells(ComboRowNumber, 1)

        ' P
        ComboPMM(ComboRowNumber, 2) = Cells(ComboRowNumber, 2)

        ' M2
        ComboPMM(ComboRowNumber, 3) = Cells(ComboRowNumber, 3)

        ' M3
        ComboPMM(ComboRowNumber, 4) = Cells(ComboRowNumber, 4)

    Next

    ' 多增加一列，並給最後一個不一樣的值，為下一步的演算法做準備，免得無法比較
    ComboPMM(ComboRowUsed + 1, 0) = 0

    Combo = ComboPMM()

End Function


Function ReadPMMCurve() As functionType




End Function


Function SelectionSelector(ComboPMM)

    ' 使輸出結果陣列與 ComboPMM 相同
    Dim SelectionSection()
    ReDim SelectionSection(UBound(ComboPMM), 4)

    Dim PMMArray()
    Dim PMMCurveName()

    ' 讀取PMMCurve最後一列
    Worksheets("PMMCurve").Activate
    PMMCurveRowUsed = Cells(Rows.Count, 3).End(xlUp).Row

    ' 傳回PMMCurve數目
    PMMNumber = (PMMCurveRowUsed - 25) / 24 + 1

    ' 讀取有幾個PMMCurve
    ReDim PMMArray(1 To PMMNumber)

    ' 多的1為例外作準備
    ReDim PMMCurveName(1 To PMMNumber + 1)
    PMMCurveName(PMMNumber + 1) = "超過所有斷面，請選擇更大的斷面！"

    ' 最佳化，不用每一次都進去跑Loop，先把所有陣列寫好
    For PMMCurveRowNumber = 5 To PMMCurveRowUsed Step 24
        i = i + 1
        PMMArray(i) = PMMCurve(PMMCurveRowNumber + 1)
        PMMCurveName(i) = Cells(PMMCurveRowNumber, 1)
    Next



    ' 從第1筆資料Loop到最後一筆
    For RowNumber = 0 To UBound(ComboPMM) - 1

        ' 看看他與下一筆資料相不相同，如果相同就是一組。
        If ComboPMM(RowNumber, 0) <> ComboPMM(RowNumber + 1, 0) Then

            EndNumber = RowNumber

            ' 每一個Column（包含很多個Combo）重新初始化
            FinalSelectionNumber = 0
            FinalRatio = 0

            ' 相同的一組
            For ColumnNumber = StartNumber To EndNumber

                ' 每一個Combo重新初始化
                SelectionNumber = 0
                Ratio = 0

                For SelectionNumber = 1 To PMMNumber

                    PMM = PMMArray(SelectionNumber)

                    ' 19條線
                    For LineNumber = 1 To 19

                        ' PMM的資料格式：
                        ' M P Angle b c
                        ' ComboPMM的資料格式：
                        ' Name M P Angle
                        If Newton(ComboPMM(ColumnNumber, 1), PMM(LineNumber, 3), ComboPMM(ColumnNumber, 2), PMM(LineNumber, 4), PMM(LineNumber - 1, 2), PMM(LineNumber, 2), ComboPMM(ColumnNumber, 3)) Then
                            Ratio = CaculateRatio(ComboPMM(ColumnNumber, 1), ComboPMM(ColumnNumber, 2), PMM(LineNumber, 3), PMM(LineNumber, 4))
                            GoTo NextCombo
                        End If

                    Next
                Next



NextCombo:
                ' Combo Loop 結束
                ' 超出所有PMMCurve，例外處理
                If SelectionNumber = 0 Then
                    SelectionNumber = PMMNumber + 1
                    SelectionSection(ColumnNumber, 4) = PMMNumber + 1
                Else
                    SelectionSection(ColumnNumber, 4) = SelectionNumber
                End If



                ' 判斷有沒有大於FinalSelectionNumber，有的話才寫入
                If FinalSelectionNumber < SelectionNumber Then
                    FinalSelectionNumber = SelectionNumber
                    FinalRatio = Ratio
                End If

                ' 判斷有沒有大於Ratio，有的話才寫入
                If FinalRatio < Ratio And FinalSelectionNumber <= SelectionNumber Then
                    FinalRatio = Ratio
                End If

            Next


            ' 斷面的Loop 結束
            ' 寫入斷面資料
            SelectionSection(SelectionSectionNumber, 0) = ComboPMM(RowNumber, 0)
            SelectionSection(SelectionSectionNumber, 1) = FinalSelectionNumber
            SelectionSection(SelectionSectionNumber, 2) = PMMCurveName(FinalSelectionNumber)
            SelectionSection(SelectionSectionNumber, 3) = FinalRatio

            ' 下一組的開始編號
            StartNumber = RowNumber + 1

            ' 下一組
            SelectionSectionNumber = SelectionSectionNumber + 1

        End If

    Next

    SelectionSelector = SelectionSection()

End Function
