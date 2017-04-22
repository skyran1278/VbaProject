Sub SRCSelectionSeltor()

' 目的：
' 由於在ETABS不會 Design SRC斷面，所以由ETABS輸出PMM。
' 以SectionBuilder建立SRC斷面，產生包絡線，看PMM有沒有在選取的斷面裡面。
'
' 演算法：
' 由SectionBuilder的20個點，產生19條方程式，與角度。
' 用角度判斷點落在哪個位置
' 再用牛頓法看有沒有和(0,0)在一起。
'
' 執行時間：
' 1.41s 7萬資料量
' 6.9s 40萬資料量
' 增加Ratio 計算後的執行時間：
' 32.36s 40萬資料量
' 重構程式碼後的執行時間：
' 21.61s 40萬資料量
'
    Time0 = Timer

    ComboPMM = Combo()

    SelectionSection = SelectionSeltor(ComboPMM)

    ' 寫入資料在EtabsPMMCombo
    Worksheets("EtabsPMMCombo").Activate
    Range(Columns(15), Columns(18)).ClearContents
    Range(Cells(2, 15), Cells(UBound(SelectionSection) + 1, 19)) = SelectionSection

    ' 寫入資料在SectionSelector
    Worksheets("SectionSelector").Activate
    Range(Columns(11), Columns(14)).ClearContents
    Range(Cells(2, 11), Cells(UBound(SelectionSection), 14)) = SelectionSection
    Cells(1, 11) = "Column"
    Cells(1, 12) = "NO."
    Cells(1, 13) = "SectionName"
    Cells(1, 14) = "Ratio"

    MsgBox "執行時間 " & Timer - Time0 & " 秒", vbOKOnly

End Sub

Function Asin(SinValue)

' 計算arcsin

    Const PI = 3.14159265358979

    If SinValue = -1 Then
        Asin = -PI / 2

    ElseIf SinValue = 1 Then
        Asin = PI / 2

    Else
        Asin = Atn(SinValue / Sqr(1 - SinValue ^ 2))

    End If

End Function

Function PMMCurve(RowNumber)

' 讀取PMMCurve
' 輸出資料矩陣格式：
' M P Angle b c

    Dim PMM(19, 4) As Double

    Worksheets("PMMCurve").Activate

    ' 讀取PMM
    For RowNumberCount = 0 To 19

        ' M
        PMM(RowNumberCount, 0) = Cells(RowNumber + RowNumberCount, 4)

        ' P
        PMM(RowNumberCount, 1) = Cells(RowNumber + RowNumberCount, 3)

        ' 計算角度
        PMM(RowNumberCount, 2) = Asin(PMM(RowNumberCount, 1) / Sqr(PMM(RowNumberCount, 0) ^ 2 + PMM(RowNumberCount, 1) ^ 2))

    Next

    ' M + b * P + c = 0
    For RowNumberCount = 1 To 19

        ' b = PMM(RowNumberCount, 3)
        PMM(RowNumberCount, 3) = -(PMM(RowNumberCount, 0) - PMM(RowNumberCount - 1, 0)) / (PMM(RowNumberCount, 1) - PMM(RowNumberCount - 1, 1))

        ' c = PMM(RowNumberCount, 4)
        PMM(RowNumberCount, 4) = -PMM(RowNumberCount - 1, 0) - PMM(RowNumberCount, 3) * PMM(RowNumberCount - 1, 1)
    Next

    PMMCurve = PMM()

End Function

Function Combo()

' 讀取每個Combo
' 資料格式：
' Name M P Angle

    Worksheets("EtabsPMMCombo").Activate
    Dim ComboPMM()
    ComboRowUsed = Cells(Rows.Count, 3).End(xlUp).Row
    ReDim ComboPMM(ComboRowUsed - 1, 3)

    ' 讀取所有的PMM
    For ComboRowNumber = 2 To ComboRowUsed

        ' Name
        ComboPMM(ComboRowNumber - 2, 0) = Cells(ComboRowNumber, 12)

        ' M
        ComboPMM(ComboRowNumber - 2, 1) = Cells(ComboRowNumber, 13)

        ' P
        ComboPMM(ComboRowNumber - 2, 2) = Cells(ComboRowNumber, 14)

        ' Angle
        ComboPMM(ComboRowNumber - 2, 3) = Asin(ComboPMM(ComboRowNumber - 2, 2) / Sqr(ComboPMM(ComboRowNumber - 2, 1) ^ 2 + ComboPMM(ComboRowNumber - 2, 2) ^ 2))
    Next

    ' 多增加一列，並給最後一個不一樣的值，為下一步的演算法做準備，免得無法比較
    ComboPMM(ComboRowUsed - 1, 0) = 0

    Combo = ComboPMM()

End Function

Function Newton(M, b, P, c, PMMSmallAngle, PMMBigAngle, ComboAngle)

' 傳回是否為True
' 判斷ComboAngle有沒有介在PMMAngle之間
' 有的話判斷與(0,0)是否在同一邊
' (M + b * P + c) * c > 0 牛頓法
' (M + b * P + c) < 0 也可以（在線的左邊）

    Newton = PMMSmallAngle < ComboAngle And ComboAngle < PMMBigAngle And (M + b * P + c) < 0

End Function

Function CaculateRatio(M, P, b, c)

' 傳回Ratio
' Ratio = 點到(0,0)距離 / (點到直線距離 + 點到(0,0)距離)

    CaculateRatio = Sqr(M ^ 2 + P ^ 2) / (Abs(M + b * P + c) / Sqr(1 + b ^ 2) + Sqr(M ^ 2 + P ^ 2))

End Function

Function SelectionSeltor(ComboPMM)

    ' 使輸出結果陣列與ComboPMM相同
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

    SelectionSeltor = SelectionSection()

End Function



