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

    time0 = Timer

    Call AutoFill

    combo = ReadCombo()

    curve = ReadCurves()

    SelectionSection = SelectionSelector(combo)

    ExecutionTime (time0)



End Function


Function AutoFill()

' 公式自動填滿

    Worksheets("EtabsPMMCombo").Activate
    comboRowUsed = Cells(Rows.Count, 1).End(xlUp).row

    Worksheets("PMM").Activate
    Range(Cells(2, 1), Cells(2, 4)).AutoFill Destination:=Range(Cells(2, 1), Cells(comboRowUsed, 4))

End Function


Function ReadCombo()

' 讀取每個 Combo
' 資料格式：
' Name P M2 M3

    Worksheets("PMM").Activate
    Dim combo()
    comboRowUsed = Cells(Rows.Count, 1).End(xlUp).row
    ReDim combo(2 To comboRowUsed, 1 To 4)

    ' 讀取所有的PMM
    For row = 2 To comboRowUsed

        ' Name
        combo(row, 1) = Cells(row, 1)

        ' P
        combo(row, 2) = Cells(row, 2)

        ' M2
        combo(row, 3) = Cells(row, 3)

        ' M3
        combo(row, 4) = Cells(row, 4)

    Next

    ReadCombo = combo()

End Function


Function ReadCurves()

    Dim curves()

    ' 定義數值意義
    nameColumn = 2

    ' 讀取PMMCurve最後一列
    Worksheets("PMMCurve").Activate
    curveRowUsed = Cells(Rows.Count, 4).End(xlUp).row

    ' 統計有幾個非空白儲存格
    curveNumber = Application.WorksheetFunction.CountA(Range(Cells(2, nameColumn), Cells(curveRowUsed, nameColumn)))

    ReDim curves(1 To curveNumber)

    For row = 2 To curveRowUsed

        If Cells(row, nameColumn) <> "" Then

            Index = Index + 1

            curves(Index) = ReadCurve(row)

        End If


    Next


End Function


Function ReadCurve(row)

    Dim curve(1 To 60, 3)

    ' 先全部讀取進來
    For Degree = 1 To 3

        Load = Degree * 4 + 1
        mement = Degree * 4 + 2

        For Point = 1 To 20

            pointCumulativeNumber = pointCumulativeNumber + 1

            ' P
            curve(pointCumulativeNumber, 0) = Cells(row + Point, Load)

            ' M
            curve(pointCumulativeNumber, Degree) = Cells(row + Point, mement)

        Next

    Next

    ' 排序
    curve = QuickSort(curve, LBound(curve), UBound(curve))



    curve = ReadCurve()

End Function


Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub


Function SelectionSelector(ComboPMM)

    ' 使輸出結果陣列與 ComboPMM 相同
    Dim SelectionSection()
    ReDim SelectionSection(UBound(ComboPMM), 4)

    Dim PMMArray()
    Dim PMMCurveName()

    ' 讀取PMMCurve最後一列
    Worksheets("PMMCurve").Activate
    PMMCurveRowUsed = Cells(Rows.Count, 3).End(xlUp).row

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


Function ExecutionTime(time0)

    If Timer - time0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - time0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - time0) / 60, 2) & " Min", vbOKOnly
    End If

End Function

