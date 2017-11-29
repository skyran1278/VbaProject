Sub SRCSelector()
' 目的：
' 由於在ETABS不會 Design SRC 斷面，所以由 ETABS 輸出 PMM。
' 以 SectionBuilder 建立 SRC 斷面，產生包絡線，檢測 ETABS PMM 有沒有在包絡線裡面。


' 演算法：
' M2 為 X，M3 為 Y
' 由 SectionBuilder 的 20 個點，產生 19 條方程式與角度。
' 用角度判斷 PMM 點落在哪兩個點之間
' 再判斷 PMM 點是否落在線的左側


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

