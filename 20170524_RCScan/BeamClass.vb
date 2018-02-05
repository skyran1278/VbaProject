Private MESSAGE(), Worksheet, GENERAL_INFORMATION, REBAR_SIZE, RAW_DATA, RATIO_DATA, DATA_ROW_END, DATA_ROW_START, FIRST_STORY, REBAR_NUMBER(), WS As Worksheet

' RAW_DATA 資料命名
Private Const STORY = 1
Private Const NUMBER = 2
Private Const BW = 3
Private Const H = 4
Private Const D = 5
Private Const REBAR_LEFT = 6
Private Const REBAR_MIDDLE = 7
Private Const REBAR_RIGHT = 8
Private Const SIDE_REBAR = 9
Private Const STIRRUP_LEFT = 10
Private Const STIRRUP_MIDDLE = 11
Private Const STIRRUP_RIGHT = 12
Private Const BEAM_LENGTH = 13
Private Const SUPPORT = 14
Private Const LOCATION = 15

' GENERAL_INFORMATION 資料命名
Private Const FY = 2
Private Const FYT = 3
Private Const FC_BEAM = 4
Private Const FC_COLUMN = 5
Private Const SDL = 6
Private Const LL = 7
Private Const SPAN_X = 8
Private Const SPAN_Y = 9

' REBAR_SIZE 資料命名
Private Const DIAMETER = 7
Private Const CROSS_AREA = 10

' 輸出資料位置
Private Const MESSAGE_POSITION = 16

' -------------------------------------------------------------------------
' -------------------------------------------------------------------------

Private Sub Class_Initialize()
' Called automatically when class is created
' GetGeneralInformation
' GetRebarSize

    Call GetGeneralInformation
    Call GetRebarSize

End Sub


Function GetGeneralInformation()

    Dim generalInformation As Worksheet
    Set generalInformation = Worksheets("General Information")

    rowStart = 1
    columnStart = 4
    rowUsed = generalInformation.Cells(Rows.Count, 5).End(xlUp).Row
    columnUsed = 12

    GENERAL_INFORMATION = generalInformation.Range(generalInformation.Cells(rowStart, columnStart), generalInformation.Cells(rowUsed, columnUsed))

    FIRST_STORY = Application.Match("1F", Application.Index(GENERAL_INFORMATION, 0, STORY), 0)

End Function


Function GetRebarSize()

    Dim rebarSize As Worksheet
    Set rebarSize = Worksheets("Rebar Size")

    rowStart = 1
    columnStart = 1
    rowUsed = rebarSize.Cells(Rows.Count, 5).End(xlUp).Row
    columnUsed = 10

    REBAR_SIZE = rebarSize.Range(rebarSize.Cells(rowStart, columnStart), rebarSize.Cells(rowUsed, columnUsed))

End Function


Function GetData(sheet)

    Worksheet = sheet

    Set WS = Worksheets(Worksheet)

    rowStart = 1
    columnStart = 1
    rowUsed = WS.Cells(Rows.Count, 5).End(xlUp).Row
    columnUsed = 16

    RAW_DATA = WS.Range(WS.Cells(rowStart, columnStart), WS.Cells(rowUsed, columnUsed))

End Function


Function NoData()
'
' 如果沒有資料，就回傳 false
'
' @returns NoData(Boolean)

    NoData = UBound(RAW_DATA) < 4

End Function


Function Initialize()
'
' DATA_ROW_START
' DATA_ROW_END
' MESSAGE
' RatioData

    WS.Columns(MESSAGE_POSITION).ClearContents
    WS.Cells(1, MESSAGE_POSITION) = "Warning Message"
    DATA_ROW_START = 3
    DATA_ROW_END = UBound(RAW_DATA)

    ReDim MESSAGE(DATA_ROW_START To DATA_ROW_END)

    ReDim RATIO_DATA(LBound(RAW_DATA, 1) To UBound(RAW_DATA, 1), LBound(RAW_DATA, 2) To UBound(RAW_DATA, 2))

    Call RatioData

End Function


Function RatioData()

    ' 樓層數字化，用以比較上下樓層。
    For i = DATA_ROW_START To DATA_ROW_END
        RATIO_DATA(i, STORY) = Application.Match(RAW_DATA(i, STORY), Application.Index(GENERAL_INFORMATION, 0, STORY), 0)
    Next

    ' 計算鋼筋面積
    For i = DATA_ROW_START To DATA_ROW_END
        For j = REBAR_LEFT To REBAR_RIGHT
            RATIO_DATA(i, j) = CalRebarArea(RAW_DATA(i, j))
        Next
    Next

    ' 一二排截面積相加
    For i = DATA_ROW_START To DATA_ROW_END Step 2
        For j = REBAR_LEFT To REBAR_RIGHT
            RATIO_DATA(i, j) = RATIO_DATA(i, j) + RATIO_DATA(i + 1, j)
        Next
    Next

    ' 計算箍筋面積
    For i = DATA_ROW_START To DATA_ROW_END Step 4
        For j = STIRRUP_LEFT To STIRRUP_RIGHT
            RATIO_DATA(i, j) = CalStirrupArea(RAW_DATA(i, j))
        Next
    Next

    ' 計算側筋面積
    For i = DATA_ROW_START To DATA_ROW_END Step 4
        RATIO_DATA(i, SIDE_REBAR) = CalSideRebarArea(RAW_DATA(i, SIDE_REBAR))
    Next

    ' 計算有效深度
    For i = DATA_ROW_START To DATA_ROW_END Step 4

        REBAR = Split(RAW_DATA(i, REBAR_LEFT), "-")
        stirrup = Split(RAW_DATA(i, STIRRUP_LEFT), "@")
        Db = Application.VLookup(REBAR(1), REBAR_SIZE, DIAMETER, False)
        tie = Application.VLookup(SplitStirrup(SplitStirrup(stirrup(0))), REBAR_SIZE, DIAMETER, False)

        ' 雙排筋
        RATIO_DATA(i, D) = RAW_DATA(i, H) - (4 + tie + Db * 1.5)

    Next

End Function

Function SplitStirrup(REBAR)

    bars = Split(REBAR, "#")

    SplitStirrup = "#" & bars(1)

End Function

Function CalRebarArea(REBAR)

    tmp = Split(REBAR, "-")

    If tmp(0) <> 0 Then

        ' 轉換鋼筋尺寸為截面積
        tmp(1) = Application.VLookup(tmp(1), REBAR_SIZE, CROSS_AREA, False)

        CalRebarArea = tmp(0) * tmp(1)
    Else
        CalRebarArea = 0
    End If

End Function

Function CalStirrupArea(REBAR)
'
' 考量雙箍
'
    tmp = Split(REBAR, "@")

    bars = Split(tmp(0), "#")

    ' 箍筋號數
    bars(1) = "#" & bars(1)

    ' 轉換鋼筋尺寸為截面積
    If bars(0) = "" Then
        CalStirrupArea = 2 * Application.VLookup(bars(1), REBAR_SIZE, CROSS_AREA, False)
    Else
        CalStirrupArea = 2 * bars(0) * Application.VLookup(bars(1), REBAR_SIZE, CROSS_AREA, False)
    End If

End Function

Function CalSideRebarArea(REBAR)

    If REBAR <> "-" Then

        REBAR = Left(REBAR, Len(REBAR) - 2)

        tmp = Split(REBAR, "#")

        ' 轉換鋼筋尺寸為截面積
        tmp(1) = Application.VLookup("#" & tmp(1), REBAR_SIZE, CROSS_AREA, False)

        ' 對稱雙排
        CalSideRebarArea = 2 * tmp(1)

    Else
        CalSideRebarArea = 0
    End If

End Function


Function GetSheetMessage(Girder, Beam, GroundBeam)

    If Worksheet = "大梁配筋" Then
        GetSheetMessage = Girder
    ElseIf Worksheet = "小梁配筋" Then
        GetSheetMessage = Beam
    ElseIf Worksheet = "地梁配筋" Then
        GetSheetMessage = GroundBeam
    End If

End Function

Function WarningMessage(warinigMessageCode, i)

    MESSAGE(i) = warinigMessageCode & vbCrLf & MESSAGE(i)

End Function

Function PrintMessage()

    ' 不知道為什麼不能直接給值，只好用 for loop
    ' Range(Cells(DATA_ROW_START, MESSAGE_POSITION), Cells(DATA_ROW_END, MESSAGE_POSITION)) = MESSAGE()
    For i = DATA_ROW_START To DATA_ROW_END Step 4

        WS.Range(WS.Cells(i, MESSAGE_POSITION), WS.Cells(i + 3, MESSAGE_POSITION)).Merge

        If MESSAGE(i) = "" Then
            MESSAGE(i) = "(S), (E), (i) - check 結果 ok"
            WS.Cells(i, MESSAGE_POSITION).Style = "好"
        Else
            WS.Cells(i, MESSAGE_POSITION).Style = "壞"
            MESSAGE(i) = Left(MESSAGE(i), Len(MESSAGE(i)) - 1)
        End If
        WS.Cells(i, MESSAGE_POSITION) = MESSAGE(i)

    Next

    Call FontSetting

End Function

Function PrintRebarRatio()

    Dim rebarRatio As Worksheet
    Set rebarRatio = Worksheets("鋼筋號數比")

    rowStart = 3
    rowUsed = UBound(REBAR_NUMBER) + 1

    If Worksheet = "大梁配筋" Then
        columnStart = 4
    ElseIf Worksheet = "小梁配筋" Then
        columnStart = 7
    ElseIf Worksheet = "地梁配筋" Then
        columnStart = 10
    End If

    columnUsed = columnStart + 2

    rebarRatio.Range(rebarRatio.Cells(rowStart, columnStart), rebarRatio.Cells(rowUsed, columnUsed)) = REBAR_NUMBER

End Function

Function FontSetting()

    WS.Cells.Font.Name = "微軟正黑體"
    WS.Cells.Font.Name = "Calibri"
    WS.Cells.HorizontalAlignment = xlCenter
    WS.Cells.VerticalAlignment = xlCenter

End Function

Private Sub Class_Terminate()

    ' Called automatically when all references to class instance are removed

End Sub

' -------------------------------------------------------------------------
' 以下為實作內容
' -------------------------------------------------------------------------

Function SafetyRebarRatioAndSpace()
'
' 安全性指標：
' 最少鋼筋比大於 0.3 %
' 鋼筋間距 25 cm 以下

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = REBAR_LEFT To REBAR_RIGHT

            code = 0.003 * RAW_DATA(i, BW) * RATIO_DATA(i, D)

            ' 請確認是否符合 上層筋下限 規定
            If RATIO_DATA(i, j) < code Then
                Call WarningMessage("【0104】請確認上層筋下限，是否符合最少鋼筋比大於 0.3 % 規定", i)
            End If

            ' 請確認是否符合 下層筋下限 規定
            If RATIO_DATA(i + 2, j) < code Then
                Call WarningMessage("【0105】請確認下層筋下限，是否符合最少鋼筋比大於 0.3 % 規定", i)
            End If

            For k = i To i + 3

                REBAR = Split(RAW_DATA(k, j), "-")

                stirrup = Split(RAW_DATA(i, j + 4), "@")

                If REBAR(0) > 1 Then

                    Db = Application.VLookup(REBAR(1), REBAR_SIZE, DIAMETER, False)
                    tie = Application.VLookup(SplitStirrup(SplitStirrup(stirrup(0))), REBAR_SIZE, DIAMETER, False)

                    Spacing = (RAW_DATA(i, BW) - 4 * 2 - tie * 2 - REBAR(0) * Db) / (REBAR(0) - 1)

                    If Spacing > 25 Then
                        Call WarningMessage("【0106】請確認鋼筋間距下限，是否符合鋼筋間距 25 cm 以下規定", i)
                    End If

                ElseIf REBAR(0) = "1" Then

                    Call WarningMessage("【0107】請確認鋼筋間距，是否符合單排支數下限規定", i)

                End If
            Next
        Next
    Next

End Function

Function Norm4_9_3()
'
' 深梁：
' 垂直剪力鋼筋面積 Av 不得小於 0.0025 * bw * s，s 不得大於 d / 5 或 30 cm。

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            tmp = Split(RAW_DATA(i, j), "@")

            isAvSmallerThanCode = RATIO_DATA(i, j) < 0.0025 * RAW_DATA(i, BW) * tmp(1)

            If isAvSmallerThanCode Then
                Call WarningMessage("【0101】請確認短梁箍筋，是否小於 0.0025 * bw * s", i)
            End If

        Next

    Next

End Function

Function Norm4_9_4()
'
' 深梁：
' 水平剪力鋼筋面積 Avh 不得小於 0.0015 * bw * s2，s2 不得大於 d / 5 或 30 cm。

    ' 版厚
    bs = 20

    ' 地基版厚
    fs = 60

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        tmp = Split(RAW_DATA(i, SIDE_REBAR), "#")

        ' 分成四種狀況
        If tmp(0) = "-" Then
            isAvhSmallerThanCode = True
        ElseIf tmp(0) = "1" Then
            isAvhSmallerThanCode = RATIO_DATA(i, SIDE_REBAR) < 0.0015 * RAW_DATA(i, BW) * (RAW_DATA(i, H) - bs - fs)
        ElseIf tmp(0) = "2" Then
            isAvhSmallerThanCode = RATIO_DATA(i, SIDE_REBAR) < 0.0015 * RAW_DATA(i, BW) * (RAW_DATA(i, H) - bs - fs) / 2
        Else
            isAvhSmallerThanCode = RATIO_DATA(i, SIDE_REBAR) < 0.0015 * RAW_DATA(i, BW) * (RAW_DATA(i, H) - bs - fs - 15 - 15) / (tmp(0) - 1)
        End If

        If isAvhSmallerThanCode Then
            Call WarningMessage("【0102】請確認短梁側筋，是否小於 0.0015 * bw * s2", i)
        End If

    Next

End Function

Function EconomicNorm4_9_4()
'
' 經濟性指標：
' Avh need to less than 1.5 * 0.0015 * BW * S2

    bs = 20
    fs = 60
    factor = 1.5

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        tmp = Split(RAW_DATA(i, SIDE_REBAR), "#")

        If tmp(0) = "-" Then
            isAvhSmallerThanCode = True
        ElseIf tmp(0) = "1" Then
            isAvhSmallerThanCode = RATIO_DATA(i, SIDE_REBAR) > factor * 0.0015 * RAW_DATA(i, BW) * (RAW_DATA(i, H) - bs - fs)
        ElseIf tmp(0) = "2" Then
            isAvhSmallerThanCode = RATIO_DATA(i, SIDE_REBAR) > factor * 0.0015 * RAW_DATA(i, BW) * (RAW_DATA(i, H) - bs - fs) / 2
        Else
            isAvhSmallerThanCode = RATIO_DATA(i, SIDE_REBAR) > factor * 0.0015 * RAW_DATA(i, BW) * (RAW_DATA(i, H) - bs - fs - 15 - 15) / (tmp(0) - 1)
        End If

        If isAvhSmallerThanCode Then
            Call WarningMessage("【0103】請確認短梁側筋，是否大於 1.5 * 0.0015 * BW * S2", i)
        End If

    Next

End Function

Function SafetyLoad()
'
' 安全性指標：
' 載重預警
' 假設帶寬 3m ,則 0.6 * 1/8 * wu * L^2 <= As * fy * d

    band = 3

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        maxRatio = Application.Max(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_MIDDLE), RATIO_DATA(i, REBAR_RIGHT), RATIO_DATA(i + 2, REBAR_LEFT), RATIO_DATA(i + 2, REBAR_MIDDLE), RATIO_DATA(i + 2, REBAR_RIGHT))

        ' 轉換 kgw-m => tf-m: * 100000
        mn = 1 / 8 * (1.2 * (0.15 * 2.4 + Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, SDL, False)) + 1.6 * Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, LL, False)) * band ^ 2 * 100000

        capacity = maxRatio * Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FY, False) * RATIO_DATA(i, D)

        If 0.6 * mn > capacity Then
            Call WarningMessage("【0312】垂直載重配筋可能不足", i)
        End If

    Next

End Function

Function SafetyRebarRatioForSB()
'
' 安全性指標：
' 小梁鋼筋比在 2.5% 以下

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = REBAR_LEFT To REBAR_RIGHT

            limit = 0.025 * RAW_DATA(i, BW) * RATIO_DATA(i, D)

            If RATIO_DATA(i, j) > limit Then
                Call WarningMessage("【0310】請確認上層筋上限，是否在 2.5% 以下", i)
            End If

            If RATIO_DATA(i + 2, j) > limit Then
                Call WarningMessage("【0311】請確認下層筋上限，是否在 2.5% 以下", i)
            End If

        Next

    Next

End Function

Function SafetyRebarRatioForGB()
'
' 安全性指標：
' 地梁鋼筋比在 2% 以下

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = REBAR_LEFT To REBAR_RIGHT

            limit = 0.02 * RAW_DATA(i, BW) * RATIO_DATA(i, D)

            If RATIO_DATA(i, j) > limit Then
                Call WarningMessage("【0108】請確認上層筋上限，是否在 2% 以下", i)
            End If

            If RATIO_DATA(i + 2, j) > limit Then
                Call WarningMessage("【0109】請確認下層筋上限，是否在 2% 以下", i)
            End If

        Next

    Next

End Function

Function Norm3_6()
'
' 受撓構材之最少鋼筋量：
' 3-3 As >= 0.8 * sqr(fc') / fy * bw * d
' 3-4 As >= 14 / fy * bw * d

For i = DATA_ROW_START To DATA_ROW_END Step 4

    code3_3 = 0.8 * Sqr(Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FC_BEAM, False)) / Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FY, False) * RAW_DATA(i, BW) * RATIO_DATA(i, D)
    code3_4 = 14 / Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FY, False) * RAW_DATA(i, BW) * RATIO_DATA(i, D)

    If RATIO_DATA(i, REBAR_LEFT) < code3_3 Or RATIO_DATA(i, REBAR_LEFT) < code3_4 Then
        Call WarningMessage(GetSheetMessage("【0201】請確認左端上層筋下限，是否符合規範 3.6 規定", "【0301】請確認左端上層筋下限，是否符合規範 3.6 規定", "請確認左端上層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If RATIO_DATA(i, REBAR_MIDDLE) < code3_3 Or RATIO_DATA(i, REBAR_MIDDLE) < code3_4 Then
        Call WarningMessage(GetSheetMessage("【0202】請確認中央上層筋下限，是否符合規範 3.6 規定", "【0302】請確認中央上層筋下限，是否符合規範 3.6 規定", "請確認中央上層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If RATIO_DATA(i, REBAR_RIGHT) < code3_3 Or RATIO_DATA(i, REBAR_RIGHT) < code3_4 Then
        Call WarningMessage(GetSheetMessage("【0203】請確認右端上層筋下限，是否符合規範 3.6 規定", "【0303】請確認右端上層筋下限，是否符合規範 3.6 規定", "請確認右端上層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If RATIO_DATA(i + 2, REBAR_LEFT) < code3_3 Or RATIO_DATA(i + 2, REBAR_LEFT) < code3_4 Then
        Call WarningMessage(GetSheetMessage("【0204】請確認左端下層筋下限，是否符合規範 3.6 規定", "【0304】請確認左端下層筋下限，是否符合規範 3.6 規定", "請確認左端下層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If RATIO_DATA(i + 2, REBAR_MIDDLE) < code3_3 Or RATIO_DATA(i + 2, REBAR_MIDDLE) < code3_4 Then
        Call WarningMessage(GetSheetMessage("【0205】請確認中央下層筋下限，是否符合規範 3.6 規定", "【0305】請確認中央下層筋下限，是否符合規範 3.6 規定", "請確認中央下層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If RATIO_DATA(i + 2, REBAR_RIGHT) < code3_3 Or RATIO_DATA(i + 2, REBAR_RIGHT) < code3_4 Then
        Call WarningMessage(GetSheetMessage("【0206】請確認右端下層筋下限，是否符合規範 3.6 規定", "【0306】請確認右端下層筋下限，是否符合規範 3.6 規定", "請確認右端下層筋下限，是否符合規範 3.6 規定"), i)
    End If

Next

End Function

Function Norm15_4_2_1()
'
' 耐震規範 (1F以下大梁不適用)：
' 拉力鋼筋比不得大於 (fc' + 100) / (4 * fy)，亦不得大於 0.025。

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        If RATIO_DATA(i, STORY) < FIRST_STORY Then

            code15_4_2_1 = Application.Min((Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FC_BEAM, False) + 100) / (4 * Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FY, False)) * RAW_DATA(i, BW) * RATIO_DATA(i, D), 0.025 * RAW_DATA(i, BW) * RATIO_DATA(i, D))

            If RATIO_DATA(i, REBAR_LEFT) > code15_4_2_1 Then
                Call WarningMessage("【0212】請確認左端上層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If RATIO_DATA(i, REBAR_MIDDLE) > code15_4_2_1 Then
                Call WarningMessage("【0213】請確認中央上層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If RATIO_DATA(i, REBAR_RIGHT) > code15_4_2_1 Then
                Call WarningMessage("【0214】請確認右端上層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If RATIO_DATA(i + 2, REBAR_LEFT) > code15_4_2_1 Then
                Call WarningMessage("【0215】請確認左端下層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If RATIO_DATA(i + 2, REBAR_MIDDLE) > code15_4_2_1 Then
                Call WarningMessage("【0216】請確認中央下層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If RATIO_DATA(i + 2, REBAR_RIGHT) > code15_4_2_1 Then
                Call WarningMessage("【0217】請確認右端下層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

        End If

    Next

End Function

Function Norm15_4_2_2()
'
' 耐震規範 (1F以下大梁不適用)：
' 規範內容：撓曲構材在梁柱交接面及其它可能產生塑鉸位置，其壓力鋼筋量不得小於拉力鋼筋量之半。在沿構材長度上任何斷面，不論正彎矩鋼筋量或負彎矩鋼筋量均不得低於兩端柱面處所具最大負彎矩鋼筋量之 1/4。
' 實作方法：最小鋼筋量需大於最大鋼筋量 1/4

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        If RATIO_DATA(i, STORY) < FIRST_STORY Then

            maxRatio = Application.Max(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_MIDDLE), RATIO_DATA(i, REBAR_RIGHT), RATIO_DATA(i + 2, REBAR_LEFT), RATIO_DATA(i + 2, REBAR_MIDDLE), RATIO_DATA(i + 2, REBAR_RIGHT))
            minRatio = Application.Min(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_MIDDLE), RATIO_DATA(i, REBAR_RIGHT), RATIO_DATA(i + 2, REBAR_LEFT), RATIO_DATA(i + 2, REBAR_MIDDLE), RATIO_DATA(i + 2, REBAR_RIGHT))
            code15_4_2_2 = minRatio < maxRatio / 4

            If code15_4_2_2 And RAW_DATA(i, STORY) <> "1F" Then
                Call WarningMessage("【0218】請確認耐震最小量鋼筋，是否符合規範 15.4.2.2 規定", i)
            End If

        End If

    Next

End Function

Function EconomicTopRebarRelativeForGB()
'
' 經濟性指標：
' 如果鋼筋支數大於3支，端部上層鋼筋量需小於中央鋼筋量的 70%。
' 淨跨度大於 400 cm，才要檢討

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        rebarLEFT = Split(RAW_DATA(i, REBAR_LEFT), "-")
        rebarRIGHT = Split(RAW_DATA(i, REBAR_RIGHT), "-")

        If RATIO_DATA(i, REBAR_MIDDLE) * 0.7 < RATIO_DATA(i, REBAR_LEFT) And rebarLEFT(0) > 3 And RATIO_DATA(i, BEAM_LENGTH) > 400 Then
            Call WarningMessage("【0111】請確認左端上層筋相對鋼筋量，是否符合端部上層鋼筋量需小於中央鋼筋量的 70% 規定", i)
        End If

        If RATIO_DATA(i, REBAR_MIDDLE) * 0.7 < RATIO_DATA(i, REBAR_RIGHT) And rebarRIGHT(0) > 3 And RATIO_DATA(i, BEAM_LENGTH) > 400 Then
            Call WarningMessage("【0112】請確認右端上層筋相對鋼筋量，是否符合端部上層鋼筋量需小於中央鋼筋量的 70% 規定", i)
        End If

    Next

End Function

Function EconomicTopRebarRelative()
'
' 經濟性指標：
' 如果鋼筋支數大於3支，中央上層鋼筋量需小於端部最小鋼筋量的 70%。
' 淨跨度大於 400 cm，才要檢討


    For i = DATA_ROW_START To DATA_ROW_END Step 4

        minRatio = Application.Min(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_RIGHT))

        REBAR = Split(RAW_DATA(i, REBAR_MIDDLE), "-")

        If RATIO_DATA(i, REBAR_MIDDLE) > minRatio * 0.7 And REBAR(0) > 3 And RATIO_DATA(i, BEAM_LENGTH) > 400 Then
            Call WarningMessage("【0221】請確認中央上層筋相對鋼筋量，是否符合中央上層鋼筋量需小於端部最小鋼筋量的 70% 規定", i)
        End If

    Next

End Function

Function EconomicBotRebarRelativeForGB()
'
' 經濟性指標：
' 如果鋼筋支數大於3支，中央下層鋼筋量需小於端部最小鋼筋量的 70%。
' 淨跨度大於 400 cm，才要檢討

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        minRatio = Application.Min(RATIO_DATA(i + 2, REBAR_LEFT), RATIO_DATA(i + 2, REBAR_RIGHT))

        REBAR = Split(RAW_DATA(i + 2, REBAR_MIDDLE), "-")

        If RATIO_DATA(i + 2, REBAR_MIDDLE) > minRatio * 0.7 And REBAR(0) > 3 And RATIO_DATA(i, BEAM_LENGTH) > 400 Then
            Call WarningMessage("【0110】請確認中央下層筋相對鋼筋量，是否符合中央下層鋼筋量需小於端部最小鋼筋量的 70% 規定", i)
        End If

    Next

End Function

Function Norm13_5_1AndSafetyRebarNumber()
'
' 鋼筋間距之限制：
' 規範內容：同層平行鋼筋間之淨距不得小於 1.0db，或粗粒料標稱最大粒徑 1.33 倍，亦不得小於 2.5 cm。
' 實作內容：單排淨距需在 1db 以上 且 單排支數需大於1支。

    For k = DATA_ROW_START To DATA_ROW_END

        For j = REBAR_LEFT To REBAR_RIGHT

            ' 重要：因為k每步都是1，所以增加一個k來計算每4步。
            ' 其實可以用 i = i + 4 比較簡單
            i = 4 * Fix((k - 3) / 4) + 3

            REBAR = Split(RAW_DATA(k, j), "-")

            stirrup = Split(RAW_DATA(i, j + 4), "@")

            ' 等於 0 直接沒做事
            If REBAR(0) > 1 Then

                Db = Application.VLookup(REBAR(1), REBAR_SIZE, DIAMETER, False)
                tie = Application.VLookup(SplitStirrup(stirrup(0)), REBAR_SIZE, DIAMETER, False)

                ' 第一種方法
                ' Max = Fix((RAW_DATA(i, BW) - 4 * 2 - tie * 2 - Db) / (2 * Db)) + 1
                ' CInt(rebar(0)) > Max
                ' 第二種方法
                ' spacing = (RAW_DATA(i, BW) - 4 * 2 - tie * 2 - Db) / (CInt(rebar(0)) - 1) - Db
                ' 可以不需要型別轉換
                ' Spacing = (RAW_DATA(i, BW) - 4 * 2 - tie * 2 - CInt(rebar(0)) * Db) / (CInt(rebar(0)) - 1)
                Spacing = (RAW_DATA(i, BW) - 4 * 2 - tie * 2 - REBAR(0) * Db) / (REBAR(0) - 1)

                ' Norm13_5_1
                ' 淨距不少於1Db
                If Spacing < Db Or Spacing < 2.5 Then
                    Call WarningMessage(GetSheetMessage("【0210】請確認單排支數上限，是否符合淨距不少於 1 Db 規定", "【0308】請確認單排支數上限，是否符合淨距不少於 1 Db 規定", "請確認單排支數上限，是否符合淨距不少於 1 Db 規定"), i)
                End If

            ElseIf REBAR(0) = "1" Then

                ' 排除掉1支的狀況，避免除以0
                ' 不少於2支
                Call WarningMessage(GetSheetMessage("【0211】請確認是否符合 單排支數下限 規定", "【0309】請確認是否符合 單排支數下限 規定", "請確認是否符合 單排支數下限 規定"), i)

            End If

        Next
    Next

End Function

Function SafetyStirrupSpace()
'
' 安全性與經濟性指標：
' 箍筋間距 10cm 以上
' 箍筋間距 30cm 以下

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            stirrup = Split(RAW_DATA(i, j), "@")

            If stirrup(1) < 10 Then
                Call WarningMessage(GetSheetMessage("【0219】請確認箍筋間距下限，是否符合 10cm 以上規定", "請確認箍筋間距下限，是否符合 10cm 以上規定", "【0113】請確認箍筋間距下限，是否符合 10cm 以上規定"), i)
            ElseIf stirrup(1) > 30 Then
                Call WarningMessage(GetSheetMessage("【0220】請確認箍筋間距上限，是否符合 30cm 以下規定", "請確認箍筋間距上限，是否符合 30cm 以下規定", "【0114】請確認箍筋間距上限，是否符合 30cm 以下規定"), i)
            End If

        Next

    Next

End Function

Function Norm4_6_6_3()
'
' 剪力鋼筋量大於3.52/fy

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            stirrup = Split(RAW_DATA(i, j), "@")

            avMin = Application.Max(0.2 * Sqr(Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FC_BEAM, False)) * data(i, BW) * stirrup(1) / Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FYT, False), 3.5 * data(i, BW) * stirrup(1) / Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FYT, False))
            av = RATIO_DATA(i, j)

            If av < avMin Then
                Call WarningMessage("請確認剪力鋼筋量下限，是否大於 3.52 / fy", i)
            End If

        Next

    Next

End Function

Function Norm4_6_7_9()
'
' 剪力鋼筋之剪力計算強度：
' 規範內容：剪力計算強度 Vs 不可大於 2.12 * fc' * bw * d。
' 實作內容：剪力鋼筋量需在 4 * Vc * 120% 以下。規範為 vs <= 4 * vc，由於取整數容易超過，所以放寬標準 120%。

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            stirrup = Split(RAW_DATA(i, j), "@")
            REBAR = Split(RAW_DATA(i, j - 4), "-")
            Db = Application.VLookup(REBAR(1), REBAR_SIZE, DIAMETER, False)
            tie = Application.VLookup(SplitStirrup(stirrup(0)), REBAR_SIZE, DIAMETER, False)
            effectiveDepth = RAW_DATA(i, H) - (4 + tie + Db / 2)
            av = RATIO_DATA(i, j)

            ' code4.4.1.1
            vc = 0.53 * Sqr(Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FC_BEAM, False)) * RAW_DATA(i, BW) * effectiveDepth

            ' code4.6.7.2
            vs = av * Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FYT, False) * effectiveDepth / stirrup(1)

            If vs > 4 * vc * 1.2 Then
                Call WarningMessage("【0209】請確認剪力鋼筋量上限，是否符合規範 4.6.7.9 規定", i)
            End If

        Next

    Next

End Function

Function Norm3_8_1()
'
' 深梁規範內容：
' 深梁為載重與支撐分別位於構材之頂面與底面，使壓桿形成於載重及支點之間，且符合：
' (1) 淨跨 ln 不大於 4 倍梁總深；或
' (2) 集中載重作用區與支承面之距離小於 2 倍梁總深。
' 深梁應依非線性應變分佈設計，或依附篇 A 設計(見第 4.9.1、5.11.6 節)；橫向屈曲必須考慮。
'
' 實作內容： L/H <= 4

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        If RAW_DATA(i, BEAM_LENGTH) <> "" And RAW_DATA(i, SUPPORT) <> "" And (RAW_DATA(i, BEAM_LENGTH) - RAW_DATA(i, SUPPORT)) <= 4 * RAW_DATA(i, H) Then
            Call WarningMessage("【0208】請確認是否為深梁", i)
        End If

    Next

End Function

Function Norm3_7_5()

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        If RAW_DATA(i, H) > 90 Then
            Call WarningMessage(GetSheetMessage("【0207】請確認是否符合 規範 3.7.5", "【0307】請確認是否符合 規範 3.7.5", "請確認是否符合 規範 3.7.5"), i)
        End If

    Next

End Function

Function CountRebarNumber()

    rowStart = 2
    rowEnd = UBound(REBAR_SIZE)
    ReDim REBAR_NUMBER(rowStart To rowEnd, 1 To 3)

    ' 主筋
    For i = DATA_ROW_START To DATA_ROW_END

        For j = REBAR_LEFT To REBAR_RIGHT

            rebarNumber = Split(RAW_DATA(i, j), "-")

            If rebarNumber(0) > 0 Then
                rebarNumber = rebarNumber(1)
            Else
                rebarNumber = ""
            End If

            For k = rowStart To rowEnd

                If rebarNumber = REBAR_SIZE(k, 1) Then
                    REBAR_NUMBER(k, 1) = REBAR_NUMBER(k, 1) + 1
                End If

            Next

        Next

    Next

    ' 腰筋
    For i = DATA_ROW_START To DATA_ROW_END Step 4

        If RAW_DATA(i, SIDE_REBAR) <> "-" Then

            sideRebar = Left(RAW_DATA(i, SIDE_REBAR), Len(RAW_DATA(i, SIDE_REBAR)) - 2)

            rebarNumber = Split(sideRebar, "#")

            rebarNumber = "#" & rebarNumber(1)

            For j = rowStart To rowEnd

                If rebarNumber = REBAR_SIZE(j, 1) Then
                    REBAR_NUMBER(j, 2) = REBAR_NUMBER(j, 2) + 1
                End If

            Next

        End If

    Next

    ' 箍筋
    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            rebarNumber = Split(RAW_DATA(i, j), "@")(0)
            rebarNumber = Split(rebarNumber, "#")
            rebarNumber = "#" & rebarNumber(1)

            For k = rowStart To rowEnd

                If rebarNumber = REBAR_SIZE(k, 1) Then
                    REBAR_NUMBER(k, 3) = REBAR_NUMBER(k, 3) + 1
                End If

            Next

        Next

    Next

    ' 轉換成比例
    Dim sum(1 To 3)
    For i = rowStart To rowEnd
        For j = 1 To 3
            sum(j) = sum(j) + REBAR_NUMBER(i, j)
        Next
    Next
    For j = 1 To 3
        For i = rowStart To rowEnd
            If REBAR_NUMBER(i, j) <> 0 Then
                REBAR_NUMBER(i, j) = REBAR_NUMBER(i, j) / sum(j)
            End If
        Next
    Next

End Function