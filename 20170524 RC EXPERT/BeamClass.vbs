Dim GENERAL_INFORMATION, REBAR_SIZE, RAW_DATA, RATIO_DATA, DATA_ROW_END, DATA_ROW_START, MESSAGE()

' RAW_DATA 資料命名
Const STORY = 1
Const NUMBER = 2
Const BW = 3
Const H = 4
Const D = 5
Const REBAR_LEFT = 6
Const REBAR_MIDDLE = 7
Const REBAR_RIGHT = 8
Const SIDE_REBAR = 9
Const STIRRUP_LEFT = 10
Const STIRRUP_MIDDLE = 11
Const STIRRUP_RIGHT = 12
Const BEAM_LENGTH = 13
Const SUPPORT = 14
Const LOCATION = 15

' GENERAL_INFORMATION 資料命名
Const FY = 2
Const FYT = 3
Const FC_BEAM = 4
Const FC_COLUMN = 5
Const SDL = 6
Const LL = 7
Const SPAN_X = 8
Const SPAN_Y = 9

' REBAR_SIZE 資料命名
Const DIAMETER = 7
Const CROSS_AREA = 10

' 輸出資料位置
Const MESSAGE_POSITION = 16

' -------------------------------------------------------------------------

' Public Property Get GENERAL_INFORMATION()
'     GENERAL_INFORMATION = vGENERAL_INFORMATION
' End Property

' Public Property Get REBAR_SIZE()
'     REBAR_SIZE = vREBAR_SIZE
' End Property

' Public Property Get RAW_DATA()
'     RAW_DATA = vRAW_DATA
' End Property

' Public Property Get DATA_ROW_END()
'     DATA_ROW_END = vDATA_ROW_END
' End Property

' Public Property Get DATA_ROW_START()
'     DATA_ROW_START = vDATA_ROW_START
' End Property

' -------------------------------------------------------------------------

' Public Property Let GENERAL_INFORMATION(value)
'     vGENERAL_INFORMATION = value
' End Property

' Public Property Let REBAR_SIZE(value)
'     vREBAR_SIZE = value
' End Property

' Public Property Let RAW_DATA(value)
'     vRAW_DATA = value
' End Property

' Public Property Let DATA_ROW_END(value)
'     vDATA_ROW_END = value
' End Property

' Public Property Let DATA_ROW_START(value)
'     vDATA_ROW_START = value
' End Property

' -------------------------------------------------------------------------

Private Sub Class_Initialize()
' Called automatically when class is created
' GetGeneralInformation
' GetRebarSize

    Call GetGeneralInformation
    Call GetRebarSize


End Sub

' -------------------------------------------------------------------------

Function GetGeneralInformation()

    Worksheets("General Information").Activate

    rowStart = 1
    columnStart = 4
    rowUsed = Cells(Rows.Count, 5).End(xlUp).Row
    columnUsed = 12

    GENERAL_INFORMATION = Range(Cells(rowStart, columnStart), Cells(rowUsed, columnUsed)).Value

End Function

Function GetRebarSize()

    Worksheets("Rebar Size").Activate

    rowStart = 1
    columnStart = 1
    rowUsed = Cells(Rows.Count, 5).End(xlUp).Row
    columnUsed = 10

    REBAR_SIZE = Range(Cells(rowStart, columnStart), Cells(rowUsed, columnUsed)).Value

End Function

Function GetData(sheet)

    Worksheets(sheet).Activate

    rowStart = 1
    columnStart = 1
    rowUsed = Cells(Rows.Count, 5).End(xlUp).Row
    columnUsed = 16

    RAW_DATA = Range(Cells(rowStart, columnStart), Cells(rowUsed, columnUsed)).Value

End Function

Function RatioData()

    ' 計算鋼筋面積
    For i = DATA_ROW_START To DATA_ROW_END

        RATIO_DATA(i, REBAR_LEFT) = CalRebarArea(RATIO_DATA(i, REBAR_LEFT))
        RATIO_DATA(i, REBAR_MIDDLE) = CalRebarArea(RATIO_DATA(i, REBAR_MIDDLE))
        RATIO_DATA(i, REBAR_RIGHT) = CalRebarArea(RATIO_DATA(i, REBAR_RIGHT))

    Next

    ' 一二排截面積相加
    For i = DATA_ROW_START To DATA_ROW_END Step 2

        RATIO_DATA(i, REBAR_LEFT) = RATIO_DATA(i, REBAR_LEFT) + RATIO_DATA(i + 1, REBAR_LEFT)
        RATIO_DATA(i, REBAR_MIDDLE) = RATIO_DATA(i, REBAR_MIDDLE) + RATIO_DATA(i + 1, REBAR_MIDDLE)
        RATIO_DATA(i, REBAR_RIGHT) = RATIO_DATA(i, REBAR_RIGHT) + RATIO_DATA(i + 1, REBAR_RIGHT)

    Next

    ' 計算有效深度
    For i = DATA_ROW_START To DATA_ROW_END Step 4

        rebar = Split(RAW_DATA(i, REBAR_LEFT), "-")
        stirrup = Split(RAW_DATA(i, STIRRUP_LEFT), "@")
        Db = Application.VLookup(rebar(1), REBAR_SIZE, DIAMETER, False)
        tie = Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)
        RATIO_DATA(i, D) = RATIO_DATA(i, H) - (4 + tie + Db / 2)

    Next

End Function

Function CalRebarArea(rebar)

    tmp = Split(rebar, "-")

    If tmp(1) <> "" Then

        ' 轉換鋼筋尺寸為截面積
        tmp(1) = Application.VLookup(tmp(1), REBAR_SIZE, CROSS_AREA, False)

        CalRebarArea = tmp(0) * tmp(1)
    Else
        CalRebarArea = 0
    End If

End Function

Function Initialize()
'
' DATA_ROW_START
' DATA_ROW_END
' MESSAGE
' RatioData

    Columns(16).ClearContents
    Cells(1, 16) = "Warning Message"
    DATA_ROW_START = 3
    DATA_ROW_END = UBound(RAW_DATA)

    ReDim MESSAGE(DATA_ROW_START To DATA_ROW_END)

    RATIO_DATA = RAW_DATA

    Call RatioData

End Function

' -------------------------------------------------------------------------

Function BeamNoLessThan03()
'
'

    For i = DATA_ROW_START To DATA_ROW_END Step 2

        For j = REBAR_LEFT To REBAR_RIGHT

            ' 重要：因為i每步都是2，所以增加一個k來計算每4步。
            k = 4 * Fix((i - 3) / 4) + 3

            a = 0.003 * RATIO_DATA(k, BW) * RATIO_DATA(k, D)
            b = 0.025 * RATIO_DATA(k, BW) * RATIO_DATA(k, D)

            If RATIO_DATA(i, j) < a Then
                Call WarningMessage("請確認是否符合 主筋比下限 規定", k)
            End If

            If RATIO_DATA(i, j) > b Then
                Call WarningMessage("請確認是否符合 主筋比上限 規定", k)
            End If

        Next

    Next

End Function

Function Norm3_6()
'
' RC規範 3-3, 3-4 最少鋼筋量大於14/fy

For i = DATA_ROW_START To DATA_ROW_END Step 4

    code3_3 = 0.8 * Sqr(Application.VLookup(RATIO_DATA(i, STORY), GENERAL_INFORMATION, FC_BEAM, False)) / Application.VLookup(RATIO_DATA(i, STORY), GENERAL_INFORMATION, FY, False) * RATIO_DATA(i, BW) * RATIO_DATA(i, D)
    code3_4 = 14 / Application.VLookup(RATIO_DATA(i, STORY), GENERAL_INFORMATION, FY, False) * RATIO_DATA(i, BW) * RATIO_DATA(i, D)

    ' 請確認是否符合 左端上層筋下限 規定
    If RATIO_DATA(i, REBAR_LEFT) < code3_3 Or RATIO_DATA(i, REBAR_LEFT) < code3_4 Then
        Call WarningMessage("請確認是否符合 左端上層筋下限 規定", i)
    End If

    ' 請確認是否符合 右端上層筋下限 規定
    If RATIO_DATA(i, REBAR_RIGHT) < code3_3 Or RATIO_DATA(i, REBAR_RIGHT) < code3_4 Then
        Call WarningMessage("請確認是否符合 右端上層筋下限 規定", i)
    End If

    ' 請確認是否符合 左端下層筋下限 規定
    If RATIO_DATA(i + 2, REBAR_LEFT) < code3_3 Or RATIO_DATA(i + 2, REBAR_LEFT) < code3_4 Then
        Call WarningMessage("請確認是否符合 左端下層筋下限 規定", i)
    End If

    ' 請確認是否符合 右端下層筋下限 規定
    If RATIO_DATA(i + 2, REBAR_RIGHT) < code3_3 Or RATIO_DATA(i + 2, REBAR_RIGHT) < code3_4 Then
        Call WarningMessage("請確認是否符合 右端下層筋下限 規定", i)
    End If

    ' 請確認是否符合 中央上層筋下限 規定
    If RATIO_DATA(i, REBAR_MIDDLE) < code3_3 Or RATIO_DATA(i, REBAR_MIDDLE) < code3_4 Then
        Call WarningMessage("請確認是否符合 中央上層筋下限 規定", i)
    End If

Next

End Function

Function Norm15_4_2_1()
'
' RC規範 15.4.2.1 耐震規範 (1F大梁不適用)：最大鋼筋量低於2.2 %

For i = DATA_ROW_START To DATA_ROW_END Step 4

    code15_4_2_1 = Application.Min((Application.VLookup(RATIO_DATA(i, STORY), GENERAL_INFORMATION, FC_BEAM, False) + 100) / (4 * Application.VLookup(RATIO_DATA(i, STORY), GENERAL_INFORMATION, FY, False)) * RATIO_DATA(i, BW) * RATIO_DATA(i, D), 0.025 * RATIO_DATA(i, BW) * RATIO_DATA(i, D))

    If RATIO_DATA(i, REBAR_LEFT) > code15_4_2_1 And RATIO_DATA(i, STORY) <> "1F" Then
        Call WarningMessage("請確認是否符合 左端上層筋上限 規定", i)
    End If

    If RATIO_DATA(i, REBAR_RIGHT) > code15_4_2_1 And RATIO_DATA(i, STORY) <> "1F" Then
        Call WarningMessage("請確認是否符合 右端上層筋上限 規定", i)
    End If

    If RATIO_DATA(i + 2, REBAR_LEFT) > code15_4_2_1 And RATIO_DATA(i, STORY) <> "1F" Then
        Call WarningMessage("請確認是否符合 左端下層筋上限 規定", i)
    End If

    If RATIO_DATA(i + 2, REBAR_RIGHT) > code15_4_2_1 And RATIO_DATA(i, STORY) <> "1F" Then
        Call WarningMessage("請確認是否符合 右端下層筋上限 規定", i)
    End If

    If RATIO_DATA(i + 2, REBAR_MIDDLE) > code15_4_2_1 And RATIO_DATA(i, STORY) <> "1F" Then
        Call WarningMessage("請確認是否符合 中央下層筋上限 規定", i)
    End If

Next

End Function

Function Norm15_4_2_2()
'
' RC規範 15.4.2.2 耐震規範 (1F大梁不適用)：任一點鋼筋量大於最大鋼筋量1/4

For i = DATA_ROW_START To DATA_ROW_END Step 4

    Max = Application.Max(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_MIDDLE), RATIO_DATA(i, REBAR_RIGHT), RATIO_DATA(i + 2, REBAR_LEFT), RATIO_DATA(i + 2, REBAR_MIDDLE), RATIO_DATA(i + 2, REBAR_RIGHT))
    Min = Application.Min(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_MIDDLE), RATIO_DATA(i, REBAR_RIGHT), RATIO_DATA(i + 2, REBAR_LEFT), RATIO_DATA(i + 2, REBAR_MIDDLE), RATIO_DATA(i + 2, REBAR_RIGHT))
    code15_4_2_2 = Min <= Max / 4

    If code15_4_2_2 And RATIO_DATA(i, STORY) <> "1F" Then
        Call WarningMessage("請確認是否符合 耐震最小量鋼筋 規定", i)
    End If

Next

End Function

Function NormMiddleNoMoreThanEndEightyPercentage()
'
' 經濟性指標 中央上層鋼筋量小於端部最小鋼筋量80%

For i = DATA_ROW_START To DATA_ROW_END Step 4

    Min = Application.Min(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_RIGHT))
    If RATIO_DATA(i, REBAR_MIDDLE) > Min * 0.8 Then
        Call WarningMessage("請確認是否符合 中央上層筋相對鋼筋量 規定", i)
    End If

Next

End Function

' -------------------------------------------------------------------------

Function Norm13_5_1AndRebarAmountNoBelowTwo()
'
' 單排淨距小於1db(不能排太多)
' 單排支數大於2支

    For i = DATA_ROW_START To DATA_ROW_END

        For j = REBAR_LEFT To REBAR_RIGHT

            ' 重要：因為i每步都是1，所以增加一個k來計算每4步。
            k = 4 * Fix((i - 3) / 4) + 3

            rebar = Split(RAW_DATA(i, j), "-")

            stirrup = Split(RAW_DATA(k, j + 4), "@")

            If rebar(0) = "1" Then

                ' 排除掉1支的狀況，避免除以0
                ' 不少於2支
                Call WarningMessage("請確認是否符合 單排支數下限 規定", k)

            ElseIf rebar(0) <> "" Then

                Db = Application.VLookup(rebar(1), REBAR_SIZE, DIAMETER, False)
                tie = Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)

                ' 第一種方法
                ' Max = Fix((RAW_DATA(k, BW) - 4 * 2 - tie * 2 - Db) / (2 * Db)) + 1
                ' CInt(rebar(0)) > Max
                ' 第二種方法
                ' spacing = (RAW_DATA(k, BW) - 4 * 2 - tie * 2 - Db) / (CInt(rebar(0)) - 1) - Db
                ' 可以不需要型別轉換
                ' Spacing = (RAW_DATA(k, BW) - 4 * 2 - tie * 2 - CInt(rebar(0)) * Db) / (CInt(rebar(0)) - 1)
                Spacing = (RAW_DATA(k, BW) - 4 * 2 - tie * 2 - rebar(0) * Db) / (rebar(0) - 1)

                ' Norm13_5_1
                ' 淨距不少於1Db
                If Spacing < Db Or Spacing < 2.5 Then
                    Call WarningMessage("請確認是否符合 單排支數上限 規定", k)
                End If

            End If

        Next
    Next

End Function

Function NormStirrupSpacingMoreThan10AndLessThan30()
'
' 箍筋間距大於10CM
' 箍筋間距小於30CM

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            stirrup = Split(RAW_DATA(i, j), "@")

            If stirrup(1) < 10 Then
                Call WarningMessage("請確認是否符合 箍筋間距下限 規定", i)
            ElseIf stirrup(1) > 30 Then
                Call WarningMessage("請確認是否符合 箍筋間距上限 規定", i)
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
            av = Application.VLookup(stirrup(0), REBAR_SIZE, CROSS_AREA, False) * 2

            If av < avMin Then
                Call WarningMessage("請確認是否符合 剪力鋼筋量下限 規定", i)
            End If

        Next

    Next

End Function

Function Norm4_6_7_9()
'
' 剪力鋼筋量小於4Vc * 120%

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            stirrup = Split(RAW_DATA(i, j), "@")
            rebar = Split(RAW_DATA(i, j - 4), "-")
            Db = Application.VLookup(rebar(1), REBAR_SIZE, DIAMETER, False)
            tie = Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)
            effectiveDepth = RAW_DATA(i, H) - (4 + tie + Db / 2)
            av = Application.VLookup(stirrup(0), REBAR_SIZE, CROSS_AREA, False) * 2

            ' code4.1.1.1
            vc = 0.53 * Sqr(Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FC_BEAM, False)) * RAW_DATA(i, BW) * effectiveDepth

            ' code4.6.7.2
            vs = av * Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FYT, False) * effectiveDepth / stirrup(1)

            ' 規範為 vs <= 4 * vc，由於取整數容易超過，所以放寬標準120%
            If vs > 4 * vc * 1.2 Then
                Call WarningMessage("請確認是否符合 剪力鋼筋量上限 規定", i)
            End If

        Next

    Next

End Function

Function Norm3_8_1()
'
' 深梁 L/H<=4

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        If RAW_DATA(i, BEAM_LENGTH) <> "" And RAW_DATA(i, SUPPORT) <> "" And (RAW_DATA(i, BEAM_LENGTH) - RAW_DATA(i, SUPPORT)) <= 4 * RAW_DATA(i, H) Then
            Call WarningMessage("請確認是否為深梁", i)
        End If

    Next

End Function

Function WarningMessage(warinigMessageCode, i)

    MESSAGE(i) = warinigMessageCode & vbCrLf & MESSAGE(i)

End Function

Function PrintMessage()

    ' Worksheets("大梁配筋").Activate

    ' 不知道為什麼不能直接給值，只好用 for loop
    ' Range(Cells(DATA_ROW_START, MESSAGE_POSITION), Cells(DATA_ROW_END, MESSAGE_POSITION)) = MESSAGE()
    For i = DATA_ROW_START To DATA_ROW_END Step 4
        If MESSAGE(i) = "" Then
            MESSAGE(i) = "(S), (E), (i) - check 結果 ok"
            Cells(i, MESSAGE_POSITION).Style = "好"
        Else
            Cells(i, MESSAGE_POSITION).Style = "壞"
            MESSAGE(i) = Left(MESSAGE(i), Len(MESSAGE(i)) - 1)
        End If
        Cells(i, MESSAGE_POSITION) = MESSAGE(i)
    Next

End Function

' -------------------------------------------------------------------------

Private Sub Class_Terminate()

    ' Called automatically when all references to class instance are removed

End Sub
