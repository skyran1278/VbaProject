Dim WARNING_MESSAGE, GENERAL_INFORMATION, REBAR_SIZE, RAW_DATA, DATA_ROW_END, DATA_ROW_START, OK_MESSAGE()

' Message position
Const WARNING_MESSAGE_POSITION = 6

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
Const OK_MESSAGE_POSITION = 16

' GENERAL_INFORMATION 資料命名
Const FY = 2
Const FYT = 3
Const FC = 4
Const SDL = 5
Const LL = 6
Const SPAN_X = 7
Const SPAN_Y = 8

' REBAR_SIZE 資料命名
Const DIAMETER = 7
Const CROSS_AREA = 10


Function GetGeneralInformation()

    Worksheets("General Information").Activate
    Dim arr()
    rowUsed = Cells(Rows.Count, 5).End(xlUp).Row
    columnUsed = 8
    ReDim arr(1 To rowUsed, 1 To columnUsed)

    For i = 1 To rowUsed
        For j = 1 To columnUsed
            arr(i, j) = Cells(i, j + 3)
        Next
    Next

    GetGeneralInformation = arr()

End Function

Function GetRebarSize()

    Worksheets("Rebar Size").Activate
    Dim arr()
    rowUsed = Cells(Rows.Count, 5).End(xlUp).Row
    columnUsed = 10
    ReDim arr(1 To rowUsed, 1 To columnUsed)

    For i = 1 To rowUsed
        For j = 1 To columnUsed
            arr(i, j) = Cells(i, j)
        Next
    Next

    GetRebarSize = arr()

End Function

Function GetData()

    Worksheets("大梁配筋").Activate
    Dim arr()
    rowUsed = Cells(Rows.Count, 5).End(xlUp).Row
    columnUsed = 15
    ReDim arr(1 To rowUsed, 1 To columnUsed)

    For i = 1 To rowUsed
        For j = 1 To columnUsed
            arr(i, j) = Cells(i, j)
        Next
    Next

    Worksheets("Echo 大梁配筋").Activate

    Range(Cells(1, 1), Cells(rowUsed, columnUsed)) = arr()

    GetData = arr()

End Function

Function Initialize()
'
' Expert Check .Activate
' DATA_ROW_START DATA_ROW_END
' OK_MESSAGE Initial Value
' Expert Check ClearContents

    Worksheets("Expert Check").Activate
    DATA_ROW_START = 3
    DATA_ROW_END = UBound(RAW_DATA)

    ReDim OK_MESSAGE(DATA_ROW_START to DATA_ROW_END)
    ' For i = DATA_ROW_START To DATA_ROW_END Step 4
    '     OK_MESSAGE(i) = "(S), (E), (i) - check 結果 ok"
    ' Next

    rowUsed = Cells(Rows.Count, WARNING_MESSAGE_POSITION).End(xlUp).Row
    Range(Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION), Cells(rowUsed, WARNING_MESSAGE_POSITION)).ClearContents

End Function

' -------------------------------------------------------------------------

Function AboutRatioNorm(data)

    ' 計算鋼筋面積
    For i = DATA_ROW_START To DATA_ROW_END

        data(i, REBAR_LEFT) = CalRebarArea(data(i, REBAR_LEFT))
        data(i, REBAR_MIDDLE) = CalRebarArea(data(i, REBAR_MIDDLE))
        data(i, REBAR_RIGHT) = CalRebarArea(data(i, REBAR_RIGHT))

    Next

    ' 一二排截面積相加
    For i = DATA_ROW_START To DATA_ROW_END Step 2

        data(i, REBAR_LEFT) = data(i, REBAR_LEFT) + data(i + 1, REBAR_LEFT)
        data(i, REBAR_MIDDLE) = data(i, REBAR_MIDDLE) + data(i + 1, REBAR_MIDDLE)
        data(i, REBAR_RIGHT) = data(i, REBAR_RIGHT) + data(i + 1, REBAR_RIGHT)

    Next

    ' 列出警告
    For i = DATA_ROW_START To DATA_ROW_END Step 4

        ' 計算有效深度
        rebar = Split(RAW_DATA(i, REBAR_LEFT), "-")
        stirrup = Split(RAW_DATA(i, STIRRUP_LEFT), "@")
        Db = Application.VLookup(rebar(1), REBAR_SIZE, DIAMETER, False)
        tie = Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)
        data(i, D) = data(i, H) - (4 + tie + Db / 2)

        Call Norm3_6(data, i)
        Call Norm15_4_2_1(data, i)
        Call Norm15_4_2_2(data, i)
        Call NormMiddleNoMoreThanEndEightyPercentage(data, i)

    Next

End Function

Function CalRebarArea(data)

    tmp = Split(data, "-")

    If tmp(1) <> "" Then

        ' 轉換鋼筋尺寸為截面積
        tmp(1) = Application.VLookup(tmp(1), REBAR_SIZE, CROSS_AREA, False)

        CalRebarArea = tmp(0) * tmp(1)
    Else
        CalRebarArea = 0
    End If

End Function

Function Norm3_6(data, i)
'
' RC規範 3-3, 3-4 不低於14/fy

    code3_3 = 0.8 * Sqr(Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FC, False)) / Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FY, False) * data(i, BW) * data(i, D)
    code3_4 = 14 / Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FY, False) * data(i, BW) * data(i, D)

    ' 請確認是否符合 左端上層筋下限 規定
    If data(i, REBAR_LEFT) < code3_3 Or data(i, REBAR_LEFT) < code3_4 Then
        call PrintWarningMessage("請確認是否符合 左端上層筋下限 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 左端上層筋下限 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

    ' 請確認是否符合 右端上層筋下限 規定
    If data(i, REBAR_RIGHT) < code3_3 Or data(i, REBAR_RIGHT) < code3_4 Then
        call PrintWarningMessage("請確認是否符合 右端上層筋下限 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 右端上層筋下限 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

    ' 請確認是否符合 左端下層筋下限 規定
    If data(i + 2, REBAR_LEFT) < code3_3 Or data(i + 2, REBAR_LEFT) < code3_4 Then
        call PrintWarningMessage("請確認是否符合 左端下層筋下限 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 左端下層筋下限 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

    ' 請確認是否符合 右端下層筋下限 規定
    If data(i + 2, REBAR_RIGHT) < code3_3 Or data(i + 2, REBAR_RIGHT) < code3_4 Then
        call PrintWarningMessage("請確認是否符合 右端下層筋下限 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 右端下層筋下限 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

    If data(i, REBAR_MIDDLE) < code3_3 Or data(i, REBAR_MIDDLE) < code3_4 Then
        call PrintWarningMessage("請確認是否符合 中央上層筋下限 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 中央上層筋下限 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

End Function

Function Norm15_4_2_1(data, i)
'
' RC規範 15.4.2.1 不高於2.2 %

    code15_4_2_1 = Application.Min((Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FC, False) + 100) / (4 * Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FY, False)) * data(i, BW) * data(i, D), 0.025 * data(i, BW) * data(i, D))

    If data(i, REBAR_LEFT) > code15_4_2_1 And data(i, STORY) <> "1F" Then
        call PrintWarningMessage("請確認是否符合 左端上層筋上限 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 左端上層筋上限 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

    If data(i, REBAR_RIGHT) > code15_4_2_1 And data(i, STORY) <> "1F" Then
        call PrintWarningMessage("請確認是否符合 右端上層筋上限 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 右端上層筋上限 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

    If data(i + 2, REBAR_LEFT) > code15_4_2_1 And data(i, STORY) <> "1F" Then
        call PrintWarningMessage("請確認是否符合 左端下層筋上限 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 左端下層筋上限 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

    If data(i + 2, REBAR_RIGHT) > code15_4_2_1 And data(i, STORY) <> "1F" Then
        call PrintWarningMessage("請確認是否符合 右端下層筋上限 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 右端下層筋上限 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

    If data(i + 2, REBAR_MIDDLE) > code15_4_2_1 And data(i, STORY) <> "1F" Then
        call PrintWarningMessage("請確認是否符合 中央下層筋上限 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 中央下層筋上限 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

End Function

Function Norm15_4_2_2(data, i)
'
' RC規範 15.4.2.2 任一點不低於1/4

    Max = Application.Max(data(i, REBAR_LEFT), data(i, REBAR_MIDDLE), data(i, REBAR_RIGHT), data(i + 2, REBAR_LEFT), data(i + 2, REBAR_MIDDLE), data(i + 2, REBAR_RIGHT))
    Min = Application.Min(data(i, REBAR_LEFT), data(i, REBAR_MIDDLE), data(i, REBAR_RIGHT), data(i + 2, REBAR_LEFT), data(i + 2, REBAR_MIDDLE), data(i + 2, REBAR_RIGHT))
    code15_4_2_2 = Min <= Max / 4

    If code15_4_2_2 And data(i, STORY) <> "1F" Then
        call PrintWarningMessage("請確認是否符合 耐震最小量鋼筋 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 耐震最小量鋼筋 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

End Function

Function NormMiddleNoMoreThanEndEightyPercentage(data, i)
'
' 經濟性指標 不多於端部小值80%

    Min = Application.Min(data(i, REBAR_LEFT), data(i, REBAR_RIGHT))
    If data(i, REBAR_MIDDLE) > Min * 0.8 Then
        call PrintWarningMessage("請確認是否符合 中央上層筋相對鋼筋量 規定", i)
        ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 中央上層筋相對鋼筋量 規定"
        ' WARNING_MESSAGE = WARNING_MESSAGE + 1
    End If

End Function

' -------------------------------------------------------------------------

Function Norm13_5_1AndRebarAmountNoBelowTwo()

    For i = DATA_ROW_START To DATA_ROW_END

        For j = REBAR_LEFT To REBAR_RIGHT

            k = 4 * Fix((i - 3) / 4) + 3

            rebar = Split(RAW_DATA(i, j), "-")

            stirrup = Split(RAW_DATA(k, j + 4), "@")

            If rebar(0) = "1" Then

                ' 排除掉1支的狀況，避免除以0
                ' 不少於2支
                call PrintWarningMessage("請確認是否符合 單排支數下限 規定", i)
                ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = RAW_DATA(k, STORY) & " " & RAW_DATA(k, NUMBER) & " 請確認是否符合 單排支數下限 規定"
                ' WARNING_MESSAGE = WARNING_MESSAGE + 1

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
                    call PrintWarningMessage("請確認是否符合 單排支數上限 規定", i)
                    ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = RAW_DATA(k, STORY) & " " & RAW_DATA(k, NUMBER) & " 請確認是否符合 單排支數上限 規定"
                    ' WARNING_MESSAGE = WARNING_MESSAGE + 1
                End If

            End If

        Next
    Next

End Function

Function StirrupSpacingMoreThan10AndLessThan30()

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            stirrup = Split(RAW_DATA(i, j), "@")

            If stirrup(1) < 10 Then
                call PrintWarningMessage("請確認是否符合 箍筋間距下限 規定", i)
                ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = RAW_DATA(i, STORY) & " " & RAW_DATA(i, NUMBER) & " 請確認是否符合 箍筋間距下限 規定"
                ' WARNING_MESSAGE = WARNING_MESSAGE + 1
            ElseIf stirrup(1) > 30 Then
                call PrintWarningMessage("請確認是否符合 箍筋間距上限 規定", i)
                ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = RAW_DATA(i, STORY) & " " & RAW_DATA(i, NUMBER) & " 請確認是否符合 箍筋間距上限 規定"
                ' WARNING_MESSAGE = WARNING_MESSAGE + 1
            End If

        Next

    Next

End Function

Function Norm4_6_6_3()
'
' 剪力鋼筋量 最小 3.52/fy

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            stirrup = Split(RAW_DATA(i, j), "@")

            avMin = Application.Max(0.2 * Sqr(Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FC, False)) * data(i, BW) * stirrup(1) / Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FYT, False), 3.5 * data(i, BW) * stirrup(1) / Application.VLookup(data(i, STORY), GENERAL_INFORMATION, FYT, False))
            av = Application.VLookup(stirrup(0), REBAR_SIZE, CROSS_AREA, False) * 2

            If av < avMin Then
                call PrintWarningMessage("請確認是否符合 剪力鋼筋量下限 規定", i)
                ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = RAW_DATA(i, STORY) & " " & RAW_DATA(i, NUMBER) & " 請確認是否符合 剪力鋼筋量下限 規定"
                ' WARNING_MESSAGE = WARNING_MESSAGE + 1
            End If

        Next

    Next

End Function

Function Norm4_6_7_9()
'
' 剪力鋼筋量 最大 Vs <=4Vc * 120%

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            stirrup = Split(RAW_DATA(i, j), "@")
            rebar = Split(RAW_DATA(i, j - 4), "-")
            Db = Application.VLookup(rebar(1), REBAR_SIZE, DIAMETER, False)
            tie = Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)
            effectiveDepth = RAW_DATA(i, H) - (4 + tie + Db / 2)
            av = Application.VLookup(stirrup(0), REBAR_SIZE, CROSS_AREA, False) * 2

            ' code4.1.1.1
            vc = 0.53 * Sqr(Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FC, False)) * RAW_DATA(i, BW) * effectiveDepth

            ' code4.6.7.2
            vs = av * Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FYT, False) * effectiveDepth / stirrup(1)

            ' 規範為 vs <= 4 * vc，由於取整數容易超過，所以放寬標準120%
            If vs > 4 * vc * 1.2 Then
                ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = RAW_DATA(i, STORY) & " " & RAW_DATA(i, NUMBER) & " 請確認是否符合 剪力鋼筋量上限 規定"
                ' WARNING_MESSAGE = WARNING_MESSAGE + 1
                call PrintWarningMessage("請確認是否符合 剪力鋼筋量上限 規定", i)
            End If

        Next

    Next

End Function

Function Norm3_8_1()
'
' 深梁 L/H<=4

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        If RAW_DATA(i, BEAM_LENGTH) <> "" And RAW_DATA(i, SUPPORT) <> "" And (RAW_DATA(i, BEAM_LENGTH) - RAW_DATA(i, SUPPORT)) <= 4 * RAW_DATA(i, H) Then
            call PrintWarningMessage("請確認是否為深梁", i)
            ' Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = RAW_DATA(i, STORY) & " " & RAW_DATA(i, NUMBER) & " 請確認是否為深梁"
            ' WARNING_MESSAGE = WARNING_MESSAGE + 1
        End If

    Next

End Function

Function PrintWarningMessage(warinigMessageCode, i)
'
' PrintWarningMessage

    Cells(WARNING_MESSAGE, WARNING_MESSAGE_POSITION) = RAW_DATA(i, STORY) & " " & RAW_DATA(i, NUMBER) & " " & warinigMessageCode
    OK_MESSAGE(i) = warinigMessageCode & vbCrLf & OK_MESSAGE(i)
    WARNING_MESSAGE = WARNING_MESSAGE + 1

End Function

Function PrintOKMessage()
'
' PrintOKMessage

    Worksheets("大梁配筋").Activate

    ' 不知道為什麼不能直接給值，只好用 for loop
    For i = DATA_ROW_START To DATA_ROW_END Step 4
        If OK_MESSAGE(i) = "" Then
            OK_MESSAGE(i) = "(S), (E), (i) - check 結果 ok"
            Cells(i, OK_MESSAGE_POSITION).Style = "好"
        Else
            Cells(i, OK_MESSAGE_POSITION).Style = "壞"
        End If
        Cells(i, OK_MESSAGE_POSITION) = OK_MESSAGE(i)
    Next
    ' Range(Cells(3, 16), Cells(106, 16)) = OK_MESSAGE()
    ' Range(Cells(3, 16), Cells(106, 16)) = OK_MESSAGE()
    ' Range(Cells(DATA_ROW_START, OK_MESSAGE_POSITION), Cells(DATA_ROW_END, OK_MESSAGE_POSITION)) = OK


End Function

' -------------------------------------------------------------------------

Sub Girder()
'
' * 目的
'       check Girder Norm
'       Norm3_6
'       Norm15_4_2_1
'       Norm15_4_2_2
'       NormMiddleNoMoreThanEndEightyPercentage
'       Norm13_5_1
'       RebarAmountNoBelowTwo
'       StirrupSpacingMoreThan10AndLessThan30
'       Norm4_6_7_9
'       Norm3_8_1

' * 環境
'       Excel

' * 輸出入格式
'       輸入：大梁配筋
'       輸出：Expert Check

' * 執行時間
'       0.21 Sec

' * 輸出結果的精確度與檢驗方式
'

    Time0 = Timer

    WARNING_MESSAGE = 3
    GENERAL_INFORMATION = GetGeneralInformation()
    REBAR_SIZE = GetRebarSize()
    RAW_DATA = GetData()
    willBeModifyToRatioData = GetData()

    Call Initialize
    Call AboutRatioNorm(willBeModifyToRatioData)
    Call Norm13_5_1AndRebarAmountNoBelowTwo
    Call StirrupSpacingMoreThan10AndLessThan30
    Call Norm4_6_7_9
    Call Norm3_8_1
    Call PrintOKMessage

    If Timer - Time0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - Time0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - Time0) / 60, 2) & " Min", vbOKOnly
    End If

End Sub
