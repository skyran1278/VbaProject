Dim GIRDER_WARNING_MESSAGE, GENERAL_INFORMATION, REBAR_SIZE, RAW_DATA

' data 資料命名
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
Const BEAM_LONG = 13
Const SUPPORT = 14
Const LOCATION = 15

Sub Girder()
'
'
'
'
    GIRDER_WARNING_MESSAGE = 7
    GENERAL_INFORMATION = ReadGeneralInformation()
    REBAR_SIZE = ReadRebarSize()
    RAW_DATA = ReadData()
    willBeModifyData = ReadData()

    Call Initialize
    Call AllRebar
    Call AboutRatioNorm(willBeModifyData)



End Sub

Function ReadGeneralInformation()
'
'
    Worksheets("General Information").Activate
    Dim arr()
    rowUsed = Cells(Rows.Count, 5).End(xlUp).Row
    columnUsed = 7
    ReDim arr(1 To rowUsed, 1 To columnUsed)

    For i = 1 To rowUsed
        For j = 1 To columnUsed
            arr(i, j) = Cells(i, j + 3)
        Next
    Next

    ReadGeneralInformation = arr()

End Function

Function ReadRebarSize()
'
'
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

    ReadRebarSize = arr()

End Function

Function ReadData()
'
'
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

    ReadData = arr()

End Function

Function Initialize()

    Worksheets("Expert Check").Activate
    Columns(3).ClearContents
    Cells(5, 3) = "Warning Message"
    Cells(6, 3) = "Girder"

End Function

Function AllRebar()
'
'
    Worksheets("Expert Check").Activate
    dataRowUsed = UBound(RAW_DATA)
    dataRowStart = 3

    For i = dataRowStart To dataRowUsed

        Call rename1(i)
        Call rename2(i)


    Next

End Function

Function rename1(i)

    For k = REBAR_LEFT To REBAR_RIGHT

        j = 4 * Fix((i - 3) / 4) + 3

        tmp = Split(RAW_DATA(i, k), "-")

        If tmp(0) <> "" And tmp(0) < 2 Then
            Cells(GIRDER_WARNING_MESSAGE, 3) = RAW_DATA(j, STORY) & " " & RAW_DATA(j, NUMBER) & " 請確認是否符合 單排支數下限 規定"
            GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
        End If

    Next

End Function

Function rename2(i)

    For k = REBAR_LEFT To REBAR_RIGHT

        j = 4 * Fix((i - 3) / 4) + 3

        rebar = Split(RAW_DATA(i, k), "-")

        stirrup = Split(RAW_DATA(j, k + 4), "@")

        If rebar(0) <> "" Then

            Db = Application.VLookup(rebar(1), REBAR_SIZE, 7, False)
            tie = Application.VLookup(stirrup(0), REBAR_SIZE, 7, False)

            Max = Fix((RAW_DATA(j, BW) - 4 * 2 - tie * 2 - Db) / (2 * Db)) + 1

            If CInt(rebar(0)) > Max Then
                Cells(GIRDER_WARNING_MESSAGE, 3) = RAW_DATA(j, STORY) & " " & RAW_DATA(j, NUMBER) & " 請確認是否符合 單排支數上限 規定"
                GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
            End If

        End If

    Next

End Function

' Fix((data(i, BW) - 4 * 2 - TieDiameter * 2 - Db) / (2 * Db)) + 1

Function AboutRatioNorm(data)
'
'
    Worksheets("Expert Check").Activate
    dataRowUsed = UBound(data)
    dataRowStart = 3

    ' 計算鋼筋面積
    For i = dataRowStart To dataRowUsed

        data(i, REBAR_LEFT) = CalRebarArea(data(i, REBAR_LEFT))
        data(i, REBAR_MIDDLE) = CalRebarArea(data(i, REBAR_MIDDLE))
        data(i, REBAR_RIGHT) = CalRebarArea(data(i, REBAR_RIGHT))

    Next

    ' 一二排截面積相加
    For i = dataRowStart To dataRowUsed Step 2

        data(i, REBAR_LEFT) = data(i, REBAR_LEFT) + data(i + 1, REBAR_LEFT)
        data(i, REBAR_MIDDLE) = data(i, REBAR_MIDDLE) + data(i + 1, REBAR_MIDDLE)
        data(i, REBAR_RIGHT) = data(i, REBAR_RIGHT) + data(i + 1, REBAR_RIGHT)

    Next

    ' 列出警告
    For i = dataRowStart To dataRowUsed Step 4

        ' 計算有效深度
        data(i, D) = data(i, H) - (4 + 1.27 + 2.54 / 2)

        Call Norm3_6(data, i)
        Call Norm15_4_2_1(data, i)
        Call Norm15_4_2_2(data, i)
        Call NormRelative(data, i)

    Next

End Function

Function CalRebarArea(data)

    tmp = Split(data, "-")

    If tmp(1) <> "" Then

        ' 轉換鋼筋尺寸為截面積
        tmp(1) = Application.VLookup(tmp(1), REBAR_SIZE, 10, False)

        CalRebarArea = tmp(0) * tmp(1)
    Else
        CalRebarArea = 0
    End If

End Function

Function Norm3_6(data, i)
'
' RC規範 3-3, 3-4 不低於14/fy

    code3_3 = 0.8 * Sqr(Application.VLookup(data(i, STORY), GENERAL_INFORMATION, 3, False)) / Application.VLookup(data(i, STORY), GENERAL_INFORMATION, 2, False) * data(i, BW) * data(i, D)
    code3_4 = 14 / Application.VLookup(data(i, STORY), GENERAL_INFORMATION, 2, False) * data(i, BW) * data(i, D)

    ' 請確認是否符合 左端上層筋下限 規定
    If data(i, REBAR_LEFT) < code3_3 Or data(i, REBAR_LEFT) < code3_4 Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 左端上層筋下限 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

    ' 請確認是否符合 右端上層筋下限 規定
    If data(i, REBAR_RIGHT) < code3_3 Or data(i, REBAR_RIGHT) < code3_4 Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 右端上層筋下限 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

    ' 請確認是否符合 左端下層筋下限 規定
    If data(i + 2, REBAR_LEFT) < code3_3 Or data(i + 2, REBAR_LEFT) < code3_4 Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 左端下層筋下限 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

    ' 請確認是否符合 右端下層筋下限 規定
    If data(i + 2, REBAR_RIGHT) < code3_3 Or data(i + 2, REBAR_RIGHT) < code3_4 Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 右端下層筋下限 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

    If data(i, REBAR_MIDDLE) < code3_3 Or data(i, REBAR_MIDDLE) < code3_4 Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 中央上層筋下限 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

End Function

Function Norm15_4_2_1(data, i)
'
' RC規範 15.4.2.1 不高於2.2 %

    code15_4_2_1 = Application.Min((Application.VLookup(data(i, STORY), GENERAL_INFORMATION, 3, False) + 100) / (4 * Application.VLookup(data(i, STORY), GENERAL_INFORMATION, 2, False)) * data(i, BW) * data(i, D), 0.025 * data(i, BW) * data(i, D))

    If data(i, REBAR_LEFT) > code15_4_2_1 And data(i, STORY) <> "1F" Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 左端上層筋上限 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

    If data(i, REBAR_RIGHT) > code15_4_2_1 And data(i, STORY) <> "1F" Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 右端上層筋上限 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

    If data(i + 2, REBAR_LEFT) > code15_4_2_1 And data(i, STORY) <> "1F" Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 左端下層筋上限 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

    If data(i + 2, REBAR_RIGHT) > code15_4_2_1 And data(i, STORY) <> "1F" Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 右端下層筋上限 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

    If data(i + 2, REBAR_MIDDLE) > code15_4_2_1 And data(i, STORY) <> "1F" Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 中央下層筋上限 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

End Function

Function Norm15_4_2_2(data, i)
'
' RC規範 15.4.2.2 任一點不低於1/4

    Max = Application.Max(data(i, REBAR_LEFT), data(i, REBAR_MIDDLE), data(i, REBAR_RIGHT), data(i + 2, REBAR_LEFT), data(i + 2, REBAR_MIDDLE), data(i + 2, REBAR_RIGHT))
    Min = Application.Min(data(i, REBAR_LEFT), data(i, REBAR_MIDDLE), data(i, REBAR_RIGHT), data(i + 2, REBAR_LEFT), data(i + 2, REBAR_MIDDLE), data(i + 2, REBAR_RIGHT))
    code15_4_2_2 = Min <= Max / 4

    If code15_4_2_2 And data(i, STORY) <> "1F" Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 耐震最小量鋼筋 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If

End Function

Function NormRelative(data, i)
'
' 經濟性指標 不多於端部小值80%

    Min = Application.Min(data(i, REBAR_LEFT), data(i, REBAR_RIGHT))
    If data(i, REBAR_MIDDLE) > Min * 0.8 Then
        Cells(GIRDER_WARNING_MESSAGE, 3) = data(i, STORY) & " " & data(i, NUMBER) & " 請確認是否符合 中央上層筋相對鋼筋量 規定"
        GIRDER_WARNING_MESSAGE = GIRDER_WARNING_MESSAGE + 1
    End If


End Function


