Dim GENERAL_INFORMATION, REBAR_SIZE, RAW_DATA, RATIO_DATA, DATA_ROW_END, DATA_ROW_START, MESSAGE()

' RAW_DATA 資料命名
Const STORY = 1
Const NUMBER = 2
Const WIDTH_X = 3
Const WIDTH_Y = 4
Const REBAR = 5
Const REBAR_X = 6
Const REBAR_Y = 7
Const BOUND_AREA = 8
Const NON_BOUND_AREA = 9
Const TIE_X = 10
Const TIE_Y = 11

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
Const MESSAGE_POSITION = 12

Function Norm15_5_4_1()

    For i = DATA_ROW_START To DATA_ROW_END

        fcColumn = Application.VLookup(RATIO_DATA(i, STORY), GENERAL_INFORMATION, FC_COLUMN, False)
        fytColumn = Application.VLookup(RATIO_DATA(i, STORY), GENERAL_INFORMATION, FYT, False)
        stirrup = Split(RAW_DATA(i, BOUND_AREA), "@")
        s = stirrup(1)
        bcX = RAW_DATA(i, WIDTH_X) - 4 * 2 - Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)
        bcY = RAW_DATA(i, WIDTH_Y) - 4 * 2 - Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)
        ashX = stirrup(0) * (RAW_DATA(i, TIE_X) + 2)
        ashY = stirrup(0) * (RAW_DATA(i, TIE_Y) + 2)
        ashDivideBc = Application.min(ashX / bcX, ashY / bcY )
        ag = WIDTH_X * WIDTH_Y
        ach = (RAW_DATA(i, WIDTH_X) - 4 * 2) * (RAW_DATA(i, WIDTH_Y) - 4 * 2)
        code15_3 = 0.3 * s * fcColumn / fytColumn * (ag / ach - 1)
        code15_4 = 0.09 * s * fcColumn / fytColumn
        If ashDivideBc < code15_3 or ashDivideBc < code15_4 Then
            Call WarningMessage("請確認是否符合 橫向鋼筋 規定", i)
        End If

    Next

End Function

' -------------------------------------------------------------------------
' -------------------------------------------------------------------------

Private Sub Class_Initialize()
' Called automatically when class is created
' GetGeneralInformation
' GetRebarSize

    ' 排序
    Worksheets("Z").Range(Cells(7, 3), Cells(zRowUsed, 10)).Sort _
        Key1:=Range(Cells(8, 10), Cells(zRowUsed, 10)), Order1:=xlAscending, _
        Key2:=Range(Cells(8, 8), Cells(zRowUsed, 8)), Order2:=xlDescending, Header:=xlYes

    Call GetGeneralInformation
    Call GetRebarSize


End Sub

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

    ' 之所以 + 1 ，是為了之後不要超出索引範圍準備
    rowUsed = Cells(Rows.Count, 5).End(xlUp).Row + 1

    columnUsed = 11

    RAW_DATA = Range(Cells(rowStart, columnStart), Cells(rowUsed, columnUsed)).Value

End Function

Function RatioData()

    ' 計算鋼筋面積
    For i = DATA_ROW_START To DATA_ROW_END
        RATIO_DATA(i, REBAR) = CalRebarArea(RATIO_DATA(i, REBAR))
    Next

    ' 計算箍筋面積
    ' For i = DATA_ROW_START To DATA_ROW_END Step 4
    '     For j = STIRRUP_LEFT To STIRRUP_RIGHT
    '         RATIO_DATA(i, j) = CalStirrupArea(RATIO_DATA(i, j))
    '     Next
    ' Next

    ' 計算側筋面積
    ' For i = DATA_ROW_START To DATA_ROW_END Step 4
    '     RATIO_DATA(i, SIDE_REBAR) = CalSideRebarArea(RATIO_DATA(i, SIDE_REBAR))
    ' Next

    ' 計算有效深度
    ' For i = DATA_ROW_START To DATA_ROW_END Step 4

    '     rebar = Split(RAW_DATA(i, REBAR_LEFT), "-")
    '     stirrup = Split(RAW_DATA(i, STIRRUP_LEFT), "@")
    '     Db = Application.VLookup(rebar(1), REBAR_SIZE, DIAMETER, False)
    '     tie = Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)
    '     RATIO_DATA(i, D) = RATIO_DATA(i, H) - (4 + tie + Db / 2)

    ' Next

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

' Function CalStirrupArea(rebar)

'     tmp = Split(rebar, "@")

'     ' 轉換鋼筋尺寸為截面積
'     tmp(0) = Application.VLookup(tmp(0), REBAR_SIZE, CROSS_AREA, False)

'     ' 字串轉為數字
'     CalStirrupArea = 2 * tmp(0)

' End Function

' Function CalSideRebarArea(rebar)

'     If rebar <> "-" Then

'         rebar = Left(rebar, Len(rebar) - 2)

'         tmp = Split(rebar, "#")

'         ' 轉換鋼筋尺寸為截面積
'         tmp(1) = Application.VLookup("#" & tmp(1), REBAR_SIZE, CROSS_AREA, False)

'         ' 對稱雙排
'         CalSideRebarArea = 2 * tmp(1)

'     Else
'         CalSideRebarArea = 0
'     End If

' End Function

Function Initialize()
'
' DATA_ROW_START
' DATA_ROW_END
' MESSAGE
' RatioData

    Columns(MESSAGE_POSITION).ClearContents
    Cells(1, MESSAGE_POSITION) = "Warning Message"
    DATA_ROW_START = 3

    ' 之所以 - 1 ，是為了還原取到的位置，讓之後不要超出索引範圍準備
    DATA_ROW_END = UBound(RAW_DATA) - 1

    ReDim MESSAGE(DATA_ROW_START To DATA_ROW_END)

    RATIO_DATA = RAW_DATA

    Call RatioData

End Function

Function WarningMessage(warinigMessageCode, i)

    MESSAGE(i) = warinigMessageCode & vbCrLf & MESSAGE(i)

End Function

Function PrintMessage()

    ' Worksheets("大梁配筋").Activate

    ' 不知道為什麼不能直接給值，只好用 for loop
    ' Range(Cells(DATA_ROW_START, MESSAGE_POSITION), Cells(DATA_ROW_END, MESSAGE_POSITION)) = MESSAGE()
    For i = DATA_ROW_START To DATA_ROW_END
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

Private Sub Class_Terminate()

    ' Called automatically when all references to class instance are removed

End Sub

' -------------------------------------------------------------------------
' -------------------------------------------------------------------------

Function EconomicSmooth()

    For i = DATA_ROW_START To DATA_ROW_END

        ' 3 case
        isUpperLimit =  RAW_DATA(i, NUMBER) <> RAW_DATA(i - 1, NUMBER) And RAW_DATA(i, NUMBER) = RAW_DATA(i + 1, NUMBER)
        isMiddle =  RAW_DATA(i, NUMBER) = RAW_DATA(i - 1, NUMBER) And RAW_DATA(i, NUMBER) = RAW_DATA(i + 1, NUMBER)
        isLowerLimit =  RAW_DATA(i, NUMBER) = RAW_DATA(i - 1, NUMBER) And RAW_DATA(i, NUMBER) <> RAW_DATA(i + 1, NUMBER)

        noSmoothDown = RATIO_DATA(i + 1, REBAR) < RATIO_DATA(i, REBAR) * 0.7 And RATIO_DATA(i + 1, REBAR) <> 0
        noSmoothUp = RATIO_DATA(i - 1, REBAR) < RATIO_DATA(i, REBAR) * 0.6 And RATIO_DATA(i - 1, REBAR) <> 0

        If isMiddle and noSmoothDown Then
            Call WarningMessage("請確認是否符合 Smooth Down 規定", i)
        elseif isMiddle and noSmoothUp Then
            Call WarningMessage("請確認是否符合 Smooth Up 規定", i)
        End If

        If isUpperLimit and noSmoothDown Then
            Call WarningMessage("請確認是否符合 Smooth Down 規定", i)
        End If

        If isLowerLimit and noSmoothUp Then
            Call WarningMessage("請確認是否符合 Smooth Up 規定", i)
        End If

    Next

End Function

