Dim MESSAGE(), GENERAL_INFORMATION, REBAR_SIZE, RAW_DATA, RATIO_DATA, DATA_ROW_END, DATA_ROW_START, REBAR_NUMBER()

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
'
' 多了排序，邊界值改變
'

    Worksheets(sheet).Activate

    rowStart = 1
    columnStart = 1

    ' 之所以 + 1 ，是為了之後不要超出索引範圍準備
    rowUsed = Cells(Rows.Count, 5).End(xlUp).Row + 1

    columnUsed = 11

    ' 排序
    Range(Cells(3, columnStart), Cells(rowUsed - 1, columnUsed)).Sort _
        Key1:=Range(Cells(3, NUMBER), Cells(rowUsed - 1, NUMBER)), Order1:=xlAscending

    For i = rowStart To rowUsed
        Cells(i, REBAR) = Trim(Cells(i, REBAR))
    Next

    RAW_DATA = Range(Cells(rowStart, columnStart), Cells(rowUsed, columnUsed)).Value

End Function

Function RatioData()
'
' 主筋比、箍筋與繫筋面積
'

    ' 計算鋼筋比
    For i = DATA_ROW_START To DATA_ROW_END
        RATIO_DATA(i, REBAR) = CalRebarArea(RATIO_DATA(i, REBAR)) / (RAW_DATA(i, WIDTH_X) * RAW_DATA(i, WIDTH_Y))
    Next

    ' 計算箍筋面積
    For i = DATA_ROW_START To DATA_ROW_END
        stirrup = Split(RAW_DATA(i, BOUND_AREA), "@")
        stirrup = Application.VLookup(stirrup(0), REBAR_SIZE, CROSS_AREA, False)
        RATIO_DATA(i, TIE_X) = stirrup * (RAW_DATA(i, TIE_X) + 2)
        RATIO_DATA(i, TIE_Y) = stirrup * (RAW_DATA(i, TIE_Y) + 2)
    Next

    ' 計算有效深度
    ' For i = DATA_ROW_START To DATA_ROW_END Step 4

    '     REBAR = Split(RAW_DATA(i, REBAR_LEFT), "-")
    '     stirrup = Split(RAW_DATA(i, STIRRUP_LEFT), "@")
    '     Db = Application.VLookup(REBAR(1), REBAR_SIZE, DIAMETER, False)
    '     tie = Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)
    '     RATIO_DATA(i, D) = RATIO_DATA(i, H) - (4 + tie + Db * 1.5)

    ' Next

End Function

Function CalRebarArea(REBAR)

    tmp = Split(REBAR, "-")

    ' 轉換鋼筋尺寸為截面積
    tmp(1) = Application.VLookup(tmp(1), REBAR_SIZE, CROSS_AREA, False)

    CalRebarArea = tmp(0) * tmp(1)

End Function

' Function CalStirrupArea(REBAR)

'     tmp = Split(REBAR, "@")

'     ' 轉換鋼筋尺寸為截面積
'     tmp(0) = Application.VLookup(tmp(0), REBAR_SIZE, CROSS_AREA, False)

'     ' 字串轉為數字
'     CalStirrupArea = 2 * tmp(0)

' End Function

' Function CalSideRebarArea(REBAR)

'     If REBAR <> "-" Then

'         REBAR = Left(REBAR, Len(REBAR) - 2)

'         tmp = Split(REBAR, "#")

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

Function PrintRebarRatio()

    rowStart = 1
    rowUsed = UBound(RATIO_DATA)
    columnStart = 13
    columnUsed = 17

    Range(Cells(rowStart, columnStart), Cells(rowUsed, columnUsed)) = RATIO_DATA
    Range(Columns(13), Columns(16)).Hidden = True

    Call FontSetting

End Function

Function PrintRebarRatioInAnotherSheets()

    Worksheets("柱鋼筋比").Activate
    rowStart = 1
    rowUsed = UBound(RATIO_DATA)
    columnStart = 1
    columnUsed = 5

    Range(Cells(rowStart, columnStart), Cells(rowUsed, columnUsed)) = RATIO_DATA

    Call FontSetting

    Worksheets("鋼筋號數比").Activate
    rowStart = 3
    rowUsed = UBound(REBAR_NUMBER) + 1
    columnStart = 2
    columnUsed = 3

    Range(Cells(rowStart, columnStart), Cells(rowUsed, columnUsed)) = REBAR_NUMBER

End Function

Function FontSetting()

    Cells.Font.Name = "微軟正黑體"
    Cells.Font.Name = "Calibri"

End Function

Private Sub Class_Terminate()

    ' Called automatically when all references to class instance are removed

End Sub

' -------------------------------------------------------------------------
' -------------------------------------------------------------------------

' FIXME: Function Name
' RC EXPERT 增加繫筋的規範  中央繫筋 >= RoundUp((主筋支數 - 1) / 2) - 1
' 修正 X Y 向隔根勾錯誤
Function Norm15_5_4_100()

    For i = DATA_ROW_START To DATA_ROW_END

        If RAW_DATA(i, TIE_X) < Int((RAW_DATA(i, REBAR_Y) - 1) / 2) - 1 Then
            Call WarningMessage("【0405】X向繫筋未符合隔根勾", i)
        End If
        If RAW_DATA(i, TIE_Y) < Int((RAW_DATA(i, REBAR_X) - 1) / 2) - 1 Then
            Call WarningMessage("【0406】Y向繫筋未符合隔根勾", i)
        End If
    Next

End Function

Function EconomicSmooth()
'
' 往上漸縮  不低於60%
' 往下漸縮  不低於70%
' 邏輯感覺蠻奇怪的，或許可以修改。2017/07/07

    For i = DATA_ROW_START To DATA_ROW_END

        ' 3 case
        isUpperLimit = RAW_DATA(i, NUMBER) <> RAW_DATA(i - 1, NUMBER) And RAW_DATA(i, NUMBER) = RAW_DATA(i + 1, NUMBER)
        isMiddle = RAW_DATA(i, NUMBER) = RAW_DATA(i - 1, NUMBER) And RAW_DATA(i, NUMBER) = RAW_DATA(i + 1, NUMBER)
        isLowerLimit = RAW_DATA(i, NUMBER) = RAW_DATA(i - 1, NUMBER) And RAW_DATA(i, NUMBER) <> RAW_DATA(i + 1, NUMBER)

        noSmoothDown = RATIO_DATA(i + 1, REBAR) < RATIO_DATA(i, REBAR) * 0.7
        noSmoothUp = RATIO_DATA(i - 1, REBAR) < RATIO_DATA(i, REBAR) * 0.6

        If isMiddle And noSmoothDown Then
            Call WarningMessage("【0401】請確認上層柱主筋量，漸縮是否過大", i)
        ElseIf isMiddle And noSmoothUp Then
            Call WarningMessage("【0402】請確認本層柱主筋量，漸縮是否過大", i)
        End If

        If isUpperLimit And noSmoothDown Then
            Call WarningMessage("【0401】請確認上層柱主筋量，漸縮是否過大", i)
        End If

        If isLowerLimit And noSmoothUp Then
            Call WarningMessage("【0402】請確認本層柱主筋量，漸縮是否過大", i)
        End If

    Next

End Function

' FIXME: X Y 好像有錯誤
Function Norm15_5_4_1()
'
' 矩形閉合箍筋及繫筋之總斷面積 Ash 不得小於式(15-3)及式(15-4)之值。

    For i = DATA_ROW_START To DATA_ROW_END

        fcColumn = Application.VLookup(RATIO_DATA(i, STORY), GENERAL_INFORMATION, FC_COLUMN, False)
        fytColumn = Application.VLookup(RATIO_DATA(i, STORY), GENERAL_INFORMATION, FYT, False)
        stirrup = Split(RAW_DATA(i, BOUND_AREA), "@")
        s = stirrup(1)
        bcX = RAW_DATA(i, WIDTH_X) - 4 * 2 - Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)
        bcY = RAW_DATA(i, WIDTH_Y) - 4 * 2 - Application.VLookup(stirrup(0), REBAR_SIZE, DIAMETER, False)
        ashDivideBc = Application.Min(RATIO_DATA(i, TIE_X) / bcX, RATIO_DATA(i, TIE_Y) / bcY)
        ag = RAW_DATA(i, WIDTH_X) * RAW_DATA(i, WIDTH_Y)
        ach = (RAW_DATA(i, WIDTH_X) - 4 * 2) * (RAW_DATA(i, WIDTH_Y) - 4 * 2)
        code15_3 = 0.3 * s * fcColumn / fytColumn * (ag / ach - 1)
        code15_4 = 0.09 * s * fcColumn / fytColumn
        If ashDivideBc < code15_3 Or ashDivideBc < code15_4 Then
            Call WarningMessage("【0403】請確認橫向鋼筋，是否符合 規範 15.5.4.1 規定", i)
        End If

    Next

End Function

Function EconomicTopStoryRebar()
'
' 一定要有 1F 和 RF
' 頂樓區鋼筋比不大於 1.2 %
'
    For i = 1 To UBound(GENERAL_INFORMATION)
        If GENERAL_INFORMATION(i, STORY) = "1F" Then
            firstStory = i
        ElseIf GENERAL_INFORMATION(i, STORY) = "RF" Then
            topStory = i
        End If
    Next

    checkStoryNumber = Fix((topStory - firstStory + 1) / 4)

    For i = DATA_ROW_START To DATA_ROW_END
        For j = topStory - checkStoryNumber + 1 To topStory

            If RAW_DATA(i, STORY) = GENERAL_INFORMATION(j, STORY) And RATIO_DATA(i, REBAR) > 0.01 * 1.2 Then
                    Call WarningMessage("【0404】請確認高樓區鋼筋比，是否超過 1.2 %", i)
            End If

        Next

    Next

End Function

Function CountRebarNumber()

    rowStart = 2
    rowEnd = UBound(REBAR_SIZE)
    ReDim REBAR_NUMBER(rowStart To rowEnd, 1 To 2)

    For i = DATA_ROW_START To DATA_ROW_END

        rebarNumber = Split(RAW_DATA(i, REBAR), "-")(1)
        boundStirrupNumber = Split(RAW_DATA(i, BOUND_AREA), "@")(0)
        nonBoundStirrupNumber = Split(RAW_DATA(i, NON_BOUND_AREA), "@")(0)

        For j = rowStart To rowEnd

            If rebarNumber = REBAR_SIZE(j, 1) Then
                REBAR_NUMBER(j, 1) = REBAR_NUMBER(j, 1) + 1
            End If

            If boundStirrupNumber = REBAR_SIZE(j, 1) Then
                REBAR_NUMBER(j, 2) = REBAR_NUMBER(j, 2) + 1
            End If

            If nonBoundStirrupNumber = REBAR_SIZE(j, 1) Then
                REBAR_NUMBER(j, 2) = REBAR_NUMBER(j, 2) + 1
            End If

        Next

    Next

    For i = rowStart To rowEnd
        sumRebarNumber = sumRebarNumber + REBAR_NUMBER(i, 1)
        sumStirrupNumber = sumStirrupNumber + REBAR_NUMBER(i, 2)
    Next

    For i = rowStart To rowEnd
        If REBAR_NUMBER(i, 1) <> 0 Then
            REBAR_NUMBER(i, 1) = REBAR_NUMBER(i, 1) / sumRebarNumber
        End If
        If REBAR_NUMBER(i, 2) <> 0 Then
            REBAR_NUMBER(i, 2) = REBAR_NUMBER(i, 2) / sumStirrupNumber
        End If
    Next

End Function
