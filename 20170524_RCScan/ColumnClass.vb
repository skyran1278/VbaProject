Private MESSAGE(), GENERAL_INFORMATION, REBAR_SIZE, RAW_DATA, RATIO_DATA, DATA_ROW_END, DATA_ROW_START, FIRST_STORY, REBAR_NUMBER(), WS As Worksheet

' RAW_DATA 資料命名
Private Const STORY = 1
Private Const NUMBER = 2
Private Const WIDTH_X = 3
Private Const WIDTH_Y = 4
Private Const REBAR = 5
Private Const REBAR_X = 6
Private Const REBAR_Y = 7
Private Const BOUND_AREA = 8
Private Const NON_BOUND_AREA = 9
Private Const TIE_X = 10
Private Const TIE_Y = 11

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
Private Const MESSAGE_POSITION = 12

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
'
' 多了排序，邊界值改變
'

    Set WS = Worksheets(sheet)

    rowStart = 1
    columnStart = 1

    ' 之所以 + 1 ，是為了之後不要超出索引範圍準備
    rowUsed = WS.Cells(Rows.Count, 5).End(xlUp).Row + 1

    columnUsed = 11

    ' 排序
    WS.Range(WS.Cells(3, columnStart), WS.Cells(rowUsed - 1, columnUsed)).Sort _
        Key1:=WS.Range(WS.Cells(3, NUMBER), WS.Cells(rowUsed - 1, NUMBER)), Order1:=xlAscending

    ' 裁掉多餘的空白
    For i = rowStart To rowUsed
        WS.Cells(i, REBAR) = Trim(WS.Cells(i, REBAR))
    Next

    RAW_DATA = WS.Range(WS.Cells(rowStart, columnStart), WS.Cells(rowUsed, columnUsed))

End Function


Function NoData()
'
' 如果沒有資料，就回傳 false
'
' @returns NoData(Boolean)

    NoData = UBound(RAW_DATA) < 4

End Function


Function CalRebarArea(REBAR)

    tmp = Split(REBAR, "-")

    ' 轉換鋼筋尺寸為截面積
    tmp(1) = Application.VLookup(tmp(1), REBAR_SIZE, CROSS_AREA, False)

    CalRebarArea = tmp(0) * tmp(1)

End Function


Function Initialize()
'
' DATA_ROW_START
' DATA_ROW_END
' MESSAGE
' RatioData

    WS.Range(WS.Columns(MESSAGE_POSITION), WS.Columns(MESSAGE_POSITION + 1)).ClearContents
    WS.Cells(1, MESSAGE_POSITION) = "Warning Message"
    WS.Cells(1, MESSAGE_POSITION + 1) = "鋼筋比"
    DATA_ROW_START = 3

    ' 之所以 - 1 ，是為了還原取到的位置，讓之後不要超出索引範圍準備
    DATA_ROW_END = UBound(RAW_DATA) - 1

    ReDim MESSAGE(DATA_ROW_START To DATA_ROW_END)

    ReDim RATIO_DATA(LBound(RAW_DATA, 1) To UBound(RAW_DATA, 1), LBound(RAW_DATA, 2) To UBound(RAW_DATA, 2))

    Call RatioData

End Function


Function RatioData()
'
' 主筋比、箍筋與繫筋面積
'
    ' 樓層數字化，用以比較上下樓層。
    For i = DATA_ROW_START To DATA_ROW_END
        RATIO_DATA(i, STORY) = Application.Match(RAW_DATA(i, STORY), Application.Index(GENERAL_INFORMATION, 0, STORY), 0)
    Next

    ' 計算鋼筋比
    For i = DATA_ROW_START To DATA_ROW_END
        RATIO_DATA(i, REBAR) = CalRebarArea(RAW_DATA(i, REBAR)) / (RAW_DATA(i, WIDTH_X) * RAW_DATA(i, WIDTH_Y))
    Next

    ' 計算箍筋與繫筋面積
    For i = DATA_ROW_START To DATA_ROW_END
        stirrup = Split(RAW_DATA(i, BOUND_AREA), "@")
        stirrup = Application.VLookup(stirrup(0), REBAR_SIZE, CROSS_AREA, False)
        RATIO_DATA(i, TIE_X) = stirrup * (RAW_DATA(i, TIE_X) + 2)
        RATIO_DATA(i, TIE_Y) = stirrup * (RAW_DATA(i, TIE_Y) + 2)
    Next

End Function


Function WarningMessage(warinigMessageCode, i)

    MESSAGE(i) = warinigMessageCode & vbCrLf & MESSAGE(i)

End Function


Function PrintMessage()

    ' 不知道為什麼不能直接給值，只好用 for loop
    ' Range(Cells(DATA_ROW_START, MESSAGE_POSITION), Cells(DATA_ROW_END, MESSAGE_POSITION)) = MESSAGE()
    For i = DATA_ROW_START To DATA_ROW_END
        If MESSAGE(i) = "" Then
            MESSAGE(i) = "(S), (E), (i) - check 結果 ok"
            WS.Cells(i, MESSAGE_POSITION).Style = "好"
        Else
            WS.Cells(i, MESSAGE_POSITION).Style = "壞"
            MESSAGE(i) = Left(MESSAGE(i), Len(MESSAGE(i)) - 1)
        End If
        WS.Cells(i, MESSAGE_POSITION) = MESSAGE(i)
    Next

End Function


Function PrintRebarRatio()

    rowStart = 1
    rowUsed = UBound(RATIO_DATA)
    columnUsed = 13

    WS.Range(WS.Cells(rowStart, columnUsed), WS.Cells(rowUsed, columnUsed)) = Application.Index(RATIO_DATA, 0, REBAR)

    Call FontSetting

End Function


Function PrintRebarRatioInAnotherSheets()

    Dim columnRatio As Worksheet
    Dim rebarRatio As Worksheet
    Set columnRatio = Worksheets("柱鋼筋比")
    Set rebarRatio = Worksheets("鋼筋號數比")

    rowStart = 1
    rowUsed = UBound(RATIO_DATA)
    columnStart = 1
    columnUsed = 5

    columnRatio.Range(columnRatio.Cells(rowStart, columnUsed), columnRatio.Cells(rowUsed, columnUsed)) = Application.Index(RATIO_DATA, 0, REBAR)

    ' 由於修改 RATIO_DATA 樓層部分，改以數字呈現，所以用 RAW_DATA 再覆蓋一次。
    columnRatio.Range(columnRatio.Cells(rowStart, columnStart), columnRatio.Cells(rowUsed, columnUsed - 1)) = RAW_DATA

    Call FontSetting

    rowStart = 3
    rowUsed = UBound(REBAR_NUMBER) + 1
    columnStart = 2
    columnUsed = 3

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
' 以下為實作規範
' -------------------------------------------------------------------------

Function Norm15_5_4_100()
' 增加繫筋的規範  中央繫筋 >= RoundUp((主筋支數 - 1) / 2) - 1
' 已修正 X Y 向隔根勾錯誤


    For i = DATA_ROW_START To DATA_ROW_END

        If RATIO_DATA(i, STORY) < FIRST_STORY Then

            If RAW_DATA(i, TIE_Y) < Int((RAW_DATA(i, REBAR_X) - 1) / 2) - 1 Then
                Call WarningMessage("【0407】Y 向繫筋未符合隔根勾", i)
            End If

            If RAW_DATA(i, TIE_X) < Int((RAW_DATA(i, REBAR_Y) - 1) / 2) - 1 Then
                Call WarningMessage("【0406】X 向繫筋未符合隔根勾", i)
            End If

        End If

    Next

End Function

Function EconomicSmooth()
'
' 往上漸縮  不低於60%
' 往下漸縮  不低於70%
' 邏輯感覺蠻奇怪的，或許可以修改。
' 已修正邏輯

    For i = DATA_ROW_START To DATA_ROW_END

        ' 3 case
        ' 判斷位置
        isUpperLimit = RAW_DATA(i, NUMBER) <> RAW_DATA(i - 1, NUMBER) And RAW_DATA(i, NUMBER) = RAW_DATA(i + 1, NUMBER)
        isMiddle = RAW_DATA(i, NUMBER) = RAW_DATA(i - 1, NUMBER) And RAW_DATA(i, NUMBER) = RAW_DATA(i + 1, NUMBER)
        isLowerLimit = RAW_DATA(i, NUMBER) = RAW_DATA(i - 1, NUMBER) And RAW_DATA(i, NUMBER) <> RAW_DATA(i + 1, NUMBER)

        ' 往下減縮超過 7 成
        sharpDown = RATIO_DATA(i + 1, REBAR) < RATIO_DATA(i, REBAR) * 0.7

        ' 往上減縮超過 6 成
        sharpUp = RATIO_DATA(i - 1, REBAR) < RATIO_DATA(i, REBAR) * 0.6

        If isMiddle And sharpDown Then
            Call WarningMessage("【0402】請確認下層柱主筋量，漸縮是否過大", i)
        End If

        If isMiddle And sharpUp Then
            Call WarningMessage("【0401】請確認上層柱主筋量，漸縮是否過大", i)
        End If

        If isUpperLimit And sharpDown Then
            Call WarningMessage("【0402】請確認下層柱主筋量，漸縮是否過大", i)
        End If

        If isLowerLimit And sharpUp Then
            Call WarningMessage("【0401】請確認上層柱主筋量，漸縮是否過大", i)
        End If

    Next

End Function


Function Norm15_5_4_1()
'
' 矩形閉合箍筋及繫筋之總斷面積 Ash 不得小於式(15-3)及式(15-4)之值。
' 增加為 X Y 向檢驗，並修正 X Y 向相反問題

    For i = DATA_ROW_START To DATA_ROW_END

        If RATIO_DATA(i, STORY) < FIRST_STORY Then

            fcColumn = Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FC_COLUMN, False)
            fytColumn = Application.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FYT, False)

            stirrup = Split(RAW_DATA(i, BOUND_AREA), "@")
            rebarSize = stirrup(0)
            s = stirrup(1)

            bcX = RAW_DATA(i, WIDTH_X) - 4 * 2 - Application.VLookup(rebarSize, REBAR_SIZE, DIAMETER, False)
            bcY = RAW_DATA(i, WIDTH_Y) - 4 * 2 - Application.VLookup(rebarSize, REBAR_SIZE, DIAMETER, False)

            ashX = RATIO_DATA(i, TIE_X)
            ashY = RATIO_DATA(i, TIE_Y)

            ag = RAW_DATA(i, WIDTH_X) * RAW_DATA(i, WIDTH_Y)
            ach = (RAW_DATA(i, WIDTH_X) - 4 * 2) * (RAW_DATA(i, WIDTH_Y) - 4 * 2)

            code15_3_X = 0.3 * s * bcX * fcColumn / fytColumn * (ag / ach - 1)
            code15_3_Y = 0.3 * s * bcY * fcColumn / fytColumn * (ag / ach - 1)
            code15_4_X = 0.09 * s * bcX * fcColumn / fytColumn
            code15_4_Y = 0.09 * s * bcY * fcColumn / fytColumn

            If ashY < code15_3_X Or ashY < code15_4_X Then
                Call WarningMessage("【0404】請確認 Y 向橫向鋼筋，是否符合 規範 15.5.4.1 規定", i)
            End If

            If ashX < code15_3_Y Or ashX < code15_4_Y Then
                Call WarningMessage("【0403】請確認 X 向橫向鋼筋，是否符合 規範 15.5.4.1 規定", i)
            End If

        End If

    Next

End Function


Function EconomicTopStoryRebar()
'
' 一定要有 1F 和 RF
' 頂樓區鋼筋比不大於 1.2 %
'

    topStory = Application.Match("RF", Application.Index(GENERAL_INFORMATION, 0, STORY), 0)

    ' 頂樓區 1/4
    checkStoryNumber = Fix((FIRST_STORY - topStory + 1) / 4) + topStory

    For i = DATA_ROW_START To DATA_ROW_END
        If RATIO_DATA(i, STORY) >= topStory And RATIO_DATA(i, STORY) <= checkStoryNumber And RATIO_DATA(i, REBAR) > 0.01 * 1.2 Then
                Call WarningMessage("【0405】請確認高樓區鋼筋比，是否超過 1.2 %", i)
        End If
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
