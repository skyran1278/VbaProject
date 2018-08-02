Private ran As UTILS_CLASS
Private APP
Private OBJ_ERR_MSG As Collection

Private OBJ_INFO
Private NUM_TOP_STORY
Private NUM_FIRST_STORY

Private OBJ_REBAR_SIZE

Private WS_OUTPUT As Worksheet
Private LB_REBAR
Private UB_REBAR
Private ARR_REBAR

Private ARR_RATIO

' 準備拋棄
' Private REBAR_NUMBER

' ARR_REBAR 資料命名
Private Const COL_STORY = 1
Private Const COL_NUMBER = 2
Private Const COL_WIDTH_X = 3
Private Const COL_WIDTH_Y = 4
Private Const COL_REBAR = 5
Private Const COL_REBAR_X = 6
Private Const COL_REBAR_Y = 7
Private Const COL_BOUND_AREA = 8
Private Const COL_NON_BOUND_AREA = 9
Private Const COL_TIE_X = 10
Private Const COL_TIE_Y = 11
' 輸出資料位置
Private Const COL_MESSAGE = 12
Private Const COL_REBAR_RATIO = 13

' GENERAL_INFORMATION 資料命名
Private Const COL_FY = 2
Private Const COL_FYT = 3
Private Const COL_FC_BEAM = 4
Private Const COL_FC_COLUMN = 5
Private Const COL_SDL = 6
Private Const COL_LL = 7
Private Const COL_BAND = 8
Private Const COL_SLAB = 9
Private Const COL_COVER = 10
Private Const COL_STORY_NUM = 11

' REBAR_SIZE 資料命名
Private Const COL_DB = 7
Private Const COL_AREA = 10



' -------------------------------------------------------------------------
' -------------------------------------------------------------------------

Private Sub Class_Initialize()
' Called automatically when class is created
' GetGeneralInformation
' GetRebarSize

    Set ran = New UTILS_CLASS
    Set APP = Application.WorksheetFunction

    Set OBJ_ERR_MSG = New Collection

    ' 輸出 OBJ_INFO
    Call GetGeneralInformation

    ' 輸出 OBJ_REBAR_SIZE
    Call GetRebarSize

    ' 輸出
    ' WS_OUTPUT
    ' LB_REBAR
    ' UB_REBAR
    ' ARR_REBAR
    Call SortRawData("柱")

    ' ReDim MESSAGE(LB_REBAR To UB_REBAR)

    ReDim ARR_RATIO(LBound(ARR_REBAR, 1) To UBound(ARR_REBAR, 1), LBound(ARR_REBAR, 2) To UBound(ARR_REBAR, 2))

    Call GetRatioData

End Sub


Function GetGeneralInformation()

    Dim wsGeneralInformation As Worksheet
    Set wsGeneralInformation = Worksheets("General Information")

    ' 後面多空出一行，以增加代號
    arrGeneralInformation = ran.GetRangeToArray(wsGeneralInformation, 1, 4, 4, 14)

    lbGeneralInformation = LBound(arrGeneralInformation, 1)
    ubGeneralInformation = UBound(arrGeneralInformation, 1)
    lbColGeneralInformation = LBound(arrGeneralInformation, 2)
    ubColGeneralInformation = UBound(arrGeneralInformation, 2)

    j = 1

    For i = ubGeneralInformation To lbGeneralInformation Step -1
        arrGeneralInformation(i, COL_STORY_NUM) = j
        j = j + 1
    Next i

    ' 掃描是否有沒輸入的數值
    For i = lbGeneralInformation To ubGeneralInformation
        For j = lbColGeneralInformation To ubColGeneralInformation

            If arrGeneralInformation(i, j) = "" Then
                OBJ_ERR_MSG.Add "General Information " & arrGeneralInformation(i, COL_STORY) & " " & arrGeneralInformation(1, j) & " 空白"
            End If

        Next j
    Next i

    Set OBJ_INFO = ran.CreateDictionary(arrGeneralInformation, 1, False)

    ' Use Cells(13, 16).Text instead of .Value
    NUM_TOP_STORY = WarnDicEmpty(OBJ_INFO.Item(wsGeneralInformation.Cells(13, 16).Text), COL_STORY_NUM, "搜尋不到頂樓樓層")
    NUM_FIRST_STORY = WarnDicEmpty(OBJ_INFO.Item(wsGeneralInformation.Cells(14, 16).Text), COL_STORY_NUM, "搜尋不到地面層")

End Function


Private Function WarnDicEmpty(ByVal arr, ByVal value, Optional ByVal warning = "Key is Empty")
'
' 如果 arr 為空，則 show error.
'
' @since 3.0.0
' @param {Array} [arr] 需要驗證的值.
' @param {Number} [value] 陣列位置.
' @param {String} [warning] 錯誤訊息.
' @return {Variant} [value] 需要驗證的值.
' @see dependencies
'

    If Not IsEmpty(arr) Then

        WarnDicEmpty = arr(value)

    Else

        OBJ_ERR_MSG.Add warning
        WarnDicEmpty = Empty

    End If

End Function


Private Function GetRebarSize()

    arrRebarSize = ran.GetRangeToArray(Worksheets("Rebar Size"), 1, 1, 5, 10)

    Set OBJ_REBAR_SIZE = ran.CreateDictionary(arrRebarSize, 1, False)

End Function


Private Function SortRawData(ByVal sheet)
'
' 排序樓層.
'
' @param {String} [sheet]大梁、小梁、地梁.
'

    ' 多抓兩行用來排序
    arrRawData = ran.GetRangeToArray(Worksheets(sheet), 1, 1, 5, 15)

    rowLbRawData = LBound(arrRawData, 1)
    colLbRawData = LBound(arrRawData, 2)
    rowUbRawData = UBound(arrRawData, 1)
    colUbRawData = UBound(arrRawData, 2)

    LB_REBAR = 3

    UB_REBAR = rowUbRawData

    colStoryNum = 14
    colNumberNoC = 15

    For i = LB_REBAR To UB_REBAR

        ' 樓層數字化，用以比較上下樓層。
        arrRawData(i, colStoryNum) = WarnDicEmpty(OBJ_INFO.Item(arrRawData(i, COL_STORY)), COL_STORY_NUM, "請確認 " & arrRawData(i, COL_STORY) & " 是否存在於 General Information")

        ' 裁掉多餘的空白
        arrRawData(i, COL_REBAR) = Trim(arrRawData(i, COL_REBAR))

        ' 去掉 大寫與小寫開頭的 C，用以排序
        arrRawData(i, colNumberNoC) = Right(arrRawData(i, COL_NUMBER), Len(arrRawData(i, COL_NUMBER)) - 1)


    Next

    Set WS_OUTPUT = ThisWorkbook.Sheets.Add(After:=Worksheets("General Information"))

    With WS_OUTPUT

        .Range(.Cells(rowLbRawData, colLbRawData), .Cells(rowUbRawData, colUbRawData)) = arrRawData

        .Range(.Cells(LB_REBAR, colLbRawData), .Cells(rowUbRawData, colUbRawData)).Sort _
            Key1:=.Range(.Cells(LB_REBAR, colNumberNoC), .Cells(rowUbRawData, colNumberNoC)), Order1:=xlAscending, DataOption1:=xlSortNormal, _
            Header:=xlNo, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin

        ' 收入資料
        ' row 之所以 + 1 ，是為了之後不要超出索引範圍準備
        ARR_REBAR = .Range(.Cells(rowLbRawData, colLbRawData), .Cells(rowUbRawData + 1, colUbRawData - 2))

    End With

    ' ' 清空前一次輸入
    WS_OUTPUT.Cells.Clear

End Function


' Function GetData(sheet)
' '
' ' 多了排序，邊界值改變
' '

'     Set WS_OUTPUT = Worksheets(sheet)

'     rowStart = 1
'     columnStart = 1

'     ' 之所以 + 1 ，是為了之後不要超出索引範圍準備
'     rowUsed = WS_OUTPUT.Cells(Rows.Count, 5).End(xlUp).Row + 1

'     columnUsed = 11

'     ' 排序
'     WS_OUTPUT.Range(WS_OUTPUT.Cells(3, columnStart), WS_OUTPUT.Cells(rowUsed - 1, columnUsed)).Sort _
'         Key1:=WS_OUTPUT.Range(WS_OUTPUT.Cells(3, COL_NUMBER), WS_OUTPUT.Cells(rowUsed - 1, COL_NUMBER)), Order1:=xlAscending

'     ' 裁掉多餘的空白
'     For i = rowStart To rowUsed
'         WS_OUTPUT.Cells(i, COL_REBAR) = Trim(WS_OUTPUT.Cells(i, COL_REBAR))
'     Next

'     ARR_REBAR = WS_OUTPUT.Range(WS_OUTPUT.Cells(rowStart, columnStart), WS_OUTPUT.Cells(rowUsed, columnUsed))

' End Function


' Function NoData()
' '
' ' 如果沒有資料，就回傳 false
' '
' ' @returns NoData(Boolean)

'     NoData = UBound(ARR_REBAR) < 4

' End Function


' Function Initialize()
' '
' ' LB_REBAR
' ' UB_REBAR
' ' MESSAGE
' ' RatioData

'     ' WS_OUTPUT.Range(WS_OUTPUT.Columns(COL_MESSAGE), WS_OUTPUT.Columns(COL_MESSAGE + 1)).ClearContents
'     ' WS_OUTPUT.Cells(1, COL_MESSAGE) = "Warning Message"
'     ' LB_REBAR = 3

'     ' 之所以 - 1 ，是為了還原取到的位置，讓之後不要超出索引範圍準備
'     ' UB_REBAR = UBound(ARR_REBAR) - 1

'     ' ReDim MESSAGE(LB_REBAR To UB_REBAR)

'     ReDim ARR_RATIO(LBound(ARR_REBAR, 1) To UBound(ARR_REBAR, 1), LBound(ARR_REBAR, 2) To UBound(ARR_REBAR, 2))

'     Call RatioData

' End Function


Function GetRatioData()
'
' 主筋比、箍筋與繫筋面積
'
    ' 樓層數字化，用以比較上下樓層。
    For i = LB_REBAR To UB_REBAR
        ARR_RATIO(i, COL_STORY) = WarnDicEmpty(OBJ_INFO.Item(ARR_REBAR(i, COL_STORY)), COL_STORY_NUM)
    Next

    ' 計算鋼筋比
    For i = LB_REBAR To UB_REBAR

        ARR_RATIO(i, COL_REBAR) = CalRebarArea(ARR_REBAR(i, COL_REBAR)) / (ARR_REBAR(i, COL_WIDTH_X) * ARR_REBAR(i, COL_WIDTH_Y))

        ARR_REBAR(i, COL_REBAR_RATIO) = ARR_RATIO(i, COL_REBAR)

        If ARR_RATIO(i, COL_REBAR) = 0 Then
            MsgBox "請確認第 " & i & " 列是否有問題.", vbOKOnly, "Error"
        End If

    Next

    ' 計算箍筋與繫筋面積
    For i = LB_REBAR To UB_REBAR
        stirrup = Split(ARR_REBAR(i, COL_BOUND_AREA), "@")
        stirrup = OBJ_REBAR_SIZE.Item(stirrup(0))(COL_AREA)
        ' stirrup = Application.VLookup(stirrup(0), REBAR_SIZE, COL_AREA, False)
        ARR_RATIO(i, COL_TIE_X) = stirrup * (ARR_REBAR(i, COL_TIE_X) + 2)
        ARR_RATIO(i, COL_TIE_Y) = stirrup * (ARR_REBAR(i, COL_TIE_Y) + 2)
    Next

End Function


Function CalRebarArea(COL_REBAR)

    tmp = Split(COL_REBAR, "-")

    If UBound(tmp) < 1 Then
        CalRebarArea = 0
    Else

        ' 轉換鋼筋尺寸為截面積
        tmp(1) = OBJ_REBAR_SIZE.Item(tmp(1))(COL_AREA)

        CalRebarArea = tmp(0) * tmp(1)

    End If

End Function


Function WarningMessage(warningMessageCode, i)

    ARR_REBAR(i, COL_MESSAGE) = warningMessageCode & vbCrLf & ARR_REBAR(i, COL_MESSAGE)

End Function


Function PrintResult()

    rowLbRawData = LBound(ARR_REBAR, 1)
    colLbRawData = LBound(ARR_REBAR, 2)
    rowUbRawData = UBound(ARR_REBAR, 1)
    colUbRawData = UBound(ARR_REBAR, 2)

    For i = LB_REBAR To UB_REBAR

        If ARR_REBAR(i, COL_MESSAGE) = "" Then
            ARR_REBAR(i, COL_MESSAGE) = "(S), (E), (i) - SCAN 結果 ok"
        Else
            WS_OUTPUT.Cells(i, COL_MESSAGE).Style = "壞"
            ARR_REBAR(i, COL_MESSAGE) = Left(ARR_REBAR(i, COL_MESSAGE), Len(ARR_REBAR(i, COL_MESSAGE)) - 1)
        End If

    Next

    With WS_OUTPUT

        .Range(.Cells(rowLbRawData, colLbRawData), .Cells(rowUbRawData, colUbRawData)) = ARR_REBAR

        .Columns(COL_MESSAGE).EntireColumn.AutoFit

        With .Columns(COL_REBAR_RATIO)

            .NumberFormatLocal = "0.00%"
            .FormatConditions.AddColorScale ColorScaleType:=2
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = 16776444
            .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueHighestValue
            .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = 8109667

        End With

    End With

    ' Columns("M:M").Select
    ' Selection.Style = "Percent"
    ' Selection.NumberFormatLocal = "0.0%"
    ' Selection.NumberFormatLocal = "0.00%"
    ' Selection.FormatConditions.AddColorScale ColorScaleType:=2
    ' Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    ' Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
    '     xlConditionValueLowestValue
    ' With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
    '     .Color = 16776444
    '     .TintAndShade = 0
    ' End With
    ' Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
    '     xlConditionValueHighestValue
    ' With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
    '     .Color = 8109667
    '     .TintAndShade = 0
    ' End With

    Call PrintError

    Call FontSetting

End Function


Function PrintError(Optional ByVal errNumber, Optional ByVal errSource, Optional ByVal errDetails)
'
' 列印錯誤.
'
' @since 1.0.0
' @param {Number} [errNumber] Err.COL_NUMBER.
' @param {String} [errSource] Err.Source.
' @param {String} [errDetails] Err.Description.
'
    Dim arrErrorMessage

    If Not IsError(errNumber) Then
        OBJ_ERR_MSG.Add "Error # " & Str(errNumber) & " was generated by " & errSource & vbCrLf & errDetails
    End If

    ubErrorMessage = OBJ_ERR_MSG.Count

    ReDim arrErrorMessage(0 To ubErrorMessage, 1 To 2)

    arrErrorMessage(0, 1) = "Number"
    arrErrorMessage(0, 2) = "Error Message"

    For i = 1 To ubErrorMessage
        arrErrorMessage(i, 1) = i
        arrErrorMessage(i, 2) = OBJ_ERR_MSG(i)
    Next i

    With Worksheets("Error")

        ' 清空資料保留格式
        .Cells.ClearContents

        .Range(.Cells(1, 1), .Cells(ubErrorMessage + 1, 2)) = arrErrorMessage

        If Not ubErrorMessage = 0 Then
            .Activate
        End If

    End With

End Function


' Function PrintRebarRatio()

'     rowStart = 1
'     rowUsed = UBound(ARR_RATIO)
'     columnUsed = 13

'     WS_OUTPUT.Range(WS_OUTPUT.Cells(rowStart, columnUsed), WS_OUTPUT.Cells(rowUsed, columnUsed)) = Application.Index(ARR_RATIO, 0, COL_REBAR)
'     WS_OUTPUT.Cells(1, COL_MESSAGE + 1) = "鋼筋比"

'     Call FontSetting

' End Function


' Function PrintRebarRatioInAnotherSheets()

'     Dim columnRatio As Worksheet
'     Dim rebarRatio As Worksheet
'     Set columnRatio = Worksheets("柱鋼筋比")
'     Set rebarRatio = Worksheets("鋼筋號數比")

'     rowStart = 1
'     rowUsed = UBound(ARR_RATIO)
'     columnStart = 1
'     columnUsed = 5

'     columnRatio.Range(columnRatio.Cells(rowStart, columnUsed), columnRatio.Cells(rowUsed, columnUsed)) = Application.Index(ARR_RATIO, 0, COL_REBAR)

'     ' 由於修改 ARR_RATIO 樓層部分，改以數字呈現，所以用 ARR_REBAR 再覆蓋一次。
'     columnRatio.Range(columnRatio.Cells(rowStart, columnStart), columnRatio.Cells(rowUsed, columnUsed - 1)) = ARR_REBAR

'     Call FontSetting

'     rowStart = 3
'     rowUsed = UBound(REBAR_NUMBER) + 1
'     columnStart = 2
'     columnUsed = 3

'     rebarRatio.Range(rebarRatio.Cells(rowStart, columnStart), rebarRatio.Cells(rowUsed, columnUsed)) = REBAR_NUMBER

' End Function


Function FontSetting()

    With WS_OUTPUT

        .Cells.Font.Name = "微軟正黑體"
        .Cells.Font.Name = "Calibri"
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter

    End With

    With Worksheets("Error")

        .Cells.Font.Name = "微軟正黑體"
        .Cells.Font.Name = "Calibri"

    End With

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


    For i = LB_REBAR To UB_REBAR

        If ARR_RATIO(i, COL_STORY) < NUM_FIRST_STORY Then

            If ARR_REBAR(i, COL_TIE_Y) < Int((ARR_REBAR(i, COL_REBAR_X) - 1) / 2) - 1 Then
                Call WarningMessage("【0407】Y 向繫筋未符合隔根勾", i)
            End If

            If ARR_REBAR(i, COL_TIE_X) < Int((ARR_REBAR(i, COL_REBAR_Y) - 1) / 2) - 1 Then
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

    For i = LB_REBAR To UB_REBAR

        ' 3 case
        ' 判斷位置
        isUpperLimit = ARR_REBAR(i, COL_NUMBER) <> ARR_REBAR(i - 1, COL_NUMBER) And ARR_REBAR(i, COL_NUMBER) = ARR_REBAR(i + 1, COL_NUMBER)
        isMiddle = ARR_REBAR(i, COL_NUMBER) = ARR_REBAR(i - 1, COL_NUMBER) And ARR_REBAR(i, COL_NUMBER) = ARR_REBAR(i + 1, COL_NUMBER)
        isLowerLimit = ARR_REBAR(i, COL_NUMBER) = ARR_REBAR(i - 1, COL_NUMBER) And ARR_REBAR(i, COL_NUMBER) <> ARR_REBAR(i + 1, COL_NUMBER)

        ' 往下減縮超過 7 成
        sharpDown = ARR_RATIO(i + 1, COL_REBAR) < ARR_RATIO(i, COL_REBAR) * 0.7

        ' 往上減縮超過 6 成
        sharpUp = ARR_RATIO(i - 1, COL_REBAR) < ARR_RATIO(i, COL_REBAR) * 0.6

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

    For i = LB_REBAR To UB_REBAR

        If ARR_RATIO(i, COL_STORY) < NUM_FIRST_STORY Then

            fcColumn = OBJ_INFO.Item(ARR_REBAR(i, COL_STORY))(COL_FC_COLUMN)
            ' fcColumn = Application.VLookup(ARR_REBAR(i, COL_STORY), GENERAL_INFORMATION, COL_FC_COLUMN, False)
            fytColumn = OBJ_INFO.Item(ARR_REBAR(i, COL_STORY))(COL_FYT)
            ' fytColumn = Application.VLookup(ARR_REBAR(i, COL_STORY), GENERAL_INFORMATION, COL_FYT, False)

            cover_ = OBJ_INFO.Item(ARR_REBAR(i, COL_STORY))(COL_COVER)

            stirrup = Split(ARR_REBAR(i, COL_BOUND_AREA), "@")
            rebarSize = stirrup(0)
            spacing_ = stirrup(1)

            bcX = ARR_REBAR(i, COL_WIDTH_X) - cover_ * 2 - OBJ_REBAR_SIZE.Item(rebarSize)(COL_DB)
            ' bcX = ARR_REBAR(i, COL_WIDTH_X) - 4 * 2 - Application.VLookup(rebarSize, REBAR_SIZE, COL_DB, False)
            bcY = ARR_REBAR(i, COL_WIDTH_Y) - cover_ * 2 - OBJ_REBAR_SIZE.Item(rebarSize)(COL_DB)
            ' bcY = ARR_REBAR(i, COL_WIDTH_Y) - 4 * 2 - Application.VLookup(rebarSize, REBAR_SIZE, COL_DB, False)

            ashX = ARR_RATIO(i, COL_TIE_X)
            ashY = ARR_RATIO(i, COL_TIE_Y)

            ag = ARR_REBAR(i, COL_WIDTH_X) * ARR_REBAR(i, COL_WIDTH_Y)
            ach = (ARR_REBAR(i, COL_WIDTH_X) - cover_ * 2) * (ARR_REBAR(i, COL_WIDTH_Y) - cover_ * 2)

            code15_3_X = 0.3 * spacing_ * bcX * fcColumn / fytColumn * (ag / ach - 1)
            code15_3_Y = 0.3 * spacing_ * bcY * fcColumn / fytColumn * (ag / ach - 1)
            code15_4_X = 0.09 * spacing_ * bcX * fcColumn / fytColumn
            code15_4_Y = 0.09 * spacing_ * bcY * fcColumn / fytColumn

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
' 頂樓區鋼筋比不大於 1.2 %
' NUM_TOP_STORY 為 RF 不含屋突
'

    ' 頂樓區 1/4
    checkStoryNumber = NUM_TOP_STORY - Fix((NUM_TOP_STORY - NUM_FIRST_STORY + 1) / 4)

    For i = LB_REBAR To UB_REBAR
        If ARR_RATIO(i, COL_STORY) >= checkStoryNumber And ARR_RATIO(i, COL_STORY) <= NUM_TOP_STORY And ARR_RATIO(i, COL_REBAR) > 0.01 * 1.2 Then
                Call WarningMessage("【0405】請確認高樓區鋼筋比，是否超過 1.2 %", i)
        End If
    Next

End Function


Function CountRebarNumber()

    rowStart = 2
    rowEnd = UBound(REBAR_SIZE)
    ReDim REBAR_NUMBER(rowStart To rowEnd, 1 To 2)

    For i = LB_REBAR To UB_REBAR

        rebarNumber = Split(ARR_REBAR(i, COL_REBAR), "-")(1)
        boundStirrupNumber = Split(ARR_REBAR(i, COL_BOUND_AREA), "@")(0)
        nonBoundStirrupNumber = Split(ARR_REBAR(i, COL_NON_BOUND_AREA), "@")(0)

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
