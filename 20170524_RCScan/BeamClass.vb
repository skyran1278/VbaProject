Private ran As UTILS_CLASS
Private APP
Private collectErrorMessage As Collection

Private BEAM_TYPE

Private objInfo
Private TOP_STORY
Private FIRST_STORY

Private objRebarSize

Private WS_OUTPUT As Worksheet
Private DATA_ROW_START
Private DATA_ROW_END
Private RAW_DATA

Private RATIO_DATA

Private REBAR_NUMBER

' Private GENERAL_INFORMATION
' Private REBAR_SIZE
' Private MESSAGE

' RAW_DATA 資料命名
Private Const STORY = 1
Private Const NUMBER = 2
Private Const BW = 3
Private Const H = 4
' 由於第幾排無用，所以放置 D 有效深度
Private Const D = 5
Private Const REBAR_LEFT = 6
Private Const REBAR_MID = 7
Private Const REBAR_RIGHT = 8
Private Const SIDE_REBAR = 9
Private Const STIRRUP_LEFT = 10
Private Const STIRRUP_MID = 11
Private Const STIRRUP_RIGHT = 12
Private Const BEAM_LENGTH = 13
Private Const SUPPORT = 14
Private Const LOCATION = 15
' 輸出資料位置
Private Const COL_MESSAGE = 16

' GENERAL_INFORMATION 資料命名
Private Const FY = 2
Private Const FYT = 3
Private Const FC_BEAM = 4
Private Const FC_COLUMN = 5
Private Const SDL = 6
Private Const LL = 7
Private Const BAND = 8
Private Const SLAB = 9
Private Const COVER = 10
Private Const STORY_NUM = 11

Private Const DIAMETER = 7
Private Const CROSS_AREA = 10

' -------------------------------------------------------------------------
' -------------------------------------------------------------------------
' REBAR_SIZE 資料命名

Private Sub Class_Initialize()
' Called automatically when class is created

    Set ran = New UTILS_CLASS
    Set APP = Application.WorksheetFunction

    Set collectErrorMessage = New Collection

End Sub


Function Initialize(ByVal sheet)
'
' 由於 VBA Class_Initialize 不能傳變數，所以這裡再做一次 Initialize.
'
' @param {String} [sheet] descrip.
' @return {type} [name] descrip.
'

    BEAM_TYPE = sheet

    ' 輸出 objInfo
    Call GetGeneralInformation

    ' 輸出 objRebarSize
    Call GetRebarSize

    ' 輸出
    ' WS_OUTPUT
    ' DATA_ROW_START
    ' DATA_ROW_END
    ' RAW_DATA
    Call SortRawData(sheet)

    ' ReDim MESSAGE(DATA_ROW_START To DATA_ROW_END)

    ReDim RATIO_DATA(LBound(RAW_DATA, 1) To UBound(RAW_DATA, 1), LBound(RAW_DATA, 2) To UBound(RAW_DATA, 2))

    Call GetRatioData

End Function


Function GetGeneralInformation()
'
'

    Dim wsGeneralInformation As Worksheet
    Set wsGeneralInformation = Worksheets("General Information")

    ' 後面多空出一行，以增加代號
    arrGeneralInformation = ran.GetRangeToArray(wsGeneralInformation, 1, 4, 4, 14)

    lbGeneralInformation = LBound(arrGeneralInformation)
    ubGeneralInformation = UBound(arrGeneralInformation)

    j = 1

    For i = ubGeneralInformation To lbGeneralInformation Step -1
        arrGeneralInformation(i, STORY_NUM) = j
        j = j + 1
    Next i

    Set objInfo = ran.CreateDictionary(arrGeneralInformation, 1, False)

    ' Use Cells(13, 16).Text instead of .Value
    TOP_STORY = WarnDicEmpty(objInfo.Item(wsGeneralInformation.Cells(13, 16).Text), STORY_NUM, "搜尋不到頂樓樓層")
    FIRST_STORY = WarnDicEmpty(objInfo.Item(wsGeneralInformation.Cells(14, 16).Text), STORY_NUM, "搜尋不到地面層")

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

        collectErrorMessage.Add warning
        WarnDicEmpty = Empty

    End If

End Function


Private Function GetRebarSize()
'
'

    arrRebarSize = ran.GetRangeToArray(Worksheets("Rebar Size"), 1, 1, 5, 10)

    Set objRebarSize = ran.CreateDictionary(arrRebarSize, 1, False)

End Function


Private Function SortRawData(ByVal sheet)
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'


    ' WS_OUTPUT.Name = sheet & "配筋 - Output"
    ' Set WS_OUTPUT = Worksheets(sheet & "配筋 - Output")

    ' 清空前一次輸入
    ' WS_OUTPUT.Cells.Clear

    ' 多抓兩行用來排序
    arrRawData = ran.GetRangeToArray(Worksheets(sheet & "配筋 - Input"), 1, 1, 5, 18)

    DATA_ROW_START = 3
    DATA_ROW_END = UBound(arrRawData)

    rowStartRawData = LBound(arrRawData, 1)
    colStartRawData = LBound(arrRawData, 2)
    colEndRawData = UBound(arrRawData, 2)

    colStoryNum = 17
    colNumberNoC = 18

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        ' 樓層數字化，用以比較上下樓層。
        arrRawData(i, colStoryNum) = WarnDicEmpty(objInfo.Item(arrRawData(i, STORY)), STORY_NUM, "請確認 " & arrRawData(i, STORY) & " 是否存在於 General Information")

        ' 去掉 大寫與小寫開頭的 C，用以排序
        If LCase(Left(arrRawData(i, NUMBER), 1)) <> "c" Then

            arrRawData(i, colNumberNoC) = arrRawData(i, NUMBER)

        Else

            arrRawData(i, colNumberNoC) = Right(arrRawData(i, NUMBER), Len(arrRawData(i, NUMBER)) - 1)

        End If

        ' 填滿以用於排序
        arrRawData(i + 1, colStoryNum) = arrRawData(i, colStoryNum)
        arrRawData(i + 2, colStoryNum) = arrRawData(i, colStoryNum)
        arrRawData(i + 3, colStoryNum) = arrRawData(i, colStoryNum)
        arrRawData(i + 1, colNumberNoC) = arrRawData(i, colNumberNoC)
        arrRawData(i + 2, colNumberNoC) = arrRawData(i, colNumberNoC)
        arrRawData(i + 3, colNumberNoC) = arrRawData(i, colNumberNoC)

    Next


    ' 排序由低到高
    ' Call ran.QuickSortArray(arrRawData, 3, , colNumberNoC)
    ' With WS_OUTPUT
    '     .Range(.Cells(rowStartRawData, colStartRawData), .Cells(DATA_ROW_END, colEndRawData)) = arrRawData
    ' End With
    ' Call ran.QuickSortArray(arrRawData, 3, , colStoryNum)

    Set WS_OUTPUT = ThisWorkbook.Sheets.Add(After:=Worksheets("General Information"))


    ' TODO:由底往上排
    With WS_OUTPUT

        .Range(.Cells(rowStartRawData, colStartRawData), .Cells(DATA_ROW_END, colEndRawData)) = arrRawData

        .Range(.Cells(DATA_ROW_START, colStartRawData), .Cells(DATA_ROW_END, colEndRawData)).Sort _
            Key1:=.Range(.Cells(DATA_ROW_START, colStoryNum), .Cells(DATA_ROW_END, colStoryNum)), Order1:=xlAscending, DataOption1:=xlSortNormal, _
            Key2:=.Range(.Cells(DATA_ROW_START, colNumberNoC), .Cells(DATA_ROW_END, colNumberNoC)), Order2:=xlAscending, DataOption2:=xlSortNormal, _
            Header:=xlNo, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin



        ' .Sort.SortFields.Clear
        ' .Sort.SortFields.Add Key:=.Range(.Cells(DATA_ROW_START, colStoryNum), .Cells(DATA_ROW_END, colStoryNum)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' .Sort.SortFields.Add Key:=.Range(.Cells(DATA_ROW_START, colNumberNoC), .Cells(DATA_ROW_END, colNumberNoC)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' .Sort.SetRange .Range(.Cells(DATA_ROW_START, colStartRawData), .Cells(DATA_ROW_END, colEndRawData))

        ' With .Sort
        '     .MatchCase = False
        '     .Orientation = xlTopToBottom
        '     .SortMethod = xlPinYin
        '     .Apply
        ' End With

        ' 收入資料
        RAW_DATA = .Range(.Cells(rowStartRawData, colStartRawData), .Cells(DATA_ROW_END, colEndRawData - 2))

    End With

    ' ' 清空前一次輸入
    WS_OUTPUT.Cells.Clear

End Function


Function GetRatioData()

    ' 樓層數字化，用以比較上下樓層。
    For i = DATA_ROW_START To DATA_ROW_END Step 4

        RATIO_DATA(i, STORY) = WarnDicEmpty(objInfo.Item(RAW_DATA(i, STORY)), STORY_NUM, "請確認 " & RAW_DATA(i, STORY) & " 是否存在於 General Information")
        ' RATIO_DATA(i, STORY) = Application.Match(RAW_DATA(i, STORY), APP.Index(GENERAL_INFORMATION, 0, STORY), 0)
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

        rebar = Split(RAW_DATA(i, REBAR_LEFT), "-")
        stirrup = Split(RAW_DATA(i, STIRRUP_LEFT), "@")
        fyDb = objRebarSize.Item(rebar(1))(DIAMETER)
        ' fyDb = APP.VLookup(rebar(1), REBAR_SIZE, DIAMETER, False)
        fytDb = objRebarSize.Item(SplitStirrup(stirrup(0)))(DIAMETER)
        ' fytDb = APP.VLookup(SplitStirrup(SplitStirrup(stirrup(0))), REBAR_SIZE, DIAMETER, False)

        ' 雙排筋
        RATIO_DATA(i, D) = RAW_DATA(i, H) - (objInfo.Item(RAW_DATA(i, STORY))(COVER) + fytDb + fyDb * 1.5)

    Next

End Function


Function CalRebarArea(rebar)

    tmp = Split(rebar, "-")

    If tmp(0) <> 0 Then

        ' 轉換鋼筋尺寸為截面積
        tmp(1) = objRebarSize.Item(tmp(1))(CROSS_AREA)
        ' tmp(1) = APP.VLookup(tmp(1), REBAR_SIZE, CROSS_AREA, False)

        CalRebarArea = tmp(0) * tmp(1)
    Else
        CalRebarArea = 0
    End If

End Function


Function CalStirrupArea(rebar)
'
' 考量雙箍
'
    tmp = Split(rebar, "@")

    bars = Split(tmp(0), "#")

    ' 箍筋號數
    bars(1) = "#" & bars(1)

    ' 轉換鋼筋尺寸為截面積
    If bars(0) = "" Then
        CalStirrupArea = 2 * objRebarSize.Item(bars(1))(CROSS_AREA)
        ' CalStirrupArea = 2 * APP.VLookup(bars(1), REBAR_SIZE, CROSS_AREA, False)
    Else
        CalStirrupArea = 2 * bars(0) * objRebarSize.Item(bars(1))(CROSS_AREA)
        ' CalStirrupArea = 2 * bars(0) * APP.VLookup(bars(1), REBAR_SIZE, CROSS_AREA, False)
    End If

End Function


Function CalSideRebarArea(rebar)

    If rebar <> "-" Then

        ' 去掉 EF
        ' 1#4EF => 1#4
        sidebarNoEF = Left(rebar, Len(rebar) - 2)

        tmp = Split(sidebarNoEF, "#")

        ' 轉換鋼筋尺寸為截面積
        tmp(1) = objRebarSize.Item("#" & tmp(1))(CROSS_AREA)
        ' tmp(1) = APP.VLookup("#" & tmp(1), REBAR_SIZE, CROSS_AREA, False)

        ' 對稱雙排
        CalSideRebarArea = 2 * tmp(1)

    Else
        CalSideRebarArea = 0
    End If

End Function


Function SplitStirrup(rebar)
'
' 處理雙箍的情況
'
    bars = Split(rebar, "#")

    SplitStirrup = "#" & bars(1)

End Function


Function GetTypeMessage(Girder, Beam, GroundBeam)

    If BEAM_TYPE = "大梁" Then
        GetTypeMessage = Girder
    ElseIf BEAM_TYPE = "小梁" Then
        GetTypeMessage = Beam
    ElseIf BEAM_TYPE = "地梁" Then
        GetTypeMessage = GroundBeam
    End If

End Function

Function WarningMessage(warningMessageCode, i)

    RAW_DATA(i, COL_MESSAGE) = warningMessageCode & vbCrLf & RAW_DATA(i, COL_MESSAGE)

End Function

Function PrintResult()

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        With WS_OUTPUT

            For j = STORY To H
                .Range(.Cells(i, j), .Cells(i + 3, j)).Merge
            Next j

            For j = SIDE_REBAR To COL_MESSAGE
                .Range(.Cells(i, j), .Cells(i + 3, j)).Merge
            Next j

            If RAW_DATA(i, COL_MESSAGE) = "" Then
                RAW_DATA(i, COL_MESSAGE) = "(S), (E), (i) - SCAN 結果 ok"
            Else
                .Cells(i, COL_MESSAGE).Style = "壞"
                RAW_DATA(i, COL_MESSAGE) = Left(RAW_DATA(i, COL_MESSAGE), Len(RAW_DATA(i, COL_MESSAGE)) - 1)
            End If

        End With

    Next

    With WS_OUTPUT
        .Range(.Cells(LBound(RAW_DATA, 1), LBound(RAW_DATA, 2)), .Cells(UBound(RAW_DATA, 1), UBound(RAW_DATA, 2))) = RAW_DATA

        .Columns(COL_MESSAGE).EntireColumn.AutoFit

    End With

    Call PrintError

    Call FontSetting

End Function


Private Function PrintError()
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'
    Dim arrErrorMessage

    ubErrorMessage = collectErrorMessage.Count

    ReDim arrErrorMessage(0 To ubErrorMessage, 1 To 2)

    arrErrorMessage(0, 1) = "Number"
    arrErrorMessage(0, 2) = "Error Message"

    For i = 1 To ubErrorMessage
        arrErrorMessage(i, 1) = i
        arrErrorMessage(i, 2) = collectErrorMessage(i)
    Next i

    With Worksheets("Error")

        .Range(.Cells(1, 1), .Cells(ubErrorMessage + 1, 2)) = arrErrorMessage

        If Not ubErrorMessage = 0 Then
            .Activate
        End If

    End With

End Function


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


' Function PrintRebarRatio()

'     Dim rebarRatio As Worksheet
'     Set rebarRatio = Worksheets("鋼筋號數比")

'     rowStart = 3
'     rowUsed = UBound(REBAR_NUMBER) + 1

'     If BEAM_TYPE = "大梁" Then
'         columnStart = 4
'     ElseIf BEAM_TYPE = "小梁" Then
'         columnStart = 7
'     ElseIf BEAM_TYPE = "地梁" Then
'         columnStart = 10
'     End If

'     columnUsed = columnStart + 2

'     rebarRatio.Range(rebarRatio.Cells(rowStart, columnStart), rebarRatio.Cells(rowUsed, columnUsed)) = REBAR_NUMBER

' End Function



' Function CountRebarNumber()

'     rowStart = 2
'     rowEnd = UBound(REBAR_SIZE)
'     ReDim REBAR_NUMBER(rowStart To rowEnd, 1 To 3)

'     ' 主筋
'     For i = DATA_ROW_START To DATA_ROW_END

'         For j = REBAR_LEFT To REBAR_RIGHT

'             rebarNumber = Split(RAW_DATA(i, j), "-")

'             If rebarNumber(0) > 0 Then
'                 rebarNumber = rebarNumber(1)
'             Else
'                 rebarNumber = ""
'             End If

'             For k = rowStart To rowEnd

'                 If rebarNumber = REBAR_SIZE(k, 1) Then
'                     REBAR_NUMBER(k, 1) = REBAR_NUMBER(k, 1) + 1
'                 End If

'             Next

'         Next

'     Next

'     ' 腰筋
'     For i = DATA_ROW_START To DATA_ROW_END Step 4

'         If RAW_DATA(i, SIDE_REBAR) <> "-" Then

'             sideRebar = Left(RAW_DATA(i, SIDE_REBAR), Len(RAW_DATA(i, SIDE_REBAR)) - 2)

'             rebarNumber = Split(sideRebar, "#")

'             rebarNumber = "#" & rebarNumber(1)

'             For j = rowStart To rowEnd

'                 If rebarNumber = REBAR_SIZE(j, 1) Then
'                     REBAR_NUMBER(j, 2) = REBAR_NUMBER(j, 2) + 1
'                 End If

'             Next

'         End If

'     Next

'     ' 箍筋
'     For i = DATA_ROW_START To DATA_ROW_END Step 4

'         For j = STIRRUP_LEFT To STIRRUP_RIGHT

'             rebarNumber = Split(RAW_DATA(i, j), "@")(0)
'             rebarNumber = Split(rebarNumber, "#")
'             rebarNumber = "#" & rebarNumber(1)

'             For k = rowStart To rowEnd

'                 If rebarNumber = REBAR_SIZE(k, 1) Then
'                     REBAR_NUMBER(k, 3) = REBAR_NUMBER(k, 3) + 1
'                 End If

'             Next

'         Next

'     Next

'     ' 轉換成比例
'     Dim sum(1 To 3)
'     For i = rowStart To rowEnd
'         For j = 1 To 3
'             sum(j) = sum(j) + REBAR_NUMBER(i, j)
'         Next
'     Next
'     For j = 1 To 3
'         For i = rowStart To rowEnd
'             If REBAR_NUMBER(i, j) <> 0 Then
'                 REBAR_NUMBER(i, j) = REBAR_NUMBER(i, j) / sum(j)
'             End If
'         Next
'     Next

' End Function



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
'

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

                rebar = Split(RAW_DATA(k, j), "-")

                stirrup = Split(RAW_DATA(i, j + 4), "@")

                If rebar(0) > 1 Then

                    fyDb = objRebarSize.Item(rebar(1))(DIAMETER)
                    ' fyDb = APP.VLookup(rebar(1), REBAR_SIZE, DIAMETER, False)
                    fytDb = objRebarSize.Item(SplitStirrup(stirrup(0)))(DIAMETER)
                    ' fytDb = APP.VLookup(SplitStirrup(SplitStirrup(stirrup(0))), REBAR_SIZE, DIAMETER, False)

                    Spacing = (RAW_DATA(i, BW) - objInfo.Item(RAW_DATA(i, STORY))(COVER) * 2 - fytDb * 2 - rebar(0) * fyDb) / (rebar(0) - 1)

                    If Spacing > 25 Then
                        Call WarningMessage("【0106】請確認鋼筋間距下限，是否符合鋼筋間距 25 cm 以下規定", i)
                    End If

                ElseIf rebar(0) = "1" Then

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

            stirrup = Split(RAW_DATA(i, j), "@")

            isAvSmallerThanCode = RATIO_DATA(i, j) < 0.0025 * RAW_DATA(i, BW) * stirrup(1)

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
' 0.6 * 1/8 * wu * L^2 <= As * fy * d

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        maxRatio = APP.Max(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_MID), RATIO_DATA(i, REBAR_RIGHT), RATIO_DATA(i + 2, REBAR_LEFT), RATIO_DATA(i + 2, REBAR_MID), RATIO_DATA(i + 2, REBAR_RIGHT))

        ' 轉換 kgw-m => tf-m: * 100000
        mn = 1 / 8 * (1.2 * (0.15 * 2.4 + objInfo.Item(RAW_DATA(i, STORY))(SDL) * objInfo.Item(RAW_DATA(i, STORY))(BAND)) + 1.6 * objInfo.Item(RAW_DATA(i, STORY))(LL) *  * objInfo.Item(RAW_DATA(i, STORY))(BAND)) * RAW_DATA(i, BEAM_LENGTH) ^ 2 * 100000
        ' mn = 1 / 8 * (1.2 * (0.15 * 2.4 + APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, SDL, False)) + 1.6 * APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, LL, False)) * BAND ^ 2 * 100000

        capacity = maxRatio * objInfo.Item(RAW_DATA(i, STORY))(FY) * RATIO_DATA(i, D)
        ' capacity = maxRatio * APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FY, False) * RATIO_DATA(i, D)

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

    code3_3 = 0.8 * Sqr(objInfo.Item(RAW_DATA(i, STORY))(FC_BEAM)) / objInfo.Item(RAW_DATA(i, STORY))(FY) * RAW_DATA(i, BW) * RATIO_DATA(i, D)
    ' code3_3 = 0.8 * Sqr(APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FC_BEAM, False)) / APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FY, False) * RAW_DATA(i, BW) * RATIO_DATA(i, D)
    code3_4 = 14 / objInfo.Item(RAW_DATA(i, STORY))(FY) * RAW_DATA(i, BW) * RATIO_DATA(i, D)
    ' code3_4 = 14 / APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FY, False) * RAW_DATA(i, BW) * RATIO_DATA(i, D)

    If RATIO_DATA(i, REBAR_LEFT) < code3_3 Or RATIO_DATA(i, REBAR_LEFT) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0201】請確認左端上層筋下限，是否符合規範 3.6 規定", "【0301】請確認左端上層筋下限，是否符合規範 3.6 規定", "請確認左端上層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If RATIO_DATA(i, REBAR_MID) < code3_3 Or RATIO_DATA(i, REBAR_MID) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0202】請確認中央上層筋下限，是否符合規範 3.6 規定", "【0302】請確認中央上層筋下限，是否符合規範 3.6 規定", "請確認中央上層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If RATIO_DATA(i, REBAR_RIGHT) < code3_3 Or RATIO_DATA(i, REBAR_RIGHT) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0203】請確認右端上層筋下限，是否符合規範 3.6 規定", "【0303】請確認右端上層筋下限，是否符合規範 3.6 規定", "請確認右端上層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If RATIO_DATA(i + 2, REBAR_LEFT) < code3_3 Or RATIO_DATA(i + 2, REBAR_LEFT) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0204】請確認左端下層筋下限，是否符合規範 3.6 規定", "【0304】請確認左端下層筋下限，是否符合規範 3.6 規定", "請確認左端下層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If RATIO_DATA(i + 2, REBAR_MID) < code3_3 Or RATIO_DATA(i + 2, REBAR_MID) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0205】請確認中央下層筋下限，是否符合規範 3.6 規定", "【0305】請確認中央下層筋下限，是否符合規範 3.6 規定", "請確認中央下層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If RATIO_DATA(i + 2, REBAR_RIGHT) < code3_3 Or RATIO_DATA(i + 2, REBAR_RIGHT) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0206】請確認右端下層筋下限，是否符合規範 3.6 規定", "【0306】請確認右端下層筋下限，是否符合規範 3.6 規定", "請確認右端下層筋下限，是否符合規範 3.6 規定"), i)
    End If

Next

End Function

Function Norm15_4_2_1()
'
' 耐震規範 (1F以下大梁不適用)：
' 拉力鋼筋比不得大於 (fc' + 100) / (4 * fy)，亦不得大於 0.025。

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        If RATIO_DATA(i, STORY) < FIRST_STORY Then

            code15_4_2_1 = APP.Min((objInfo.Item(RAW_DATA(i, STORY))(FC_BEAM) + 100) / (4 * objInfo.Item(RAW_DATA(i, STORY))(FY)) * RAW_DATA(i, BW) * RATIO_DATA(i, D), 0.025 * RAW_DATA(i, BW) * RATIO_DATA(i, D))
            ' code15_4_2_1 = APP.Min((APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FC_BEAM, False) + 100) / (4 * APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FY, False)) * RAW_DATA(i, BW) * RATIO_DATA(i, D), 0.025 * RAW_DATA(i, BW) * RATIO_DATA(i, D))

            If RATIO_DATA(i, REBAR_LEFT) > code15_4_2_1 Then
                Call WarningMessage("【0212】請確認左端上層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If RATIO_DATA(i, REBAR_MID) > code15_4_2_1 Then
                Call WarningMessage("【0213】請確認中央上層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If RATIO_DATA(i, REBAR_RIGHT) > code15_4_2_1 Then
                Call WarningMessage("【0214】請確認右端上層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If RATIO_DATA(i + 2, REBAR_LEFT) > code15_4_2_1 Then
                Call WarningMessage("【0215】請確認左端下層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If RATIO_DATA(i + 2, REBAR_MID) > code15_4_2_1 Then
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

            maxRatio = APP.Max(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_MID), RATIO_DATA(i, REBAR_RIGHT), RATIO_DATA(i + 2, REBAR_LEFT), RATIO_DATA(i + 2, REBAR_MID), RATIO_DATA(i + 2, REBAR_RIGHT))
            minRatio = APP.Min(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_MID), RATIO_DATA(i, REBAR_RIGHT), RATIO_DATA(i + 2, REBAR_LEFT), RATIO_DATA(i + 2, REBAR_MID), RATIO_DATA(i + 2, REBAR_RIGHT))
            code15_4_2_2 = minRatio < maxRatio / 4

            If code15_4_2_2 Then
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

        If RATIO_DATA(i, REBAR_MID) * 0.7 < RATIO_DATA(i, REBAR_LEFT) And rebarLEFT(0) > 3 And RATIO_DATA(i, BEAM_LENGTH) > 400 Then
            Call WarningMessage("【0111】請確認左端上層筋相對鋼筋量，是否符合端部上層鋼筋量需小於中央鋼筋量的 70% 規定", i)
        End If

        If RATIO_DATA(i, REBAR_MID) * 0.7 < RATIO_DATA(i, REBAR_RIGHT) And rebarRIGHT(0) > 3 And RATIO_DATA(i, BEAM_LENGTH) > 400 Then
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

        minRatio = APP.Min(RATIO_DATA(i, REBAR_LEFT), RATIO_DATA(i, REBAR_RIGHT))

        rebar = Split(RAW_DATA(i, REBAR_MID), "-")

        If RATIO_DATA(i, REBAR_MID) > minRatio * 0.7 And rebar(0) > 3 And RATIO_DATA(i, BEAM_LENGTH) > 400 Then
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

        minRatio = APP.Min(RATIO_DATA(i + 2, REBAR_LEFT), RATIO_DATA(i + 2, REBAR_RIGHT))

        rebar = Split(RAW_DATA(i + 2, REBAR_MID), "-")

        If RATIO_DATA(i + 2, REBAR_MID) > minRatio * 0.7 And rebar(0) > 3 And RATIO_DATA(i, BEAM_LENGTH) > 400 Then
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

            rebar = Split(RAW_DATA(k, j), "-")

            stirrup = Split(RAW_DATA(i, j + 4), "@")

            ' 等於 0 直接沒做事
            If rebar(0) > 1 Then

                fyDb = objRebarSize.Item(rebar(1))(DIAMETER)
                ' fyDb = APP.VLookup(rebar(1), REBAR_SIZE, DIAMETER, False)
                fytDb = objRebarSize.Item(SplitStirrup(stirrup(0)))(DIAMETER)
                ' fytDb = APP.VLookup(SplitStirrup(stirrup(0)), REBAR_SIZE, DIAMETER, False)

                ' 第一種方法
                ' Max = Fix((RAW_DATA(i, BW) - 4 * 2 - fytDb * 2 - fyDb) / (2 * fyDb)) + 1
                ' CInt(rebar(0)) > Max
                ' 第二種方法
                ' spacing = (RAW_DATA(i, BW) - 4 * 2 - fytDb * 2 - fyDb) / (CInt(rebar(0)) - 1) - fyDb
                ' 可以不需要型別轉換
                ' Spacing = (RAW_DATA(i, BW) - 4 * 2 - fytDb * 2 - CInt(rebar(0)) * fyDb) / (CInt(rebar(0)) - 1)
                Spacing = (RAW_DATA(i, BW) - 4 * 2 - fytDb * 2 - rebar(0) * fyDb) / (rebar(0) - 1)

                ' Norm13_5_1
                ' 淨距不少於1Db
                If Spacing < fyDb Or Spacing < 2.5 Then
                    Call WarningMessage(GetTypeMessage("【0210】請確認單排支數上限，是否符合淨距不少於 1 Db 規定", "【0308】請確認單排支數上限，是否符合淨距不少於 1 Db 規定", "請確認單排支數上限，是否符合淨距不少於 1 Db 規定"), i)
                End If

            ElseIf rebar(0) = "1" Then

                ' 排除掉1支的狀況，避免除以0
                ' 不少於2支
                Call WarningMessage(GetTypeMessage("【0211】請確認是否符合 單排支數下限 規定", "【0309】請確認是否符合 單排支數下限 規定", "請確認是否符合 單排支數下限 規定"), i)

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
                Call WarningMessage(GetTypeMessage("【0219】請確認箍筋間距下限，是否符合 10cm 以上規定", "請確認箍筋間距下限，是否符合 10cm 以上規定", "【0113】請確認箍筋間距下限，是否符合 10cm 以上規定"), i)
            ElseIf stirrup(1) > 30 Then
                Call WarningMessage(GetTypeMessage("【0220】請確認箍筋間距上限，是否符合 30cm 以下規定", "請確認箍筋間距上限，是否符合 30cm 以下規定", "【0114】請確認箍筋間距上限，是否符合 30cm 以下規定"), i)
            End If

        Next

    Next

End Function

Function Norm4_6_6_3()
'
' 剪力鋼筋量大於 3.52/fy

    For i = DATA_ROW_START To DATA_ROW_END Step 4

        For j = STIRRUP_LEFT To STIRRUP_RIGHT

            stirrup = Split(RAW_DATA(i, j), "@")

            avMin = APP.Max(0.2 * Sqr(objInfo.Item(RAW_DATA(i, STORY))(FC_BEAM)) * RAW_DATA(i, BW) * stirrup(1) / objInfo.Item(RAW_DATA(i, STORY))(FYT), 3.5 * RAW_DATA(i, BW) * stirrup(1) / objInfo.Item(RAW_DATA(i, STORY))(FYT))
            ' avMin = APP.Max(0.2 * Sqr(APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FC_BEAM, False)) * RAW_DATA(i, BW) * stirrup(1) / APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FYT, False), 3.5 * RAW_DATA(i, BW) * stirrup(1) / APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FYT, False))
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
            ' rebar = Split(RAW_DATA(i, j - 4), "-")

            ' fyDb = objRebarSize.Item(rebar(1))(DIAMETER)
            ' fyDb = APP.VLookup(rebar(1), REBAR_SIZE, DIAMETER, False)
            ' fytDb = objRebarSize.Item(SplitStirrup(stirrup(0)))(DIAMETER)
            ' fytDb = APP.VLookup(SplitStirrup(stirrup(0)), REBAR_SIZE, DIAMETER, False)
            ' effectiveDepth = RAW_DATA(i, H) - (4 + fytDb + fyDb / 2)
            av = RATIO_DATA(i, j)

            ' code4.4.1.1
            vc = 0.53 * Sqr(objInfo.Item(RAW_DATA(i, STORY))(FC_BEAM)) * RAW_DATA(i, BW) * RATIO_DATA(i, D)
            ' vc = 0.53 * Sqr(APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FC_BEAM, False)) * RAW_DATA(i, BW) * RATIO_DATA(i, D)

            ' code4.6.7.2
            vs = av * objInfo.Item(RAW_DATA(i, STORY))(FYT) * RATIO_DATA(i, D) / stirrup(1)
            ' vs = av * APP.VLookup(RAW_DATA(i, STORY), GENERAL_INFORMATION, FYT, False) * RATIO_DATA(i, D) / stirrup(1)

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
            Call WarningMessage(GetTypeMessage("【0207】請確認是否符合 規範 3.7.5", "【0307】請確認是否符合 規範 3.7.5", "請確認是否符合 規範 3.7.5"), i)
        End If

    Next

End Function
