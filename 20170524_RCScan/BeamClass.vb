' TODO: 輸出 expect != actual
Private ran As UTILS_CLASS
Private APP
Private OBJ_ERR_MSG As Collection

Private S_BEAM_TYPE

Private OBJ_INFO
Private NUM_TOP_STOREY
Private NUM_FIRST_STOREY

Private OBJ_REBAR_SIZE

Private WS_OUTPUT As Worksheet
Private LB_REBAR
Private UB_REBAR
Private ARR_REBAR

Private ARR_RATIO

' 準備拋棄
' Private REBAR_NUMBER

' ARR_REBAR 資料命名
Private Const COL_STOREY = 1
Private Const COL_NUMBER = 2
Private Const COL_BW = 3
Private Const COL_H = 4
' 由於第幾排無用，所以放置 COL_D 有效深度，用於 ARR_RATIO
Private Const COL_D = 5
Private Const COL_REBAR_LEFT = 6
Private Const COL_REBAR_MID = 7
Private Const COL_REBAR_RIGHT = 8
Private Const COL_SIDEBAR = 9
Private Const COL_STIRRUP_LEFT = 10
Private Const COL_STIRRUP_MID = 11
Private Const COL_STIRRUP_RIGHT = 12
Private Const COL_SPAN = 13
Private Const COL_SUPPORT = 14
Private Const COL_NOTE = 15
' 輸出資料位置
Private Const COL_LAP_LEFT = 17
Private Const COL_LAP_MID = 18
Private Const COL_LAP_RIGHT = 19
Private Const COL_MESSAGE = 20

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
Private Const COL_STIRRUP_10 = 11
Private Const COL_STOREY_NUM = 12

Private Const COL_DB = 7
Private Const COL_AREA = 10

' -------------------------------------------------------------------------
' -------------------------------------------------------------------------
' REBAR_SIZE 資料命名

Private Sub Class_Initialize()
' Called automatically when class is created

    Set ran = New UTILS_CLASS
    Set APP = Application.WorksheetFunction

    Set OBJ_ERR_MSG = New Collection

End Sub


Function Initialize(ByVal sheet)
'
' 由於 VBA Class_Initialize 不能傳變數，所以這裡再做一次 Initialize.
'
' @param {String} [sheet] 大梁、小梁、地梁.
'

    S_BEAM_TYPE = sheet

    ' 輸出 OBJ_INFO
    Call GetGeneralInformation

    ' 輸出 OBJ_REBAR_SIZE
    Call GetRebarSize

    ' 輸出
    ' WS_OUTPUT
    ' LB_REBAR
    ' UB_REBAR
    ' ARR_REBAR
    Call SortRawData(sheet)

    ' ReDim MESSAGE(LB_REBAR To UB_REBAR)

    ReDim ARR_RATIO(LBound(ARR_REBAR, 1) To UBound(ARR_REBAR, 1), LBound(ARR_REBAR, 2) To UBound(ARR_REBAR, 2))

    Call GetRatioData

End Function


Function GetGeneralInformation()
'
'

    Dim wsGeneralInformation As Worksheet
    Set wsGeneralInformation = Worksheets("General Information")

    ' 後面多空出一行，以增加代號
    arrGeneralInformation = ran.GetRangeToArray(wsGeneralInformation, 1, 4, 4, 15)

    lbGeneralInformation = LBound(arrGeneralInformation, 1)
    ubGeneralInformation = UBound(arrGeneralInformation, 1)
    lbColGeneralInformation = LBound(arrGeneralInformation, 2)
    ubColGeneralInformation = UBound(arrGeneralInformation, 2)

    j = 1

    For i = ubGeneralInformation To lbGeneralInformation Step -1
        arrGeneralInformation(i, COL_STOREY_NUM) = j
        j = j + 1
    Next i

    ' 掃描是否有沒輸入的數值
    For i = lbGeneralInformation To ubGeneralInformation
        For j = lbColGeneralInformation To ubColGeneralInformation

            If arrGeneralInformation(i, j) = "" Then
                OBJ_ERR_MSG.Add "General Information " & arrGeneralInformation(i, COL_STOREY) & " " & arrGeneralInformation(1, j) & " 是否空白？"
            End If

        Next j
    Next i

    Set OBJ_INFO = ran.CreateDictionary(arrGeneralInformation, 1, False)

    ' Use Cells(13, 16).Text instead of .Value
    NUM_TOP_STOREY = DicIsEmpty(OBJ_INFO.Item(wsGeneralInformation.Cells(1, 17).Text), COL_STOREY_NUM, "搜尋不到頂樓樓層")
    NUM_FIRST_STOREY = DicIsEmpty(OBJ_INFO.Item(wsGeneralInformation.Cells(2, 17).Text), COL_STOREY_NUM, "搜尋不到地面樓層")

End Function


Private Function DicIsEmpty(ByVal arr, ByVal value, Optional ByVal warning = "Dictionary is Empty")
'
' 如果 arr 為空，則 show error.
'
' @since 3.0.0
' @param {Array} [arr] 需要驗證的值.
' @param {Number} [value] 陣列位置.
' @param {String} [warning] 錯誤訊息.
' @return {Variant} [value] 空 或是 查詢到的值.
'

    If Not IsEmpty(arr) Then

        DicIsEmpty = arr(value)

    Else

        OBJ_ERR_MSG.Add warning
        DicIsEmpty = Empty

    End If

End Function


Private Function GetRebarSize()
'
'

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
    arrRawData = ran.GetRangeToArray(Worksheets(sheet), 1, 1, 5, COL_MESSAGE + 2)

    rowLbRawData = LBound(arrRawData, 1)
    colLbRawData = LBound(arrRawData, 2)
    rowUbRawData = UBound(arrRawData, 1)
    colUbRawData = UBound(arrRawData, 2)

    ' 確認有沒有少貼資料
    If (rowUbRawData - 2) Mod 4 Then
        OBJ_ERR_MSG.Add "列數與預期不同，請確認資料是否齊全。"
    End If

    LB_REBAR = 3
    UB_REBAR = rowUbRawData

    ' 利用最後兩行來做排序處理
    colStoreyNum = colUbRawData - 1
    colNumberNoC = colUbRawData

    ' clear output area
    For i = 1 To UB_REBAR Step 1
        For j = COL_NOTE + 1 To colUbRawData Step 1
            arrRawData(i, j) = ""
        Next j
    Next i

    ' title
    arrRawData(1, COL_LAP_LEFT) = "搭接長度"
    arrRawData(2, COL_LAP_LEFT) = "左端"
    arrRawData(2, COL_LAP_MID) = "中央"
    arrRawData(2, COL_LAP_RIGHT) = "右端"
    arrRawData(1, COL_MESSAGE) = "Warning"

    For i = LB_REBAR To UB_REBAR Step 4

        ' 樓層數字化，用以比較上下樓層。
        arrRawData(i, colStoreyNum) = DicIsEmpty(OBJ_INFO.Item(arrRawData(i, COL_STOREY)), COL_STOREY_NUM, "請確認 " & arrRawData(i, COL_STOREY) & " 是否存在於 General Information")

        ' 去掉 大寫與小寫開頭的 C，用以排序
        If LCase(Left(arrRawData(i, COL_NUMBER), 1)) <> "c" Then

            arrRawData(i, colNumberNoC) = arrRawData(i, COL_NUMBER)

        Else

            arrRawData(i, colNumberNoC) = Right(arrRawData(i, COL_NUMBER), Len(arrRawData(i, COL_NUMBER)) - 1)

        End If

        ' 填滿以用於排序
        arrRawData(i + 1, colStoreyNum) = arrRawData(i, colStoreyNum)
        arrRawData(i + 2, colStoreyNum) = arrRawData(i, colStoreyNum)
        arrRawData(i + 3, colStoreyNum) = arrRawData(i, colStoreyNum)
        arrRawData(i + 1, colNumberNoC) = arrRawData(i, colNumberNoC)
        arrRawData(i + 2, colNumberNoC) = arrRawData(i, colNumberNoC)
        arrRawData(i + 3, colNumberNoC) = arrRawData(i, colNumberNoC)

    Next

    ' 新增一個工作表
    Set WS_OUTPUT = ThisWorkbook.Sheets.Add(After:=Worksheets("General Information"))


    With WS_OUTPUT

        ' 輸出到 excel 利用內建函數進行排序
        .Range(.Cells(rowLbRawData, colLbRawData), .Cells(rowUbRawData, colUbRawData)) = arrRawData

        ' 以樓層排序，再以去掉 c 的文字排序
        .Range(.Cells(LB_REBAR, colLbRawData), .Cells(rowUbRawData, colUbRawData)).Sort _
            Key1:=.Range(.Cells(LB_REBAR, colStoreyNum), .Cells(rowUbRawData, colStoreyNum)), Order1:=xlDescending, DataOption1:=xlSortNormal, _
            Key2:=.Range(.Cells(LB_REBAR, colNumberNoC), .Cells(rowUbRawData, colNumberNoC)), Order2:=xlAscending, DataOption2:=xlSortNormal, _
            Header:=xlNo, MatchCase:=True, Orientation:=xlTopToBottom, SortMethod:=xlPinYin

        ' 收入資料進 Array
        ARR_REBAR = .Range(.Cells(rowLbRawData, colLbRawData), .Cells(rowUbRawData, colUbRawData - 2))

    End With

    ' 清空輸入
    WS_OUTPUT.Cells.Clear

End Function


Function GetRatioData()

    ' 樓層數字化，用以比較上下樓層。
    For i = LB_REBAR To UB_REBAR Step 4
        ARR_RATIO(i, COL_STOREY) = DicIsEmpty(OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY)), COL_STOREY_NUM)
    Next

    ' 計算鋼筋面積
    For i = LB_REBAR To UB_REBAR
        For j = COL_REBAR_LEFT To COL_REBAR_RIGHT
            ARR_RATIO(i, j) = CalRebarArea(ARR_REBAR(i, j))
        Next
    Next

    ' 一二排截面積相加
    For i = LB_REBAR To UB_REBAR Step 2
        For j = COL_REBAR_LEFT To COL_REBAR_RIGHT
            ARR_RATIO(i, j) = ARR_RATIO(i, j) + ARR_RATIO(i + 1, j)
        Next
    Next

    ' 計算箍筋面積
    For i = LB_REBAR To UB_REBAR Step 4
        For j = COL_STIRRUP_LEFT To COL_STIRRUP_RIGHT
            ARR_RATIO(i, j) = CalStirrupArea(ARR_REBAR(i, j))
        Next
    Next

    ' 計算側筋面積
    For i = LB_REBAR To UB_REBAR Step 4
        ARR_RATIO(i, COL_SIDEBAR) = CalSideRebarArea(ARR_REBAR(i, COL_SIDEBAR))
    Next

    ' 計算有效深度
    For i = LB_REBAR To UB_REBAR Step 4

        rebar_ = Split(ARR_REBAR(i, COL_REBAR_LEFT), "-")
        stirrup = Split(ARR_REBAR(i, COL_STIRRUP_LEFT), "@")
        fyDb = OBJ_REBAR_SIZE.Item(rebar_(1))(COL_DB)
        ' fyDb = APP.VLookup(rebar_(1), REBAR_SIZE, COL_DB, False)
        fytDb = OBJ_REBAR_SIZE.Item(SplitStirrup(stirrup(0)))(COL_DB)
        ' fytDb = APP.VLookup(SplitStirrup(SplitStirrup(stirrup(0))), REBAR_SIZE, COL_DB, False)
        cover_ = OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_COVER)

        ' 雙排筋
        ARR_RATIO(i, COL_D) = ARR_REBAR(i, COL_H) - (cover_ + fytDb + fyDb * 1.5)

    Next

End Function


Function CalRebarArea(rebar_)

    tmp = Split(rebar_, "-")

    If tmp(0) <> 0 Then

        ' 轉換鋼筋尺寸為截面積
        tmp(1) = DicIsEmpty(OBJ_REBAR_SIZE.Item(tmp(1)), COL_AREA, rebar_ & "主筋尺寸搜尋不到，請確認格式是否有誤。")
        ' tmp(1) = APP.VLookup(tmp(1), REBAR_SIZE, COL_AREA, False)

        CalRebarArea = tmp(0) * tmp(1)
    Else
        CalRebarArea = 0
    End If

End Function


Function CalStirrupArea(rebar_)
'
' 考量雙箍
'
    tmp = Split(rebar_, "@")

    bars = Split(tmp(0), "#")

    ' 箍筋號數
    bars(1) = "#" & bars(1)

    stirrupArea = DicIsEmpty(OBJ_REBAR_SIZE.Item(bars(1)), COL_AREA, rebar_ & "箍筋尺寸搜尋不到，請確認格式是否有誤。")

    ' 轉換鋼筋尺寸為截面積
    If bars(0) = "" Then
        CalStirrupArea = 2 * stirrupArea
        ' CalStirrupArea = 2 * APP.VLookup(bars(1), REBAR_SIZE, COL_AREA, False)
    Else
        CalStirrupArea = 2 * bars(0) * stirrupArea
        ' CalStirrupArea = 2 * bars(0) * APP.VLookup(bars(1), REBAR_SIZE, COL_AREA, False)
    End If

End Function


Function CalSideRebarArea(rebar_)

    If rebar_ <> "-" Then

        ' 去掉 EF
        ' 1#4EF => 1#4
        sidebarNoEF = Left(rebar_, Len(rebar_) - 2)

        tmp = Split(sidebarNoEF, "#")

        ' 轉換鋼筋尺寸為截面積
        tmp(1) = DicIsEmpty(OBJ_REBAR_SIZE.Item("#" & tmp(1)), COL_AREA, rebar_ & "側筋尺寸搜尋不到，請確認格式是否有誤。")
        ' tmp(1) = APP.VLookup("#" & tmp(1), REBAR_SIZE, COL_AREA, False)

        ' 對稱雙排
        CalSideRebarArea = 2 * tmp(1)

    Else
        CalSideRebarArea = 0
    End If

End Function


Function SplitStirrup(rebar_)
'
' 處理雙箍的情況
'
    bars = Split(rebar_, "#")

    SplitStirrup = "#" & bars(1)

End Function


Function GetTypeMessage(Girder, Beam, GroundBeam)

    If S_BEAM_TYPE = "大梁" Then
        GetTypeMessage = Girder

    ElseIf S_BEAM_TYPE = "小梁" Then
        GetTypeMessage = Beam

    ElseIf S_BEAM_TYPE = "地梁" Then
        GetTypeMessage = GroundBeam

    End If

End Function

Function WarningMessage(warningMessageCode, i)

    ARR_REBAR(i, COL_MESSAGE) = warningMessageCode & vbCrLf & ARR_REBAR(i, COL_MESSAGE)

End Function

Function PrintResult()

    For i = LB_REBAR To UB_REBAR Step 4

        With WS_OUTPUT

            For j = COL_STOREY To COL_H
                .Range(.Cells(i, j), .Cells(i + 3, j)).Merge
            Next j

            For j = COL_SIDEBAR To COL_NOTE
                .Range(.Cells(i, j), .Cells(i + 3, j)).Merge
            Next j

            .Range(.Cells(i, COL_MESSAGE), .Cells(i + 3, COL_MESSAGE)).Merge

            If ARR_REBAR(i, COL_MESSAGE) = "" Then
                ARR_REBAR(i, COL_MESSAGE) = "(S), (E), (i) - SCAN 結果 ok"
            Else
                .Cells(i, COL_MESSAGE).Style = "壞"
                ARR_REBAR(i, COL_MESSAGE) = Left(ARR_REBAR(i, COL_MESSAGE), Len(ARR_REBAR(i, COL_MESSAGE)) - 1)
            End If

        End With

    Next

    With WS_OUTPUT
        .Range(.Cells(LBound(ARR_REBAR, 1), LBound(ARR_REBAR, 2)), .Cells(UBound(ARR_REBAR, 1), UBound(ARR_REBAR, 2))) = ARR_REBAR

        .Columns(COL_MESSAGE).EntireColumn.AutoFit

    End With

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

'     If S_BEAM_TYPE = "大梁" Then
'         columnStart = 4
'     ElseIf S_BEAM_TYPE = "小梁" Then
'         columnStart = 7
'     ElseIf S_BEAM_TYPE = "地梁" Then
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
'     For i = LB_REBAR To UB_REBAR

'         For j = COL_REBAR_LEFT To COL_REBAR_RIGHT

'             rebarNumber = Split(ARR_REBAR(i, j), "-")

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
'     For i = LB_REBAR To UB_REBAR Step 4

'         If ARR_REBAR(i, COL_SIDEBAR) <> "-" Then

'             sideRebar = Left(ARR_REBAR(i, COL_SIDEBAR), Len(ARR_REBAR(i, COL_SIDEBAR)) - 2)

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
'     For i = LB_REBAR To UB_REBAR Step 4

'         For j = COL_STIRRUP_LEFT To COL_STIRRUP_RIGHT

'             rebarNumber = Split(ARR_REBAR(i, j), "@")(0)
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

    For i = LB_REBAR To UB_REBAR Step 4

        For j = COL_REBAR_LEFT To COL_REBAR_RIGHT

            code = 0.003 * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D)

            ' 請確認是否符合 上層筋下限 規定
            If ARR_RATIO(i, j) < code Then
                Call WarningMessage("【0104】上層筋鋼筋比不得小於 0.3% (" & 0.003 * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D) & "cm^2)", i)
            End If

            ' 請確認是否符合 下層筋下限 規定
            If ARR_RATIO(i + 2, j) < code Then
                Call WarningMessage("【0105】下層筋鋼筋比不得小於 0.3% (" & 0.003 * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D) & "cm^2)", i)
            End If

            For k = i To i + 3

                rebar_ = Split(ARR_REBAR(k, j), "-")

                stirrup = Split(ARR_REBAR(i, j + 4), "@")

                If rebar_(0) > 1 Then

                    fyDb = OBJ_REBAR_SIZE.Item(rebar_(1))(COL_DB)
                    ' fyDb = APP.VLookup(rebar_(1), REBAR_SIZE, COL_DB, False)
                    fytDb = OBJ_REBAR_SIZE.Item(SplitStirrup(stirrup(0)))(COL_DB)
                    ' fytDb = APP.VLookup(SplitStirrup(SplitStirrup(stirrup(0))), REBAR_SIZE, COL_DB, False)

                    spacing = (ARR_REBAR(i, COL_BW) - OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_COVER) * 2 - fytDb * 2 - rebar_(0) * fyDb) / (rebar_(0) - 1)

                    If Not spacing < 25 Then
                        Call WarningMessage("【0106】主筋淨間距不得大於 25 cm (主筋淨間距 = " & spacing & " cm)", i)
                    End If

                ElseIf rebar_(0) = "1" Then

                    Call WarningMessage("【0107】主筋單排支數不得小於 2 支", i)

                End If
            Next
        Next
    Next

End Function

Function Norm4_9_3()
'
' 深梁：
' 由於地梁都很深，所以大部分都是深梁
' 垂直剪力鋼筋面積 Av 不得小於 0.0025 * bw * s，s 不得大於 d / 5 或 30 cm。

    For i = LB_REBAR To UB_REBAR Step 4

        For j = COL_STIRRUP_LEFT To COL_STIRRUP_RIGHT

            stirrup = Split(ARR_REBAR(i, j), "@")

            code = ARR_RATIO(i, j) > 0.0025 * ARR_REBAR(i, COL_BW) * stirrup(1)

            If Not code Then
                Call WarningMessage("【0101】短梁箍筋面積不得小於 0.0025 * bw * s (" & 0.0025 * ARR_REBAR(i, COL_BW) * stirrup(1) & "cm^2)", i)
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

    For i = LB_REBAR To UB_REBAR Step 4

        tmp = Split(ARR_REBAR(i, COL_SIDEBAR), "#")

        ' 分成四種狀況
        If tmp(0) = "-" Then
            code = False
        ElseIf tmp(0) = "1" Then
            code = ARR_RATIO(i, COL_SIDEBAR) > 0.0015 * ARR_REBAR(i, COL_BW) * (ARR_REBAR(i, COL_H) - bs - fs)
        ElseIf tmp(0) = "2" Then
            code = ARR_RATIO(i, COL_SIDEBAR) > 0.0015 * ARR_REBAR(i, COL_BW) * (ARR_REBAR(i, COL_H) - bs - fs) / 2
        Else
            code = ARR_RATIO(i, COL_SIDEBAR) > 0.0015 * ARR_REBAR(i, COL_BW) * (ARR_REBAR(i, COL_H) - bs - fs - 15 - 15) / (tmp(0) - 1)
        End If

        If Not code Then
            Call WarningMessage("【0102】短梁側筋面積不得小於 0.0015 * bw * s2", i)
        End If

    Next

End Function

Function EconomicNorm4_9_4()
'
' 經濟性指標：
' 深梁
' Avh need to less than 1.5 * 0.0015 * COL_BW * S2

    bs = 20
    fs = 60
    factor = 1.5

    For i = LB_REBAR To UB_REBAR Step 4

        tmp = Split(ARR_REBAR(i, COL_SIDEBAR), "#")

        If tmp(0) = "-" Then
            code = True
        ElseIf tmp(0) = "1" Then
            code = ARR_RATIO(i, COL_SIDEBAR) < factor * 0.0015 * ARR_REBAR(i, COL_BW) * (ARR_REBAR(i, COL_H) - bs - fs)
        ElseIf tmp(0) = "2" Then
            code = ARR_RATIO(i, COL_SIDEBAR) < factor * 0.0015 * ARR_REBAR(i, COL_BW) * (ARR_REBAR(i, COL_H) - bs - fs) / 2
        Else
            code = ARR_RATIO(i, COL_SIDEBAR) < factor * 0.0015 * ARR_REBAR(i, COL_BW) * (ARR_REBAR(i, COL_H) - bs - fs - 15 - 15) / (tmp(0) - 1)
        End If

        If Not code Then
            Call WarningMessage("【0103】短梁側筋面積不得大於 1.5 * 0.0015 * bw * S2", i)
        End If

    Next

End Function

Function SafetyLoad()
'
' 安全性指標：
' 載重預警
' 0.6 * 1/8 * wu * L^2 <= As * fy * d

    For i = LB_REBAR To UB_REBAR Step 4

        ' m2
        maxRatio = APP.Max(ARR_RATIO(i, COL_REBAR_LEFT), ARR_RATIO(i, COL_REBAR_MID), ARR_RATIO(i, COL_REBAR_RIGHT), ARR_RATIO(i + 2, COL_REBAR_LEFT), ARR_RATIO(i + 2, COL_REBAR_MID), ARR_RATIO(i + 2, COL_REBAR_RIGHT)) / 10000

        ' m
        slab = OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_SLAB)

        ' tf/m2
        SDL = OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_SDL)

        ' m
        band = OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_BAND)

        ' tf/m2
        LL = OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_LL)

        ' kgf/cm2
        fy = OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FY) / 1000 * 10000

        ' m
        span = ARR_REBAR(i, COL_SPAN) / 100

        ' m
        d = ARR_RATIO(i, COL_D) / 100

        ' 在上面就轉換成一致的單位
        ' 轉換 kgw-m => tf-m: * 100000 我覺得有錯
        mn = 1 / 8 * (1.2 * (slab * 2.4 + SDL) + 1.6 * LL) * band * span ^ 2
        ' mn = 1 / 8 * (1.2 * (0.15 * 2.4 + APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_SDL, False)) + 1.6 * APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_LL, False)) * COL_BAND ^ 2 * 100000

        capacity = maxRatio * fy * d
        ' capacity = maxRatio * APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FY, False) * ARR_RATIO(i, COL_D)

        If 0.6 * mn > capacity Then
            Call WarningMessage("【0312】垂直載重配筋可能不足", i)
        End If

    Next

End Function

Function SafetyRebarRatioForSB()
'
' 安全性指標：
' 小梁鋼筋比在 2.5% 以下

    For i = LB_REBAR To UB_REBAR Step 4

        For j = COL_REBAR_LEFT To COL_REBAR_RIGHT

            limit = 0.025 * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D)

            If ARR_RATIO(i, j) > limit Then
                Call WarningMessage("【0310】請確認上層筋上限，是否在 2.5% 以下", i)
            End If

            If ARR_RATIO(i + 2, j) > limit Then
                Call WarningMessage("【0311】請確認下層筋上限，是否在 2.5% 以下", i)
            End If

        Next

    Next

End Function

Function SafetyRebarRatioForGB()
'
' 安全性指標：
' 地梁鋼筋比在 2% 以下

    For i = LB_REBAR To UB_REBAR Step 4

        For j = COL_REBAR_LEFT To COL_REBAR_RIGHT

            limit = 0.02 * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D)

            If ARR_RATIO(i, j) > limit Then
                Call WarningMessage("【0108】上層筋鋼筋比不得大於 2% (" & limit & "cm^2)", i)
            End If

            If ARR_RATIO(i + 2, j) > limit Then
                Call WarningMessage("【0109】下層筋鋼筋比不得大於 2% (" & limit & "cm^2)", i)
            End If

        Next

    Next

End Function

Function Norm3_6()
'
' 受撓構材之最少鋼筋量：
' 3-3 As >= 0.8 * sqr(fc') / fy * bw * d
' 3-4 As >= 14 / fy * bw * d

For i = LB_REBAR To UB_REBAR Step 4

    code3_3 = 0.8 * Sqr(OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FC_BEAM)) / OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FY) * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D)
    ' code3_3 = 0.8 * Sqr(APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FC_BEAM, False)) / APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FY, False) *ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D)
    code3_4 = 14 / OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FY) * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D)
    ' code3_4 = 14 / APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FY, False) * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D)

    If ARR_RATIO(i, COL_REBAR_LEFT) < code3_3 Or ARR_RATIO(i, COL_REBAR_LEFT) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0201】請確認左端上層筋下限，是否符合規範 3.6 規定", "【0301】請確認左端上層筋下限，是否符合規範 3.6 規定", "請確認左端上層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If ARR_RATIO(i, COL_REBAR_MID) < code3_3 Or ARR_RATIO(i, COL_REBAR_MID) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0202】請確認中央上層筋下限，是否符合規範 3.6 規定", "【0302】請確認中央上層筋下限，是否符合規範 3.6 規定", "請確認中央上層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If ARR_RATIO(i, COL_REBAR_RIGHT) < code3_3 Or ARR_RATIO(i, COL_REBAR_RIGHT) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0203】請確認右端上層筋下限，是否符合規範 3.6 規定", "【0303】請確認右端上層筋下限，是否符合規範 3.6 規定", "請確認右端上層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If ARR_RATIO(i + 2, COL_REBAR_LEFT) < code3_3 Or ARR_RATIO(i + 2, COL_REBAR_LEFT) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0204】請確認左端下層筋下限，是否符合規範 3.6 規定", "【0304】請確認左端下層筋下限，是否符合規範 3.6 規定", "請確認左端下層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If ARR_RATIO(i + 2, COL_REBAR_MID) < code3_3 Or ARR_RATIO(i + 2, COL_REBAR_MID) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0205】請確認中央下層筋下限，是否符合規範 3.6 規定", "【0305】請確認中央下層筋下限，是否符合規範 3.6 規定", "請確認中央下層筋下限，是否符合規範 3.6 規定"), i)
    End If

    If ARR_RATIO(i + 2, COL_REBAR_RIGHT) < code3_3 Or ARR_RATIO(i + 2, COL_REBAR_RIGHT) < code3_4 Then
        Call WarningMessage(GetTypeMessage("【0206】請確認右端下層筋下限，是否符合規範 3.6 規定", "【0306】請確認右端下層筋下限，是否符合規範 3.6 規定", "請確認右端下層筋下限，是否符合規範 3.6 規定"), i)
    End If

Next

End Function

Function Norm15_4_2_1()
'
' 耐震規範 (1F以下大梁不適用)：
' 拉力鋼筋比不得大於 (fc' + 100) / (4 * fy)，亦不得大於 0.025。

    For i = LB_REBAR To UB_REBAR Step 4

        If ARR_RATIO(i, COL_STOREY) > NUM_FIRST_STOREY Then

            code15_4_2_1 = APP.Min((OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FC_BEAM) + 100) / (4 * OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FY)) * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D), 0.025 * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D))
            ' code15_4_2_1 = APP.Min((APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FC_BEAM, False) + 100) / (4 * APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FY, False)) * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D), 0.025 * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D))

            If ARR_RATIO(i, COL_REBAR_LEFT) > code15_4_2_1 Then
                Call WarningMessage("【0212】請確認左端上層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If ARR_RATIO(i, COL_REBAR_MID) > code15_4_2_1 Then
                Call WarningMessage("【0213】請確認中央上層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If ARR_RATIO(i, COL_REBAR_RIGHT) > code15_4_2_1 Then
                Call WarningMessage("【0214】請確認右端上層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If ARR_RATIO(i + 2, COL_REBAR_LEFT) > code15_4_2_1 Then
                Call WarningMessage("【0215】請確認左端下層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If ARR_RATIO(i + 2, COL_REBAR_MID) > code15_4_2_1 Then
                Call WarningMessage("【0216】請確認中央下層筋上限，是否符合規範 15.4.2.1 規定", i)
            End If

            If ARR_RATIO(i + 2, COL_REBAR_RIGHT) > code15_4_2_1 Then
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

    For i = LB_REBAR To UB_REBAR Step 4

        If ARR_RATIO(i, COL_STOREY) > NUM_FIRST_STOREY Then

            maxRatio = APP.Max(ARR_RATIO(i, COL_REBAR_LEFT), ARR_RATIO(i, COL_REBAR_MID), ARR_RATIO(i, COL_REBAR_RIGHT), ARR_RATIO(i + 2, COL_REBAR_LEFT), ARR_RATIO(i + 2, COL_REBAR_MID), ARR_RATIO(i + 2, COL_REBAR_RIGHT))
            minRatio = APP.Min(ARR_RATIO(i, COL_REBAR_LEFT), ARR_RATIO(i, COL_REBAR_MID), ARR_RATIO(i, COL_REBAR_RIGHT), ARR_RATIO(i + 2, COL_REBAR_LEFT), ARR_RATIO(i + 2, COL_REBAR_MID), ARR_RATIO(i + 2, COL_REBAR_RIGHT))
            code15_4_2_2 = minRatio >= maxRatio / 4

            If Not code15_4_2_2 Then
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

    ' 梁長
    span = 400

    For i = LB_REBAR To UB_REBAR Step 4

        rebarLEFT = Split(ARR_REBAR(i, COL_REBAR_LEFT), "-")
        rebarRIGHT = Split(ARR_REBAR(i, COL_REBAR_RIGHT), "-")

        If ARR_RATIO(i, COL_REBAR_MID) * 0.7 < ARR_RATIO(i, COL_REBAR_LEFT) And rebarLEFT(0) > 3 And ARR_REBAR(i, COL_SPAN) > span Then
            Call WarningMessage("【0111】左端上層鋼筋量不得大於中央鋼筋量的 70% (" & ARR_RATIO(i, COL_REBAR_MID) * 0.7 & " cm^2)", i)
        End If

        If ARR_RATIO(i, COL_REBAR_MID) * 0.7 < ARR_RATIO(i, COL_REBAR_RIGHT) And rebarRIGHT(0) > 3 And ARR_REBAR(i, COL_SPAN) > span Then
            Call WarningMessage("【0112】右端上層鋼筋量不得大於中央鋼筋量的 70% (" & ARR_RATIO(i, COL_REBAR_MID) * 0.7 & " cm^2)", i)
        End If

    Next

End Function

Function EconomicTopRebarRelative()
'
' 經濟性指標：
' 如果鋼筋支數大於3支，中央上層鋼筋量需小於端部最小鋼筋量的 70%。
' 淨跨度大於 400 cm，才要檢討


    For i = LB_REBAR To UB_REBAR Step 4

        minRatio = APP.Min(ARR_RATIO(i, COL_REBAR_LEFT), ARR_RATIO(i, COL_REBAR_RIGHT))

        rebar_ = Split(ARR_REBAR(i, COL_REBAR_MID), "-")

        If ARR_RATIO(i, COL_REBAR_MID) > minRatio * 0.7 And rebar_(0) > 3 And ARR_REBAR(i, COL_SPAN) > 400 Then
            Call WarningMessage("【0221】請確認中央上層筋相對鋼筋量，是否符合中央上層鋼筋量需小於端部最小鋼筋量的 70% 規定", i)
        End If

    Next

End Function

Function EconomicBotRebarRelativeForGB()
'
' 經濟性指標：
' 如果鋼筋支數大於3支，中央下層鋼筋量需小於端部最小鋼筋量的 70%。
' 淨跨度大於 400 cm，才要檢討

    ' 跨度
    span = 400

    For i = LB_REBAR To UB_REBAR Step 4

        bot = i + 2

        minRatio = APP.Min(ARR_RATIO(bot, COL_REBAR_LEFT), ARR_RATIO(bot, COL_REBAR_RIGHT))

        rebar_ = Split(ARR_REBAR(bot, COL_REBAR_MID), "-")

        If ARR_RATIO(bot, COL_REBAR_MID) > minRatio * 0.7 And rebar_(0) > 3 And ARR_REBAR(i, COL_SPAN) > span Then
            Call WarningMessage("【0110】中央下層鋼筋量不得大於端部最小鋼筋量 70% (" & minRatio * 0.7 & " cm^2)", i)
        End If

    Next

End Function

Function Norm13_5_1AndSafetyRebarNumber()
'
' 鋼筋間距之限制：
' 規範內容：同層平行鋼筋間之淨距不得小於 1.0db，或粗粒料標稱最大粒徑 1.33 倍，亦不得小於 2.5 cm。
' 實作內容：單排淨距需在 1db 以上 且 單排支數需大於1支。

    For k = LB_REBAR To UB_REBAR

        For j = COL_REBAR_LEFT To COL_REBAR_RIGHT

            ' 重要：因為k每步都是1，所以增加一個k來計算每4步。
            ' 其實可以用 i = i + 4 比較簡單
            i = 4 * Fix((k - 3) / 4) + 3

            rebar_ = Split(ARR_REBAR(k, j), "-")

            stirrup = Split(ARR_REBAR(i, j + 4), "@")

            ' 等於 0 直接沒做事
            If rebar_(0) > 1 Then

                fyDb = OBJ_REBAR_SIZE.Item(rebar_(1))(COL_DB)
                ' fyDb = APP.VLookup(rebar_(1), REBAR_SIZE, COL_DB, False)
                fytDb = OBJ_REBAR_SIZE.Item(SplitStirrup(stirrup(0)))(COL_DB)
                ' fytDb = APP.VLookup(SplitStirrup(stirrup(0)), REBAR_SIZE, COL_DB, False)
                cover_ = OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_COVER)

                ' 第一種方法
                ' Max = Fix((ARR_REBAR(i, COL_BW) - 4 * 2 - fytDb * 2 - fyDb) / (2 * fyDb)) + 1
                ' CInt(rebar_(0)) > Max
                ' 第二種方法
                ' spacing = (ARR_REBAR(i, COL_BW) - 4 * 2 - fytDb * 2 - fyDb) / (CInt(rebar_(0)) - 1) - fyDb
                ' 可以不需要型別轉換
                ' Spacing = (ARR_REBAR(i, COL_BW) - 4 * 2 - fytDb * 2 - CInt(rebar_(0)) * fyDb) / (CInt(rebar_(0)) - 1)
                spacing = (ARR_REBAR(i, COL_BW) - cover_ * 2 - fytDb * 2 - rebar_(0) * fyDb) / (rebar_(0) - 1)

                ' Norm13_5_1
                ' 淨距不少於1Db
                If spacing < fyDb Or spacing < 2.5 Then
                    Call WarningMessage(GetTypeMessage("【0210】請確認單排支數上限，是否符合淨距不少於 1 Db 規定", "【0308】請確認單排支數上限，是否符合淨距不少於 1 Db 規定", "請確認單排支數上限，是否符合淨距不少於 1 Db 規定"), i)
                End If

            ElseIf rebar_(0) = "1" Then

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

    For i = LB_REBAR To UB_REBAR Step 4

        For j = COL_STIRRUP_LEFT To COL_STIRRUP_RIGHT

            stirrup = Split(ARR_REBAR(i, j), "@")

            If stirrup(1) < 10 Then
                Call WarningMessage(GetTypeMessage("【0219】請確認箍筋間距下限，是否符合 10cm 以上規定", "請確認箍筋間距下限，是否符合 10cm 以上規定", "【0113】箍筋間距不得小於 10cm"), i)
            ElseIf stirrup(1) > 30 Then
                Call WarningMessage(GetTypeMessage("【0220】請確認箍筋間距上限，是否符合 30cm 以下規定", "請確認箍筋間距上限，是否符合 30cm 以下規定", "【0114】箍筋間距不得大於 30cm"), i)
            End If

        Next

    Next

End Function

Function Norm4_6_6_3()
'
' 剪力鋼筋量大於 3.52/fy

    For i = LB_REBAR To UB_REBAR Step 4

        For j = COL_STIRRUP_LEFT To COL_STIRRUP_RIGHT

            stirrup = Split(ARR_REBAR(i, j), "@")

            avMin = APP.Max(0.2 * Sqr(OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FC_BEAM)) * ARR_REBAR(i, COL_BW) * stirrup(1) / OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FYT), 3.5 * ARR_REBAR(i, COL_BW) * stirrup(1) / OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FYT))
            ' avMin = APP.Max(0.2 * Sqr(APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FC_BEAM, False)) * ARR_REBAR(i, COL_BW) * stirrup(1) / APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FYT, False), 3.5 * ARR_REBAR(i, COL_BW) * stirrup(1) / APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FYT, False))
            av = ARR_RATIO(i, j)

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
' Vc = 0.53 * sqrt(f’c) * bw * d
' Vs = Av * fyt * d / s
' 實作內容：剪力鋼筋量需在 4 * Vc * 120% 以下。規範為 vs <= 4 * vc，由於取整數容易超過，所以放寬標準 120%。

    For i = LB_REBAR To UB_REBAR Step 4

        For j = COL_STIRRUP_LEFT To COL_STIRRUP_RIGHT

            stirrup = Split(ARR_REBAR(i, j), "@")
            ' rebar_ = Split(ARR_REBAR(i, j - 4), "-")

            ' fyDb = OBJ_REBAR_SIZE.Item(rebar_(1))(COL_DB)
            ' fyDb = APP.VLookup(rebar_(1), REBAR_SIZE, COL_DB, False)
            ' fytDb = OBJ_REBAR_SIZE.Item(SplitStirrup(stirrup(0)))(COL_DB)
            ' fytDb = APP.VLookup(SplitStirrup(stirrup(0)), REBAR_SIZE, COL_DB, False)
            ' effectiveDepth = ARR_REBAR(i, COL_H) - (4 + fytDb + fyDb / 2)
            av = ARR_RATIO(i, j)

            ' code4.4.1.1
            vc = 0.53 * Sqr(OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FC_BEAM)) * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D)
            ' vc = 0.53 * Sqr(APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FC_BEAM, False)) * ARR_REBAR(i, COL_BW) * ARR_RATIO(i, COL_D)

            ' code4.6.7.2
            vs = av * OBJ_INFO.Item(ARR_REBAR(i, COL_STOREY))(COL_FYT) * ARR_RATIO(i, COL_D) / stirrup(1)
            ' vs = av * APP.VLookup(ARR_REBAR(i, COL_STOREY), GENERAL_INFORMATION, COL_FYT, False) * ARR_RATIO(i, COL_D) / stirrup(1)

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
' 實作內容： L <= 4 * COL_H

    For i = LB_REBAR To UB_REBAR Step 4

        If (ARR_REBAR(i, COL_SPAN) - ARR_REBAR(i, COL_SUPPORT)) <= 4 * ARR_REBAR(i, COL_H) Then
            Call WarningMessage("【0208】請確認是否為深梁", i)
        End If

    Next

End Function

Function Norm3_7_5()

    For i = LB_REBAR To UB_REBAR Step 4

        If ARR_REBAR(i, COL_H) > 90 Then
            Call WarningMessage(GetTypeMessage("【0207】請確認是否符合 規範 3.7.5", "【0307】請確認是否符合 規範 3.7.5", "請確認是否符合 規範 3.7.5"), i)
        End If

    Next

End Function

Function CalcLap()
'
' calc lap.
'
' @see Ld
'

    For top_row = LB_REBAR To UB_REBAR Step 4

        bot_row = top_row + 3

        B = ARR_REBAR(top_row, COL_BW)
        fc = OBJ_INFO.Item(ARR_REBAR(top_row, COL_STOREY))(COL_FC_BEAM)
        fy = OBJ_INFO.Item(ARR_REBAR(top_row, COL_STOREY))(COL_FY)
        fyh = OBJ_INFO.Item(ARR_REBAR(top_row, COL_STOREY))(COL_FYT)
        cover = OBJ_INFO.Item(ARR_REBAR(top_row, COL_STOREY))(COL_COVER)

        Call AssignLapWithDifferentRowAndWithDifferentStirrup(top_row, "top", top_row, B, fc, fy, fyh, cover)
        Call AssignLapWithDifferentRowAndWithDifferentStirrup(bot_row, "bot", top_row, B, fc, fy, fyh, cover)

    Next top_row

End Function

Function AssignLapWithDifferentRowAndWithDifferentStirrup(ByVal loc_row, ByVal top_bot_string, ByVal top_row, ByVal B, ByVal fc, ByVal fy, ByVal fyh, ByVal cover)
'
'

    rebar = Split(ARR_REBAR(loc_row, COL_REBAR_LEFT), "-")
    fyDb = OBJ_REBAR_SIZE.Item(rebar(1))(COL_DB)
    num = rebar(0)

    If num > 1 Then
        ARR_REBAR(loc_row, COL_LAP_LEFT) = ran.Max(LapWithDifferentStirrup(COL_STIRRUP_LEFT, top_bot_string, top_row, B, fc, fy, fyh, cover, fyDb, num), LapWithDifferentStirrup(COL_STIRRUP_MID, top_bot_string, top_row, B, fc, fy, fyh, cover, fyDb, num))
    Else
        OBJ_ERR_MSG.Add "Cells(" & loc_row & ", " & COL_REBAR_LEFT & ") rebar number smaller than 2."
    End If

    rebar = Split(ARR_REBAR(loc_row, COL_REBAR_MID), "-")
    fyDb = OBJ_REBAR_SIZE.Item(rebar(1))(COL_DB)
    num = rebar(0)

    If num > 1 Then
        ARR_REBAR(loc_row, COL_LAP_MID) = ran.Max(LapWithDifferentStirrup(COL_STIRRUP_LEFT, top_bot_string, top_row, B, fc, fy, fyh, cover, fyDb, num), LapWithDifferentStirrup(COL_STIRRUP_MID, top_bot_string, top_row, B, fc, fy, fyh, cover, fyDb, num), LapWithDifferentStirrup(COL_STIRRUP_RIGHT, top_bot_string, top_row, B, fc, fy, fyh, cover, fyDb, num))
    Else
        OBJ_ERR_MSG.Add "Cells(" & loc_row & ", " & COL_LAP_MID & ") rebar number smaller than 2."
    End If

    rebar = Split(ARR_REBAR(loc_row, COL_REBAR_RIGHT), "-")
    fyDb = OBJ_REBAR_SIZE.Item(rebar(1))(COL_DB)
    num = rebar(0)

    If num > 1 Then
        ARR_REBAR(loc_row, COL_LAP_RIGHT) = ran.Max(LapWithDifferentStirrup(COL_STIRRUP_RIGHT, top_bot_string, top_row, B, fc, fy, fyh, cover, fyDb, num), LapWithDifferentStirrup(COL_STIRRUP_MID, top_bot_string, top_row, B, fc, fy, fyh, cover, fyDb, num))
    Else
        OBJ_ERR_MSG.Add "Cells(" & loc_row & ", " & COL_LAP_RIGHT & ") rebar number smaller than 2."
    End If

End Function

Function LapWithDifferentStirrup(ByVal col_stirrup, ByVal top_bot_string, ByVal top_row, ByVal B, ByVal fc, ByVal fy, ByVal fyh, ByVal cover, ByVal db, ByVal num)
'
'

    stirrup = Split(ARR_REBAR(top_row, col_stirrup), "@")
    dh = OBJ_REBAR_SIZE.Item(SplitStirrup(stirrup(0)))(COL_DB)
    avh = ARR_RATIO(top_row, col_stirrup)
    spacing = stirrup(1)

    ' Girder stirrups is 10cm
    If S_BEAM_TYPE = "大梁" And OBJ_INFO.Item(ARR_REBAR(top_row, COL_STOREY))(COL_STIRRUP_10) Then
        spacing = 10
    End If

    LapWithDifferentStirrup = ran.RoundUp(1.3 * Ld(B, fc, fy, fyh, cover, db, num, dh, avh, spacing, top_bot_string))

End Function

Function Ld(ByVal B, ByVal fc, ByVal fy, ByVal fyh, ByVal cover, ByVal db, ByVal num, ByVal dh, ByVal avh, ByVal spacing, ByVal top_bot_string)
'
' ld.
'
' @since 1.0.0
' @param {number} [B] Girder/Beam Width (cm).
' @param {number} [fc] 28-days concrete compressive strength (kgf/cm2).
' @param {number} [fy] Rebar Nominal Yielding Streng (kgf/cm2).
' @param {number} [fyh] Stirrup Nominal Yielding Streng (kgf/cm2).
' @param {number} [cover] Clear Cover THK (cm).
' @param {number} [db] diameter of flexural rebar (cm).
' @param {number} [num] numbers of flexural rebar at lap location.
' @param {number} [dh] stirrup diameter (cm).
' @param {number} [avh] stirrup area (cm2).
' @param {number} [spacing] space of stirrup at lap location of flexural rebar (cm).
' @param {string} [top_bot_string] "top" or "bot".
' @return {number} [Ld] ld.
'

    ' 5.2.2
    If Sqr(fc) > 26.5 Then fc = 700

    ' R5.3.4.1.1
    cc = dh + cover

    ' R5.3.4.1.1
    cs = (B - db * num - dh * 2 - cover * 2) / (num - 1) / 2

    ' Vertical splitting failure / Horizontal splitting failure
    cb = ran.Min(cc, cs) + db / 2

    ' R5.3.4.1.2
    If cc <= cs Then
        ktr = avh * fyh / 105 / spacing
    Else
        ktr = 2 / num * avh * fyh / 105 / spacing
    End If

    ' 5.3.4.1
    Ld = 0.28 * fy / Sqr(fc) * db / ran.Min((cb + ktr) / db, 2.5)

    ' 5.3.4.1
    simple_ld = 0.19 * fy / Sqr(fc) * db

    Ld = ran.Min(Ld, simple_ld)

    ' phi_s factor
    If db < 2.2 Then Ld = 0.8 * Ld

    ' phi_t factor
    If top_bot_string = "top" Then Ld = 1.3 * Ld

    ' 5.3.1
    If Ld < 30 Then Ld = 30

End Function
