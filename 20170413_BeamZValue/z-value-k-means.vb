Private UTIL As UTILS_CLASS
Private APP
Private WS_Z
Private ROW_INPUT
Private ROW_LOOP
Private ROW_SUM
Private ROW_Z_START
Private ROW_NOR
Private ROW_OUTPUT_START
Private COL_OUTPUT
Private COL_INPUT
Private COL_RATIO
Private COL_DATA_START
Private COL_Z
Private COL_L
Private COL_AS
Private COL_REPLACE
Private COL_SELECT
Private COL_SELECT_START
Private COL_SELECT_END
Private ROW_RATIO_END
Private ROW_Z_END

Private SUM_ARRAY()
' Private ROW_Z_START
' Private ROW_Z_END
' Private TIME0
Private CANCEL
Private SELECT_VALUE
Private SELECT_VALUE_LENGTH
Private SELECT_VALUE_COUNT
Private MAX_LOOP_VALUE
Private ROW_OUTPUT_VALUE_COUNT
Private CONST_MAX_LOOP_VALUE
Private COL_OUTPUT_POINT

' DATA 資料命名
' Private Const COL_OUTPUT = 1
' Private Const COL_RATIO = 2
' Private Const COL_Z = 13
' Private Const COL_REPLACE = 21
' Private Const COL_SELECT = 22

Function LoopSelectValue()
'
' 多次執行
'

    For i = 0 To SELECT_VALUE_LENGTH

        SELECT_VALUE_COUNT = SELECT_VALUE(i)
        MAX_LOOP_VALUE = CONST_MAX_LOOP_VALUE
        Cells(ROW_OUTPUT_VALUE_COUNT, COL_OUTPUT) = SELECT_VALUE(i)
        ROW_OUTPUT_VALUE_COUNT = ROW_OUTPUT_VALUE_COUNT + 1
        Cells(ROW_Z_START, COL_SELECT) = "*"
        Cells(ROW_SUM, COL_SELECT) = APP.sum(Range(Cells(ROW_Z_START, COL_REPLACE), Cells(ROW_Z_END, COL_REPLACE)))

        Call Controller

        Range(Cells(ROW_Z_START, COL_OUTPUT_POINT), Cells(ROW_Z_END, COL_OUTPUT_POINT)) = Range(Cells(ROW_Z_START, COL_SELECT), Cells(ROW_Z_END, COL_SELECT)).Value
        ' Columns(COL_OUTPUT_POINT) = Columns(COL_SELECT).Value
        Cells(ROW_SUM, COL_OUTPUT_POINT) = SELECT_VALUE(i)
        Columns(COL_SELECT).ClearContents

        COL_OUTPUT_POINT = COL_OUTPUT_POINT + 1

    Next

End Function

Function DoMoreThings(sum)
'
' 顯示動畫
'

    If Cells(ROW_SUM, COL_SELECT) > sum Then
        Cells(ROW_SUM, COL_SELECT) = sum
    End If

    Cells(ROW_SUM, COL_REPLACE) = sum

End Function

Function Controller()

    ' 第一次最佳化
    Do While SELECT_VALUE_COUNT > 1

        For i = ROW_Z_START To ROW_Z_END

            If Cells(i, COL_SELECT) = "" Then

                Cells(i, COL_SELECT) = "*"
                SUM_ARRAY(i) = APP.sum(Range(Cells(ROW_Z_START, COL_REPLACE), Cells(ROW_Z_END, COL_REPLACE)))
                DoMoreThings (SUM_ARRAY(i))
                Cells(i, COL_SELECT) = ""

            End If

        Next

        Cells(APP.Match(APP.Min(SUM_ARRAY), SUM_ARRAY, 0) + ROW_SUM, COL_SELECT) = "*"

        SELECT_VALUE_COUNT = SELECT_VALUE_COUNT - 1

    Loop

    ' 多次最佳化
    Do
        selectBefore = Range(Cells(1, COL_SELECT), Cells(ROW_Z_END, COL_SELECT))

        For i = ROW_Z_START + 1 To ROW_Z_END

            If Cells(i, COL_SELECT) = "*" Then

                Cells(i, COL_SELECT) = ""
                ReDim SUM_ARRAY(ROW_Z_START To ROW_Z_END)

                For j = ROW_Z_START To ROW_Z_END

                    If Cells(j, COL_SELECT) = "" Then

                        Cells(j, COL_SELECT) = "*"
                        SUM_ARRAY(j) = APP.sum(Range(Cells(ROW_Z_START, COL_REPLACE), Cells(ROW_Z_END, COL_REPLACE)))
                        DoMoreThings (SUM_ARRAY(j))
                        Cells(j, COL_SELECT) = ""

                    End If

                Next

                Cells(APP.Match(APP.Min(SUM_ARRAY), SUM_ARRAY, 0) + ROW_SUM, COL_SELECT) = "*"

            End If

        Next

        selectAfter = Range(Cells(1, COL_SELECT), Cells(ROW_Z_END, COL_SELECT))

        Call PrintEachLoopValue

        doLoop = False

        For i = ROW_Z_START To ROW_Z_END

            If selectBefore(i, 1) <> selectAfter(i, 1) Then
                doLoop = True
                Exit For
            End If

        Next

    Loop While doLoop And MAX_LOOP_VALUE > 0

End Function

Function PrintEachLoopValue()
'
' 輸出 RATIO
'

    Cells(ROW_OUTPUT_VALUE_COUNT, COL_OUTPUT) = APP.sum(Range(Cells(ROW_Z_START, COL_REPLACE), Cells(ROW_Z_END, COL_REPLACE)))

    Cells(ROW_OUTPUT_VALUE_COUNT, COL_RATIO) = Cells(ROW_OUTPUT_VALUE_COUNT, COL_OUTPUT) / Cells(ROW_NOR, COL_OUTPUT)

    ROW_OUTPUT_VALUE_COUNT = ROW_OUTPUT_VALUE_COUNT + 1

    MAX_LOOP_VALUE = MAX_LOOP_VALUE - 1

End Function


Function DimVaribale()
'
' 宣告可能會用到的變數，並集中到同一地方，增加維護性。
'
' @param
' @returns

    ROW_INPUT = 5
    ROW_LOOP = 6
    ROW_SUM = 7
    ROW_Z_START = 8
    ROW_NOR = 16
    ROW_OUTPUT_START = 18

    COL_OUTPUT = 1
    COL_INPUT = 2
    COL_RATIO = 2
    COL_DATA_START = 8
    COL_Z = 13
    COL_L = 14
    COL_AS = 19
    COL_REPLACE = 21
    COL_SELECT = 22
    COL_SELECT_START = 23
    COL_SELECT_END = 200

    ROW_RATIO_END = UTIL.GetRowEnd(WS_Z, COL_RATIO)
    ROW_Z_END = UTIL.GetRowEnd(WS_Z, COL_Z)

End Function


Function ClearData()
'
' 清除先前的資料
'
' @param
' @returns

    With WS_Z
        ' 清除 output ratio
        .Range(.Cells(ROW_OUTPUT_START, COL_OUTPUT), .Cells(ROW_RATIO_END, COL_RATIO)).ClearContents

        ' Range(Cells(ROW_REPLACE_NUMBER, COL_REPLACE), Cells(Cells(Rows.Count, COL_REPLACE).End(xlUp).Row, COL_REPLACE)).ClearContents

        ' clear select z value
        .Range(.Columns(COL_SELECT_START), .Columns(COL_SELECT_END)).ClearContents

    End With

End Function


Sub Main()
'
' * 目的
'       找出最佳化數值

' * 環境
'       Excel

' * 輸出入格式
'       輸入：Z Value
'       輸出：

' * 執行時間
'       0.06 Sec

' * 輸出結果的精確度與檢驗方式
'

    ' Initialize
    time0 = Timer

    Set UTIL = New UTILS_CLASS
    Set APP = Application.WorksheetFunction
    Set WS_Z = Worksheets("z-value-k-means")

    WS_Z.Activate

    ' ROW_Z_START = 8
    ' ROW_Z_END = Cells(Rows.Count, COL_Z).End(xlUp).Row
    ' COL_START = 8
    ' ROW_OUTPUT_VALUE_START = 18
    ' ROW_REPLACE_NUMBER = 9
    ' ROW_INPUT_VALUE = 5
    ' ROW_LOOP_INPUT_VALUE = 6

    Call DimVaribale

    Call ClearData

    With WS_Z

        ' 排序
        .Range(.Cells(ROW_Z_START, COL_DATA_START), .Cells(ROW_Z_END, COL_Z)).Sort _
        Key1:=.Range(.Cells(ROW_Z_START, COL_Z), .Cells(ROW_Z_END, COL_Z)), Order1:=xlDescending

        ' autofill formula
        .Cells(ROW_Z_START, COL_REPLACE).AutoFill Destination:=.Range(.Cells(ROW_Z_START, COL_REPLACE), .Cells(ROW_Z_END, COL_REPLACE))
        .Range(.Cells(ROW_Z_START, COL_L), .Cells(ROW_Z_START, COL_AS)).AutoFill Destination:=.Range(.Cells(ROW_Z_START, COL_L), .Cells(ROW_Z_END, COL_AS))

        ' input select
        SELECT_VALUE = Split(.Cells(ROW_INPUT, COL_INPUT), ",")

        ' input loop
        CONST_MAX_LOOP_VALUE = .Cells(ROW_LOOP, COL_INPUT)

        .Cells(ROW_NOR, COL_OUTPUT) = APP.sum(.Range(.Cells(ROW_Z_START, COL_Z), .Cells(ROW_Z_END, COL_Z)))

    End With

    ReDim SUM_ARRAY(ROW_Z_START To ROW_Z_END)

    ' Range(Cells(ROW_OUTPUT_VALUE_START, COL_OUTPUT), Cells(Cells(Rows.Count, COL_OUTPUT).End(xlUp).Row, COL_RATIO)).ClearContents
    ' Range(Cells(ROW_REPLACE_NUMBER, COL_REPLACE), Cells(Cells(Rows.Count, COL_REPLACE).End(xlUp).Row, COL_REPLACE)).ClearContents
    ' Range(Columns(COL_SELECT), Columns(100)).ClearContents


    SELECT_VALUE_LENGTH = UBound(SELECT_VALUE)

    ROW_OUTPUT_VALUE_COUNT = 18
    COL_OUTPUT_POINT = COL_SELECT_START


    Call LoopSelectValue

    ' Terminate
    Cells(ROW_SUM, COL_REPLACE).ClearContents

    UTIL.ExecutionTimeVBA(time0)

End Sub



