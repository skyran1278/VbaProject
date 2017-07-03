Dim SUM_ARRAY(), ROW_START, ROW_END, TIME0, CANCEL, SELECT_VALUE, SELECT_VALUE_LENGTH, SELECT_VALUE_COUNT, MAX_LOOP_VALUE, OUTPUT_VALUE_COUNT, CONST_MAX_LOOP_VALUE, OUTPUT_POINT

' DATA 資料命名
Const OUTPUT_VALUE = 1
Const INPUT_VALUE = 2
Const Name = 3
Const STORY = 4
Const LABEL = 5
Const MAX_M = 6
Const FY = 7
Const Z = 8
Const LENGTH = 9
Const GROUP = 10
Const REPLACE_NUMBER = 11
Const SELECT_NUMBER = 12

Function Initialize()

    TIME0 = Timer

    Worksheets("Z").Activate

    ROW_START = 8
    ROW_END = Cells(Rows.Count, Z).End(xlUp).Row

    ' 排序
    Worksheets("Z").Range(Cells(7, 3), Cells(ROW_END, 10)).Sort _
        Key1:=Range(Cells(ROW_START, Z), Cells(ROW_END, Z)), Order1:=xlDescending, Header:=xlYes

    ReDim SUM_ARRAY(ROW_START To ROW_END)

    Range(Cells(18, OUTPUT_VALUE), Cells(Cells(Rows.Count, OUTPUT_VALUE).End(xlUp).Row, INPUT_VALUE)).ClearContents
    Range(Cells(9, REPLACE_NUMBER), Cells(Cells(Rows.Count, REPLACE_NUMBER).End(xlUp).Row, REPLACE_NUMBER)).ClearContents
    Range(Columns(SELECT_NUMBER), Columns(100)).ClearContents


    Cells(ROW_START, REPLACE_NUMBER).AutoFill Destination:=Range(Cells(ROW_START, REPLACE_NUMBER), Cells(ROW_END, REPLACE_NUMBER))

    SELECT_VALUE = Split(Cells(5, INPUT_VALUE).Value, ",")
    SELECT_VALUE_LENGTH = UBound(SELECT_VALUE)

    CONST_MAX_LOOP_VALUE = Cells(6, INPUT_VALUE)

    OUTPUT_VALUE_COUNT = 18
    OUTPUT_POINT = 13

    Cells(16, OUTPUT_VALUE) = Application.sum(Range(Cells(ROW_START, Z), Cells(ROW_END, Z)))

End Function

Function LoopSelectValue()
'
' 多次執行
'

    For i = 0 To SELECT_VALUE_LENGTH

        SELECT_VALUE_COUNT = SELECT_VALUE(i)
        MAX_LOOP_VALUE = CONST_MAX_LOOP_VALUE
        Cells(OUTPUT_VALUE_COUNT, OUTPUT_VALUE) = SELECT_VALUE(i)
        OUTPUT_VALUE_COUNT = OUTPUT_VALUE_COUNT + 1
        Cells(ROW_START, SELECT_NUMBER) = "*"
        Cells(7, SELECT_NUMBER) = Application.sum(Range(Cells(ROW_START, REPLACE_NUMBER), Cells(ROW_END, REPLACE_NUMBER)))

        Call Controller

        Columns(OUTPUT_POINT) = Columns(SELECT_NUMBER).Value
        Cells(7, OUTPUT_POINT) = SELECT_VALUE(i)
        Columns(SELECT_NUMBER).ClearContents

        OUTPUT_POINT = OUTPUT_POINT + 1

    Next

End Function

Function DoMoreThings(sum)
'
' 顯示動畫
'

    If Cells(7, SELECT_NUMBER) > sum Then
        Cells(7, SELECT_NUMBER) = sum
    End If

    Cells(7, REPLACE_NUMBER) = sum

End Function

Function Controller()

    ' 第一次最佳化
    Do While SELECT_VALUE_COUNT > 1

        For i = ROW_START To ROW_END

            If Cells(i, SELECT_NUMBER) = "" Then

                Cells(i, SELECT_NUMBER) = "*"
                SUM_ARRAY(i) = Application.sum(Range(Cells(ROW_START, REPLACE_NUMBER), Cells(ROW_END, REPLACE_NUMBER)))
                DoMoreThings (SUM_ARRAY(i))
                Cells(i, SELECT_NUMBER) = ""

            End If

        Next

        Cells(Application.Match(Application.Min(SUM_ARRAY), SUM_ARRAY, 0) + ROW_START - 1, SELECT_NUMBER) = "*"

        SELECT_VALUE_COUNT = SELECT_VALUE_COUNT - 1

    Loop

    ' 多次最佳化
    Do
        selectBefore = Range(Cells(1, SELECT_NUMBER), Cells(ROW_END, SELECT_NUMBER)).Value

        For i = ROW_START + 1 To ROW_END

            If Cells(i, SELECT_NUMBER) = "*" Then

                Cells(i, SELECT_NUMBER) = ""
                ReDim SUM_ARRAY(ROW_START To ROW_END)

                For j = ROW_START To ROW_END

                    If Cells(j, SELECT_NUMBER) = "" Then

                        Cells(j, SELECT_NUMBER) = "*"
                        SUM_ARRAY(j) = Application.sum(Range(Cells(ROW_START, REPLACE_NUMBER), Cells(ROW_END, REPLACE_NUMBER)))
                        DoMoreThings (SUM_ARRAY(j))
                        Cells(j, SELECT_NUMBER) = ""

                    End If

                Next

                Cells(Application.Match(Application.Min(SUM_ARRAY), SUM_ARRAY, 0) + ROW_START - 1, SELECT_NUMBER) = "*"

            End If

        Next

        selectAfter = Range(Cells(1, SELECT_NUMBER), Cells(ROW_END, SELECT_NUMBER)).Value

        Call PrintEachLoopValue

        doLoop = False

        For i = ROW_START To ROW_END

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

    Cells(OUTPUT_VALUE_COUNT, OUTPUT_VALUE) = Application.sum(Range(Cells(ROW_START, REPLACE_NUMBER), Cells(ROW_END, REPLACE_NUMBER)))

    Cells(OUTPUT_VALUE_COUNT, INPUT_VALUE) = Cells(OUTPUT_VALUE_COUNT, OUTPUT_VALUE) / Cells(16, OUTPUT_VALUE)

    OUTPUT_VALUE_COUNT = OUTPUT_VALUE_COUNT + 1

    MAX_LOOP_VALUE = MAX_LOOP_VALUE - 1

End Function

Function Terminate()

    Cells(7, REPLACE_NUMBER).ClearContents

    If Timer - TIME0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - TIME0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - TIME0) / 60, 2) & " Min", vbOKOnly
    End If

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
    Call Initialize

    Call LoopSelectValue

    ' Terminate
    Call Terminate

End Sub



