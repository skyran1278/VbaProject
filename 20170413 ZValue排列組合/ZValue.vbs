Dim SUM_ARRAY(), ROW_START, ROW_END, TIME_0, CANCEL, SELECT_VALUE, SELECT_VALUE_MIN, SELECT_VALUE_MAX, SELECT_VALUE_COUNT, MAX_LOOP_VALUE, OUTPUT_VALUE_COUNT

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
Const SELECT_NUMBER = 11
Const REPLACE_NUMBER = 12

Function Initialize()

    TIME_0 = Timer

    Worksheets("Z").Activate

    ROW_START = 8
    ROW_END = Cells(Rows.Count, Z).End(xlUp).Row

    ' If Application.Sum(Range(Cells(ROW_START, REPLACE_NUMBER), Cells(ROW_END, REPLACE_NUMBER))) <> 0 Then
    '     CANCEL = MsgBox("請確認第 11 欄不為 0", vbOKCancel Or vbExclamation)
    ' End If

    ' Application.ScreenUpdating = False

    ' 排序
    Worksheets("Z").Range(Cells(7, 3), Cells(ROW_END, 10)).Sort _
        Key1:=Range(Cells(ROW_START, Z), Cells(ROW_END, Z)), Order1:=xlDescending, Header:=xlYes

    ReDim SUM_ARRAY(ROW_START To ROW_END)

    Range(Cells(19, OUTPUT_VALUE), Cells(1000, INPUT_VALUE)).ClearContents

    Cells(ROW_START, REPLACE_NUMBER).AutoFill Destination:=Range(Cells(ROW_START, REPLACE_NUMBER), Cells(ROW_END, REPLACE_NUMBER))

    SELECT_VALUE = split(Cells(5, INPUT_VALUE),"-")
    MAX_LOOP_VALUE = Cells(6, INPUT_VALUE)

    OUTPUT_VALUE_COUNT = 18

    Cells(16, OUTPUT_VALUE) = Application.Sum(Range(Cells(ROW_START, Z), Cells(ROW_END, Z)))

End Function

Function LoopSelectValue()

    SELECT_VALUE_MIN = SELECT_VALUE(0)
    SELECT_VALUE_MAX = SELECT_VALUE(1)

    SELECT_VALUE = SELECT_VALUE_MIN

    Do
        SELECT_VALUE_COUNT = SELECT_VALUE

        columns(SELECT_NUMBER).ClearContents
        Cells(ROW_START, SELECT_NUMBER) = "*"

        Call Controller

        SELECT_VALUE = SELECT_VALUE + 1
        OUTPUT_VALUE_COUNT = OUTPUT_VALUE_COUNT + 1

    Loop While SELECT_VALUE + 1 < SELECT_VALUE_MAX

End Function

Function Controller()

    ' 第一次最佳化
    Do While SELECT_VALUE_COUNT > 1

        For i = ROW_START To ROW_END

            If Cells(i, SELECT_NUMBER) = "" Then

                Cells(i, SELECT_NUMBER) = "*"
                SUM_ARRAY(i) = Application.Sum(Range(Cells(ROW_START, REPLACE_NUMBER), Cells(ROW_END, REPLACE_NUMBER)))
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
                        SUM_ARRAY(j) = Application.Sum(Range(Cells(ROW_START, REPLACE_NUMBER), Cells(ROW_END, REPLACE_NUMBER)))
                        Cells(j, SELECT_NUMBER) = ""

                    End If

                Next

                Cells(Application.Match(Application.Min(SUM_ARRAY), SUM_ARRAY, 0) + ROW_START - 1, SELECT_NUMBER) = "*"

            End If

        Next

        selectAfter = Range(Cells(1, SELECT_NUMBER), Cells(ROW_END, SELECT_NUMBER)).Value

        Call PrintEachLoopValue

        For i = ROW_START To ROW_END

            If selectBefore(i, 1) <> selectAfter(i, 1) Then
                doloop = True
                Exit For
            End If

            doloop = False

        Next

    Loop While doloop And MAX_LOOP_VALUE > 0



End Function

Function PrintEachLoopValue()

    Cells(OUTPUT_VALUE_COUNT, OUTPUT_VALUE) = Application.Sum(Range(Cells(ROW_START, REPLACE_NUMBER), Cells(ROW_END, REPLACE_NUMBER)))

    OUTPUT_VALUE_COUNT = OUTPUT_VALUE_COUNT + 1

    MAX_LOOP_VALUE = MAX_LOOP_VALUE - 1

End Function

Function Terminate()

    ' Application.ScreenUpdating = True

    Cells(18, INPUT_VALUE).AutoFill Destination:=Range(Cells(18, INPUT_VALUE), Cells(OUTPUT_VALUE_COUNT - 1, INPUT_VALUE))

    If Timer - TIME_0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - TIME_0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - TIME_0) / 60, 2) & " Min", vbOKOnly
    End If

End Function

sub Main()
'
' * 目的
'       Check Norm

' * 環境
'       Excel

' * 輸出入格式
'       輸入：
'       輸出：

' * 執行時間
'       0.06 Sec

' * 輸出結果的精確度與檢驗方式
'

    ' Initialize
    Call Initialize
    ' If CANCEL = 2 Then
    '     CANCEL = Empty
    '     Exit Sub
    ' End If

    call LoopSelectValue

    ' Terminate
    Call Terminate

End sub