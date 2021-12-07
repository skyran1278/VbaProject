Attribute VB_Name = "Main"
Private ran As UTILS_CLASS

Sub MAIN()
'
' @purpose:
'
'
'
' @algorithm:
'
'
'
' @test:
'
'
'
    Dim arrOutput()
    Dim time0 As Double
    Dim prevTime As Date
    Dim nextTime As Date
    Dim weekdayOverTime As Date

    ' Global Var
    Set ran = New UTILS_CLASS
    Set ws = Worksheets("工時")
    ' 平日算加班時間
    weekdayOverTime = TimeValue("18:30")

    time0 = Timer
    Call ran.PerformanceVBA(True)

    ' Model
    arrInput = ran.GetRangeToArray(ws, 1, 1, 5, 5)
    uBoundInput = UBound(arrInput)
    colAttType = 4
    colDayTime = 5

    errorMsg = ""

    ' check input
    For i = uBoundInput - 1 To 2 Step -1

        prevAttType = arrInput(i, colAttType)

        ' 忽略公出公返
        j = i + 1
        While arrInput(j, colAttType) = "公出" Or arrInput(j, colAttType) = "公出返回"
            j = j + 1
        Wend

        nextAttType = arrInput(j, colAttType)

        If prevAttType = "上班" Then
            If nextAttType <> "下班" Then
                errorMsg = errorMsg + "第 " & i & " 列為上班卡。" & vbNewLine & "第 " & i + 1 & " 列應該要打下班卡。" & vbNewLine & vbNewLine
            End If
        ElseIf prevAttType = "加班" Then
            If nextAttType <> "加班結束" Then
                errorMsg = errorMsg + "第 " & i & " 列為加班卡。" & vbNewLine & "第 " & i + 1 & " 列應該要打加班結束卡。" & vbNewLine & vbNewLine
            End If
        End If

    Next i

    If errorMsg <> "" Then
        MsgBox errorMsg, 0, "ERROR"
        Exit Sub
    End If

    ' controller
    ' arrOutput = arrInput
    ReDim Preserve arrOutput(1 To uBoundInput + 1, 1 To 16)
    colTime = 6
    colHour = 7
    colIndex = 8
    colDate = 9
    colWeek = 10
    colRealHour = 11
    colLeave = 12
    colOverTime = 13
    colOverTime34 = 15
    colOverTime67 = 16
    ' 由於後面做比較，所以需要插入一列不同的，這是比較 hack 的部分
    arrOutput(UBound(arrOutput), colDayTime) = Day(arrInput(uBoundInput, colDayTime)) + 1

    ' 星期幾
    ' 時間
    For i = 2 To uBoundInput

        dayTime = arrInput(i, colDayTime)

        arrOutput(i, colIndex) = i - 1
        arrOutput(i, colDate) = dayTime
        arrOutput(i, colWeek) = dayTime
        arrOutput(i, colTime) = Hour(dayTime) & ":" & Minute(dayTime)

    Next i


    For i = 2 To uBoundInput - 1

        attType = arrInput(i, colAttType)
        prevTime = arrOutput(i, colTime)

        ' 忽略公出公返
        j = i + 1
        While arrInput(j, colAttType) = "公出" Or arrInput(j, colAttType) = "公出返回"
            j = j + 1
        Wend

        nextTime = arrOutput(j, colTime)

        ' 時數
        ' 真實時數扣 1.5
        If attType = "上班" Then

            ' 四捨五入
            arrOutput(i, colHour) = Round((nextTime - prevTime) * 24, 3)

            If arrOutput(i, colHour) - 1.5 > 8 Then
                arrOutput(i, colRealHour) = 8
            Else
                arrOutput(i, colRealHour) = arrOutput(i, colHour) - 1.5
            End If

        Else

            arrOutput(i, colHour) = "-"
            arrOutput(i, colRealHour) = "-"

        End If

        ' 請假時數
        If arrOutput(i, colRealHour) < 8 Then
            arrOutput(i, colLeave) = 8 - arrOutput(i, colRealHour)
        Else
            arrOutput(i, colLeave) = "-"
        End If

        ' 加班時數
        If attType = "加班" Then
            dayTime = arrInput(i, colDayTime)

            ' 如果是週一到週五放假就無法處理了，需要人工判斷
            If Weekday(dayTime, 2) < 5 And prevTime < weekdayOverTime And nextTime > weekdayOverTime Then
                overTime = (nextTime - weekdayOverTime) * 24

            Else
                overTime = (nextTime - prevTime) * 24

            End If

            ' 四捨五入
            arrOutput(i, colOverTime) = Round(overTime, 3)

        End If

    Next i

    ' 處理加班是 1.34 還是 1.67
    lower = 2
    For i = 2 To uBoundInput - 1

        prevDay = Day(arrInput(i, colDayTime))
        nextDay = Day(arrInput(i + 1, colDayTime))

        If prevDay <> nextDay Or i + 1 = uBoundInput Then

            upper = i

            ' 計算當天總時數
            overTime = 0
            For j = lower To upper
                overTime = overTime + arrOutput(j, colOverTime)
            Next j

            ' 判斷是 1.34 或 1.67
            If overTime > 2 Then
                overTime34 = 2
                overTime67 = overTime - 2
            Else
                overTime34 = overTime
                overTime67 = 0
            End If

            ' 放在第一個加班處
            For j = lower To upper

                attType = arrInput(j, colAttType)

                If attType = "加班" Then

                    arrOutput(j, colOverTime34) = overTime34
                    arrOutput(j, colOverTime67) = overTime67

                    j = upper

                End If

            Next j

            lower = i + 1

        End If

    Next i

    ' view
    With Worksheets("工時計算結果")

        .Range(.Columns(1), .Columns(16)).ClearContents
        .Range(.Cells(1, 1), .Cells(uBoundInput, UBound(arrOutput, 2))) = arrOutput
        .Range(.Cells(1, 1), .Cells(uBoundInput, UBound(arrInput, 2))) = arrInput
        .Activate

    End With

    Call FontSetting

    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTimeVBA(time0)

End Sub


Function FontSetting()

    With Worksheets("工時")
        .Range(.Columns(1), .Columns(5)).Copy
    End With

    With Worksheets("工時計算結果")

        ' Output 顏色同步 Input
        .Cells(1, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        .Range(.Columns(1), .Columns(5)).Borders.LineStyle = xlContinuous
        .Cells.Font.Name = "微軟正黑體"
        .Cells.Font.Name = "Calibri"
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter

    End With

End Function

