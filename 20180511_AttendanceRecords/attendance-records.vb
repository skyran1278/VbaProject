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

    ' Golobal Var
    Set ran = New UTILS_CLASS
    Set ws = Worksheets("修改後DATA")

    time0 = Timer
    Call ran.PerformanceVBA(True)

    ' Model
    arrInput = ran.GetRangeToArray(ws, 1, 1, 1, 5)
    uBoundInput = UBound(arrInput)
    colAttType = 4
    colDayTime = 5

    ' condtroller
    arrOutput = arrInput
    ReDim Preserve arrOutput(1 To uBoundInput, 1 To 14)
    colWeek = 6
    colTime = 7
    colHour = 8
    colRealHour = 9
    colLeave = 10
    colOverTime = 11
    colOverTime34 = 13
    colOverTime67 = 14

    ' 星期幾
    ' 時間
    For i = 2 To uBoundInput

        dayTime = arrInput(i, colDayTime)

        arrOutput(i, colWeek) = dayTime
        arrOutput(i, colTime) = Hour(dayTime) & ":" & Minute(dayTime)

    Next i


    For i = 2 To uBoundInput - 1

        attType = arrInput(i, colAttType)
        prevTime = arrOutput(i, colTime)
        nextTime = arrOutput(i + 1, colTime)

        ' 時數
        ' 真實時數扣 1.5
        If attType = "上班" Then

            arrOutput(i, colHour) = (nextTime - prevTime) * 24

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
            arrOutput(i, colOverTime) = (nextTime - prevTime) * 24
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
    Set outputWS = Worksheets("VBA Output")
    outputWS.Range(outputWS.Columns(1), outputWS.Columns(14))
    outputWS.Range(outputWS.Cells(1, 1), outputWS.Cells(UBound(arrOutput), UBound(arrOutput, 2))) = arrOutput

    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTimeVBA(time0)

End Sub

