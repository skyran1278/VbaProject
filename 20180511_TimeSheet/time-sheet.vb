Private ran As UTILS_CLASS
Private ws As Worksheet
Private APP

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

    ' Golobal Var
    Set ran = New UTILS_CLASS
    Set APP = Application.WorksheetFunction
    Set ws = Worksheets("修改後DATA")

    time0 = Timer
    Call ran.PerformanceVBA(True)

    ' Model
    arrInput = ran.GetRangeToArray(ws, 1, 1, 1, 5)
    uBoundInput = UBound(arrInput)
    colAttType = 4
    colDayTime = 5

    ' view
    ReDim arrOutput(1 To uBoundInput, 1 To 14)
    colWeek = 6
    colTime = 7
    colHour = 8
    colRealHour = 9
    colLeave = 10
    colOverTime = 11
    colOverTime34 = 13
    colOverTime67 = 14

    ' condtroller
    ' prior two column
    For i = 2 To uBoundInput

        dayTime = arrInput(i, colDayTime)

        arrOutput(i, colWeek) = dayTime
        arrOutput(i, colTime) = APP.Hour(dayTime) & ":" & APP.Minute(dayTime)

    Next i

    For i = 2 To uBoundInput

        attType = arrInput(i, colAttType)
        prevTime = arrOutput(i, colTime)
        nextTime = arrOutput(i + 1, colTime)

        If attType = "上班" Then

            arrOutput(i, colHour) = (nextTime - prevTime) * 24

            If arrOutput(i, colHour) - 1.5 > 8 Then
                arrOutput(i, colRealHour) = 8
            Else
                arrOutput(i, colRealHour) = arrOutput(i, colHour) - 1.5
            End If

        ElseIf attType = "加班" Then
            arrOutput(i, colOverTime) = (nextTime - prevTime) * 24
        Else

            arrOutput(i, colHour) = "-"
            arrOutput(i, colRealHour) = "-"

        End If

        If arrOutput(i, colRealHour) < 8 Then
            arrOutput(i, colLeave) = 8 - arrOutput(i, colRealHour)
        Else
            arrOutput(i, colLeave) = "-"
        End If

    Next i

    lower = 2
    overTime = 0

    For i = 2 To uBoundInput

        attType = arrInput(i, colAttType)
        prevDay = Day(arrInput(i - 1, colDayTime))
        nextDay = Day(arrInput(i, colDayTime))

        If prevDay <> nextDay Then

            upper = i - 1

            For j = lower To upper
                overTime = overTime + arrOutput(i, colOverTime)
            Next j

            If overTime > 2 Then
                overTime34 = 2
                overTime67 = overTime - 2
            Else
                overTime34 = overTime
                overTime67 = 0
            End If

            For j = lower To upper

                If attType = "加班" Then

                    arrOutput(i, colOverTime34) = overTime34
                    arrOutput(i, colOverTime67) = overTime67

                    j = upper

                End If

            Next j

            lower = i

        End If

    Next i

    ' view
    ws.Range(ws.Cells(1, colWeek), ws.Cells(UBound(arrOutput), colOverTime67)) = arrOutput

    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTimeVBA(time0)

End Sub

