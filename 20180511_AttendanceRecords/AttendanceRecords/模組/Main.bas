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
    Set ws = Worksheets("�u��")
    ' �����[�Z�ɶ�
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

        ' �������X����
        j = i + 1
        While arrInput(j, colAttType) = "���X" Or arrInput(j, colAttType) = "���X��^"
            j = j + 1
        Wend

        nextAttType = arrInput(j, colAttType)

        If prevAttType = "�W�Z" Then
            If nextAttType <> "�U�Z" Then
                errorMsg = errorMsg + "�� " & i & " �C���W�Z�d�C" & vbNewLine & "�� " & i + 1 & " �C���ӭn���U�Z�d�C" & vbNewLine & vbNewLine
            End If
        ElseIf prevAttType = "�[�Z" Then
            If nextAttType <> "�[�Z����" Then
                errorMsg = errorMsg + "�� " & i & " �C���[�Z�d�C" & vbNewLine & "�� " & i + 1 & " �C���ӭn���[�Z�����d�C" & vbNewLine & vbNewLine
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
    ' �ѩ�᭱������A�ҥH�ݭn���J�@�C���P���A�o�O��� hack ������
    arrOutput(UBound(arrOutput), colDayTime) = Day(arrInput(uBoundInput, colDayTime)) + 1

    ' �P���X
    ' �ɶ�
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

        ' �������X����
        j = i + 1
        While arrInput(j, colAttType) = "���X" Or arrInput(j, colAttType) = "���X��^"
            j = j + 1
        Wend

        nextTime = arrOutput(j, colTime)

        ' �ɼ�
        ' �u��ɼƦ� 1.5
        If attType = "�W�Z" Then

            ' �|�ˤ��J
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

        ' �а��ɼ�
        If arrOutput(i, colRealHour) < 8 Then
            arrOutput(i, colLeave) = 8 - arrOutput(i, colRealHour)
        Else
            arrOutput(i, colLeave) = "-"
        End If

        ' �[�Z�ɼ�
        If attType = "�[�Z" Then
            dayTime = arrInput(i, colDayTime)

            ' �p�G�O�g�@��g���񰲴N�L�k�B�z�F�A�ݭn�H�u�P�_
            If Weekday(dayTime, 2) < 5 And prevTime < weekdayOverTime And nextTime > weekdayOverTime Then
                overTime = (nextTime - weekdayOverTime) * 24

            Else
                overTime = (nextTime - prevTime) * 24

            End If

            ' �|�ˤ��J
            arrOutput(i, colOverTime) = Round(overTime, 3)

        End If

    Next i

    ' �B�z�[�Z�O 1.34 �٬O 1.67
    lower = 2
    For i = 2 To uBoundInput - 1

        prevDay = Day(arrInput(i, colDayTime))
        nextDay = Day(arrInput(i + 1, colDayTime))

        If prevDay <> nextDay Or i + 1 = uBoundInput Then

            upper = i

            ' �p�����`�ɼ�
            overTime = 0
            For j = lower To upper
                overTime = overTime + arrOutput(j, colOverTime)
            Next j

            ' �P�_�O 1.34 �� 1.67
            If overTime > 2 Then
                overTime34 = 2
                overTime67 = overTime - 2
            Else
                overTime34 = overTime
                overTime67 = 0
            End If

            ' ��b�Ĥ@�ӥ[�Z�B
            For j = lower To upper

                attType = arrInput(j, colAttType)

                If attType = "�[�Z" Then

                    arrOutput(j, colOverTime34) = overTime34
                    arrOutput(j, colOverTime67) = overTime67

                    j = upper

                End If

            Next j

            lower = i + 1

        End If

    Next i

    ' view
    With Worksheets("�u�ɭp�⵲�G")

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

    With Worksheets("�u��")
        .Range(.Columns(1), .Columns(5)).Copy
    End With

    With Worksheets("�u�ɭp�⵲�G")

        ' Output �C��P�B Input
        .Cells(1, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        .Range(.Columns(1), .Columns(5)).Borders.LineStyle = xlContinuous
        .Cells.Font.Name = "�L�n������"
        .Cells.Font.Name = "Calibri"
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter

    End With

End Function

