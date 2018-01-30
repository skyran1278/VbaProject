' If Then is FASTER than IIF
' 若想要最佳化效能 還是需要自己寫一個針對的最快
' 呼叫 function 比本地直接執行慢 3.5 倍左右，但是通常都還是會拆分 function，所以我認為沒差。


Public Sub ExecutionTimeVBA(time0 As Double)
'
' 計算執行時間
'
' @param time0(Double)

    If Timer - time0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - time0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - time0) / 60, 2) & " Min", vbOKOnly
    End If

End Sub


Public Sub PerformanceVBA(isOn As Boolean)
'
' 提升執行效能
'
' @param isOn(Boolean)

    Application.ScreenUpdating = Not (isOn) ' 37.26

    Application.DisplayStatusBar = Not (isOn) ' 57.29

    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic) ' 57

    Application.EnableEvents = Not (isOn) ' 58.75

    ' FIXME: 這裡需要再想一下
    ' displayPageBreakState = ActiveSheet.DisplayPageBreaks
    ' ActiveSheet.DisplayPageBreaks = False
    ' ActiveSheet.DisplayPageBreaks = IIf(isOn, False, displayPageBreaksState)
    ' ActiveSheet.DisplayPageBreaks = displayPageBreaksState
    ' ThisWorkbook.ActiveSheet.DisplayPageBreaks = Not(isOn) 'note this is a sheet-level setting 53.51

    ' .Value2

End Sub


Public Function Min(ParamArray values() As Variant) As Variant
   Dim minValue, Value As Variant

   minValue = values(0)

   For Each Value In values
       If Value < minValue Then minValue = Value
   Next

   Min = minValue

End Function


Public Function Max(ParamArray values() As Variant) As Variant
   Dim maxValue, Value As Variant

   maxValue = values(0)

   For Each Value In values
       If Value > maxValue Then maxValue = Value
   Next

   Max = maxValue

End Function


Public Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortArray(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortArray(SortArray, i, lngMax, lngColumn)

End Sub

' Private Sub Workbook_Open()
'     application.onkey("^+v", TextOnly)
' End Sub

' Sub TextOnly()
' '
' ' 純文字貼上
' '

'     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

' End Sub

Public Sub SpeedTest()
'
' 測試速度用
'

    Dim time0 As Double



    topStory = Application.Match("RF", Application.Index(GENERAL_INFORMATION, 0, STORY), 0)
    firstStory = Application.Match("1F", Application.Index(GENERAL_INFORMATION, 0, STORY), 0)

    ' 頂樓區 1/4
    checkStoryNumber = Fix((firstStory - topStory + 1) / 4) + topStory - 1


    For i = DATA_ROW_START To DATA_ROW_END
        If RATIO_DATA(i, STORY) >= topStory And checkStoryNumber >= RATIO_DATA(i, STORY) And RATIO_DATA(i, REBAR) > 0.01 * 1.2 Then
                Call WarningMessage("【0405】請確認高樓區鋼筋比，是否超過 1.2 %", i)
        End If
    Next

    time0 = Timer

    numStory = UBound(GENERAL_INFORMATION)
    For i = 1 To numStory
        If GENERAL_INFORMATION(numStory - i + 1, STORY) = "1F" Then
            firstStory = i
        ElseIf GENERAL_INFORMATION(numStory - i + 1, STORY) = "RF" Then
            topStory = i
        End If
    Next

    checkStoryNumber = Fix((topStory - firstStory + 1) / 4)

    For i = DATA_ROW_START To DATA_ROW_END
        For j = topStory - checkStoryNumber + 1 To topStory

            If RAW_DATA(i, STORY) = GENERAL_INFORMATION(numStory - j + 1, STORY) And RATIO_DATA(i, REBAR) > 0.01 * 1.2 Then
                    Call WarningMessage("【0405】請確認高樓區鋼筋比，是否超過 1.2 %", i)
            End If

        Next

    Next


    Call ExecutionTimeVBA(time0)

End Sub
