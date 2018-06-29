' @license UTILS_CLASS v2.0.0
' UTILS_CLASS.vb
'
' Copyright (c) 2016-present, skyran
'
' This source code is licensed under the MIT license found in the
' LICENSE file in the root directory of this source tree.

' TODO: 參考 lodash 文檔
' If Then is FASTER than IIF
' 若想要最佳化效能 還是需要自己寫一個針對的最快
' 呼叫 function 比本地直接執行慢 3.5 倍左右，但是通常都還是會拆分 function，所以我認為沒差。
'
' 教學
' Private ran As UTILS_CLASS
' Set ran = New UTILS_CLASS


Function CreateDictionary(arr, colKey, colValue)
'
' 取代內建的 VLookup.
'
' @since 2.0.0
' @param {array} [arr] to create dictionary table.
' @param {number} [colKey] key column.
' @param {number} [colValue] value column.
' @return {object} [CreateDictionary] descrip.
'

    ' 設定 Dictionary
    Set objDictionary = CreateObject("Scripting.Dictionary")

    lbArr = LBound(arr, 1)
    ubArr = UBound(arr, 1)

    For rowArr = lbArr To ubArr
        If Not objDictionary.Exists(arr(rowArr, colKey)) Then
            Call objDictionary.Add(arr(rowArr, colKey), arr(rowArr, colValue))
        End If
    Next rowArr

    Set CreateDictionary = objDictionary

End Function


Function GetRangeToArray(ws, rowStart, colStart, rowEnd, colEnd)
'
' 取得表格資料
'
' @returns GetRangeToArray(Array)

    With ws
        GetRangeToArray = .Range(.Cells(rowStart, colStart), .Cells(.Cells(Rows.Count, rowEnd).End(xlUp).Row, colEnd))
    End With

End Function


Function GetRowEnd(ws, col)
'
' 回傳最後一列 row 值
'
' @param
' @returns

    GetRowEnd = ws.Cells(Rows.Count, col).End(xlUp).Row

End Function


Function FontSetting(ws)
'
' 美化格式
'

    With ws
        .Cells.Font.Name = "微軟正黑體"
        .Cells.Font.Name = "Calibri"
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter
    End With

End Function


Function RoundUp(ByVal Value As Double)
    If Int(Value) = Value Then
        RoundUp = Value
    Else
        RoundUp = Int(Value) + 1
    End If
End Function


Sub ExecutionTimeVBA(time0 As Double)
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


Sub PerformanceVBA(isOn As Boolean)
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


Function Min(ParamArray values() As Variant) As Variant
   Dim minValue, Value As Variant

   minValue = values(0)

   For Each Value In values
       If Value < minValue Then minValue = Value
   Next

   Min = minValue

End Function


Function Max(ParamArray values() As Variant) As Variant
   Dim maxValue, Value As Variant

   maxValue = values(0)

   For Each Value In values
       If Value > maxValue Then maxValue = Value
   Next

   Max = maxValue

End Function


Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
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

Sub SpeedTest()
'
' 測試速度用
'

    Dim time0 As Double
    Dim a As Double

    time0 = Timer
    Set app = Application.WorksheetFunction
    For i = 1 To 1000000
        a = app.Max(11, 2, 3)
    Next i
    Debug.Print Timer - time0

    time0 = Timer
    For i = 1 To 1000000
        a = Application.WorksheetFunction.Max(11, 2, 3)
    Next i
    Debug.Print Timer - time0

    time0 = Timer
    For i = 1 To 1000000
        a = Application.Max(11, 2, 3)
    Next i
    Debug.Print Timer - time0

    time0 = Timer
    For i = 1 To 1000000
        a = Max(11, 2, 3)
    Next i
    Debug.Print Timer - time0

End Sub
