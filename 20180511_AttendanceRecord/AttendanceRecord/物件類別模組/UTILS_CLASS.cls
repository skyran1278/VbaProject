VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UTILS_CLASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @license UTILS_CLASS v3.1.0
' UTILS_CLASS.vb
'
' Copyright (c) 2016-present, ran
'
' This source code is licensed under the MIT license found in the
' LICENSE file in the root directory of this source tree.
'
' - Getting Start
' Dim ran As New UTILS_CLASS
'
' - All API
' CreateDictionary
' GetRangeToArray
' GetRowEnd
' RoundUp
' ExecutionTime
' PerformanceVBA
' Min
' Max
' QuickSortArray
' ParseJSON
' ListPaths
' GetFilteredValues
' GetFilteredTable
' OpenTextFile

Option Explicit

Private time0 As Double
Private p&, token, dic

Function CreateDictionary(ByVal arr, ByVal colKey, ByVal colValue)
'
' 取代內建的 VLookup.
' 也可用作取得 Unique Array
'
' @since 2.0.0
' @param {array} [arr] to create dictionary table.
' @param {number} [colKey] key column.
' @param {number|boolean} [colValue] value column or false to use all value.
' @return {object} [CreateDictionary] descrip.
' @example
' objDictionary.Item(key)

    ' 設定 Dictionary
    Set objDictionary = CreateObject("Scripting.Dictionary")

    lbArr = LBound(arr, 1)
    ubArr = UBound(arr, 1)

    If colValue Then
        For rowArr = lbArr To ubArr
            If Not objDictionary.Exists(arr(rowArr, colKey)) Then
                Call objDictionary.Add(arr(rowArr, colKey), arr(rowArr, colValue))
            End If
        Next rowArr
    Else
        For rowArr = lbArr To ubArr
            If Not objDictionary.Exists(arr(rowArr, colKey)) Then
                ' VBA 不能方便的存取整列整欄，所以用 Index
                ' Application.WorksheetFunction.Index(array, 0, columnYouWant)
                ' Application.WorksheetFunction.Index(array, rowYouWant, 0)
                Call objDictionary.Add(arr(rowArr, colKey), Application.WorksheetFunction.Index(arr, rowArr, 0))
            End If
        Next rowArr
    End If


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

' 使用這個要小心，會怪怪的
' 估計是遇到了 VBA 底層的問題
' 好像沒有其他辦法了
' 0.5 + 0.5 可能為 0.999999999，這是規範的問題，比較難處理
Function RoundUp(ByVal value As Double)
'
' 取代內建的 RoundUp
'

    If Int(value) = value Then
        RoundUp = value
    Else
        RoundUp = Int(value) + 1
    End If

End Function


Sub ExecutionTime(ByVal isOn As Boolean)
'
' 計算執行時間，取代 ExecutionTimeVBA
'
' @since 2.2.0
' @param {Boolean} [isOn] True = time0, False = show Msg.
'
    If isOn Then
        time0 = Timer
    Else
        If Timer - time0 < 60 Then
            MsgBox "Execution Time " & Application.Round((Timer - time0), 2) & " Sec", vbOKOnly
        Else
            MsgBox "Execution Time " & Application.Round((Timer - time0) / 60, 2) & " Min", vbOKOnly
        End If
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

' 速度是 APP.Min 的 10 倍
' FIXME: 不知道為甚麼 ran 無法
Function Min(ParamArray values() As Variant) As Variant
   Dim minValue, value As Variant

   minValue = values(0)

   For Each value In values
       If value < minValue Then minValue = value
   Next

   Min = minValue

End Function


Function Max(ParamArray values() As Variant) As Variant
   Dim maxValue, value As Variant

   maxValue = values(0)

   For Each value In values
       If value > maxValue Then maxValue = value
   Next

   Max = maxValue

End Function


Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin = -1, Optional lngMax = -1, Optional lngColumn = 0)
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

    Set APP = Application.WorksheetFunction
    Dim ran As New UTILS_CLASS

    Dim generalInformation As Worksheet
    Set generalInformation = Worksheets("General Information")

    GENERAL_INFORMATION = ran.GetRangeToArray(generalInformation, 1, 4, 4, 12)

    Set objDictionary = CreateDictionary(GENERAL_INFORMATION, 1, False)

    Debug.Print objDictionary.Item("RF")(1)


    Debug.Print Timer - time0

End Sub

'-------------------------------------------------------------------
' VBA JSON Parser
' https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a
'-------------------------------------------------------------------
Function ParseJSON(json$, Optional key$ = "obj") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If token(p) = "{" Then ParseObj key Else ParseArr key
    Set ParseJSON = dic
End Function

Function ParseObj(key$)
    Do: p = p + 1
        Select Case token(p)
            Case "]"
            Case "[":  ParseArr key
            Case "{"
                       If token(p + 1) = "}" Then
                           p = p + 1
                           dic.Add key, "null"
                       Else
                           ParseObj key
                       End If

            Case "}":  key = ReducePath(key): Exit Do
            Case ":":  key = key & "." & token(p - 1)
            Case ",":  key = ReducePath(key)
            Case Else: If token(p + 1) <> ":" Then dic.Add key, token(p)
        End Select
    Loop
End Function

Function ParseArr(key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
            Case "}"
            Case "{":  ParseObj key & ArrayID(e)
            Case "[":  ParseArr key
            Case "]":  Exit Do
            Case ":":  key = key & ArrayID(e)
            Case ",":  e = e + 1
            Case Else: dic.Add key & ArrayID(e), token(p)
        End Select
    Loop
End Function

'-------------------------------------------------------------------
' Support Functions
'-------------------------------------------------------------------
Function Tokenize(s$)
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = RExtract(s, Pattern, True)
End Function

Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n, v
  With CreateObject("vbscript.regexp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .TEST(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.value = """""" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  RExtract = v
End Function

Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function

Function ReducePath$(key$)
    If InStr(key, ".") Then ReducePath = Left(key, InStrRev(key, ".") - 1) Else ReducePath = key
End Function

Function ListPaths(dic)
    Dim s$, v
    For Each v In dic
        s = s & v & " --> " & dic(v) & vbLf
    Next
    Debug.Print s
End Function

Function GetFilteredValues(dic, match)
    Dim c&, i&, v, w
    v = dic.keys
    ReDim w(1 To dic.Count)
    For i = 0 To UBound(v)
        If v(i) Like match Then
            c = c + 1
            w(c) = dic(v(i))
        End If
    Next
    ReDim Preserve w(1 To c)
    GetFilteredValues = w
End Function

Function GetFilteredTable(dic, cols)
    Dim c&, i&, j&, v, w, z
    v = dic.keys
    z = GetFilteredValues(dic, cols(0))
    ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
    For j = 1 To UBound(cols) + 1
         z = GetFilteredValues(dic, cols(j - 1))
         For i = 1 To UBound(z)
            w(i, j) = z(i)
         Next
    Next
    GetFilteredTable = w
End Function

Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
    End With
End Function
