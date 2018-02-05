Private WS_ANGLE, WS_DATA


Function GetAngle()
'
' 取得 safe beam length 資料
'
' @returns GetAngle(Array)

    With WS_ANGLE
        rowStart = 4
        columnStart = 1
        rowUsed = .Cells(Rows.Count, 1).End(xlUp).Row
        columnUsed = 5

        GetAngle = .Range(.Cells(rowStart, columnStart), .Cells(rowUsed, columnUsed))
    End With

End Function


Function GetData()
'
' 取得 safe node displacement 資料
'
' @returns GetData(Array)

    With WS_DATA
        rowStart = 2
        columnStart = 1
        rowUsed = .Cells(Rows.Count, 1).End(xlUp).Row
        columnUsed = 5

        GetData = .Range(.Cells(rowStart, columnStart), .Cells(rowUsed, columnUsed))
    End With

End Function


Function CombinedData(dataArray)
'
'
'
' @param
' @returns

    Dim combinedArray()

    dataLBound = LBound(dataArray, 1)
    dataUBound = UBound(dataArray, 1)

    ReDim combinedArray(dataLBound To dataUBound, 1 To 2)

    For dataRow = dataLBound To dataUBound
        combinedArray(dataRow, 1) = dataArray(dataRow, 1) & dataArray(dataRow, 4)
        combinedArray(dataRow, 2) = dataArray(dataRow, 5)
    Next dataRow

    CombinedData = combinedArray

    ' Application.Index(combinedArray, , 1) = Application.Index(dataArray, , 1) & Application.Index(dataArray, , 4)
    ' Application.Index(combinedArray, , 2) = Application.Index(dataArray, , 5)

End Function

Function VBALookup(lookup, dictionary)
'
'
'
' @param
' @returns

    If dictionary.Exists(lookup) Then
        VBALookup = dictionary.Item(lookup)
    End If

End Function
' Function vbalookup(lookupRange As Range, refRange As Range, dataCol As Long) As Variant
'   Dim dict As New Scripting.Dictionary
'   Dim myRow As Range
'   Dim I As Long, J As Long
'   Dim vResults() As Variant

'   ' 1. Build a dictionnary
'   For Each myRow In refRange.Columns(1).Cells
'     ' Append A : B to dictionnary
'     dict.Add myRow.Value, myRow.Offset(0, dataCol - 1).Value
'   Next myRow

'   ' 2. Use it over all lookup data
'   ReDim vResults(1 To lookupRange.Rows.Count, 1 To lookupRange.Columns.Count) As Variant
'   For I = 1 To lookupRange.Rows.Count
'     For J = 1 To lookupRange.Columns.Count
'       If dict.Exists(lookupRange.Cells(I, J).Value) Then
'         vResults(I, J) = dict(lookupRange.Cells(I, J).Value)
'       End If
'     Next J
'   Next I

'   vbalookup = vResults
' End Function

' Function vbalookup2(lookupRangepart As Range, refRange As Range, dataCol As Long) As Variant
'   Dim dict As New Scripting.Dictionary
'   Dim myRow As Range
'   Dim I As Long, J As Long
'   Dim vResults() As Variant
'   Dim LastRow As Long
'   Dim Columnselect As String
'   Dim lookupRangepartString As String
'   Dim lookupRangefull As String
'   Dim lookupRange As Range
'   Dim FirstPart As String

' ' Finds last entry on any column
'   LastRow = ActiveSheet.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row

' ' Generates variable for the targeted Column e.g. A
'   Columnselect = Left(lookupRangepart, 1)

' ' Build complete range to lookup
'     strColumn = Replace(lookupRangepart.Address, "$", "")
'     lookupRangefull = strColumn & ":" & Left(strColumn, Len(strColumn) - 1)
'     If IsNumeric(Right(lookupRangefull, 1)) Then lookupRangefull = Left(lookupRangefull, Len(lookupRangefull) - 1)
'     lookupRangefull = lookupRangefull & LastRow

'   Set lookupRange = Range(lookupRangefull)

'   ' 1. Build a dictionnary
'   On Error Resume Next
'   For Each myRow In refRange.Columns(1).Cells
'     ' Append A : B to dictionnary
'     dict.Add UCase(myRow.Value), myRow.Offset(0, dataCol - 1).Value
'   Next myRow

'   ' 2. Use it over all lookup data
'   ReDim vResults(1 To lookupRange.Rows.Count, 1 To lookupRange.Columns.Count) As Variant
'   For I = 1 To lookupRange.Rows.Count
'     For J = 1 To lookupRange.Columns.Count
'       If dict.Exists(UCase(lookupRange.Cells(I, J).Value)) Then
'         vResults(I, J) = dict(UCase(lookupRange.Cells(I, J).Value))
'       End If
'     Next J
'   Next I

'   vbalookup = vResults

'   End Function

Sub Main()
'
' @purpose:
' check 角變量 是否符合規範
'
'
' @algorithm:
' 桿件兩點的沈陷量除以桿件長度
'
'
' @test:
'
'
'
    Dim result()
    Dim time0 As Double

    Call PerformanceVBA(True)

    Set dictionary = CreateObject("Scripting.Dictionary")

    time0 = Timer

    Set WS_ANGLE = Worksheets("Angle")
    Set WS_DATA = Worksheets("Data")

    angleArray = GetAngle()
    dataArray = GetData()
    idAndLoadArray = CombinedData(dataArray)

    angleLBound = LBound(angleArray, 1)
    angleUBound = UBound(angleArray, 1)
    dataLBound = LBound(dataArray, 1)
    dataUBound = UBound(dataArray, 1)

    ReDim result(angleLBound To angleUBound, 1 To 108)


    For dataRow = dataLBound To dataUBound
        If not dictionary.Exists(dataArray(dataRow, 1)) Then
            Call dictionary.Add(dataArray(dataRow, 1), dataArray(dataRow, 2))
        End If
    Next dataRow

    For ASD = 1 To 36
        loadCombo = "ASD" & Format(ASD, "00")
        id1 = (ASD - 1) * 3 + 1
        id2 = (ASD - 1) * 3 + 2
        angleChange = (ASD - 1) * 3 + 3
        For angleRow = angleLBound To angleUBound
            id1AndLoad = angleArray(angleRow, 2) & loadCombo
            id2AndLoad = angleArray(angleRow, 3) & loadCombo
            result(angleRow, id1) = VBALookup(id1AndLoad, dictionary)
            result(angleRow, id2) = VBALookup(id2AndLoad, dictionary)
            result(angleRow, angleChange) = Abs(result(angleRow, id1) - result(angleRow, id2)) / angleArray(angleRow, 5)
        Next angleRow
    Next ASD

    rowStart = 4
    rowEnd = rowStart + angleUBound
    colStart = 6
    colEnd = colStart + 108 - 1

    WS_ANGLE.Range(WS_ANGLE.Cells(rowStart, colStart), WS_ANGLE.Cells(rowEnd, colEnd)) = result

    Call FontSetting
    Call PerformanceVBA(False)
    Call ExecutionTimeVBA(time0)

End Sub

