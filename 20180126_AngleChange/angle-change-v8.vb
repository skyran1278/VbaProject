Private WS_LINE, WS_DISP


Function GetLines()
'
' 取得 safe beam length 資料
'
' @returns GetLines(Array)

    With WS_LINE
        rowStart = 4
        colStart = 1
        rowEnd = .Cells(Rows.Count, 1).End(xlUp).Row
        colEnd = 5

        GetLines = .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd))
    End With

End Function


Function GetDisp()
'
' 取得 safe node displacement 資料
'
' @returns GetDisp(Array)

    With WS_DISP
        rowStart = 2
        colStart = 1
        rowEnd = .Cells(Rows.Count, 1).End(xlUp).Row
        colEnd = 5

        GetDisp = .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd))
    End With

End Function


Function CombinedDispArray(dispArray)
'
' 合併 point 和 ASD，使之成為唯一ID。
' 取出 Z displacement
'
' @param dispArray(Array)
' @returns CombinedDispArray(Array)

    Dim combinedArray()

    dispLBound = LBound(dispArray, 1)
    dispUBound = UBound(dispArray, 1)

    ReDim combinedArray(dispLBound To dispUBound, 1 To 2)

    id = 1
    ASD = 4
    zDisp = 5

    For dispRow = dispLBound To dispUBound
        combinedArray(dispRow, 1) = dispArray(dispRow, id) & dispArray(dispRow, ASD)
        combinedArray(dispRow, 2) = dispArray(dispRow, zDisp)
    Next dispRow

    CombinedDispArray = combinedArray

End Function


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

    time0 = Timer

    Call PerformanceVBA(True)

    Set dictionary = CreateObject("Scripting.Dictionary")

    Set WS_LINE = Worksheets("Lines-v8")
    Set WS_DISP = Worksheets("Nodal Displacements-v8")

    lineArray = GetLines()
    dispArray = GetDisp()
    idAndLoadArray = CombinedDispArray(dispArray)

    lineLBound = LBound(lineArray, 1)
    lineUBound = UBound(lineArray, 1)
    idAndLoadLBound = LBound(idAndLoadArray, 1)
    idAndLoadUBound = UBound(idAndLoadArray, 1)

    ReDim result(lineLBound To lineUBound, 1 To 108)

    For idAndLoadRow = idAndLoadLBound To idAndLoadUBound
        If Not dictionary.Exists(idAndLoadArray(idAndLoadRow, 1)) Then
            Call dictionary.Add(idAndLoadArray(idAndLoadRow, 1), idAndLoadArray(idAndLoadRow, 2))
        End If
    Next idAndLoadRow

    For ASD = 1 To 36
        loadCombo = "ASD" & Format(ASD, "00")
        id1 = (ASD - 1) * 3 + 1
        id2 = (ASD - 1) * 3 + 2
        angleChange = (ASD - 1) * 3 + 3
        For lineRow = lineLBound To lineUBound
            id1AndLoad = lineArray(lineRow, 2) & loadCombo
            id2AndLoad = lineArray(lineRow, 3) & loadCombo
            result(lineRow, id1) = dictionary.Item(id1AndLoad)
            result(lineRow, id2) = dictionary.Item(id2AndLoad)
            result(lineRow, angleChange) = Abs(result(lineRow, id1) - result(lineRow, id2)) / lineArray(lineRow, 5)
        Next lineRow
    Next ASD

    rowStart = 4
    rowEnd = rowStart + lineUBound - 1
    colStart = 6
    colEnd = colStart + 108 - 1

    WS_LINE.Range(WS_LINE.Cells(rowStart, colStart), WS_LINE.Cells(rowEnd, colEnd)) = result

    Call FontSetting(WS_LINE)
    Call PerformanceVBA(False)
    Call ExecutionTimeVBA(time0)

End Sub
