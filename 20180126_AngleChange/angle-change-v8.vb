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
    idAndLoadLBound = LBound(idAndLoadArray, 1)
    idAndLoadUBound = UBound(idAndLoadArray, 1)

    ReDim result(angleLBound To angleUBound, 1 To 108)

    For dataRow = idAndLoadLBound To idAndLoadUBound
        If Not dictionary.Exists(idAndLoadArray(dataRow, 1)) Then
            Call dictionary.Add(idAndLoadArray(dataRow, 1), idAndLoadArray(dataRow, 2))
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
            result(angleRow, id1) = dictionary.Item(id1AndLoad)
            result(angleRow, id2) = dictionary.Item(id2AndLoad)
            result(angleRow, angleChange) = Abs(result(angleRow, id1) - result(angleRow, id2)) / angleArray(angleRow, 5)
        Next angleRow
    Next ASD

    rowStart = 4
    rowEnd = rowStart + angleUBound - 1
    colStart = 6
    colEnd = colStart + 108 - 1

    WS_ANGLE.Range(WS_ANGLE.Cells(rowStart, colStart), WS_ANGLE.Cells(rowEnd, colEnd)) = result

    Call FontSetting
    Call PerformanceVBA(False)
    Call ExecutionTimeVBA(time0)

End Sub
