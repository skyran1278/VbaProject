Private WS_ANGLE, WS_DATA


Sub Macro1()
'
' Macro1 Macro

 Set sht = ActiveSheet
'Set sht = Worksheets("P9851")
total_no = Application.WorksheetFunction.CountA(sht.Range("A:A")) ' '點數

total_no2 = Application.WorksheetFunction.CountA(Worksheets("data").Range("A:A")) ' '點數


  For k = 1 To 36  '36組CHECK
    For i = 4 To total_no

        chk = Cells(2, (k - 1) * 3 + 6)
        For j = 2 To total_no2 '需依輸入檔修改
            If Worksheets("data").Cells(j, 4) = chk And Worksheets("data").Cells(j, 1) = Cells(i, 2) Then
                Cells(i, (k - 1) * 3 + 6) = Worksheets("data").Cells(j, 5)
            End If
            If Worksheets("data").Cells(j, 4) = chk And Worksheets("data").Cells(j, 1) = Cells(i, 3) Then
                Cells(i, (k - 1) * 3 + 7) = Worksheets("data").Cells(j, 5)
            End If
        Next
        Cells(i, (k - 1) * 3 + 8) = Abs(Cells(i, (k - 1) * 3 + 6) - Cells(i, (k - 1) * 3 + 7)) / Cells(i, 5)

    Next

  Next


End Sub


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

    For data = dataLBound To dataUBound
        combinedArray(data, 1) = dataArray(data, 1) & dataArray(data, 4)
        combinedArray(data, 2) = dataArray(data, 5)
    Next data

    CombinedData = combinedArray

    ' Application.Index(combinedArray, , 1) = Application.Index(dataArray, , 1) & Application.Index(dataArray, , 4)
    ' Application.Index(combinedArray, , 2) = Application.Index(dataArray, , 5)

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

    time0 = Timer

    Set WS_ANGLE = Worksheets("Angle")
    Set WS_DATA = Worksheets("Data")

    angleArray = GetAngle()
    dataArray = GetData()
    idAndLoadArray = CombinedData(dataArray)

    angleLBound = LBound(angleArray, 1)
    angleUBound = UBound(angleArray, 1)

    ReDim result(angleLBound To angleUBound, 1 To 108)

    For ASD = 1 To 36

        loadCombo = "ASD" & Format(ASD, "00")
        id1 = (ASD - 1) * 3 + 1
        id2 = (ASD - 1) * 3 + 2
        angleChange = (ASD - 1) * 3 + 3
        For angleRow = angleLBound To angleUBound
            id1AndLoad = angleArray(angleRow, 2) & loadCombo
            id2AndLoad = angleArray(angleRow, 3) & loadCombo
            result(angleRow, id1) = application.VLookup(id1AndLoad, idAndLoadArray, 2, false)
            result(angleRow, id2) = application.VLookup(id2AndLoad, idAndLoadArray, 2, false)
            result(angleRow, angleChange) = abs(result(angleRow, id1) - result(angleRow, id2)) / angleArray(angleRow, 5)
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
