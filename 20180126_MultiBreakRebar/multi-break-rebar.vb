Private WS_BEAM As Worksheet
Private WS_RESULT As Worksheet
Private OBJ_REBAR_SIZE_TO_DB As Object
Private ARR_INFO


Function GetRebarSizeToDb()
'
' 取得 rebar size area 資料
' 取代內建的 VLookup，效能大幅提升。
'
' @returns GetRebarSizeToDb(Object)

    Dim wsRebarSize As Worksheet
    Set wsRebarSize = Worksheets("Rebar Size")

    ' 取資料
    With wsRebarSize
        rowStart = 1
        colStart = 1
        rowEnd = .Cells(Rows.Count, 1).End(xlUp).Row
        colEnd = 10

        arrRebar = .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd))
    End With

    ' 設定 Dictionary
    Set objDictionary = CreateObject("Scripting.Dictionary")

    lbRebar = LBound(arrRebar, 1)
    ubRebar = UBound(arrRebar, 1)
    varSize = 1
    varDb = 7

    For rowRebar = lbRebar To ubRebar
        If Not objDictionary.Exists(arrRebar(rowRebar, varSize)) Then
            Call objDictionary.Add(arrRebar(rowRebar, varSize), arrRebar(rowRebar, varDb))
        End If
    Next rowRebar

    Set GetRebarSizeToDb = objDictionary

End Function


Function ClearBeforeOutputData()
'
' 清空前次輸出的資料
'

    WS_RESULT.Cells.Clear

End Function


Function CalRebarNumber(arrRawData)
'
' 計算上下排總支數
' 計算單排最大支數
'
' @param
' @returns

    Dim arrRebarNumber()
    Redim arrRebarNumber(1 To UBound(arrRawData), 1 To 3)

    rowStart = 1
    rowEnd = UBound(arrRawData)
    colStart = 6
    colEnd = 8

    ' 一二排相加
    For i = rowStart To rowEnd Step 2
        For j = colStart To colEnd

            ' 計算上下排總支數
            arrRebarNumber(i, j - 5) = Int(Split(arrRawData(i, j), "-")(0)) + Int(Split(arrRawData(i + 1, j), "-")(0))

            ' 計算單排最大支數
            arrRebarNumber(i + 1, j - 5) = Max(Int(Split(arrRawData(i, j), "-")(0)), Int(Split(arrRawData(i + 1, j), "-"))(0))

        Next
    Next

    CalRebarNumber = arrRebarNumber

End Function


Sub Main()
'
' @purpose:
' reduce 鋼筋量
'
'
' @algorithm:
' 上層筋由耐震控制
' 下層筋由重力與耐震共同控制
' 計算完成後加上延伸長度
'
' @test:
'
'
'

    time0 = Timer
    Call PerformanceVBA(True)

    ' Golobal Var
    Set WS_BEAM = Worksheets("小梁配筋")
    Set WS_RESULT = Worksheets("最佳化斷筋點")
    ARR_INFO = GetRangeToArray(Worksheets("General Information"), 2, 4, 4, 12)
    Set OBJ_REBAR_SIZE_TO_DB = GetRebarSizeToDb()

    Call ClearBeforeOutputData

    arrBeam = GetRangeToArray(WS_BEAM, 3, 1, 5, 16)
    arrRebarNumber = CalRebarNumber(arrBeam)
    arrMultiBreakRebar = CalMultiBreakPoint(arrRebarNumber)

    Call PrintResult(arrMultiBreakRebar)

    Call FontSetting(WS_RESULT)
    Call PerformanceVBA(False)
    Call ExecutionTimeVBA(time0)

End Sub
