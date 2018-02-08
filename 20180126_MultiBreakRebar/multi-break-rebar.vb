Private OBJ_REBAR_AREA As Object


Function GetRawData()
'
' 取得 beam rebar 資料
'
' @returns GetRawData(Array)

    Set wsBeam = Worksheets("小梁配筋")

    With wsBeam
        rowStart = 1
        colStart = 1
        rowEnd = .Cells(Rows.Count, 1).End(xlUp).Row
        colEnd = 16

        GetRawData = .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd))
    End With

End Function


Function GetRebarArea()
'
' 取得 rebar size area 資料
' 取代內建的 VLookup，效能大幅提升。
'
' @returns GetRebarArea(Object)

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
    varArea = 10

    For rowRebar = lbRebar To ubRebar
        If Not objDictionary.Exists(arrRebar(rowRebar, varSize)) Then
            Call objDictionary.Add(arrRebar(rowRebar, varSize), arrRebar(rowRebar, varArea))
        End If
    Next rowRebar

    Set GetRebarArea = objDictionary

End Function


Function GetRatioData(arrRawData)
'
'
'
' @param
' @returns

    Dim arrOnlyRationData()
    Redim arrOnlyRationData(1 To UBound(arrRawData), 1 To 1)

    arrRatioData = arrRawData

    rowStart = 3
    rowEnd = UBound(arrRawData)
    colStart = 6
    colEnd = 8

    ' 計算鋼筋面積
    For i = rowStart To rowEnd
        For j = colStart To colEnd
            arrRatioData(i, j) = CalRebarArea(arrRawData(i, j))
        Next
    Next

    ' 一二排截面積相加
    For i = rowStart To rowEnd Step 2
        For j = colStart To colEnd
            arrOnlyRationData(i, j) = arrRatioData(i, j) + arrRatioData(i + 1, j)
        Next
    Next

    GetRatioData = arrOnlyRationData

End Function


Function CalRebarArea(rebar)

    tmp = Split(rebar, "-")

    ' 排除為 0 的狀況
    If tmp(0) = 0 Then
        CalRebarArea = 0
    Else
        numberOfRebar = tmp(0)
        rebarSize = tmp(1)

        ' 轉換鋼筋尺寸為截面積
        rebarArea = OBJ_REBAR_AREA.Item(rebarSize)

        CalRebarArea = numberOfRebar * rebarArea

    End If

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
'
' @test:
'
'
'

    Set OBJ_REBAR_AREA = GetRebarArea()

    arrBeam = GetData()
    arrRatioData = GetRatioData()

End Sub
