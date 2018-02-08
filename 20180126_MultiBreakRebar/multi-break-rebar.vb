Function GetData()
'
' 取得 beam rebar 資料
'
' @returns GetData(Array)

    Set wsBeam = Worksheets("小梁配筋")

    With wsBeam
        rowStart = 1
        colStart = 1
        rowEnd = .Cells(Rows.Count, 1).End(xlUp).Row
        colEnd = 16

        GetData = .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd))
    End With

End Function


Function GetRebarSize()
'
' 取得 rebar size 資料
'
' @returns GetData(Array)

    Dim wsRebarSize As Worksheet
    Set wsRebarSize = Worksheets("Rebar Size")

    With wsRebarSize
        rowStart = 1
        colStart = 1
        rowEnd = .Cells(Rows.Count, 1).End(xlUp).Row
        colEnd = 10

        GetRebarSize = .Range(.Cells(rowStart, colStart), .Cells(rowEnd, colEnd))
    End With

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

    beam = GetData()
    rebarSize = GetRebarSize()

End Sub
