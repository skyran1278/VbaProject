Option Explicit

Sub Pretreatment()
    Dim Time0#, BeamPositionColumn As Integer
    Time0 = Timer
    Application.ScreenUpdating = False
    
    '排列組合產生表格
    Worksheets("輸入").Activate
    Worksheets("輸入").Range(Rows(29), Rows(1000)).ClearContents
    SortTable 4
    CountNumberTable 4
    
    Application.ScreenUpdating = True
    MsgBox "執行時間 " & Timer - Time0 & " 秒", vbOKOnly
End Sub

Sub SortTable(BeamPositionColumn)
    Dim WsInput As Worksheet, I As Integer, J As Integer
    Set WsInput = Worksheets("輸入")

    Cells.HorizontalAlignment = xlCenter
    Cells.Font.Name = "微軟正黑體"

    For J = 1 To 10
        If BeamPositionColumn < 39 Then
            For I = 1 To 6
                Select Case I
                    Case 2, 3
                        WsInput.Range(Cells(6, BeamPositionColumn + I), Cells(100, BeamPositionColumn + I)).NumberFormatLocal = """D""0"
                    Case Else
                        WsInput.Range(Cells(6, BeamPositionColumn + I), Cells(100, BeamPositionColumn + I)).NumberFormatLocal = "G/通用格式"
                End Select
                WsInput.Range(Cells(6, BeamPositionColumn + I), Cells(28, BeamPositionColumn + I)).Sort Key1:=Range(Cells(6, BeamPositionColumn + I), Cells(28, BeamPositionColumn + I)), _
                order1:=xlAscending
            Next
            BeamPositionColumn = BeamPositionColumn + 7
        Else
            For I = 1 To 5
                Select Case I
                    Case 1
                        WsInput.Range(Cells(6, BeamPositionColumn + I), Cells(100, BeamPositionColumn + I)).NumberFormatLocal = """D""0"
                    Case Else
                        WsInput.Range(Cells(6, BeamPositionColumn + I), Cells(100, BeamPositionColumn + I)).NumberFormatLocal = "G/通用格式"
                End Select
                WsInput.Range(Cells(6, BeamPositionColumn + I), Cells(28, BeamPositionColumn + I)).Sort Key1:=Range(Cells(6, BeamPositionColumn + I), Cells(28, BeamPositionColumn + I)), _
                order1:=xlAscending
            Next
            BeamPositionColumn = BeamPositionColumn + 6
        End If
    Next

End Sub

Sub CountNumberTable(BeamPositionColumn)
    Dim lastrow(4) As Integer, Diameter As Double, TieDiameter As Double, Fyt As Double
    Dim Concrete As Double, TieSpacing As Double, CountNumber As Double
    Dim Fy As Double, Cover As Double, RebarSpacing As Double
    Dim I As Integer, a As Integer, b As Integer, c As Integer, d As Integer, e As Integer
    Dim CountAllTable As Integer, CountFirstTable As Integer, CountLastTable As Integer

    For CountFirstTable = 1 To 5
        For I = 0 To 4
            lastrow(I) = Cells(Rows.Count, I + BeamPositionColumn + 2).End(xlUp).Row - 5
        Next
        CountNumber = 0
        For a = 1 To lastrow(0)
            Diameter = Cells(a + 5, BeamPositionColumn + 2)
            For b = 1 To lastrow(1)
                TieDiameter = Cells(b + 5, BeamPositionColumn + 3)
                For c = 1 To lastrow(2)
                    TieSpacing = Cells(c + 5, BeamPositionColumn + 4)
                    For d = 1 To lastrow(3)
                        Fyt = Cells(d + 5, BeamPositionColumn + 5)
                        For e = 1 To lastrow(4)
                            Concrete = Cells(e + 5, BeamPositionColumn + 6)
                            CountNumber = CountNumber + 1
                            Cells(29 + CountNumber, BeamPositionColumn) = CountNumber
                            Cells(29 + CountNumber, BeamPositionColumn + 2) = Diameter
                            Cells(29 + CountNumber, BeamPositionColumn + 3) = TieDiameter
                            Cells(29 + CountNumber, BeamPositionColumn + 4) = TieSpacing
                            Cells(29 + CountNumber, BeamPositionColumn + 5) = Fyt
                            Cells(29 + CountNumber, BeamPositionColumn + 6) = Concrete
                        Next
                    Next
                Next
            Next
        Next
        BeamPositionColumn = BeamPositionColumn + 7
    Next
    
    For CountLastTable = 1 To 5
        For I = 0 To 4
            lastrow(I) = Cells(Rows.Count, I + BeamPositionColumn + 1).End(xlUp).Row - 5
        Next
        CountNumber = 0
        For a = 1 To lastrow(0)
            Diameter = Cells(a + 5, BeamPositionColumn + 1)
            For b = 1 To lastrow(1)
                Cover = Cells(b + 5, BeamPositionColumn + 2)
                For c = 1 To lastrow(2)
                    RebarSpacing = Cells(c + 5, BeamPositionColumn + 3)
                    For d = 1 To lastrow(3)
                        Fy = Cells(d + 5, BeamPositionColumn + 4)
                        For e = 1 To lastrow(4)
                            Concrete = Cells(e + 5, BeamPositionColumn + 5)
                            CountNumber = CountNumber + 1
                            Cells(29 + CountNumber, BeamPositionColumn) = CountNumber
                            Cells(29 + CountNumber, BeamPositionColumn + 1) = Diameter
                            Cells(29 + CountNumber, BeamPositionColumn + 2) = Cover
                            Cells(29 + CountNumber, BeamPositionColumn + 3) = RebarSpacing
                            Cells(29 + CountNumber, BeamPositionColumn + 4) = Fy
                            Cells(29 + CountNumber, BeamPositionColumn + 5) = Concrete
                        Next
                    Next
                Next
            Next
        Next
        BeamPositionColumn = BeamPositionColumn + 6
    Next

End Sub



