Public NameListForWeb As Object, namecounterForWeb As Integer, GRsize As Double, Vsize As Double, maxspacingM As Double, maxspacingLNR As Double, netprotect As Double ''RC4.9.3用的到

Dim TOP_BAR_SIZE
Dim TOP_BAR_AREA
Dim TOP_DB

Sub WITH4_9_3()

    ' 增加可上下層選擇的功能
    Call AddTopBarSize

    Call Cal_S1
    Call Web_0124_S2
    Call RC4_9_3_S4
    Call rebardist_S3
    Call replaceRebar_S5
End Sub

' ---------------------------------------------------------------------------------------------------------------------

Function AddTopBarSize()
'
'
'
' @param
' @returns

    TOP_BAR_SIZE = InputBox("Please enter the size of TOP reinforcements in girder.", "Girder Reinforcement Size", "8")

    If TOP_BAR_SIZE = 3 Then
        TOP_BAR_AREA = 0.71
        TOP_DB = 0.95
    ElseIf TOP_BAR_SIZE = 4 Then
        TOP_BAR_AREA = 1.27
        TOP_DB = 1.27
    ElseIf TOP_BAR_SIZE = 5 Then
        TOP_BAR_AREA = 1.99
        TOP_DB = 1.59
    ElseIf TOP_BAR_SIZE = 6 Then
        TOP_BAR_AREA = 2.87
        TOP_DB = 1.91
    ElseIf TOP_BAR_SIZE = 7 Then
        TOP_BAR_AREA = 3.87
        TOP_DB = 2.22
    ElseIf TOP_BAR_SIZE = 8 Then
        TOP_BAR_AREA = 5.07
        TOP_DB = 2.54
    ElseIf TOP_BAR_SIZE = 9 Then
        TOP_BAR_AREA = 6.47
        TOP_DB = 2.87
    ElseIf TOP_BAR_SIZE = 10 Then
        TOP_BAR_AREA = 8.14
        TOP_DB = 3.22
    ElseIf TOP_BAR_SIZE = 11 Then
        TOP_BAR_AREA = 10.07
        TOP_DB = 3.58
    Else
        MsgBox ("請確認主筋尺寸")
    End If


End Function


Sub Cal_S1()
    Sheets("計算表").Select
    Application.ScreenUpdating = False



    Dim d As Double
    Dim Area As Double
    Dim cnt As Integer

    Range("E1") = "VArea cm^2/cm"
    Range("F1") = "BarSize"
    Range("G1") = "BarArea "
    Range("H1") = "TopNumber"
    Range("I1") = "BotNumber"
    Range("J1") = "BName"
    Range("K1") = "AdjBname"
    Range("L1") = "VSize"
    Range("M1") = "VBarArea"
    Range("N1") = "VSpacing"
    Range("O1") = "maxspacingM"
    Range("P1") = "maxspacingLNR"


    cnt = Range("A1").End(xlDown).Row

    GRsize = InputBox("Please enter the size of BOTTOM reinforcements in girder.", "Girder Reinforcement Size", "8")
    Vsize = InputBox("Please enter the size of reinforcements for shear.", "Shear Reinforcement Size", "4")
    maxspacingM = InputBox("Please enter the maxspacing for shear reinforcements in tie region.", "maxspacing for tie region", "25")
    maxspacingLNR = InputBox("Please enter the maxspacing for shear reinforcements in Conf. region.", "maxspacing for Conf region", "15")


    For i = 2 To cnt
        Cells(i, "F") = GRsize
    Next

    For i = 2 To cnt
        Cells(i, "L") = Vsize
    Next

    For i = 2 To cnt
        Cells(i, "O") = maxspacingM
    Next

    For i = 2 To cnt
        Cells(i, "P") = maxspacingLNR
    Next

    Dim GRArea As Double
    If GRsize = 3 Then
        GRArea = 0.71
    ElseIf GRsize = 4 Then
        GRArea = 1.27
    ElseIf GRsize = 5 Then
        GRArea = 1.99
    ElseIf GRsize = 6 Then
        GRArea = 2.87
    ElseIf GRsize = 7 Then
        GRArea = 3.87
    ElseIf GRsize = 8 Then
        GRArea = 5.07
    ElseIf GRsize = 9 Then
        GRArea = 6.47
    ElseIf GRsize = 10 Then
        GRArea = 8.14
    ElseIf GRsize = 11 Then
        GRArea = 10.07
    Else
    MsgBox ("請確認主筋尺寸")
    End If

    For i = 2 To cnt
       Cells(i, "G") = GRArea
    Next

    Dim VArea As Double
    If Vsize = 3 Then
    VArea = 0.71
    ElseIf Vsize = 4 Then
    VArea = 1.27
    ElseIf Vsize = 5 Then
    VArea = 1.99
    ElseIf Vsize = 6 Then
    VArea = 2.87
    ElseIf Vsize = 7 Then
    VArea = 3.87
    ElseIf Vsize = 8 Then
    VArea = 5.07
    ElseIf Vsize = 9 Then
    VArea = 6.47
    ElseIf Vsize = 10 Then
    VArea = 8.14
    ElseIf Vsize = 11 Then
    VArea = 10.07
    Else
    MsgBox ("請確認箍筋尺寸")
    End If

    For i = 2 To cnt
       Cells(i, "M") = VArea
    Next


    For i = 2 To cnt
       Cells(i, "G").Value = Application.WorksheetFunction.RoundUp(Cells(i, "G"), 2)
    Next

    For i = 2 To cnt
       Cells(i, "M").Value = Application.WorksheetFunction.RoundUp(Cells(i, "M"), 2)
    Next

    For i = 2 To cnt
        Cells(i, "H") = Cells(i, "C") / TOP_BAR_AREA
    Next


    For i = 2 To cnt
        Cells(i, "H").Value = Application.WorksheetFunction.RoundUp(Cells(i, "H"), 0)
        If Cells(i, "H") < 3 Then
            Cells(i, "H") = 3
        End If
    Next

    For i = 2 To cnt
        Cells(i, "I") = Cells(i, "D") / Cells(i, "G")
    Next


    For i = 2 To cnt
        Cells(i, "I").Value = Application.WorksheetFunction.RoundUp(Cells(i, "I"), 0)
        If Cells(i, "I") < 3 Then
            Cells(i, "I") = 3
        End If
    Next


'剪力筋根數 & 更改Vsize顯示 方便後續

    For i = 2 To cnt
    '''0需求
        If Cells(i, "E") = 0 Then
            If StrComp(Cells(i, "B"), "Middle") = 0 Then
                Cells(i, "N") = maxspacingM
            Else
                Cells(i, "N") = maxspacingLNR
            End If

'''''''有需求''''''''''check this equation
        Else
            Cells(i, "N") = Cells(i, "M") * 1 * 2 / Cells(i, "E")
        End If
''''''''''''''''''''''''''''''''''
    Next

    Dim looptime As Integer, done As Integer
    looptime = 0
    done = 0
    Dim Vsizeup() As Integer, VAreaup As Double
    ReDim Vsizeup(cnt - 2 + 1 - 1)

    For i = 0 To cnt - 2
    Vsizeup(i) = Vsize
    Next

    For i = 2 To cnt
        Cells(i, "L") = Chr(35) & Vsizeup(i - 2)
        '''SPACING <10 改雙箍 spacing過大改至上限值 其餘無條件進位至整數
        While (Cells(i, "N") < 10 And done <> 1) '先改雙箍
            Cells(i, "L") = "2" & Chr(35) & Vsizeup(i - 2)
            Cells(i, "M") = Cells(i, "M") * 2
            Cells(i, "N") = Cells(i, "M") * 1 / Cells(i, "E")
            If Cells(i, "N") < 10 Then '再升號數
                looptime = looptime + 1
                Vsizeup(i - 2) = Vsize + looptime

                If Vsizeup(i - 2) = 3 Then
                VAreaup = 0.71
                Cells(i, "L") = Chr(35) & Vsizeup(i - 2)
                Cells(i, "M") = VAreaup
                Cells(i, "N") = Cells(i, "M") * 1 * 2 / Cells(i, "E")
                ElseIf Vsizeup(i - 2) = 4 Then
                VAreaup = 1.27
                Cells(i, "L") = Chr(35) & Vsizeup(i - 2)
                Cells(i, "M") = VAreaup
                Cells(i, "N") = Cells(i, "M") * 1 * 2 / Cells(i, "E")
                ElseIf Vsizeup(i - 2) = 5 Then
                VAreaup = 1.99
                Cells(i, "L") = Chr(35) & Vsizeup(i - 2)
                Cells(i, "M") = VAreaup
                Cells(i, "N") = Cells(i, "M") * 1 * 2 / Cells(i, "E")
                ElseIf Vsizeup(i - 2) = 6 Then
                VAreaup = 2.87
                Cells(i, "L") = Chr(35) & Vsizeup(i - 2)
                Cells(i, "M") = VAreaup
                Cells(i, "N") = Cells(i, "M") * 1 * 2 / Cells(i, "E")
                Else
                Cells(i, "L") = Chr(35) & Vsizeup(i - 2)
                Cells(i, "M") = "號數過大"

                done = 1
                End If


            End If
        Wend
        done = 0
        looptime = 0
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''間距限制在max內
        If StrComp(Cells(i, "M"), "號數過大") <> 0 Then ''沒加會被STR轉ASCII
            If StrComp(Cells(i, "B"), "Middle") = 0 And Cells(i, "N") > maxspacingM Then
                Cells(i, "N") = maxspacingM
            ElseIf StrComp(Cells(i, "B"), "Middle") <> 0 And Cells(i, "N") > maxspacingLNR Then
                Cells(i, "N") = maxspacingLNR
            End If
        End If
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''間距限制在max內
        If StrComp(Cells(i, "M"), "號數過大") <> 0 Then ''沒加會被STR轉ASCII
            Cells(i, "N") = Application.WorksheetFunction.RoundUp(Cells(i, "N"), 0)
        End If
    Next
    '修正箍筋間距 僅留10,12,15,18,20,25
    For i = 2 To cnt
        If StrComp(Cells(i, "M"), "號數過大") <> 0 Then '沒排除會STR轉ASCII
            If Cells(i, "N") < 0 Then
                Cells(i, "N") = 25
            ElseIf Cells(i, "N") > 25 Then
                Cells(i, "N") = 25
            ElseIf Cells(i, "N") < 25 And Cells(i, "N") > 20 Then
                Cells(i, "N") = 20
            ElseIf Cells(i, "N") < 20 And Cells(i, "N") > 18 Then
                Cells(i, "N") = 18
            ElseIf Cells(i, "N") < 18 And Cells(i, "N") > 15 Then
                Cells(i, "N") = 15
            ElseIf Cells(i, "N") < 15 And Cells(i, "N") > 12 Then
                Cells(i, "N") = 12
            ElseIf Cells(i, "N") < 12 And Cells(i, "N") > 10 Then
                Cells(i, "N") = 10
            End If
        End If
    Next

     ' 以下將修改梁名
    Sheets("計算表").Select
    Dim RowNum(1) As Integer
    Dim FB As String
    FB = InputBox("Please enter the name of story.", "NAME", "FB")
    RowNum(0) = Range("A1").End(xlDown).Row
    For i = 2 To RowNum(0)
         name = Cells(i, 10)
         '開頭不一定是FB 若B1 B2 等等 則需再修改
         Cells(i, "K") = Chr(34) & FB & name & Chr(34) & Chr(40)
    Next

  Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------------------------------------------------------------

Sub Web_0124_S2()
    Sheets("RCAD").Select
    Application.ScreenUpdating = False
    Set NameListForWeb = CreateObject("Scripting.Dictionary")

    Dim RowNum As Integer
    RowNum = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row

    Dim ColNum As Integer
    ColNum = 200

    Dim k As Integer
    k = (RowNum + 2) / 17
    Dim NameL(2) As String
    '抓1.8+17c 的row 為名字與stir
    m = 0
    Dim g As Integer
    namecounterForWeb = 0
    Dim n As Integer


     '先抓樑名&SIZE 12/13 debug for diff. name length
 For R = 0 To k
     For c = 4 To ColNum
         If StrComp(Cells(1 + R * 17, c), "") <> 0 Then
            n = 0
            g = 0
            While n <> 1
                If Cells(1 + R * 17, c) <> 0 Then
                    NameL(g) = Cells(1 + R * 17, c).Value
                    c = c + 1
                    g = g + 1
                Else
                    n = 1
                    namecounterForWeb = namecounterForWeb + 1
                End If
            Wend
            NameListForWeb.Add (namecounterForWeb), Join(NameL, "")
        End If
     Next
 Next
    '抓到粱名了 再把SIZE隔離出來

Dim size As Integer
size = namecounterForWeb - 1
Dim b() As String '製造梁寬array
ReDim b(size)
Dim h() As String '梁深array
ReDim h(size)
Dim Wsize As Integer
Dim WArea As Double
Dim BS As Double
Dim FS As Double
FS = InputBox("Please enter the thickness of FS(cm).", "FS", "75")
BS = InputBox("Please enter the thickness of BS(cm).", "BS", "15")
Wsize = InputBox("Please enter the size of reinforcements for web.", "Web Reinforcement Size", "4")
    If Wsize = 3 Then
    WArea = 0.71
    ElseIf Wsize = 4 Then
    WArea = 1.27
    ElseIf Wsize = 5 Then
    WArea = 1.99
    ElseIf Wsize = 6 Then
    WArea = 2.87
    ElseIf Wsize = 7 Then
    WArea = 3.87
    ElseIf Wsize = 8 Then
    WArea = 5.07
    ElseIf Wsize = 9 Then
    WArea = 6.47
    ElseIf Wsize = 10 Then
    WArea = 8.14
    ElseIf Wsize = 11 Then
    WArea = 10.07
    Else
    MsgBox ("請確認腰筋尺寸")
    End If
Dim s2() As Double
ReDim s2(size)
Dim num() As Double
ReDim num(size)
For i = 1 To namecounterForWeb
    b(i - 1) = NameListForWeb.Item(i)
    h(i - 1) = NameListForWeb.Item(i)
    Dim position(2) As Integer
    position(0) = InStr(b(i - 1), "(")
    position(1) = InStr(b(i - 1), "*")
    position(2) = InStr(b(i - 1), ")")
    b(i - 1) = Mid(b(i - 1), position(0) + 1, position(1) - position(0) - 1) '個別的梁寬
    h(i - 1) = Mid(h(i - 1), position(1) + 1, position(2) - position(1) - 1) '個別的梁深
    s2(i - 1) = WArea * 2 / (0.0015 * b(i - 1)) '''不得大於d/5或30cm 目前寫為不大於d/5且不大於30cm
    Dim upper As Integer
    upper = 30
    If (h(i - 1) - netprotect) / 5 < 30 Then
        upper = (h(i - 1) - netprotect) / 5
    End If
    If s2(i - 1) > upper Then
        s2(i - 1) = upper
    End If
    num(i - 1) = Round(0.5 + (h(i - 1) - BS - FS - 10 - 10) / s2(i - 1)) + 1
Next



Dim PutNum As Integer
PutNum = 0

For R = 0 To k
    For c = 4 To ColNum
            If Cells(9 + R * 17, c) <> "" Then
                Cells(9 + R * 17, c) = num(PutNum) & "#" & Wsize
                PutNum = PutNum + 1
            End If
        Next
Next


'處理NameListForWeb , namecounterForWeb
Sheets("計算表").Select
Dim cnt As Integer
cnt = Range("A1").End(xlDown).Row
Dim name493() As String, width493() As String, depth493() As String, VAreaLower493() As Double
ReDim VAreaLower493(namecounterForWeb - 1)  ''個數要注意
ReDim depth493(namecounterForWeb - 1)
ReDim name493(namecounterForWeb - 1)
ReDim width493(namecounterForWeb - 1)
For i = 1 To namecounterForWeb
    name493(i - 1) = NameListForWeb(i)
    Dim pos493(2) As Integer
    pos493(0) = InStr(name493(i - 1), "(")
    pos493(1) = InStr(name493(i - 1), "*")
    pos493(2) = InStr(name493(i - 1), ")")
    width493(i - 1) = Mid(name493(i - 1), pos493(0) + 1, pos493(1) - pos493(0) - 1) '個別的寬度
    depth493(i - 1) = Mid(name493(i - 1), pos493(1) + 1, pos493(2) - pos493(1) - 1) '個別的深度
    name493(i - 1) = Left(name493(i - 1), pos493(0)) '此與Adjname相同
    For j = 2 To cnt '後綴個別深度與寬度
        If StrComp(Cells(j, "K"), name493(i - 1)) = 0 Then
            Cells(j, "Q") = width493(i - 1)
            Cells(j, "R") = depth493(i - 1)
        End If
    Next
    VAreaLower493(i - 1) = 0.0025 * width493(i - 1)
Next


Application.ScreenUpdating = True

End Sub

' ---------------------------------------------------------------------------------------------------------------------

Sub rebardist_S3()
 Application.ScreenUpdating = False
Sheets("計算表").Select
Dim cnt As Integer
cnt = Range("A1").End(xlDown).Row
Dim DistCoef As Double
    DistCoef = InputBox("請輸入主筋中心距。若為2db則輸入2，2.5db則輸入2.5。", "主筋間距限制", "2.5")
    netprotect = InputBox("請輸入淨保護層厚度(cm)。", "淨保護層厚度", "7.5")
Dim dv As Double
For i = 2 To cnt
    If Cells(i, "L") = "#3" Then
        dv = 0.95
    ElseIf Cells(i, "L") = "#4" Then
        dv = 1.27
    ElseIf Cells(i, "L") = "#5" Then
        dv = 1.59
    ElseIf Cells(i, "L") = "#6" Then
        dv = 1.91
    End If
Next
Dim db As Double
For i = 2 To cnt
    If Cells(i, "F") = 3 Then
        db = 0.95
    ElseIf Cells(i, "F") = 4 Then
        db = 1.27
    ElseIf Cells(i, "F") = 5 Then
        db = 1.59
    ElseIf Cells(i, "F") = 6 Then
        db = 1.91
    ElseIf Cells(i, "F") = 7 Then
        db = 2.22
    ElseIf Cells(i, "F") = 8 Then
        db = 2.54
    ElseIf Cells(i, "F") = 9 Then
        db = 2.87
    ElseIf Cells(i, "F") = 10 Then
        db = 3.22
    ElseIf Cells(i, "F") = 11 Then
        db = 3.58
    ElseIf Cells(i, "F") = 12 Then
        db = 3.81
    End If
Next
Dim netdist() As Double

ReDim netdist(cnt - 2)
Dim correct As Integer
correct = 0
For i = 2 To cnt '先把全部放在外層 內層放0
    Cells(i, "S") = Cells(i, "H")
    Cells(i, "T") = 0
    Cells(i, "U") = Cells(i, "I")
    Cells(i, "V") = 0
Next

For i = 2 To cnt '先做TOP
    correct = 0
    Do While correct <> 1
        If Cells(i, "J") = "" Then
            Exit Do
        End If
        netdist(i - 2) = (Cells(i, "Q") - 2 * netprotect - 2 * dv - TOP_DB) / (Cells(i, "S") - 1)
        If netdist(i - 2) < DistCoef * TOP_DB Then '若淨間距過小 則外面少放一根 裡面多放一根
            Cells(i, "S") = Cells(i, "S") - 1
            If Cells(i, "S") = 0 Then
                MsgBox ("有鋼筋擺不下喔!")
            End If
            Cells(i, "T") = Cells(i, "T") + 1
        Else
            correct = 1 '淨間距合格 離開while
        End If
    Loop
Next

For i = 2 To cnt '再做BOT
    correct = 0
    Do While correct <> 1
        If Cells(i, "J") = "" Then
            Exit Do
        End If
        netdist(i - 2) = (Cells(i, "Q") - 2 * netprotect - 2 * dv - db) / (Cells(i, "U") - 1)
        If netdist(i - 2) < DistCoef * db Then '若淨間距過小 則外面少放一根 裡面多放一根
            Cells(i, "U") = Cells(i, "U") - 1
            If Cells(i, "U") = 0 Then
                MsgBox ("有鋼筋擺不下喔!")
            End If
            Cells(i, "V") = Cells(i, "V") + 1
        Else
            correct = 1 '淨間距合格 離開while
        End If
    Loop
Next
'''完成淨間距檢核與主筋分層

''修正內層只放1根的分層 讓內層若有放至少放兩根
For i = 2 To cnt
    If Cells(i, "T") = 1 Then
        Cells(i, "T") = Cells(i, "T") + 1
        Cells(i, "S") = Cells(i, "S") - 1
    End If
    If Cells(i, "V") = 1 Then
        Cells(i, "V") = Cells(i, "V") + 1
        Cells(i, "U") = Cells(i, "U") - 1
    End If
Next
''修正內層只放1根的分層 讓內層若有放至少放兩根


 Application.ScreenUpdating = True

End Sub

' ---------------------------------------------------------------------------------------------------------------------

Sub RC4_9_3_S4()
 Application.ScreenUpdating = False
    Sheets("計算表").Select




    Dim d As Double
    Dim Area As Double
    Dim cnt As Integer



    cnt = Range("A1").End(xlDown).Row


    For i = 2 To cnt
        Cells(i, "F") = GRsize
    Next

    For i = 2 To cnt
        Cells(i, "L") = Vsize
    Next

    For i = 2 To cnt
        Cells(i, "O") = maxspacingM
    Next

    For i = 2 To cnt
        Cells(i, "P") = maxspacingLNR
    Next

    Dim GRArea As Double
    If GRsize = 3 Then
    GRArea = 0.71
    ElseIf GRsize = 4 Then
    GRArea = 1.27
    ElseIf GRsize = 5 Then
    GRArea = 1.99
    ElseIf GRsize = 6 Then
    GRArea = 2.87
    ElseIf GRsize = 7 Then
    GRArea = 3.87
    ElseIf GRsize = 8 Then
    GRArea = 5.07
    ElseIf GRsize = 9 Then
    GRArea = 6.47
    ElseIf GRsize = 10 Then
    GRArea = 8.14
    ElseIf GRsize = 11 Then
    GRArea = 10.07
    Else
    MsgBox ("請確認主筋尺寸")
    End If

    For i = 2 To cnt
       Cells(i, "G") = GRArea
    Next

    Dim VArea As Double
    If Vsize = 3 Then
    VArea = 0.71
    ElseIf Vsize = 4 Then
    VArea = 1.27
    ElseIf Vsize = 5 Then
    VArea = 1.99
    ElseIf Vsize = 6 Then
    VArea = 2.87
    ElseIf Vsize = 7 Then
    VArea = 3.87
    ElseIf Vsize = 8 Then
    VArea = 5.07
    ElseIf Vsize = 9 Then
    VArea = 6.47
    ElseIf Vsize = 10 Then
    VArea = 8.14
    ElseIf Vsize = 11 Then
    VArea = 10.07
    Else
    MsgBox ("請確認箍筋尺寸")
    End If

    For i = 2 To cnt
       Cells(i, "M") = VArea
    Next


    For i = 2 To cnt
       Cells(i, "G").Value = Application.WorksheetFunction.RoundUp(Cells(i, "G"), 2)
    Next

    For i = 2 To cnt
       Cells(i, "M").Value = Application.WorksheetFunction.RoundUp(Cells(i, "M"), 2)
    Next

    For i = 2 To cnt
        Cells(i, "H") = Cells(i, "C") / TOP_BAR_AREA
    Next


    For i = 2 To cnt
        Cells(i, "H").Value = Application.WorksheetFunction.RoundUp(Cells(i, "H"), 0)
        If Cells(i, "H") < 3 Then
            Cells(i, "H") = 3
        End If
    Next

    For i = 2 To cnt
        Cells(i, "I") = Cells(i, "D") / Cells(i, "G")
    Next


    For i = 2 To cnt
        Cells(i, "I").Value = Application.WorksheetFunction.RoundUp(Cells(i, "I"), 0)
        If Cells(i, "I") < 3 Then
            Cells(i, "I") = 3
        End If
    Next


'處理NameListForWeb , namecounterForWeb
Sheets("計算表").Select
Dim name493() As String, width493() As String, depth493() As String, VAreaLower493() As Double
ReDim VAreaLower493(namecounterForWeb - 1)  ''個數要注意
ReDim depth493(namecounterForWeb - 1)
ReDim name493(namecounterForWeb - 1)
ReDim width493(namecounterForWeb - 1)
For i = 1 To namecounterForWeb
    name493(i - 1) = NameListForWeb(i)
    Dim pos493(2) As Integer
    pos493(0) = InStr(name493(i - 1), "(")
    pos493(1) = InStr(name493(i - 1), "*")
    pos493(2) = InStr(name493(i - 1), ")")
    width493(i - 1) = Mid(name493(i - 1), pos493(0) + 1, pos493(1) - pos493(0) - 1) '個別的寬度
    depth493(i - 1) = Mid(name493(i - 1), pos493(1) + 1, pos493(2) - pos493(1) - 1) '個別的深度
    name493(i - 1) = Left(name493(i - 1), pos493(0)) '此與Adjname相同
    For j = 2 To cnt '後綴個別深度與寬度
        If StrComp(Cells(j, "K"), name493(i - 1)) = 0 Then
            Cells(j, "Q") = width493(i - 1)
            Cells(j, "R") = depth493(i - 1)
        End If
    Next

    VAreaLower493(i - 1) = 0.0025 * width493(i - 1)
Next
'處理NameListForWeb , namecounterForWeb
    '''0.0025*b*s<=VArea ---> Area/s>=Lower
    '''s<=d/5 and s<=30
Dim a As Double
Dim b As Double
'剪力筋根數 & 更改Vsize顯示 方便後續

    For i = 2 To cnt
    '''0需求
        If Cells(i, "E") = 0 Then
            ' 簡化
            For j = 1 To namecounterForWeb ''對每一根都做
                If Cells(i, "K") = name493(j - 1) Then ''名字對到的時候
                    Cells(i, "N") = VArea * 2 / VAreaLower493(j - 1) ''反求一個可以符合493敘述的間距並帶回
                End If
            Next
            ' ' Cells(i, "N") = VArea * 2 / VAreaLower493(j - 1) ' 簡化 20170712
            ' If StrComp(Cells(i, "B"), "Middle") = 0 Then
            '     Cells(i, "N") = maxspacingM
            '     ''''''493
            '     a = VArea / Cells(i, "N").Value ''面積除以間距
            '     For j = 1 To namecounterForWeb ''對每一根都做
            '         If Cells(i, "K") = name493(j - 1) Then ''名字對到的時候
            '             If VAreaLower493(j - 1) > a Then ''若鋼筋間距不符合493的敘述 (已經同除間距)
            '                 Cells(i, "N") = VArea * 2 / VAreaLower493(j - 1) ''反求一個可以符合493敘述的間距並帶回
            '             End If
            '             b = WorksheetFunction.Min((Cells(i, "R") - netprotect) / 5, 30) ''S不得大於b b是有效深度的五分之一 或三十公分
            '             If Cells(i, "N") > b Then ''如果S大於b
            '                 Cells(i, "N") = b ''縮到b
            '             End If
            '         End If
            '     Next
            '     ''''''493
            ' Else
            '     Cells(i, "N") = maxspacingLNR
            '     ''''''493
            '     a = VArea / Cells(i, "N").Value
            '     For j = 1 To namecounterForWeb
            '         If Cells(i, "K") = name493(j - 1) Then
            '             If VAreaLower493(j - 1) > a Then
            '                 Cells(i, "N") = VArea / VAreaLower493(j - 1)
            '             End If
            '             b = WorksheetFunction.Min((Cells(i, "R") - netprotect) / 5, 30)
            '             If Cells(i, "N") > b Then
            '                 Cells(i, "N") = b
            '             End If
            '         End If
            '     Next
            '     ''''''493
            ' End If

'''''''有需求''''''''''
        Else
            ' Cells(i, "N") = Cells(i, "M") * 2 / Cells(i, "E")
                ''''''493
                ' a = VArea  * 2 / Cells(i, "N").Value
                a = Cells(i, "E")
                For j = 1 To namecounterForWeb
                    If Cells(i, "K") = name493(j - 1) Then
                        If VAreaLower493(j - 1) > a Then
                            Cells(i, "N") = Cells(i, "M") * 2 / VAreaLower493(j - 1)
                        Else
                            Cells(i, "N") = Cells(i, "M") * 2 / a
                        End If
                        b = WorksheetFunction.Min((Cells(i, "R") - netprotect) / 5, 30)
                        If Cells(i, "N") > b Then
                            Cells(i, "N") = b
                        End If
                    End If
                Next
                ''''''493
        End If
''''''''''''''''''''''''''''''''''
    Next

    Dim looptime As Integer, done As Integer
    looptime = 0
    done = 0
    Dim Vsizeup() As Integer, VAreaup As Double
    ReDim Vsizeup(cnt - 2 + 1 - 1)
    For i = 0 To cnt - 2
    Vsizeup(i) = Vsize
    Next

    For i = 2 To cnt
        Cells(i, "L") = Chr(35) & Vsizeup(i - 2)
        '''SPACING <10 改雙箍 spacing過大改至上限值 其餘無條件進位至整數
        While (Cells(i, "N") < 10 And StrComp(Cells(i, "M"), "號數過大") <> 0)  '先改雙箍
            Cells(i, "L") = "2" & Chr(35) & Vsizeup(i - 2)
            Cells(i, "M") = Cells(i, "M") * 2
            If Cells(i, "E") <> 0 Then
                Cells(i, "N") = Cells(i, "M") * 1 * 2 / Cells(i, "E") '改雙箍之後間距的計算方法
            Else ''若有無剪力需求 SPACING卻小於10 代表是被鋼筋號數太小害的(A>=0.0025bws)
                VAreaup = Cells(i, "M") * 2
                For j = 1 To namecounterForWeb
                    If Cells(i, "K") = name493(j - 1) Then
                        Cells(i, "N") = VAreaup / VAreaLower493(j - 1) ''將間距提升到改雙箍後的上限間距
                    End If
                Next
            End If

            If Cells(i, "E") <> 0 Then
                ''''''493
                VAreaup = Cells(i, "M") * 2
                a = VAreaup / Cells(i, "N").Value
                For j = 1 To namecounterForWeb
                    If Cells(i, "K") = name493(j - 1) Then
                        If VAreaLower493(j - 1) > a Then
                            Cells(i, "N") = VAreaup / VAreaLower493(j - 1)
                        End If
                        b = WorksheetFunction.Min((Cells(i, "R") - netprotect) / 5, 30)
                        '''''netprotect should be user define
                        If Cells(i, "N") > b Then
                            Cells(i, "N") = b
                        End If
                    End If
                Next
                ''''''493
            End If

            If Cells(i, "N") < 10 Then '再升號數
                looptime = looptime + 1
                Vsizeup(i - 2) = Vsize + looptime

                If Vsizeup(i - 2) = 3 Then
                VAreaup = 0.71
                Cells(i, "M") = VAreaup
                ElseIf Vsizeup(i - 2) = 4 Then
                VAreaup = 1.27
                Cells(i, "M") = VAreaup
                ElseIf Vsizeup(i - 2) = 5 Then
                VAreaup = 1.99
                Cells(i, "M") = VAreaup
                ElseIf Vsizeup(i - 2) = 6 Then
                VAreaup = 2.87
                Cells(i, "M") = VAreaup
                Else
                Cells(i, "M") = "號數過大"
                End If

                Cells(i, "L") = "2" & Chr(35) & Vsizeup(i - 2)
                If StrComp(Cells(i, "M"), "號數過大") <> 0 Then
                    If Cells(i, "E") <> 0 Then
                        Cells(i, "N") = Cells(i, "M") * 1 * 2 / Cells(i, "E")
                    Else ''若有無剪力需求 SPACING卻小於10 代表是被鋼筋號數太小害的(A>=0.0025bws)
                        For j = 1 To namecounterForWeb
                            If Cells(i, "K") = name493(j - 1) Then
                                Cells(i, "N") = VAreaup / VAreaLower493(j - 1) ''將間距提升到改雙箍後的上限間距
                            End If
                        Next
                    End If
                    a = VAreaup / Cells(i, "N").Value
                    For j = 1 To namecounterForWeb
                        If Cells(i, "K") = name493(j - 1) Then
                            If VAreaLower493(j - 1) > a Then
                                Cells(i, "N") = VAreaup / VAreaLower493(j - 1)
                            End If
                            b = WorksheetFunction.Min((Cells(i, "R") - netprotect) / 5, 30)
                            If Cells(i, "N") > b Then
                                Cells(i, "N") = b
                            End If
                        End If
                    Next
                ''''''493
                End If

            End If
        Wend
        looptime = 0
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''間距限制在max內
        If StrComp(Cells(i, "M"), "號數過大") <> 0 Then ''沒排除會STR轉ASCII
            VAreaup = Cells(i, "M") * 2
            If StrComp(Cells(i, "B"), "Middle") = 0 And Cells(i, "N") > maxspacingM Then
                Cells(i, "N") = maxspacingM
                    ''''''493
                    a = VAreaup / Cells(i, "N").Value
                    For j = 1 To namecounterForWeb
                        If Cells(i, "K") = name493(j - 1) Then
                            If VAreaLower493(j - 1) > a Then
                                Cells(i, "N") = VAreaup / VAreaLower493(j - 1)
                            End If
                            b = WorksheetFunction.Min((Cells(i, "R") - netprotect) / 5, 30)
                            If Cells(i, "N") > b Then
                                Cells(i, "N") = b
                            End If
                        End If
                    Next
                    ''''''493
            ElseIf StrComp(Cells(i, "B"), "Middle") <> 0 And Cells(i, "N") > maxspacingLNR Then
                Cells(i, "N") = maxspacingLNR
                    ''''''493
                    a = VAreaup / Cells(i, "N").Value
                    For j = 1 To namecounterForWeb
                        If Cells(i, "K") = name493(j - 1) Then
                            If VAreaLower493(j - 1) > a Then
                                Cells(i, "N") = VAreaup / VAreaLower493(j - 1)
                            End If
                            b = WorksheetFunction.Min((Cells(i, "R") - netprotect) / 5, 30)
                            If Cells(i, "N") > b Then
                                Cells(i, "N") = b
                            End If
                        End If
                    Next
                    ''''''493
            End If
        End If
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''間距限制在max內

    Next
    '修正箍筋間距 僅留10,12,15,18,20,25
    For i = 2 To cnt
        If StrComp(Cells(i, "M"), "號數過大") <> 0 Then '沒排除會STR轉ASCII
            If Cells(i, "N") < 0 Then
                Cells(i, "N") = 25
            ElseIf Cells(i, "N") > 25 Then
                Cells(i, "N") = 25
            ElseIf Cells(i, "N") < 25 And Cells(i, "N") > 20 Then
                Cells(i, "N") = 20
            ElseIf Cells(i, "N") < 20 And Cells(i, "N") > 18 Then
                Cells(i, "N") = 18
            ElseIf Cells(i, "N") < 18 And Cells(i, "N") > 15 Then
                Cells(i, "N") = 15
            ElseIf Cells(i, "N") < 15 And Cells(i, "N") > 12 Then
                Cells(i, "N") = 12
            ElseIf Cells(i, "N") < 12 And Cells(i, "N") > 10 Then
                Cells(i, "N") = 10
            End If
        End If
    Next




     Application.ScreenUpdating = True



End Sub

' ---------------------------------------------------------------------------------------------------------------------

Sub replaceRebar_S5()
    Application.ScreenUpdating = False
    Sheets("RCAD").Select
    Dim name As String
    Dim num(3) As String
    Dim shearspacing As String
    Dim search As String
    Dim RowNum(1) As Integer
    Dim col As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer





    Sheets("計算表").Select
    RowNum(0) = Range("A1").End(xlDown).Row
    web = InputBox("Please enter the default value of web. Format: (number of bars)#(bar size) EX:3#4 or 4#5", "web", "3#4") '新增程式碼 腰筋


    '以下將梁名輸入nameList
    Dim NameList() As String
    ReDim NameList(RowNum(0) - 2)
    For i = 2 To RowNum(0)
        NameList(i - 2) = Cells(i, 11)
    Next

    '以下開始進行取代主筋、剪力筋、WEB

    col = 200
    Dim equal As Integer
    Dim t(2) As Integer
    Dim b(2) As Integer
    Dim s(2) As Integer
    Dim Vsize_S5 As String

    '先找資料底端
    Sheets("RCAD").Select
    RowNum(1) = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row

    '再清空2.3.5.6.7.8.9+17c 的row
    Dim c As Integer
    c = (RowNum(1) + 2) / 17
    For i = 0 To c
    Worksheets("RCAD").Range(Cells(2 + 17 * i, 4), Cells(9 + 17 * i, 300)).Clear
    Next


    Dim q As Integer
    Dim cnt As Integer
    cnt = 0

    Sheets("RCAD").Select
    For i = 1 To RowNum(1)
        q = 0
        cnt = 0
        For j = 1 To col
            For k = 0 To UBound(NameList)
                equal = StrComp(Cells(i, j), NameList(k))
                '上面的CELLS的I.J就是待會要用來抓貼上位置的參考點


                '''''''''''''''梁名長度不一之問題已解決
                If equal = 0 Then
                    If j > cnt Then
                        cnt = j
                        q = q + 1
                    End If

                '''''''''''''''梁名長度不一之問題已解決


                'q就是 該row "第幾根樑"
                '決定主筋剪力筋要貼的COL位置

                    t(0) = (1 + 1 + 1 + 7) + 13 + 1 + 20 * (q - 1)
                    t(1) = (1 + 1 + 1 + 7) + 13 + 1 + 3 + 20 * (q - 1)
                    t(2) = (1 + 1 + 1 + 7) + 13 + 1 + 3 + 3 + 20 * (q - 1)

                    b(0) = (1 + 1 + 1 + 7) + 13 + 1 + 20 * (q - 1)
                    b(1) = (1 + 1 + 1 + 7) + 13 + 1 + 3 + 20 * (q - 1)
                    b(2) = (1 + 1 + 1 + 7) + 13 + 1 + 3 + 3 + 20 * (q - 1)

                    s(0) = 16 + 8 * (q - 1)
                    s(1) = 18 + 8 * (q - 1)
                    s(2) = 20 + 8 * (q - 1)

                    W = 16 + 20 * (q - 1) '新增程式碼 腰筋
                '''''''''''''''''''''



                 ''''取代主筋剪力筋
                    Sheets("計算表").Select

                    If StrComp(Cells(k + 2, 2), "Start") = 0 Then
                        num(0) = Cells(k + 2, "S")
                        num(1) = Cells(k + 2, "U")
                        num(2) = Cells(k + 2, "T")
                        num(3) = Cells(k + 2, "V")
                        shearspacing = Cells(k + 2, "N")
                        Vsize_S5 = Cells(k + 2, "L")
                        Sheets("RCAD").Select

                        Cells(i + 8, W) = web '新增程式碼 腰筋

                        If num(0) > 10 Then
                            Cells(i + 1, t(0) - 1) = num(0)
                            t(1) = t(1) - 1
                            t(2) = t(2) - 1

                        Else
                            Cells(i + 1, t(0)) = num(0)
                        End If
                        If num(1) > 10 Then
                            Cells(i + 5, b(0) - 1) = num(1)
                            b(1) = b(1) - 1
                            b(2) = b(2) - 1
                        Else
                            Cells(i + 5, b(0)) = num(1)
                        End If


                        '''''t1t2的修正目前沒用
                        '''''因為修正之後K就換了 t1t2又被重製
                        '''''若t1t2沒修正會造成資料不能讀取 就必須在一個k內完成S M E 的鋼筋量覆寫
                        '''''就是若namelist(k) == namelist(k-1) 就 跳下一個k
                        '''''若沒影響讀取就不用修正了
                        If num(2) > 10 Then
                            Cells(i + 2, t(0) - 1) = num(2)
                            t(1) = t(1) - 1
                            t(2) = t(2) - 1
                        Else
                            Cells(i + 2, t(0)) = num(2)
                        End If
                        If num(3) > 10 Then
                            Cells(i + 4, b(0) - 1) = num(3)
                            b(1) = b(1) - 1
                            b(2) = b(2) - 1
                        Else
                            Cells(i + 4, b(0)) = num(3)
                        End If

                        Cells(i + 7, s(0)) = Vsize_S5 & Chr(64) & shearspacing
                    End If

                    If StrComp(Cells(k + 2, 2), "Middle") = 0 Then
                        num(0) = Cells(k + 2, "S")
                        num(1) = Cells(k + 2, "U")
                        num(2) = Cells(k + 2, "T")
                        num(3) = Cells(k + 2, "V")
                        shearspacing = Cells(k + 2, "N")
                        Vsize_S5 = Cells(k + 2, "L")
                        Sheets("RCAD").Select

                        Cells(i + 8, W) = web '新增程式碼 腰筋


                        If num(0) > 10 Then
                            Cells(i + 1, t(1) - 1) = num(0)
                            t(2) = t(2) - 1
                        Else
                            Cells(i + 1, t(1)) = num(0)
                        End If
                        If num(1) > 10 Then
                            Cells(i + 5, b(1) - 1) = num(1)
                            b(2) = b(2) - 1
                        Else
                            Cells(i + 5, b(1)) = num(1)
                        End If

                        If num(2) > 10 Then
                            Cells(i + 2, t(1) - 1) = num(2)
                            t(2) = t(2) - 1
                        Else
                            Cells(i + 2, t(1)) = num(2)
                        End If
                        If num(3) > 10 Then
                            Cells(i + 4, b(1) - 1) = num(3)
                            b(2) = b(2) - 1
                        Else
                            Cells(i + 4, b(1)) = num(3)
                        End If


                        Cells(i + 7, s(1)) = Vsize_S5 & Chr(64) & shearspacing
                    End If

                    If StrComp(Cells(k + 2, 2), "End") = 0 Then
                        num(0) = Cells(k + 2, "S")
                        num(1) = Cells(k + 2, "U")
                        num(2) = Cells(k + 2, "T")
                        num(3) = Cells(k + 2, "V")
                        shearspacing = Cells(k + 2, "N")
                        Vsize_S5 = Cells(k + 2, "L")
                        Sheets("RCAD").Select

                        Cells(i + 8, W) = web '新增程式碼 腰筋

                        If num(0) > 10 Then
                            Cells(i + 1, t(2) - 1) = num(0)
                        Else
                            Cells(i + 1, t(2)) = num(0)
                        End If
                        If num(1) > 10 Then
                             Cells(i + 5, b(2) - 1) = num(1)
                        Else
                            Cells(i + 5, b(2)) = num(1)
                        End If
                        If num(2) > 10 Then
                            Cells(i + 2, t(2) - 1) = num(2)
                        Else
                            Cells(i + 2, t(2)) = num(2)
                        End If
                        If num(3) > 10 Then
                             Cells(i + 4, b(2) - 1) = num(3)
                        Else
                            Cells(i + 4, b(2)) = num(3)
                        End If

                        Cells(i + 7, s(2)) = Vsize_S5 & Chr(64) & shearspacing
                    End If
                End If

            Next
        Next
    Next


'開始取代鋼筋號數
Dim sizeadj As String

Sheets("計算表").Select
sizeadj = Chr(35) & Cells(2, 6)

Sheets("RCAD").Select

' Search for "#".
Dim SearchChar As String
SearchChar = Chr(35)

Dim TestPos As Integer

SearchString = "1"

'searching, casesensitive

For i = 1 To RowNum(1)
    TestPos = InStr(Cells(i, 3), SearchChar)
    If TestPos <> 0 Then
        Cells(i, 3) = sizeadj
    End If
Next
For i = 1 To RowNum(1)
    TestPos = InStr(Cells(i, 2), SearchChar)
    If TestPos <> 0 Then
        Cells(i, 2) = sizeadj
    End If
Next


Application.ScreenUpdating = True




End Sub



