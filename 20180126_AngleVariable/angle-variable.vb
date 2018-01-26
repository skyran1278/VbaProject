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
