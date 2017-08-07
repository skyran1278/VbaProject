Sub SuperSlab()

    Cells.Font.Name = "微軟正黑體"
    Cells.Font.Name = "Calibri"

    '// add declarations
    On Error GoTo catchError
exitSub:
    Exit Sub
catchError:
    '// add error handling
    GoTo exitSub
End Sub