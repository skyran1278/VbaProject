Sub main()
'
'
    position = Array("A3:AG35", "AK3:BQ35", "BU3:DA35", "DE3:EK35", "EO3:FU35", "A43:AG75", "AK43:BQ75", "BU43:DA75", "DE43:EK75", "EO43:FU75", "A83:AG115", "AK83:BQ115", "BU83:DA115", "DE83:EK115", "EO83:FU115", "A123:AG155", "AK123:BQ155", "BU123:DA155", "DE123:EK155", "EO123:FU155", "A163:AG195", "AK163:BQ195", "BU163:DA195", "DE163:EK195", "EO163:FU195", "A203:AG235", "AK203:BQ235", "BU203:DA235", "DE203:EK235", "EO203:FU235", "A243:AG275", "AK243:BQ275", "BU243:DA275", "DE243:EK275", "EO243:FU275", "A283:AG315", "AK283:BQ315", "BU283:DA315", "DE283:EK315", "EO283:FU315", "A323:AG355", "AK323:BQ355", "BU323:DA355", "DE323:EK355", "EO323:FU355", "A363:AG395", "AK363:BQ395", "BU363:DA395", "DE363:EK395", "EO363:FU395")

    For worksheets_index = 5 To 10 Step 5
        For position_index = LBound(position) To UBound(position) Step 1
            Call capture(worksheets_index, position_index, position(position_index))
        Next position_index
    Next worksheets_index

End Sub

Sub capture(worksheets_index, position_index, position)
'
'

    Sheets(worksheets_index).Activate

    Set Rng = Range(position)

    Rng.CopyPicture

    Name = Split(ThisWorkbook.Name, ".xlsm")(0)

    Set co = Sheets(worksheets_index).ChartObjects.Add(0, 0, Rng.Width, Rng.Height)

    co.Activate

    With co
        .Chart.Paste
        .Chart.Export ThisWorkbook.Path & "\\" & Name & "-" & ActiveSheet.Name & "-" & position_index & ".JPEG"
        .Delete
    End With

    Application.CutCopyMode = False

End Sub
