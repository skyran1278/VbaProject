Sub TestSortRawData()
'
'

    On Error GoTo ErrorHandler

    Dim Beam As New BeamClass

    Call Beam.Initialize("列數不合")

    Exit Sub

ErrorHandler:
    Call Beam.PrintError

End Sub
