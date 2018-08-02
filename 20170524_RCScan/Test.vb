Sub TestSortRawData()
'
' TODO:處理筏基版厚、上下版厚
'

    Dim ran As New UTILS_CLASS
    Dim Beam As New BeamClass
    ' Set Beam = New BeamClass

    Call ran.ExecutionTime(True)
    Call ran.PerformanceVBA(True)

    On Error GoTo ErrorHandler

    Call Beam.Initialize("SortRawData")

    ' 實作規範
    Beam.Norm4_9_3
    Beam.Norm4_9_4
    Beam.EconomicNorm4_9_4
    Beam.SafetyRebarRatioAndSpace
    Beam.SafetyRebarRatioForGB
    Beam.EconomicBotRebarRelativeForGB
    Beam.EconomicTopRebarRelativeForGB
    Beam.SafetyStirrupSpace

    Beam.PrintResult

    ' Beam.CountRebarNumber
    ' Beam.PrintRebarRatio

    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTime(False)

    Exit Sub

ErrorHandler:
    Call ran.PerformanceVBA(False)
    Call Beam.PrintError(Err.NUMBER, Err.Source, Err.Description)

End Sub
