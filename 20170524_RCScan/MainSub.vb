Sub ScanColumnNorm()

    Dim ran As New UTILS_CLASS
    Dim Column As New ColumnClass

    Call ran.ExecutionTime(True)
    ' Call ran.PerformanceVBA(True)

    Column.GetData ("柱配筋")

    ' 沒有資料就跳出
    ' If Column.NoData Then
    '     Exit Sub
    ' End If

    Column.Initialize

    ' 實作規範
    Column.EconomicSmooth
    Column.Norm15_5_4_1
    Column.EconomicTopStoryRebar

    ' FIXME: Function Name
    Column.Norm15_5_4_100

    Column.PrintMessage
    Column.PrintRebarRatio

    Column.CountRebarNumber
    Column.PrintRebarRatioInAnotherSheets

    ' Call ran.PerformanceVBA(False)
    Call ran.ExecutionTime(False)

End Sub

Sub ScanFoundationBeamNorm()
'
' TODO:處理筏基版厚、上下版厚
'

    Dim ran As New UTILS_CLASS
    Dim FoundationBeam As New BeamClass
    ' Set FoundationBeam = New BeamClass

    Call ran.ExecutionTime(True)
    ' Call ran.PerformanceVBA(True)

    FoundationBeam.Initialize("地梁")

    ' 實作規範
    FoundationBeam.Norm4_9_3
    FoundationBeam.Norm4_9_4
    FoundationBeam.EconomicNorm4_9_4
    FoundationBeam.SafetyRebarRatioAndSpace
    FoundationBeam.SafetyRebarRatioForGB
    FoundationBeam.EconomicBotRebarRelativeForGB
    FoundationBeam.EconomicTopRebarRelativeForGB
    FoundationBeam.SafetyStirrupSpace

    FoundationBeam.PrintResult

    ' FoundationBeam.CountRebarNumber
    ' FoundationBeam.PrintRebarRatio

    ' Call ran.PerformanceVBA(False)
    Call ran.ExecutionTime(False)

End Sub

Sub ScanBeamNorm()

    Dim ran As New UTILS_CLASS
    Dim Beam As New BeamClass

    Call ran.ExecutionTime(True)
    ' Call ran.PerformanceVBA(True)

    Beam.Initialize("小梁")

    ' 實作規範
    Beam.Norm3_6
    Beam.Norm3_7_5
    Beam.Norm13_5_1AndSafetyRebarNumber
    Beam.SafetyRebarRatioForSB
    Beam.SafetyLoad

    Beam.PrintResult

    ' Beam.CountRebarNumber
    ' Beam.PrintRebarRatio

    ' Call ran.PerformanceVBA(False)
    Call ran.ExecutionTime(False)

End Sub

Sub ScanGirderNorm()

    Dim ran As New UTILS_CLASS
    Dim Girder As New BeamClass

    Call ran.ExecutionTime(True)
    ' Call ran.PerformanceVBA(True)

    Girder.Initialize("大梁")

    ' 實作規範
    Girder.Norm3_6
    Girder.Norm3_7_5
    Girder.Norm3_8_1
    Girder.Norm4_6_7_9
    Girder.Norm13_5_1AndSafetyRebarNumber

    Girder.Norm15_4_2_1
    Girder.Norm15_4_2_2
    Girder.SafetyStirrupSpace
    Girder.EconomicTopRebarRelative

    Girder.PrintResult

    ' Girder.CountRebarNumber
    ' Girder.PrintRebarRatio

    ' Call ran.PerformanceVBA(False)
    Call ran.ExecutionTime(False)

End Sub
