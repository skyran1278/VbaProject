Sub CheckColumnNorm()

    Dim Column As ColumnClass
    Set Column = New ColumnClass

    Column.GetData ("柱配筋")

    ' 沒有資料就跳出
    If Column.NoData Then
        Exit Sub
    End If

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

End Sub

Sub CheckGroundBeamNorm()

    Dim GroundBeam As BeamClass
    Set GroundBeam = New BeamClass

    GroundBeam.GetData ("地梁")

    GroundBeam.Initialize

    ' 實作規範
    GroundBeam.Norm4_9_3
    GroundBeam.Norm4_9_4
    GroundBeam.EconomicNorm4_9_4
    GroundBeam.SafetyRebarRatioAndSpace
    GroundBeam.SafetyRebarRatioForGB
    GroundBeam.EconomicBotRebarRelativeForGB
    GroundBeam.EconomicTopRebarRelativeForGB
    GroundBeam.SafetyStirrupSpace

    GroundBeam.PrintMessage

    GroundBeam.CountRebarNumber
    GroundBeam.PrintRebarRatio

End Sub

Sub CheckBeamNorm()

    Dim Beam As BeamClass
    Set Beam = New BeamClass

    Beam.GetData ("小梁")

    Beam.Initialize

    ' 實作規範
    Beam.Norm3_6
    Beam.Norm3_7_5
    Beam.Norm13_5_1AndSafetyRebarNumber
    Beam.SafetyRebarRatioForSB
    Beam.SafetyLoad

    Beam.PrintMessage

    Beam.CountRebarNumber
    Beam.PrintRebarRatio

End Sub

Sub CheckGirderNorm()

    Dim ran As New UTILS_CLASS
    Dim Girder As New BeamClass

    Call ran.ExecutionTime(True)
    Call ran.PerformanceVBA(True)

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

    Girder.PrintMessage

    Girder.CountRebarNumber
    ' Girder.PrintRebarRatio

    Call ran.PerformanceVBA(False)
    Call ran.ExecutionTime(False)

End Sub
