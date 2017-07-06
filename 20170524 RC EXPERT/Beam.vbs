Sub CheckGroundBeamNorm()

    Dim GroundBeam As BeamClass
    Set GroundBeam = New BeamClass

    GroundBeam.GetData ("地梁配筋")

    GroundBeam.Initialize

    GroundBeam.Norm4_9_3
    GroundBeam.Norm4_9_4
    GroundBeam.EconomicNorm4_9_4
    GroundBeam.SafetyRebarRatioAndSpace
    GroundBeam.SafetyRebarRatioGB
    GroundBeam.EconomicBotMidRelativeEnd
    GroundBeam.EconomicTopEndRelativeMid
    GroundBeam.SafetyStirrupSpace
    GroundBeam.CountRebarNumber

    GroundBeam.PrintMessage

End Sub

Sub CheckBeamNorm()

    Dim Beam As BeamClass
    Set Beam = New BeamClass

    Beam.GetData ("小梁配筋")

    Beam.Initialize

    Beam.Norm3_6
    Beam.Norm3_7_5
    Beam.Norm13_5_1AndSafetyRebarNumber
    Beam.SafetyRebarRatioSB
    Beam.SafetyLoad
    Beam.CountRebarNumber

    Beam.PrintMessage

End Sub

Sub CheckGirderNorm()

    Dim Girder As BeamClass
    Set Girder = New BeamClass

    Girder.GetData ("大梁配筋")

    Girder.Initialize

    Girder.Norm3_6
    Girder.Norm3_7_5
    Girder.Norm3_8_1
    Girder.Norm4_6_7_9
    Girder.Norm13_5_1AndSafetyRebarNumber
    Girder.Norm15_4_2_1
    Girder.Norm15_4_2_2
    Girder.SafetyStirrupSpace
    Girder.EconomicTopMidRelativeEnd
    Girder.CountRebarNumber

    Girder.PrintMessage

End Sub

Function ExecutionTime(Time0)

    If Timer - Time0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - Time0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - Time0) / 60, 2) & " Min", vbOKOnly
    End If

End Function

Sub Main()
'
' * 目的
'       Check Norm

' * 環境
'       Excel

' * 輸出入格式
'       輸入：
'       輸出：

' * 執行時間
'       0.06 Sec

' * 輸出結果的精確度與檢驗方式
'

    Time0 = Timer

    Call CheckGirderNorm
    Call CheckBeamNorm
    Call CheckGroundBeamNorm

    Call ExecutionTime(Time0)

End Sub
