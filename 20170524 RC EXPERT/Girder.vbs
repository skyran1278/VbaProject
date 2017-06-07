Sub Main()
'
' * 目的
'       check Girder Norm
'       Norm3_6
'       Norm15_4_2_1
'       Norm15_4_2_2
'       NormMiddleNoMoreThanEndEightyPercentage
'       Norm13_5_1
'       RebarAmountNoBelowTwo
'       StirrupSpacingMoreThan10AndLessThan30
'       Norm4_6_7_9
'       Norm3_8_1

' * 環境
'       Excel

' * 輸出入格式
'       輸入：大梁配筋
'       輸出：大梁配筋 Message

' * 執行時間
'       0.06 Sec

' * 輸出結果的精確度與檢驗方式
'

    Time0 = Timer

    Dim Girder As BeamClass
    Set Girder = New BeamClass

    Girder.GetData ("大梁配筋")

    Girder.Initialize

    ' Girder.RatioData (Girder.RAW_DATA)
    Girder.Norm3_6
    Girder.Norm15_4_2_1
    Girder.Norm15_4_2_2
    Girder.NormMiddleNoMoreThanEndEightyPercentage
    Girder.Norm13_5_1AndRebarAmountNoBelowTwo
    Girder.StirrupSpacingMoreThan10AndLessThan30
    Girder.Norm4_6_7_9
    Girder.Norm3_8_1

    Girder.PrintMessage


    If Timer - Time0 < 60 Then
        MsgBox "Execution Time " & Application.Round((Timer - Time0), 2) & " Sec", vbOKOnly
    Else
        MsgBox "Execution Time " & Application.Round((Timer - Time0) / 60, 2) & " Min", vbOKOnly
    End If

End Sub
