Sub Main()

' * 目的
'       從多個力量中提取三段區間的力量輸出

' * 環境
'       Excel

' * 輸出入格式
'       輸入：Beam_Forces Frame_Assignments_Summary
'       輸出：Summary

' * 執行時間
'       1.08 Min

' * 輸出結果的精確度與檢驗方式


    Time0 = Timer

    length = CatchBeamLength()

    allBeamForces = CatchBeamForces(length)

    shearForces = CatchShearForces(allBeamForces)
    moment = CatchMoment(allBeamForces)

    shearForces = SplitArray(shearForces)
    moment = SplitArray(moment)

    Call WriteDownSummery(shearForces, moment)

    MsgBox "Execution Time " & Application.Round((Timer - Time0) / 60, 2) & " Min", vbOKOnly

End Sub



Function SplitArray(havaToSplitArray)

' 依據空白分割字串

    Dim haveSplitArray()
    uBoundAllBeamForces = UBound(havaToSplitArray)
    ReDim haveSplitArray(2 To uBoundAllBeamForces)

    i = 2

    Do While havaToSplitArray(i, 1) <> Empty

        haveSplitArray(i) = Split(havaToSplitArray(i, 1))
        havaToSplitArray(i, 1) = haveSplitArray(i)(0)
        havaToSplitArray(i, 2) = haveSplitArray(i)(1)
        havaToSplitArray(i, 3) = haveSplitArray(i)(2)
        i = i + 1

    Loop

    SplitArray = havaToSplitArray

End Function



Sub WriteDownSummery(shearForces, moment)

' 寫入Summery

    Worksheets("Summary").Activate
    Column = 6
    Range(Columns(6), Columns(21)).ClearContents

    Cells(1, Column) = "Story"
    Cells(1, Column + 1) = "Beam"
    Cells(1, Column + 2) = "Load"
    Cells(1, Column + 3) = "Left"
    Cells(1, Column + 4) = "Middle"
    Cells(1, Column + 5) = "Right"

    Cells(1, Column + 7) = "Story"
    Cells(1, Column + 8) = "Beam"
    Cells(1, Column + 9) = "Load"
    Cells(1, Column + 10) = "Left"
    Cells(1, Column + 11) = "Middle"
    Cells(1, Column + 12) = "Right"
    Cells(1, Column + 13) = "Left"
    Cells(1, Column + 14) = "Middle"
    Cells(1, Column + 15) = "Right"

    Range(Cells(2, Column), Cells(UBound(shearForces), Column + 5)) = shearForces
    Range(Cells(2, Column + 7), Cells(UBound(moment), Column + 15)) = moment

End Sub



Function MaxBeamShearForces(lowerLength, upperLength, length, shearForces, priorMaxShearForces)

' 取最大值

    If length > lowerLength And length < upperLength And shearForces > priorMaxShearForces Then
        MaxBeamShearForces = shearForces
    Else
        MaxBeamShearForces = priorMaxShearForces
    End If

End Function



Function MaxMoment(lowerLength, upperLength, length, moment, priorMaxMoment)

' 取最大值

    If length > lowerLength And length < upperLength And priorMaxMoment = Empty Then
        MaxMoment = moment
    ElseIf length > lowerLength And length < upperLength And moment > priorMaxMoment Then
        MaxMoment = moment
    Else
        MaxMoment = priorMaxMoment
    End If

End Function



Function MinMoment(lowerLength, upperLength, length, moment, priorMinMoment)

' 取最小值

    If length > lowerLength And length < upperLength And priorMinMoment = Empty Then
        MinMoment = moment
    ElseIf length > lowerLength And length < upperLength And moment < priorMinMoment Then
        MinMoment = moment
    Else
        MinMoment = priorMinMoment
    End If

End Function



Function CatchMoment(allBeamForces)

' 取M3

    momentNumber = 2
    Dim moment()
    momentRowUsed = UBound(allBeamForces)
    ReDim moment(2 To momentRowUsed, 1 To 9)

    For allBeamForcesNumber = 2 To momentRowUsed - 1

        moment(momentNumber, 4) = MaxMoment(0, 1 / 3, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 4))
        moment(momentNumber, 5) = MaxMoment(1 / 4, 3 / 4, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 5))
        moment(momentNumber, 6) = MaxMoment(2 / 3, 1, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 6))
        moment(momentNumber, 7) = MinMoment(0, 1 / 3, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 7))
        moment(momentNumber, 8) = MinMoment(1 / 4, 3 / 4, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 8))
        moment(momentNumber, 9) = MinMoment(2 / 3, 1, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 9))

        If allBeamForces(allBeamForcesNumber, 1) <> allBeamForces(allBeamForcesNumber + 1, 1) Then
            moment(momentNumber, 1) = allBeamForces(allBeamForcesNumber, 1)
            momentNumber = momentNumber + 1
        End If
    Next

    CatchMoment = moment()

End Function



Function CatchShearForces(allBeamForces)

' 取V2

    beamShearForcesNumber = 2
    Dim shearForces()
    beamForcesRowUsed = UBound(allBeamForces)
    ReDim shearForces(2 To beamForcesRowUsed, 1 To 6)

    For allBeamForcesNumber = 2 To beamForcesRowUsed - 1

        shearForces(beamShearForcesNumber, 4) = MaxBeamShearForces(0, 1 / 3, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 3), shearForces(beamShearForcesNumber, 4))
        shearForces(beamShearForcesNumber, 5) = MaxBeamShearForces(1 / 4, 3 / 4, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 3), shearForces(beamShearForcesNumber, 5))
        shearForces(beamShearForcesNumber, 6) = MaxBeamShearForces(2 / 3, 1, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 3), shearForces(beamShearForcesNumber, 6))

        If allBeamForces(allBeamForcesNumber, 1) <> allBeamForces(allBeamForcesNumber + 1, 1) Then
            shearForces(beamShearForcesNumber, 1) = allBeamForces(allBeamForcesNumber, 1)
            beamShearForcesNumber = beamShearForcesNumber + 1
        End If
    Next

    CatchShearForces = shearForces()

End Function



Function PercentageOfLoc(i, length)

' 計算長度百分比

    PercentageOfLoc = Cells(i, 4) / Application.VLookup(Cells(i, 1) & Cells(i, 2), length, 2, False)

End Function



Function CatchBeamForces(length)

' 抓力量

    Worksheets("Beam_Forces").Activate
    Dim allBeamForces()
    beamForcesRowUsed = Cells(Rows.Count, 1).End(xlUp).Row

    ' + 1 是為了準備之後遇到邊界值的問題
    ReDim allBeamForces(2 To beamForcesRowUsed + 1, 1 To 4)

    For i = 2 To beamForcesRowUsed

            ' "1FB1"
            ' allBeamForces(i, 1) = Cells(i, 1) & Cells(i, 2)

            ' "1FB1DL"
            allBeamForces(i, 1) = Cells(i, 1) & " " & Cells(i, 2) & " " & Cells(i, 3)

            ' Absolute Loc
            ' allBeamForces(i, 3) = Cells(i, 4)

            ' Percentage Loc
            ' Debug.Print Application.VLookup(allBeamForces(i, 1), length, 2, False)
            allBeamForces(i, 2) = PercentageOfLoc(i, length)

            ' V2
            allBeamForces(i, 3) = Abs(Cells(i, 6))

            ' M3
            allBeamForces(i, 4) = Cells(i, 10)
    Next

    CatchBeamForces = allBeamForces()

End Function



Function CatchBeamLength()

' 抓總長度

    Worksheets("Frame_Assignments_Summary").Activate
    Dim length()
    lengthRowUsed = Cells(Rows.Count, 1).End(xlUp).Row
    ReDim length(2 To lengthRowUsed, 1 To 2)

    For i = 2 To lengthRowUsed

            ' "1FB1"
            length(i, 1) = Cells(i, 1) & Cells(i, 2)

            ' Length
            length(i, 2) = Cells(i, 4)
    Next

    CatchBeamLength = length()


End Function