Sub Main()

' * 目的
'       從多個力量中提取三段區間的力量輸出

' * 環境
'       Excel

' * 輸出入格式
'       輸入：Beam_Forces Frame_Assignments_Summary
'       輸出：Combo

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

    Call WriteDownCombo(shearForces, moment, "Combo")

    shearForces = ConcatArray(shearForces)
    moment = ConcatArray(moment)
    shearForces = OneShearForces(shearForces)

    Call WriteDownCombo(shearForces, moment, "Summary")


    MsgBox "Execution Time " & Application.Round((Timer - Time0) / 60, 2) & " Min", vbOKOnly

End Sub

Function ConcatArray(havaToConcatArray)

' 合併第一欄與第二欄陣列

    ' Dim haveConcatArray()
    ' uBoundhavaToConcatArray = UBound(havaToConcatArray)
    ' ReDim haveSplitArray(2 To uBoundhavaToConcatArray,)

    i = 2

    Do While havaToConcatArray(i, 1) <> Empty

        havaToConcatArray(i,3) = havaToConcatArray(i,1) & havaToConcatArray(i,2)
        i = i + 1

    Loop

    ' ReDim preserve haveSplitArray(2 To uBoundhavaToConcatArray, 3 to 9)

    ConcatArray = havaToConcatArray

End Function

Function SplitArray(havaToSplitArray)

' 依據逗號分割字串

    Dim haveSplitArray()
    uBoundAllBeamForces = UBound(havaToSplitArray)
    ReDim haveSplitArray(2 To uBoundAllBeamForces)

    i = 2

    Do While havaToSplitArray(i, 1) <> Empty

        haveSplitArray(i) = Split(havaToSplitArray(i, 1), ",")
        havaToSplitArray(i, 1) = haveSplitArray(i)(0)
        havaToSplitArray(i, 2) = haveSplitArray(i)(1)
        havaToSplitArray(i, 3) = haveSplitArray(i)(2)
        i = i + 1

    Loop

    SplitArray = havaToSplitArray

End Function



Sub WriteDownCombo(shearForces, moment, workSheetName)

' 寫入Combo

    Worksheets(workSheetName).Activate
    Column = 6
    Range(Columns(6), Columns(21)).ClearContents

    Cells(1, Column) = "Shear Forces"
    Cells(2, Column) = "Story"
    Cells(2, Column + 1) = "Beam"
    Cells(2, Column + 2) = "Load"
    Cells(2, Column + 3) = "Left"
    Cells(2, Column + 4) = "Middle"
    Cells(2, Column + 5) = "Right"

    Cells(1, Column + 7) = "Moment"
    Cells(2, Column + 7) = "Story"
    Cells(2, Column + 8) = "Beam"
    Cells(2, Column + 9) = "Load"
    Cells(2, Column + 10) = "Left"
    Cells(2, Column + 11) = "Middle"
    Cells(2, Column + 12) = "Right"
    Cells(2, Column + 13) = "Left"
    Cells(2, Column + 14) = "Middle"
    Cells(2, Column + 15) = "Right"

    Range(Cells(3, Column), Cells(UBound(shearForces), Column + 5)) = shearForces
    Range(Cells(3, Column + 7), Cells(UBound(moment), Column + 15)) = moment

End Sub



Function MaxForces(lowerLength, upperLength, length, forces, priorMaxForces)

' 取最大值

    If length > lowerLength And length < upperLength And forces > priorMaxForces Then
        MaxForces = forces
    Else
        MaxForces = priorMaxForces
    End If

End Function



Function MinForces(lowerLength, upperLength, length, forces, priorMinForces)

' 取最小值

    If length > lowerLength And length < upperLength And forces < priorMinForces Then
        MinForces = forces
    Else
        MinForces = priorMinForces
    End If

End Function

Function OneMoment(moment)

' 取M3

    momentNumber = 2
    Dim moment()
    momentRowUsed = UBound(allBeamForces)
    ReDim moment(2 To momentRowUsed, 2 To 9)

    For allBeamForcesNumber = 2 To momentRowUsed - 1

        For LMR = 4 To 6
            moment(beamShearForcesNumber, LMR) = MaxValue(allBeamForces(allBeamForcesNumber, LMR), moment(beamShearForcesNumber, LMR))
        Next

        For LMR = 7 To 9
            moment(beamShearForcesNumber, LMR) = MinValue(allBeamForces(allBeamForcesNumber, LMR), moment(beamShearForcesNumber, LMR))
        Next

        If BooleanSame(allBeamForces(allBeamForcesNumber, 1), allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber + 1, 1), allBeamForces(allBeamForcesNumber + 1, 2)) Then
            moment(momentNumber, 2) = allBeamForces(allBeamForcesNumber, 1)
            moment(momentNumber, 3) = allBeamForces(allBeamForcesNumber, 2)
            momentNumber = momentNumber + 1
        End If
    Next

    OneMoment = moment()

End Function

Function CatchMoment(allBeamForces)

' 取M3

    momentNumber = 2
    Dim moment()
    momentRowUsed = UBound(allBeamForces)
    ReDim moment(2 To momentRowUsed, 1 To 9)

    For allBeamForcesNumber = 2 To momentRowUsed - 1

        moment(momentNumber, 4) = MaxForces(0, 1 / 3, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 4))
        moment(momentNumber, 5) = MaxForces(1 / 4, 3 / 4, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 5))
        moment(momentNumber, 6) = MaxForces(2 / 3, 1, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 6))
        moment(momentNumber, 7) = MinForces(0, 1 / 3, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 7))
        moment(momentNumber, 8) = MinForces(1 / 4, 3 / 4, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 8))
        moment(momentNumber, 9) = MinForces(2 / 3, 1, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 4), moment(momentNumber, 9))

        If allBeamForces(allBeamForcesNumber, 1) <> allBeamForces(allBeamForcesNumber + 1, 1) Then
            moment(momentNumber, 1) = allBeamForces(allBeamForcesNumber, 1)
            momentNumber = momentNumber + 1
        End If
    Next

    CatchMoment = moment()

End Function

' Function MaxValue(shearForces, beamShearForcesNumber, allBeamForces, allBeamForcesNumber)

' ' 取最大值


'     If allBeamForces(allBeamForcesNumber, LMR) > shearForces(beamShearForcesNumber, LMR) Then
'         shearForces(beamShearForcesNumber, LMR) = allBeamForces(allBeamForcesNumber, LMR)
'     End If

'     MaxValue = shearForces()

' End Function

Function MaxValue(value, priorMaxValue)

' 取最大值

    If value > priorMaxValue Then
        MaxValue = value
    Else
        MaxValue = priorMaxValue
    End If

End Function

Function BooleanSame(allBeamForces01, allBeamForces02, allBeamForces11, allBeamForces12)

' 判斷一二欄合起來是否相同

    BooleanSame = allBeamForces01 & allBeamForces02 <> allBeamForces11 & allBeamForces12

End Function

Function OneShearForces(allBeamForces)

' 多個載重組合取一個

    beamShearForcesNumber = 2
    Dim shearForces()
    beamForcesRowUsed = UBound(allBeamForces)
    ReDim shearForces(2 To beamForcesRowUsed, 2 To 6)

    For allBeamForcesNumber = 2 To beamForcesRowUsed - 1

        For LMR = 4 To 6
            shearForces(beamShearForcesNumber, LMR) = MaxValue(allBeamForces(allBeamForcesNumber, LMR), shearForces(beamShearForcesNumber, LMR))
        Next

        If BooleanSame(allBeamForces(allBeamForcesNumber, 1), allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber + 1, 1), allBeamForces(allBeamForcesNumber + 1, 2)) Then
            shearForces(beamShearForcesNumber, 2) = allBeamForces(allBeamForcesNumber, 1)
            shearForces(beamShearForcesNumber, 3) = allBeamForces(allBeamForcesNumber, 2)
            beamShearForcesNumber = beamShearForcesNumber + 1
        End If
    Next

    OneShearForces = shearForces()

End Function

Function CatchShearForces(allBeamForces)

' 取V2

    beamShearForcesNumber = 2
    Dim shearForces()
    beamForcesRowUsed = UBound(allBeamForces)
    ReDim shearForces(2 To beamForcesRowUsed, 1 To 6)

    For allBeamForcesNumber = 2 To beamForcesRowUsed - 1

        shearForces(beamShearForcesNumber, 4) = MaxForces(0, 1 / 3, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 3), shearForces(beamShearForcesNumber, 4))
        shearForces(beamShearForcesNumber, 5) = MaxForces(1 / 4, 3 / 4, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 3), shearForces(beamShearForcesNumber, 5))
        shearForces(beamShearForcesNumber, 6) = MaxForces(2 / 3, 1, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 3), shearForces(beamShearForcesNumber, 6))

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
            allBeamForces(i, 1) = Cells(i, 1) & "," & Cells(i, 2) & "," & Cells(i, 3)

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