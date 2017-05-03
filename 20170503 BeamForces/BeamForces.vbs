Sub Main()

    ' Debug.Print Now
    ' Application.SendKeys "^g ^a {DEL}"
    Time0 = Timer



    length = CatchBeamLength()

    allBeamForces = CatchBeamForces(length)

    Range(Cells(2, 12), Cells(UBound(allBeamForces), 15)) = allBeamForces

    beamShearForces = CatchShearForces(allBeamForces)



    MsgBox "執行時間 " & Timer - Time0 & " 秒", vbOKOnly
End Sub

Function MaxBeamShearForces(lowerLength, upperLength, length, shearForces, priorMaxShearForces)

    If length > lowerLength And length < upperLength And shearForces > priorMaxShearForces Then
        MaxBeamShearForces = shearForces
    else
        MaxBeamShearForces = priorMaxShearForces
    End If

End Function

Function CatchShearForces(allBeamForces)

    uBoundAllBeamForces = UBound(allBeamForces)
    beamShearForcesNumber = 2
    Dim beamShearForces()
    beamForcesRowUsed = Cells(Rows.Count, 1).End(xlUp).Row
    ReDim beamShearForces(2 To beamForcesRowUsed, 1 To 4)

    For allBeamForcesNumber = 2 To beamForcesRowUsed - 1

        beamShearForces(beamShearForcesNumber, 2) = MaxBeamShearForces(0, 1 / 3, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 3), beamShearForces(beamShearForcesNumber, 2))
        beamShearForces(beamShearForcesNumber, 3) = MaxBeamShearForces(1 / 4, 3 / 4, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 3), beamShearForces(beamShearForcesNumber, 3))
        beamShearForces(beamShearForcesNumber, 4) = MaxBeamShearForces(2 / 3, 1, allBeamForces(allBeamForcesNumber, 2), allBeamForces(allBeamForcesNumber, 3), beamShearForces(beamShearForcesNumber, 4))

        If allBeamForces(allBeamForcesNumber, 1) <> allBeamForces(allBeamForcesNumber + 1, 1) Then
            beamShearForces(beamShearForcesNumber, 1) = allBeamForces(allBeamForcesNumber, 1)
            beamShearForcesNumber = beamShearForcesNumber + 1
        End If
    Next

    CatchShearForces = beamShearForces()

End Function

Function PercentageOfLoc(i, length)

    PercentageOfLoc = Cells(i, 4) / Application.VLookup(Cells(i, 1) & Cells(i, 2), length, 2, False)

End Function

Function CatchBeamForces(length)

    Worksheets("Beam_Forces").Activate
    Dim allBeamForces()
    beamForcesRowUsed = Cells(Rows.Count, 1).End(xlUp).Row
    ReDim allBeamForces(2 To beamForcesRowUsed, 1 To 4)

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




