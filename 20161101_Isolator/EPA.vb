Function BS(Damping)
'

    If Damping <= 0.02 Then
        BS = 0.8

    ElseIf 0.02 < Damping And Damping <= 0.05 Then
        BS = (Damping - 0.02) / 0.03 * (1 - 0.8) + 0.8

    ElseIf 0.05 < Damping And Damping <= 0.1 Then
        BS = (Damping - 0.05) / 0.05 * (1.33 - 1) + 1

    ElseIf 0.1 < Damping And Damping <= 0.2 Then
        BS = (Damping - 0.1) / 0.1 * (1.6 - 1.33) + 1.33

    ElseIf 0.2 < Damping And Damping <= 0.3 Then
        BS = (Damping - 0.2) / 0.1 * (1.79 - 1.6) + 1.6

    ElseIf 0.3 < Damping And Damping <= 0.4 Then
        BS = (Damping - 0.3) / 0.1 * (1.87 - 1.79) + 1.79

    ElseIf 0.4 < Damping And Damping <= 0.5 Then
        BS = (Damping - 0.4) / 0.1 * (1.93 - 1.87) + 1.87

    ElseIf 0.5 < Damping Then
        BS = 1.93

    End If

End Function

Function B1(Damping)
'

    If Damping <= 0.02 Then
        B1 = 0.8

    ElseIf 0.02 < Damping And Damping <= 0.05 Then
        B1 = (Damping - 0.02) / 0.03 * (1 - 0.8) + 0.8

    ElseIf 0.05 < Damping And Damping <= 0.1 Then
        B1 = (Damping - 0.05) / 0.05 * (1.25 - 1) + 1

    ElseIf 0.1 < Damping And Damping <= 0.2 Then
        B1 = (Damping - 0.1) / 0.1 * (1.5 - 1.25) + 1.25

    ElseIf 0.2 < Damping And Damping <= 0.3 Then
        B1 = (Damping - 0.2) / 0.1 * (1.63 - 1.5) + 1.5

    ElseIf 0.3 < Damping And Damping <= 0.4 Then
        B1 = (Damping - 0.3) / 0.1 * (1.7 - 1.63) + 1.63

    ElseIf 0.4 < Damping And Damping <= 0.5 Then
        B1 = (Damping - 0.4) / 0.1 * (1.75 - 1.7) + 1.7

    ElseIf 0.5 < Damping Then
        B1 = 1.75

    End If

End Function

Function EPA(Sa, Teff, SDS, T0D, Damping)
'

    T0 = T0D * BS(Damping) / B1(Damping)

    If Teff <= 0.2 * T0 Then
        EPA = Sa * BS(Damping) / (1 + 3 * Teff / 0.4 / T0D)

    ElseIf Teff < T0 Then
        EPA = Sa * BS(Damping) / 2.5

    ElseIf Teff < 2.5 * T0 Then
        EPA = Sa * B1(Damping) * Teff / (2.5 * T0D)

    ElseIf 2.5 * T0 < Teff Then
        EPA = Sa * B1(Damping) / (2.5 * T0D / Teff)

    Else
        EPA = 10000

    End If

End Function
