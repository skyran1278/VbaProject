Function Ld_cal(B, Cv, Fyh, fc, dh, s, db, N)
'This function calculates the develope length of flexural rebar used in Girder/Beam.
'It is used for nominal concrete in case of phi_e=1.0 & phi_t=1.0.
'Reference:土木401-93

'B   : Girder/Beam Width (cm)
'Cv  : Clear Cover THK (cm)
'Fyh : Stirrup Nominal Yielding Streng (kgf/cm2)
'fc  : 28-days concrete compressive strength (kgf/cm2)
'dh  : stirrup diameter (mm)
's   : space of stirrup at lap location of flexural rebar (cm)
'db  : diameter of flexural rebar (mm)
'N   : numbers of flexural rebar at lap location

    db = db / 10  'change unit from mm to cm
    dh = dh / 10  'change unit from mm to cm

    If (fc) ^ 0.5 > 26.5 Then fc = 700                                 '5.2.2

    Cc = dh + Cv                                                       'R5.3.4.1.1
    Cs = (B - db * N - dh * 2 - Cv * 2) / 2 / (N - 1)                  'R5.3.4.1.1

    If Cs > Cc Then                                                    'Vertical splitting failure
        cb = db / 2 + Cc
        Ktr = (3.14159 * dh ^ 2 / 4) * Fyh / 105 / s                    'R5.3.4.1.2
    Else                                                               'Horizontal splitting failure
        cb = db / 2 + Cs
        Ktr = 2 * (3.14159 * dh ^ 2 / 4) * Fyh / 105 / s / N            'R5.3.4.1.2
    End If

    If ((cb + Ktr) / db) <= 2.5 Then                                   '5.3.4.1
        Ld_cal = 0.28 * 4200 / (fc) ^ 0.5 * db / ((cb + Ktr) / db)
    Else
        Ld_cal = 0.28 * 4200 / (fc) ^ 0.5 * db / 2.5
    End If

    If db < 2.2 Then Ld_cal = 0.8 * Ld_cal                             'phi_s factor

    If Ld_cal < 30 Then Ld_cal = 30                                    '5.3.1

End Function
