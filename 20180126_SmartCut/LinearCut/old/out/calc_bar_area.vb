Function calc_bar_area(ByVal cell)
'
' calc bar area.
'
' @since 1.0.0
' @param {range} [cell] to calc rebar area cell.
' @return {number} [area] rebar area.
' @see dependencies
'
  tmp = Split(cell, "-")

  If tmp(0) = "0" Then
    calc_bar_area = 0

  Else
    Set dict = CreateObject("Scripting.Dictionary")

    dict.Add "#2", 0.32258
    dict.Add "#3", 0.709676
    dict.Add "#4", 1.29032
    dict.Add "#5", 1.999996
    dict.Add "#6", 2.838704
    dict.Add "#7", 3.87096
    dict.Add "#8", 5.096764
    dict.Add "#9", 6.4516
    dict.Add "#10", 8.193532
    dict.Add "#11", 10.0645

    num = tmp(0)
    Size = tmp(1)

    calc_bar_area = num * dict(Size)

  End If

End Function

