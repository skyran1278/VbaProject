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

    Dim Column As ColumnClass
    Set Column = New ColumnClass

    Column.GetData ("柱配筋")

    Column.Initialize

    Column.EconomicSmooth
    Column.EconomicSmooth

    Column.PrintMessage

    Call ExecutionTime(Time0)

End Sub