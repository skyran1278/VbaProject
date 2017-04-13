Option Explicit

Sub OnlyNeedToRunOnce()
    Dim Time0#, Number As Integer
    Time0 = Timer
    Application.ScreenUpdating = False
    
    '複製表格
    CopyAndPaste 5
    CopyAndPaste 6
    CopyAndPaste 7
    
    Application.ScreenUpdating = True
    MsgBox "執行時間 " & Timer - Time0 & " 秒", vbOKOnly
End Sub

Sub CopyAndPaste(Number)
    Dim TableNumber As Integer, TableRowNumber  As Integer, I  As Integer
    Worksheets(Number).Activate
    Worksheets(Number).Range(Cells(11, 3), Cells(11, 34)).Copy
    TableNumber = Worksheets(8).Cells(13, 18)
    TableRowNumber = Worksheets(8).Cells(14, 18)

    For I = 12 To TableRowNumber
        Worksheets(Number).Cells(I, 3).Select
        ActiveSheet.Paste
    Next
    
    Range(Columns(1), Columns(35)).Copy
    
    For I = 1 To TableNumber
        Columns(I * 35 + 1).Select
        ActiveSheet.Paste
    Next
    
End Sub

