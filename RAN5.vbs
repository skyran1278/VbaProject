Option Explicit

Sub RemarkText()
    Dim WsExplanation As Worksheet, FirstRowNumber As Long
    Set WsExplanation = Worksheets(8)
    FirstRowNumber = 28
    'Application.ScreenUpdating = False
    If Len(WsExplanation.Cells(FirstRowNumber, 9)) <> 0 Then
        WsExplanation.Cells(FirstRowNumber, 9) = ""
        WsExplanation.Cells(FirstRowNumber + 1, 9) = ""
        WsExplanation.Cells(FirstRowNumber + 2, 9) = ""
        WsExplanation.Cells(FirstRowNumber + 3, 9) = ""
    Else
        WsExplanation.Cells(FirstRowNumber, 9) = "以上執行時間依電腦效能差 2~5 倍不等。如果遇到問題，或有不清楚的地方，歡迎與我聯絡^_^"
        'WsExplanation.Cells(FirstRowNumber+1, 9) = "http://www.evernote.com/l/Aagkf6QougdALJ4-BNrdXTuyC9KhgAfimCI/"
        'WsExplanation.Cells(FirstRowNumber+2, 9) = "Email：skyran1278@gmail.com"
        WsExplanation.Cells(FirstRowNumber + 3, 9) = "Paul"
        
        WsExplanation.Hyperlinks.Add _
        Anchor:=WsExplanation.Cells(FirstRowNumber + 1, 9), _
        Address:="https://www.evernote.com/shard/s424/sh/247fa428-ba07-402c-9e3e-04dadd5d3bb2/0bd2a18007e29822", _
        ScreenTip:="Hi~我不是病毒喔^_^", _
        TextToDisplay:="搭接長度精細計算實作歷程"
        
        WsExplanation.Hyperlinks.Add Anchor:=WsExplanation.Cells(FirstRowNumber + 2, 9), _
        Address:="mailto:skyran1278@gmail.com?subject=【搭接長度精細計算】", _
        ScreenTip:="寄信給我~感覺很好玩", _
        TextToDisplay:="Please Email to me"
        
        Cells(FirstRowNumber + 1, 9).Font.Name = "微軟正黑體"
        Cells(FirstRowNumber + 2, 9).Font.Name = "微軟正黑體"
        
    End If
    
    'Application.ScreenUpdating = True
    
End Sub

