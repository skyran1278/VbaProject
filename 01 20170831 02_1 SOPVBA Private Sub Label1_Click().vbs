Private Sub Label1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub 確定_Click()
    Dim EQfloor As Variant
    EQfloor = TextBox1.Text
    
    '
' 整理
' SOP
'
' 快速鍵: Ctrl+j

    Set Assemble = Worksheets("Assembled Point Masses")
    Set Material = Worksheets("Material List By Story")
    Set StoryShears = Worksheets("Story Shears")
    Set Check = Worksheets("Check")
    Set DL = Worksheets("DL")
    Set LL = Worksheets("LL")
    Set StaticSeismic = Worksheets("靜態地震力")
    Set Dynamic = Worksheets("動態地震力修正")
    
    AssembleLastrow = Sheets("Assembled Point Masses").UsedRange.Rows.Count '------------抓取最後一行
    MaterialLastrow = Sheets("Material List By Story").UsedRange.Rows.Count '------------抓取最後一行
    StoryShearsLastrow = Sheets("Story Shears").UsedRange.Rows.Count '------------抓取最後一行

' ---------先刪除原有的表格，DEBUG
    DL.Cells.ClearContents
    LL.Cells.ClearContents
    StaticSeismic.Cells.ClearContents
    Dynamic.Cells.ClearContents
    DL.Cells.ClearContents
    LL.Cells.ClearContents
    StaticSeismic.Cells.ClearContents
    Dynamic.Cells.ClearContents
    For i = 3 To 17 Step 7
        Check.Range(Cells(14, i), Cells(10000, i + 1)).ClearContents
        Check.Range(Cells(14, i + 3), Cells(10000, i + 3)).ClearContents
    Next
    
'    Sheets("DL").Select
'    Cells.Select
'    Selection.ClearContents
'    Selection.ClearContents
'    Sheets("LL").Select
'    Cells.Select
'    Selection.ClearContents
'    Selection.ClearContents
'    Sheets("靜態地震力").Select
'    Cells.Select
'    Selection.ClearContents
'    Selection.ClearContents
'    Sheets("動態地震力修正").Select
'    Cells.Select
'    Selection.ClearContents
'    Selection.ClearContents

'---------複製Story Shears貼上到各Sheet

    Sheets("Story Shears").Select
    Cells.Select
    Selection.Copy
    Sheets("DL").Select
    Cells.Select
    ActiveSheet.Paste
    Sheets("LL").Select
    Cells.Select
    ActiveSheet.Paste
    Sheets("靜態地震力").Select
    Cells.Select
    ActiveSheet.Paste
    Sheets("動態地震力修正").Select
    Cells.Select
    ActiveSheet.Paste
    
'---------篩選

'    Sheets("Material List By Story").Select
'    Cells.Select
'    Selection.AutoFilter
'    ActiveSheet.Range("$A$1:$H$24").AutoFilter Field:=2, Criteria1:="Floor"
'    Sheets("DL").Select
'    Cells.Select
'    Selection.AutoFilter
'    Sheets("LL").Select
'    Cells.Select
'    Selection.AutoFilter
'    Sheets("靜態地震力").Select
'    Cells.Select
'    Selection.AutoFilter
'    Sheets("動態地震力修正").Select
'    Cells.Select
'    Selection.AutoFilter
'    Sheets("DL").Select
'    ActiveSheet.Range("$A$1:$I$1585").AutoFilter Field:=2, Criteria1:="DL"
'    ActiveSheet.Range("$A$1:$I$1585").AutoFilter Field:=3, Criteria1:="Bottom"
'    Sheets("LL").Select
'    ActiveSheet.Range("$A$1:$I$1585").AutoFilter Field:=2, Criteria1:="LL"
'    ActiveSheet.Range("$A$1:$I$1585").AutoFilter Field:=3, Criteria1:="Bottom"
'    Sheets("靜態地震力").Select
'    ActiveSheet.Range("$A$1:$I$1585").AutoFilter Field:=2, Criteria1:=Array( _
'        "DL", "EQXN", "EQXP", "EQYN", "EQYP"), Operator:=xlFilterValues
'    ActiveSheet.Range("$A$1:$I$1585").AutoFilter Field:=3, Criteria1:="Bottom"
'    ActiveSheet.Range("$A$1:$I$1585").AutoFilter Field:=1, Criteria1:=EQfloor
'    Sheets("動態地震力修正").Select
'    ActiveSheet.Range("$A$1:$I$1585").AutoFilter Field:=2, Criteria1:=Array( _
'        "0SPECX", "0SPECY", "EQV", "EQXN", "EQXP", "EQYN", "EQYP", "SPECXF MAX", "SPECYF MAX") _
'        , Operator:=xlFilterValues
'    ActiveSheet.Range("$A$1:$I$1585").AutoFilter Field:=3, Criteria1:="Bottom"
'    ActiveSheet.Range("$A$1:$I$1585").AutoFilter Field:=1, Criteria1:=EQfloor

    Sheets("Material List By Story").Select
    Cells.Select
    Selection.AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(MaterialLastrow, 8)).AutoFilter Field:=2, Criteria1:="Floor"
    Sheets("DL").Select
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:="DL"
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"
    Sheets("LL").Select
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:="LL"
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"
    Sheets("靜態地震力").Select
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:=Array( _
        "DL", "EQXN", "EQXP", "EQYN", "EQYP"), Operator:=xlFilterValues
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=1, Criteria1:=EQfloor
    Sheets("動態地震力修正").Select
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:=Array( _
        "0SPECX", "0SPECY", "EQV", "EQXN", "EQXP", "EQYN", "EQYP", "SPECXF MAX", "SPECYF MAX") _
        , Operator:=xlFilterValues
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=1, Criteria1:=EQfloor
    
'-----貼上CHECK

    '------------------貼上STORY
    Sheets("Assembled Point Masses").Select
    Range(Cells(2, 1), Cells(AssembleLastrow, 1)).Select
    Selection.Copy
    Sheets("Check").Select
    Range("C14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("J14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("Q14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    '------------------貼上P
    Sheets("DL").Select
    Range(Cells(3, 4), Cells(StoryShearsLastrow, 4)).Select
    Selection.Copy
    Sheets("Check").Select
    Range("D14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    '------------------貼上P
    Sheets("LL").Select
    Range(Cells(3, 4), Cells(StoryShearsLastrow, 4)).Select
    Selection.Copy
    Sheets("Check").Select
    Range("R14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    '------------------貼上M
    Sheets("Assembled Point Masses").Select
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
        Cells.Select
        ActiveSheet.Range(Cells(1, 1), Cells(AssembleLastrow, 11)).AutoFilter Field:=2, Criteria1:="All"
    Else
    Cells.Select
    Selection.AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(AssembleLastrow, 11)).AutoFilter Field:=2, Criteria1:="All"
    End If
    
    Range(Cells(2, 3), Cells(AssembleLastrow, 3)).Select
    Selection.Copy
    Sheets("Check").Select
    Range("K14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    '------------------貼上FloorArea
    Sheets("Material List By Story").Select
    Range(Cells(4, 5), Cells(MaterialLastrow, 5)).Select
    Selection.Copy
    Sheets("Check").Select
    Range("F14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    '-----------------修正FloorArea
    FloorAreaLastrow = Check.Cells(Check.Rows.Count, "F").End(xlUp).Row
    StoryLastrow = Check.Cells(Check.Rows.Count, "C").End(xlUp).Row
    Cells(StoryLastrow, 6) = Cells(FloorAreaLastrow, 6)
    Cells(FloorAreaLastrow, 6) = ""
    Range(Cells(14, 6), Cells(StoryLastrow, 6)).Select
    Selection.Copy
    '-----------------重貼FloorArea
    Range("M14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("T14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    '------------------貼上靜態地震力
    Sheets("靜態地震力").Select
    Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).Select
    Selection.Copy
    Sheets("Check").Select
    Range("W14").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    '------------------貼上動態地震力修正
    Sheets("動態地震力修正").Select
    Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).Select
    Selection.Copy
    Sheets("Check").Select
    Range("AG14").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        
    Unload Me
    
End Sub

Private Sub 篩選樓層_Change()

End Sub

Private Sub 重新篩選樓層_Initialize()
    篩選樓層.RowSource = Worksheets("DL").Range("C3:C100").Address
End Sub




