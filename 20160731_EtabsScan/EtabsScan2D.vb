Sub FilterData()

    Set Assemble = Worksheets("Assembled Point Masses")
    Set StoryShears = Worksheets("Story Shears")
    Set Scan = Worksheets("Scan")
    Set DL = Worksheets("DL")
    Set LL = Worksheets("LL")
    Set StaticSeismic = Worksheets("靜態地震力")
    Set Dynamic = Worksheets("動態地震力修正")

    AssembleLastrow = Sheets("Assembled Point Masses").UsedRange.Rows.Count '------------抓取最後一行
    StoryShearsLastrow = Sheets("Story Shears").UsedRange.Rows.Count '------------抓取最後一行

' ---------先刪除原有的表格，DEBUG
    DL.Cells.ClearContents
    LL.Cells.ClearContents
    StaticSeismic.Cells.ClearContents
    Dynamic.Cells.ClearContents
    For i = 3 To 17 Step 7
        Scan.Range(Cells(14, i), Cells(10000, i + 1)).ClearContents
    Next

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

    EQfloor = Scan.Cells(2, 4)

    Sheets("DL").Select
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:="DL"
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"
    Sheets("LL").Select
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:="LL"
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"
    Sheets("靜態地震力").Select
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:=Array( _
        "DL", "EQV", "EQX"), Operator:=xlFilterValues
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=1, Criteria1:=EQfloor
    Sheets("動態地震力修正").Select
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:=Array( _
        "0SPECX", "EQX", "SPECXF MAX") _
        , Operator:=xlFilterValues
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"
    ActiveSheet.Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).AutoFilter Field:=1, Criteria1:=EQfloor
    AssembledSub

'-----貼上Scan

    '------------------貼上STORY
    Sheets("Assembled Point Masses").Select
    Range(Cells(2, 1), Cells(AssembleLastrow, 1)).Select
    Selection.Copy
    Sheets("Scan").Select
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
    Sheets("Scan").Select
    Range("D14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '------------------貼上P
    Sheets("LL").Select
    Range(Cells(3, 4), Cells(StoryShearsLastrow, 4)).Select
    Selection.Copy
    Sheets("Scan").Select
    Range("R14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '------------------貼上M
    Sheets("Assembled Point Masses").Select

    Range(Cells(2, 3), Cells(AssembleLastrow, 3)).Select
    Selection.Copy
    Sheets("Scan").Select
    Range("K14").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '------------------貼上靜態地震力
    Sheets("靜態地震力").Select
    Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).Select
    Selection.Copy
    Sheets("Scan").Select
    Range("W14").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '------------------貼上動態地震力修正
    Sheets("動態地震力修正").Select
    Range(Cells(1, 1), Cells(StoryShearsLastrow, 9)).Select
    Selection.Copy
    Sheets("Scan").Select
    Range("AG14").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

End Sub


