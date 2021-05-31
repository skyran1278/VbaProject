Attribute VB_Name = "filter"
Sub FilterData()

    Set AssembledPointMasses = Worksheets("Assembled Point Masses")
    Set MaterialListByStory = Worksheets("Material List By Story")
    Set StoryShears = Worksheets("Story Shears")
    Set Scan = Worksheets("Scan")

    Set Mass = Worksheets("Mass")
    Set Area = Worksheets("Area")
    Set DL = Worksheets("DL")
    Set LL = Worksheets("LL")
    Set StaticSeismic = Worksheets("�R�A�a�_�O")
    Set DynamicSeismic = Worksheets("�ʺA�a�_�O�ץ�")

    AssembledPointMassesLastrow = AssembledPointMasses.UsedRange.Rows.Count '------------����̫�@��
    MaterialListByStoryLastrow = MaterialListByStory.UsedRange.Rows.Count '------------����̫�@��
    StoryShearsLastrow = StoryShears.UsedRange.Rows.Count '------------����̫�@��

' ---------���R���즳�����ADEBUG�A���n clear �⦸�~��������b
    DL.Cells.Clear
    DL.Cells.Clear
    LL.Cells.Clear
    LL.Cells.Clear
    StaticSeismic.Cells.Clear
    StaticSeismic.Cells.Clear
    DynamicSeismic.Cells.Clear
    DynamicSeismic.Cells.Clear
    Area.Cells.Clear
    Area.Cells.Clear
    Mass.Cells.Clear
    Mass.Cells.Clear

    ' �M���e�T��
    For i = 3 To 17 Step 7
        Scan.Range(Scan.Cells(14, i), Scan.Cells(10000, i + 1)).ClearContents
        Scan.Range(Scan.Cells(14, i + 3), Scan.Cells(10000, i + 3)).ClearContents
    Next

'--------- �ƻs�K�W

    '�ƻs Story Shears �K�W��U Sheet
    StoryShears.Cells.Copy
    DL.Paste (DL.Cells)
    LL.Paste (LL.Cells)
    StaticSeismic.Paste (StaticSeismic.Cells)
    DynamicSeismic.Paste (DynamicSeismic.Cells)

    '�ƻs Floor �K�W�� Area
    MaterialListByStory.Cells.Copy
    Area.Paste (Area.Cells)

    '�ƻs AssembleMass �K�W�� Mass
    AssembledPointMasses.Cells.Copy
    Mass.Paste (Mass.Cells)

'---------�z��

    EQfloor = Scan.Cells(2, 4)

    Area.Cells.AutoFilter
    Area.Range(Area.Cells(1, 1), Area.Cells(MaterialListByStoryLastrow, 8)).AutoFilter Field:=2, Criteria1:="Floor"

    DL.Range(DL.Cells(1, 1), DL.Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:="DL"
    DL.Range(DL.Cells(1, 1), DL.Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"

    LL.Range(LL.Cells(1, 1), LL.Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:="LL"
    LL.Range(LL.Cells(1, 1), LL.Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"

    StaticSeismic.Cells.Sort Key1:=StaticSeismic.Range(StaticSeismic.Cells(2, 2), StaticSeismic.Cells(StoryShearsLastrow, 2)), Order1:=xlAscending, Header:=xlYes
    StaticSeismic.Range(StaticSeismic.Cells(1, 1), StaticSeismic.Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:=Array( _
        "DL", "EQXN", "EQXP", "EQYN", "EQYP"), Operator:=xlFilterValues
    StaticSeismic.Range(StaticSeismic.Cells(1, 1), StaticSeismic.Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"
    StaticSeismic.Range(StaticSeismic.Cells(1, 1), StaticSeismic.Cells(StoryShearsLastrow, 9)).AutoFilter Field:=1, Criteria1:=EQfloor

    DynamicSeismic.Cells.Sort Key1:=DynamicSeismic.Range(DynamicSeismic.Cells(2, 2), DynamicSeismic.Cells(StoryShearsLastrow, 2)), Order1:=xlAscending, Header:=xlYes
    DynamicSeismic.Range(DynamicSeismic.Cells(1, 1), DynamicSeismic.Cells(StoryShearsLastrow, 9)).AutoFilter Field:=2, Criteria1:=Array( _
        "0SPECX", "0SPECY", "EQV", "EQXN", "EQXP", "EQYN", "EQYP", "SPECXF MAX", "SPECYF MAX") _
        , Operator:=xlFilterValues
    DynamicSeismic.Range(DynamicSeismic.Cells(1, 1), DynamicSeismic.Cells(StoryShearsLastrow, 9)).AutoFilter Field:=3, Criteria1:="Bottom"
    DynamicSeismic.Range(DynamicSeismic.Cells(1, 1), DynamicSeismic.Cells(StoryShearsLastrow, 9)).AutoFilter Field:=1, Criteria1:=EQfloor

    Mass.Cells.Sort Key1:=Mass.Range(Mass.Cells(2, 2), Mass.Cells(AssembledPointMassesLastrow, 2)), Order1:=xlDescending, Header:=xlYes
    Mass.Cells.AutoFilter
    Mass.Range(Mass.Cells(1, 1), Mass.Cells(AssembledPointMassesLastrow, 11)).AutoFilter Field:=2, Criteria1:="All"

'-----�K�WScan

    '------------------�K�WSTORY
    Mass.Range(Mass.Cells(2, 1), Mass.Cells(AssembledPointMassesLastrow, 1)).Copy
    Scan.Range("C14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Scan.Range("J14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Scan.Range("Q14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '------------------�K�WM
    Mass.Range(Mass.Cells(2, 3), Mass.Cells(AssembledPointMassesLastrow, 3)).Copy
    Scan.Range("D14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '------------------�K�WP
    DL.Range(DL.Cells(3, 4), DL.Cells(StoryShearsLastrow, 4)).Copy
    Scan.Range("K14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '------------------�K�WP
    LL.Range(LL.Cells(3, 4), LL.Cells(StoryShearsLastrow, 4)).Copy
    Scan.Range("R14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '------------------�K�WFloorArea
    Area.Range(Area.Cells(4, 5), Area.Cells(MaterialListByStoryLastrow, 5)).Copy
    Scan.Range("F14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '-----------------�ץ�FloorArea
    FloorAreaLastrow = Scan.Cells(Scan.Rows.Count, "F").End(xlUp).Row
    StoryLastrow = Scan.Cells(Scan.Rows.Count, "C").End(xlUp).Row
    Scan.Cells(StoryLastrow, 6) = Scan.Cells(FloorAreaLastrow, 6)
    Scan.Cells(FloorAreaLastrow, 6) = ""
    Scan.Range(Scan.Cells(14, 6), Scan.Cells(StoryLastrow, 6)).Copy
    '-----------------���KFloorArea
    Scan.Range("M14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Scan.Range("T14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '------------------�K�W�R�A�a�_�O
    StaticSeismic.Range(StaticSeismic.Cells(1, 1), StaticSeismic.Cells(StoryShearsLastrow, 9)).Copy
    Scan.Range("W14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    '------------------�K�W�ʺA�a�_�O�ץ�
    DynamicSeismic.Range(DynamicSeismic.Cells(1, 1), DynamicSeismic.Cells(StoryShearsLastrow, 9)).Copy
    Scan.Range("AG14").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    Scan.Select

End Sub


