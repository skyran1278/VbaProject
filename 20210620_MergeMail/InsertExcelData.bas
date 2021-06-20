Option Explicit

Sub insertExcelData()
    Dim mailMerge As Object
    Dim excelApplication As Object
    Dim excelDatabase As Object
    Dim excelSheet As Object
    Dim excelPath As String
    Dim excelSheetName As String
    Dim xlup As Long
    Dim lastRow As Long
    Dim Index As Long

    ' �o�̥i�H���� excel �W�٩άO�ϥ� �l�� > �������� > �ϥβ{���M�椤�� excel
    ' excelPath = "C:\Users\skyra\Downloads\�j��110.05.21.xlsx"
    excelPath = ActiveDocument.mailMerge.DataSource.Name
    excelSheetName = "�u�@��2"

    Set mailMerge = ActiveDocument.mailMerge
    Set excelApplication = CreateObject("Excel.Application")

    Set excelDatabase = excelApplication.Workbooks.Open(excelPath)

    Set excelSheet = excelDatabase.Worksheets(excelSheetName)

    ' https://www.reddit.com/r/vba/comments/altr3h/ms_project_xlup_and_variable_not_defined_error/
    xlup = -4162
    lastRow = excelSheet.Cells(excelSheet.Rows.Count, 1).End(xlup).Row

    With Selection
        For Index = 2 To 10
        ' For Index = 2 To lastRow
            If excelSheet.Cells(Index, 2) <> excelSheet.Cells(Index - 1, 2) Then
                If Index <> 2 Then
                    .InsertBreak Type:=wdPageBreak
                End If

                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeText Text:="       "
                .TypeText Text:=excelSheet.Cells(Index, 10)
                .TypeText Text:="                                                                           "
                ' mailMerge.Fields.Add Range:=.Range, Name:="�p���"
                .TypeText Text:="  "
                ' mailMerge.Fields.Add Range:=.Range, Name:="��"
                .TypeParagraph
                .TypeText Text:="       "
                ' mailMerge.Fields.Add Range:=.Range, Name:="�ιq�a�}"
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeText Text:="                  "
                ' mailMerge.Fields.Add Range:=.Range, Name:="����"
                .TypeText Text:="  "
                ' mailMerge.Fields.Add Range:=.Range, Name:="�ۧO"
                .TypeText Text:="  "
                ' mailMerge.Fields.Add Range:=.Range, Name:="�q���"
                .TypeText Text:="  "
                ' mailMerge.Fields.Add Range:=.Range, Name:="����"
            Else
                If Index <> 2 Then
                    .TypeParagraph
                End If
                .TypeParagraph
                .TypeText Text:="                  "
                ' mailMerge.Fields.Add Range:=.Range, Name:="����"
                .TypeText Text:="  "
                ' mailMerge.Fields.Add Range:=.Range, Name:="�ۧO"
                .TypeText Text:="  "
                ' mailMerge.Fields.Add Range:=.Range, Name:="�q���"
                .TypeText Text:="  "
                ' mailMerge.Fields.Add Range:=.Range, Name:="����"
            End If
        Next Index

    End With

    excelDatabase.Close False
    Set excelApplication = Nothing
    Set excelSheet = Nothing

End Sub

