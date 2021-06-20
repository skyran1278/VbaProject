Option Explicit

Sub insertExcelData()
    Dim excelApplication As Object
    Dim excelDatabase As Object
    Dim excelSheet As Object
    Dim excelPath As String
    Dim excelSheetName As String
    Dim xlup As Long
    Dim lastRow As Long
    Dim row As Long
    Dim userName As String
    Dim userNameCharacter As String
    Dim userNameSpaceWidth As Long
    Dim userNameIndex As Long

    ' �o�̦��T�ؼg�k
    ' 1. �H������|���� excel �W��
    ' excelPath = "C:\Users\skyra\Downloads\�j��110.05.21.xlsx"
    ' 2. �M word �ۦP��Ƨ��� excel �W��
    excelPath = ActiveDocument.Path & Application.PathSeparator & "�j��110.05.21.xlsx"
    ' 3. �ϥ� �l�� > �������� > �ϥβ{���M�椤�� excel
    ' excelPath = ActiveDocument.MailMerge.DataSource.Name

    ' excel �u�@��W��
    excelSheetName = "�u�@��2"

    Set excelApplication = CreateObject("Excel.Application")
    Set excelDatabase = excelApplication.Workbooks.Open(excelPath)
    Set excelSheet = excelDatabase.Worksheets(excelSheetName)

    ' https://www.reddit.com/r/vba/comments/altr3h/ms_project_xlup_and_variable_not_defined_error/
    xlup = -4162
    lastRow = excelSheet.Cells(excelSheet.Rows.Count, 1).End(xlup).row

    With Selection
        For row = 2 To lastRow
            If excelSheet.Cells(row, 2) <> excelSheet.Cells(row - 1, 2) Then
                If row <> 2 Then
                    .InsertBreak Type:=wdPageBreak
                End If

                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeText Text:="              "
                ' �Τ�W��
                Selection.Range.CharacterWidth = wdWidthHalfWidth
                userName = excelSheet.Cells(row, 10)
                .TypeText Text:=userName

                ' �r���e�פ��@�˳y�������վ��m
                ' �Q�� unicode �i��P�_
                .TypeText Text:="                                                                                      "
                userNameSpaceWidth = 0
                For userNameIndex = 1 To Len(userName)
                    userNameSpaceWidth = userNameSpaceWidth + 1
                    'do something to each character in string
                    'here we'll msgbox each character
                    userNameCharacter = Mid(userName, userNameIndex, 1)
                    If Asc(userNameCharacter) <> AscW(userNameCharacter) Then
                        userNameSpaceWidth = userNameSpaceWidth + 1
                    End If
                Next
                While userNameSpaceWidth > 0
                    .TypeBackspace
                    userNameSpaceWidth = userNameSpaceWidth - 1
                Wend

                ' �p���
                .TypeText Text:=excelSheet.Cells(row, 1)
                .TypeText Text:="     "
                ' ��
                .TypeText Text:=excelSheet.Cells(row, 2)
                .TypeParagraph
                .TypeText Text:="              "
                ' �ιq�a�}
                .TypeText Text:=excelSheet.Cells(row, 11)
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                ' �ۧO
                .TypeText Text:="                     "
                .TypeText Text:=excelSheet.Cells(row, 5)
            End If

            .TypeParagraph
            .TypeText Text:="                  "
            ' ����
            .TypeText Text:=excelSheet.Cells(row, 4)
            .TypeText Text:="           "
            ' �q���
            .TypeText Text:=excelSheet.Cells(row, 6)
            .TypeText Text:="    "
            ' ����
            .TypeText Text:=excelSheet.Cells(row, 8)
            .TypeText Text:="       "
            ' �˩w����
            .TypeText Text:=excelSheet.Cells(row, 9)

        Next row

    End With

    excelDatabase.Close False
    Set excelApplication = Nothing
    Set excelSheet = Nothing

End Sub


