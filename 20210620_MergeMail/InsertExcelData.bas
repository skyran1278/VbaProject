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
    Dim pageRow As Long
    Dim pageContainRowNumber As Long
    Dim pageFirstRow As Long

    Dim userName As String
    Dim userNameCharacter As String
    Dim userNameSpaceWidth As Long
    Dim userNameIndex As Long

    Dim userNumber As String
    Dim userBusinessNumber As String
    Dim userAccountNumber As String
    Dim userBranchNumber As String
    Dim userCheckNumber As String

    Dim tableType As String

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

    pageContainRowNumber = 23

    pageFirstRow = 0

    ' �M���Ҧ����e
    ActiveDocument.Content.Select
    Selection.Delete

    With Selection
        With .ParagraphFormat
            .LeftIndent = CentimetersToPoints(-0.75)
            .RightIndent = CentimetersToPoints(-1.38)
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
        End With
        For row = 2 To lastRow
            If excelSheet.Cells(row, 2) <> excelSheet.Cells(row - 1, 2) Then
                If row <> 2 Then
                    ' ��W�@�Ӫ��̫� row
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + pageContainRowNumber
                    ' �W�[�s�� page
                    .TypeParagraph
                    ' �s�� page row �_�l
                    pageFirstRow = pageFirstRow + pageContainRowNumber
                End If

                ' �ɻ��Ҧ� page row�A��K���ᰵ����
                For pageRow = 1 To pageContainRowNumber - 1
                    .TypeParagraph
                Next pageRow

                Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 4
                .TypeText Text:="                 "
                ' �Τ�W��
                userName = excelSheet.Cells(row, 10)
                .TypeText Text:=userName

                ' https://www.tek-tips.com/viewthread.cfm?qid=1056901
                ' �r���e�פ��@�˳y�������վ��m
                ' �Q�� unicode �i��P�_
                .TypeText Text:="                                                                                       "
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
                .TypeText Text:="    "
                ' ��
                userNumber = excelSheet.Cells(row, 2)
                userBusinessNumber = Mid(userNumber, 1, 2)
                userAccountNumber = Mid(userNumber, 3, 4)
                userBranchNumber = Mid(userNumber, 7, 2)
                userCheckNumber = Mid(userNumber, 9, 1)
                .TypeText Text:=userBusinessNumber
                .TypeText Text:=" "
                .TypeText Text:=userAccountNumber
                .TypeText Text:="   "
                .TypeText Text:=userBranchNumber
                .TypeText Text:=" "
                .TypeText Text:=userCheckNumber

                Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 5
                .TypeText Text:="                 "
                ' �ιq�a�}
                .TypeText Text:=excelSheet.Cells(row, 11)

                Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 11
                ' �ۧO
                .TypeText Text:="                         "
                .TypeText Text:=excelSheet.Cells(row, 5)
            End If

            tableType = excelSheet.Cells(row, 3)

            Select Case tableType
                Case "1"
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 12
                Case "2"
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 13
                Case "3"
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 14
                Case "4"
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 16
                Case "6"
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 17
                Case "8"
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 19
                Case "9"
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 20
                Case "10"
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 21
                Case "11"
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 22
                Case "12"
                    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=pageFirstRow + 23
            End Select

            .TypeText Text:="                      "
            ' ����
            .TypeText Text:=excelSheet.Cells(row, 4)
            .TypeText Text:="           "
            ' �q���
            .TypeText Text:=excelSheet.Cells(row, 6)
            .TypeText Text:="       "
            ' ����
            .TypeText Text:=excelSheet.Cells(row, 8)
            .TypeText Text:="     "
            ' �˩w����
            .TypeText Text:=excelSheet.Cells(row, 9)

        Next row

    End With

    excelDatabase.Close False
    Set excelApplication = Nothing
    Set excelSheet = Nothing

End Sub

