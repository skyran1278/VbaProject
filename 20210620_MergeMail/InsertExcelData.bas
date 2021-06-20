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

    ' 這裡有三種寫法
    ' 1. 以絕對路徑替換 excel 名稱
    ' excelPath = "C:\Users\skyra\Downloads\大表110.05.21.xlsx"
    ' 2. 和 word 相同資料夾的 excel 名稱
    excelPath = ActiveDocument.Path & Application.PathSeparator & "大表110.05.21.xlsx"
    ' 3. 使用 郵件 > 選取收件者 > 使用現有清單中的 excel
    ' excelPath = ActiveDocument.MailMerge.DataSource.Name

    ' excel 工作表名稱
    excelSheetName = "工作表2"

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
                ' 用戶名稱
                Selection.Range.CharacterWidth = wdWidthHalfWidth
                userName = excelSheet.Cells(row, 10)
                .TypeText Text:=userName

                ' 字元寬度不一樣造成很難調整位置
                ' 利用 unicode 進行判斷
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

                ' 計算日
                .TypeText Text:=excelSheet.Cells(row, 1)
                .TypeText Text:="     "
                ' 號
                .TypeText Text:=excelSheet.Cells(row, 2)
                .TypeParagraph
                .TypeText Text:="              "
                ' 用電地址
                .TypeText Text:=excelSheet.Cells(row, 11)
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                .TypeParagraph
                ' 相別
                .TypeText Text:="                     "
                .TypeText Text:=excelSheet.Cells(row, 5)
            End If

            .TypeParagraph
            .TypeText Text:="                  "
            ' 型式
            .TypeText Text:=excelSheet.Cells(row, 4)
            .TypeText Text:="           "
            ' 電表表號
            .TypeText Text:=excelSheet.Cells(row, 6)
            .TypeText Text:="    "
            ' 倍數
            .TypeText Text:=excelSheet.Cells(row, 8)
            .TypeText Text:="       "
            ' 檢定期限
            .TypeText Text:=excelSheet.Cells(row, 9)

        Next row

    End With

    excelDatabase.Close False
    Set excelApplication = Nothing
    Set excelSheet = Nothing

End Sub


