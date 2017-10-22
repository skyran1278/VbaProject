Private Sub Workbook_Open()
'
' * 目的: 檢查程式最新版本，並自動提示更新
'
' * 隨工作簿不同而需更改的參數:
'       VERSION_URL: 該工作簿 version.txt
'       DOWNLOAD_URL: 該工作簿 下載檔案位置
'
' * 重要且通常不會更動數值:
'       工作表位置: 版本資訊
'       名稱: Cells(2, 3)
'       目前版本號: Cells(3, 3)
'       最新版本號: Cells(4, 3)


    ' 此程序包含的變數
    Dim DOWNLOAD_URL As String
    Dim VERSION_URL As String

    Dim sheet As String
    Dim project As String
    Dim currentVersion As String
    Dim latestVersion As String


    ' 依據不同工作簿有不同值
    VERSION_URL = "https://raw.githubusercontent.com/skyran1278/VbaProject/master/20170524%20RC%20Scan/version.txt"
    DOWNLOAD_URL = "https://github.com/skyran1278/VbaProject/raw/master/20170524%20RC%20Scan/RC%20Scan.xlsm"


    sheet = "版本資訊"
    Worksheets(sheet).Activate


    ' 位置在 Cells(4, 3)
    With ActiveSheet.QueryTables.Add(Connection:= "URL;" & VERSION_URL, _
        Destination:= Cells(4, 3))
        .Name = "version"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        ' .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells '覆蓋文字
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        ' .RefreshPeriod = 0
        ' .WebSelectionType = xlEntirePage
        ' .WebFormatting = xlWebFormattingNone
        ' .WebPreFormattedTextToColumns = True
        ' .WebConsecutiveDelimitersAsOne = True
        ' .WebSingleBlockTextImport = False
        ' .WebDisableDateRecognition = False
        ' .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With


    project = Cells(2, 3)
    currentVersion = Cells(3, 3)
    latestVersion = Cells(4, 3)

    If latestVersion > currentVersion Then

        intMessage = MsgBox("下載最新版本...", vbYesNo, project)

        If intMessage = vbYes Then
            Set OBJ_SHELL = CreateObject("Wscript.Shell")
            OBJ_SHELL.Run (DOWNLOAD_URL)
        End If

    End If


    ' 移除連線
    ActiveWorkbook.Connections("連線").Delete

    ' 移除名稱
    ' 第二次執行會出現錯誤，但一般來說不會出現第二次。所以先註解掉。
    ' On Error Resume Next
    ActiveWorkbook.Names("版本資訊!version").Delete


End Sub
