' @license version_control v1.0.0
' version_control.vb
'
' Copyright (c) 2016-present, skyran
'
' This source code is licensed under the MIT license found in the
' LICENSE file in the root directory of this source tree.


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
'
' * 測試環境:
'       office 2016 in windows 10
'       Mac 版本容易出現錯誤，不推薦在 Mac 執行


    ' 此程序包含的變數
    Dim DOWNLOAD_URL As String
    Dim VERSION_URL As String

    Dim VERSION_SHEET As Worksheet
    Dim project As String
    Dim currentVersion As String
    Dim latestVersion As String


    ' 依據不同工作簿有不同值
    VERSION_URL = "https://github.com/skyran1278/VbaProject/raw/master/20170226_LapLength/lap-length-version.txt"
    DOWNLOAD_URL = "https://github.com/skyran1278/VbaProject/raw/master/20170226_LapLength/lap-length.xlsm"


    Set VERSION_SHEET = ThisWorkbook.Worksheets("版本資訊")

    ' 位置在 Cells(4, 3)
    With VERSION_SHEET.QueryTables.Add(Connection:="URL;" & VERSION_URL, _
        Destination:=VERSION_SHEET.Cells(4, 3))
        .NAME = "version"
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


    ' 移除連線
    ' Mac 版本 Connections 錯誤，所以增加下面一行
    ' On Error Resume Next
    ThisWorkbook.Connections("連線").Delete

    ' 移除名稱
    ' 第二次執行會出現錯誤，但一般來說不會出現第二次。所以先註解掉。
    ' On Error Resume Next
    ThisWorkbook.Names("版本資訊!version").Delete


    project = VERSION_SHEET.Cells(2, 3)
    currentVersion = VERSION_SHEET.Cells(3, 3)
    latestVersion = VERSION_SHEET.Cells(4, 3)

    If latestVersion > currentVersion Then

        intMessage = MsgBox("下載最新版本...", vbYesNo, project)

        If intMessage = vbYes Then

            ' Mac 版本出現錯誤，不推薦在 Mac 執行
            Set OBJ_SHELL = CreateObject("Wscript.Shell")
            OBJ_SHELL.Run (DOWNLOAD_URL)
            MsgBox "請關閉此檔案，並使用從瀏覽器下載的最新版本。", vbOKOnly, project

        Else

            MsgBox "使用舊版程式具有無法預期的風險，建議下載最新版程式。" & vbCrLf & "若需下載新版程式請重開檔案。", vbOKOnly, project

        End If
    End If

    VERSION_SHEET.Cells.Font.NAME = "微軟正黑體"
    VERSION_SHEET.Cells.Font.NAME = "Calibri"
    VERSION_SHEET.Activate

End Sub


