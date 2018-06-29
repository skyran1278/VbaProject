' @license version_control v1.1.0
' version_control.vb
'
' Copyright (c) 2016-present, skyran
'
' This source code is licensed under the MIT license found in the
' LICENSE file in the root directory of this source tree.

' Private Sub Workbook_Open()
' '
' ' * 目的: 檢查程式最新版本，並自動提示更新
' '
' ' * 隨工作簿不同而需更改的參數:
' '       VERSION_URL: 該工作簿 version.txt
' '       DOWNLOAD_URL: 該工作簿 下載檔案位置
' '
' ' * 重要且通常不會更動數值:
' '       工作表位置: 版本資訊
' '       名稱: Cells(2, 3)
' '       目前版本號: Cells(3, 3)
' '       最新版本號: Cells(4, 3)
' '
' ' * 測試環境:
' '       office 2016 in windows 10
' '       Mac 版本容易出現錯誤，不推薦在 Mac 執行


'     ' 此程序包含的變數
'     Dim ws_version As Worksheet
'     Dim project As String
'     Dim currentVersion As String
'     Dim latestVersion As String
'     Dim srvXmlHttp As Object
'     Dim inputPwd As String
'     Dim cloud_pwd As String

'     Set srvXmlHttp = CreateObject("MSXML2.serverXMLHTTP")

'     srvXmlHttp.Open "GET", VERSION_URL, False

'     Set ws_version = ThisWorkbook.Worksheets("版本資訊")

'     ' 位置在 Cells(4, 3)
'     ' With ws_version.QueryTables.Add(Connection:="URL;" & VERSION_URL, _
'     '     Destination:=ws_version.Cells(4, 3))
'     '     .NAME = "version"
'     '     .FieldNames = True
'     '     .RowNumbers = False
'     '     .FillAdjacentFormulas = False
'     '     ' .PreserveFormatting = True
'     '     .RefreshOnFileOpen = False
'     '     .BackgroundQuery = True
'     '     .RefreshStyle = xlOverwriteCells '覆蓋文字
'     '     .SavePassword = False
'     '     .SaveData = True
'     '     .AdjustColumnWidth = False
'     '     ' .RefreshPeriod = 0
'     '     ' .WebSelectionType = xlEntirePage
'     '     ' .WebFormatting = xlWebFormattingNone
'     '     ' .WebPreFormattedTextToColumns = True
'     '     ' .WebConsecutiveDelimitersAsOne = True
'     '     ' .WebSingleBlockTextImport = False
'     '     ' .WebDisableDateRecognition = False
'     '     ' .WebDisableRedirections = False
'     '     .Refresh BackgroundQuery:=False
'     ' End With


'     ' 移除連線
'     ' Mac 版本 Connections 錯誤，所以增加下面一行
'     ' On Error Resume Next
'     ' ThisWorkbook.Connections("連線").Delete

'     ' 移除名稱
'     ' 第二次執行會出現錯誤，但一般來說不會出現第二次。所以先註解掉。
'     ' On Error Resume Next
'     ' ThisWorkbook.Names("版本資訊!version").Delete


'     project = ws_version.Cells(2, 3)
'     currentVersion = ws_version.Cells(3, 3)
'     ' latestVersion = ws_version.Cells(4, 3)

'     srvXmlHttp.send

'     latestVersion = srvXmlHttp.ResponseText

'     ' 消除空白行
'     latestVersion = Trim(Replace(latestVersion, Chr(10), ""))

'     If latestVersion > currentVersion Then

'         intMessage = MsgBox("下載最新版本...", vbYesNo, project)

'         If intMessage = vbYes Then

'             ' Mac 版本出現錯誤，不推薦在 Mac 執行
'             Set OBJ_SHELL = CreateObject("Wscript.Shell")
'             OBJ_SHELL.Run (DOWNLOAD_URL)
'             MsgBox "請關閉此檔案，並使用從瀏覽器下載的最新版本。", vbOKOnly, project

'         Else

'             MsgBox "使用舊版程式具有無法預期的風險，建議下載最新版程式。" & vbCrLf & "若需下載新版程式請重開檔案。", vbOKOnly, project

'         End If
'     End If

'     ws_version.Cells.Font.NAME = "微軟正黑體"
'     ws_version.Cells.Font.NAME = "Calibri"
'     ws_version.Activate

' End Sub
' 隨工作簿不同而需更改的參數:
' PASSWORD_URL: 該工作簿 pwd.txt
' VERSION_URL: 該工作簿 version.txt
' DOWNLOAD_URL: 該工作簿 下載檔案位置
Private Const PASSWORD_URL = "https://github.com/skyran1278/VbaProject/raw/master/01%20utils/example-pwd.txt"
Private Const VERSION_URL = "https://github.com/skyran1278/VbaProject/raw/master/01%20utils/example-version.txt"
' Private Const RELEASE_URL = "https://github.com/skyran1278/VbaProject/raw/master/20170413_BeamZValue/z-value-release.txt"
Private Const DOWNLOAD_URL = "https://github.com/skyran1278/VbaProject/raw/master/01%20utils/example.xlsm"


Private Sub VerifyPassword()
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'

    Dim srvXmlHttp As Object
    Dim inputPwd As String
    Dim cloudPwd As String

    Set srvXmlHttp = CreateObject("MSXML2.serverXMLHTTP")

    srvXmlHttp.Open "GET", PASSWORD_URL, False

    inputPwd = Trim(Application.InputBox("Please Input Passward.", "Verify User Identity", type:=2))

    srvXmlHttp.send

    cloudPwd = srvXmlHttp.ResponseText

    ' 消除空白行
    cloudPwd = Trim(Replace(cloudPwd, Chr(10), ""))

    If inputPwd <> cloudPwd Then

        MsgBox "Wrong Password"
        ThisWorkbook.Close SaveChanges:=False

    Else

        Application.StatusBar = "Sign In Success."

    End If

End Sub

' Function Test()
' '
' '
' '
' ' @param
' ' @returns

'     ' Dim srvXmlHttp As Object
'     ' Dim srvXmlHttp3 As Object
'     ' Dim srvXmlHttp6 As Object
'     ' Dim time0 As Double

'     ' Set srvXmlHttp = CreateObject("MSXML2.serverXMLHTTP")
'     ' Set srvXmlHttp3 = CreateObject("MSXML2.serverXMLHTTP.3.0")
'     ' Set srvXmlHttp6 = CreateObject("MSXML2.serverXMLHTTP.6.0")

'     ' time0 = Timer
'     ' srvXmlHttp.Open "GET", VERSION_URL, False
'     ' srvXmlHttp.send
'     ' latestVersionAndReleaseNotes = srvXmlHttp.ResponseText
'     ' Debug.Print Timer - time0

'     ' time0 = Timer
'     ' srvXmlHttp3.Open "GET", VERSION_URL, False
'     ' srvXmlHttp3.send
'     ' latestVersionAndReleaseNotes = srvXmlHttp3.ResponseText
'     ' Debug.Print Timer - time0

'     ' time0 = Timer
'     ' srvXmlHttp6.Open "GET", VERSION_URL, False
'     ' srvXmlHttp6.send
'     ' latestVersionAndReleaseNotes = srvXmlHttp6.ResponseText
'     ' Debug.Print Timer - time0

'     Dim WinHttpReq As Object
'     Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
'     WinHttpReq.Open "GET", DOWNLOAD_URL, False
'     WinHttpReq.send

'     DOWNLOAD = WinHttpReq.responseBody
'     If WinHttpReq.Status = 200 Then
'         Set oStream = CreateObject("ADODB.Stream")
'         oStream.Open
'         oStream.Type = 1
'         oStream.Write WinHttpReq.responseBody
'         oStream.SaveToFile "下載\file.xlsm", 2 ' 1 = no overwrite, 2 = overwrite
'         oStream.Close
'     End If


' End Function

Private Function CompareVersion(currentVersion As String, latestVersion As String)
'
' compare which version is latest.
'
' @since 1.0.0
' @param {string} [currentVersion] currentVersion.
' @param {string} [latestVersion] latestVersion.
' @return {boolean} [CompareVersion] latestVersion > currentVersion return true.
'

    arrCurrentVersion = Split(currentVersion, ".")
    arrLatestVersion = Split(latestVersion, ".")

    If arrLatestVersion(0) > arrCurrentVersion(0) Then
        CompareVersion = True
    ElseIf arrLatestVersion(1) > arrCurrentVersion(1) Then
        CompareVersion = True
    ElseIf arrLatestVersion(2) > arrCurrentVersion(2) Then
        CompareVersion = True
    Else
        CompareVersion = False
    End If

End Function

Private Sub CheckVersion()
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'

    ' 此程序包含的變數
    Dim srvXmlHttp As Object
    Dim ws_version As Worksheet
    Dim project As String
    Dim currentVersion As String
    Dim latestVersion As String

    Set srvXmlHttp = CreateObject("MSXML2.serverXMLHTTP")

    srvXmlHttp.Open "GET", VERSION_URL, False

    MsgBox "Click To Check Latest Version."

    Application.StatusBar = "Checking Latest Version..."

    Set ws_version = ThisWorkbook.Worksheets("版本資訊")

    currentVersion = ws_version.Cells(3, 3)

    srvXmlHttp.send

    latestVersionAndReleaseNotes = srvXmlHttp.ResponseText

    ' 區分版本號和更新說明
    latestVersionAndReleaseNotes = Split(latestVersionAndReleaseNotes, Chr(10) & "===" & Chr(10))
    latestVersion = latestVersionAndReleaseNotes(0)
    releaseNotes = latestVersionAndReleaseNotes(1)
    ' MsgBox latestVersion
    ' MsgBox releaseNote

    If CompareVersion(currentVersion, latestVersion) Then

        ' srvXmlHttp.Open "GET", DOWNLOAD_URL, False

        intMessage = MsgBox(releaseNotes, vbYesNo, "Please Download Latest Version.")

        If intMessage = vbYes Then

            Set OBJ_SHELL = CreateObject("Wscript.Shell")
            OBJ_SHELL.Run (DOWNLOAD_URL)
            MsgBox "請關閉此檔案，並使用從瀏覽器下載的最新版本。", vbOKOnly

        Else

            MsgBox "使用舊版程式具有無法預期的風險，建議下載最新版程式。" & vbCrLf & "若需下載新版程式請重開檔案。", vbOKOnly

        End If
    End If

    ws_version.Cells.Font.NAME = "微軟正黑體"
    ws_version.Cells.Font.NAME = "Calibri"
    ws_version.Activate

End Sub


'


