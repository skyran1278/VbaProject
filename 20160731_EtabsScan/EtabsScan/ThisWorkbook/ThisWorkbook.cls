VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' @license Version v4.0.0
Option Explicit

' 每次更新版本都需要修改
' 由於改成強制更新，所以要拉到 private，不讓 user 可以自己更改
Private Const CURRENT_VERSION = "3.0.0"

' 隨工作簿不同而需更改的參數:
Private Const WORKBOOK_NAME = "EtabsScan"
Private Const VERSION_URL = "https://github.com/skyran1278/elements-vba-service/raw/main/" + WORKBOOK_NAME + "/VERSION.json"
Private Const DOWNLOAD_URL = "https://github.com/skyran1278/elements-vba-service/raw/main/" + WORKBOOK_NAME + "/" + WORKBOOK_NAME + ".xlsm"

Private ran As New UTILS_CLASS

Private Sub Workbook_Open()
'
' * 目的: 驗證密碼，檢查程式最新版本，並自動提示更新
'
' * 隨工作簿不同而需更改的參數:
'       VERSION_URL: 該工作簿 version.txt
'       DOWNLOAD_URL: 該工作簿 下載檔案位置
'
' * 測試環境:
'       office 2016 in windows 10
'       Mac 版本容易出現錯誤，不推薦在 Mac 執行

    Dim inputPassword As String
    Dim strongPassword As String

    On Error GoTo ErrorHandler

    If isOutdatedVersion() Then
        Call downloadLatestFileToCurrentFolder
        MsgBox "將關閉此檔案，並開啟新版程式。", vbOKOnly
        Workbooks.Open Filename:=Application.ActiveWorkbook.Path & Application.PathSeparator & WORKBOOK_NAME & " " & CURRENT_VERSION & ".xlsm"
        ThisWorkbook.Close SaveChanges:=False
    End If

    Call verifyPassword

    Exit Sub

' something error happen, for example
' no internet connect
' parse error
' 存在相同檔案名稱，寫入檔案失敗
ErrorHandler:
    MsgBox (Err.Description)

    inputPassword = Trim(Application.InputBox("Please input strong passward.", "No internet connect", Type:=2))

    strongPassword = "7U5uE+SMKg^@?qSJ"

    If inputPassword <> strongPassword Then
        MsgBox "Wrong strong password, excel will be close."
        ThisWorkbook.Close SaveChanges:=False
    End If

End Sub


Sub verifyPassword()
'
' 驗證密碼.
'
' @since 1.0.0
'

    Dim winHttpRequset As Object
    Dim res As Object
    Dim inputPassword As String
    Dim securityURL As String

    securityURL = "https://tammkyq18g.execute-api.ap-southeast-1.amazonaws.com/login"

    inputPassword = Trim(Application.InputBox("Please input passward.", "Verify user identity", Type:=2))

    Set winHttpRequset = CreateObject("WinHttp.WinHttpRequest.5.1")
    winHttpRequset.Open "POST", securityURL, False
    winHttpRequset.send "{""password"":" + inputPassword + "}"

    If winHttpRequset.Status <> 200 Then
        Set res = ran.ParseJSON(winHttpRequset.ResponseText)
        MsgBox res("obj.message")
        ThisWorkbook.Close SaveChanges:=False
    End If

End Sub


Function isOutdatedVersion()
'
' check version is latest.
'

    Dim res As Object
    Dim winHttpRequset As Object
    Dim latestVersion As String

    Set winHttpRequset = CreateObject("WinHttp.WinHttpRequest.5.1")

    winHttpRequset.Open "GET", VERSION_URL, False
    winHttpRequset.send
    Set res = ran.ParseJSON(winHttpRequset.ResponseText)

    latestVersion = res("obj.version")

    isOutdatedVersion = True
    If CURRENT_VERSION = latestVersion Then
        isOutdatedVersion = False
    End If

End Function

Sub downloadLatestFileToCurrentFolder()
'
' 下載新版本.
'
' @since 1.0.0
'

    Dim oStream As Object
    Dim winHttpRequset As Object

    Set winHttpRequset = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set oStream = CreateObject("ADODB.Stream")

    winHttpRequset.Open "GET", DOWNLOAD_URL, False
    winHttpRequset.send

    oStream.Open
    oStream.Type = 1
    oStream.Write winHttpRequset.responseBody
    oStream.SaveToFile(Application.ActiveWorkbook.Path & Application.PathSeparator & WORKBOOK_NAME & " " & CURRENT_VERSION & ".xlsm")
    oStream.Close
End Sub
