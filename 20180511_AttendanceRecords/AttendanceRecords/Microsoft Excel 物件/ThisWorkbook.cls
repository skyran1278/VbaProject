VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' @license Version v4.1.0
' In the VBE Editor set a reference to "Microsoft Visual Basic For Applications Extensibility 5.3" and to "Microsoft Scripting Runtime" and then save the file.
' 依賴 UTILS_CLASS.cls, version.bas
Option Explicit

' 隨工作簿不同而需更改的參數:
Private Const SERVICE_NAME = "AttendanceRecord"

Private Const BASE_URL = "https://tammkyq18g.execute-api.ap-southeast-1.amazonaws.com"
Private ran As New UTILS_CLASS

Private Sub Workbook_Open()
'
' 1. 驗證是否是最新版本，如果是舊版本就下載新版本、刪除舊版的程式，並且不再需要密碼登入
' 2. 驗證雲端密碼，錯誤就關閉 excel
' 3. 驗證強密碼，錯誤就關閉 excel
'
' * 測試環境:
'       office 2016 in windows 10
'       Mac 版本容易出現錯誤，不推薦在 Mac 執行

    Dim inputPassword As String
    Dim strongPassword As String
    Dim latestVersion As String
    Dim newFilePath As String

    On Error GoTo ErrorHandler

    latestVersion = getLatestVersion()

    If getVersion() <> latestVersion Then
        newFilePath = Application.ActiveWorkbook.Path & Application.PathSeparator & SERVICE_NAME & " " & latestVersion & ".xlsm"

        downloadLatestFileToCurrentFolder (newFilePath)

        MsgBox "檢查到新版本, 將開啟新版程式 (舊版本將無法執行程式).", vbOKOnly

        Workbooks.Open Filename:=newFilePath

        'Delete all modules/Userforms from ThisWorkbook
        Call DeleteVBAModulesAndUserForms

        ' 由於已經刪除程式了，所以不需要使用者再輸入雲端密碼
        Exit Sub
    End If

    If isInvalid() Then
        MsgBox ActiveWorkbook.Name & " will be close."
        ThisWorkbook.Close SaveChanges:=False
    End If

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
        MsgBox "Wrong strong password, " & ActiveWorkbook.Name & " will be close."
        ThisWorkbook.Close SaveChanges:=False
    End If

End Sub


Function isInvalid()
'
' 驗證密碼.
'
' @since 1.0.0
'

    Dim winHttpRequset As Object
    Dim res As Object
    Dim inputPassword As String
    Dim securityURL As String

    securityURL = BASE_URL & "/login"

    inputPassword = Trim(Application.InputBox("Please input passward.", "Verify user identity", Type:=2))

    Set winHttpRequset = CreateObject("WinHttp.WinHttpRequest.5.1")
    winHttpRequset.Open "POST", securityURL, False
    winHttpRequset.send "{""password"":" & inputPassword & "}"

    isInvalid = True
    If winHttpRequset.Status = 200 Then
        isInvalid = False
    Else
        Set res = ran.ParseJSON(winHttpRequset.ResponseText)
        MsgBox res("obj.message")
    End If

End Function


Function getLatestVersion()
'
' check version is latest.
'

    Dim res As Object
    Dim winHttpRequset As Object
    Dim versionURL As String

    versionURL = getFileURL(SERVICE_NAME & "/VERSION.json")

    Set winHttpRequset = CreateObject("WinHttp.WinHttpRequest.5.1")
    winHttpRequset.Open "GET", versionURL, False
    winHttpRequset.send

    Set res = ran.ParseJSON(winHttpRequset.ResponseText)
    getLatestVersion = res("obj.version")

End Function

Function downloadLatestFileToCurrentFolder(newFilePath)
'
' 下載新版本.
'
' @since 1.0.0
'

    Dim oStream As Object
    Dim winHttpRequset As Object
    Dim fileURL As String

    fileURL = getFileURL(SERVICE_NAME & "/" & SERVICE_NAME & ".xlsm")

    Set winHttpRequset = CreateObject("WinHttp.WinHttpRequest.5.1")
    winHttpRequset.Open "GET", fileURL, False
    winHttpRequset.send

    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write winHttpRequset.responseBody
    oStream.SaveToFile (newFilePath)
    oStream.Close
End Function

Function getFileURL(filePath)
'

    Dim winHttpRequset As Object

    Set winHttpRequset = CreateObject("WinHttp.WinHttpRequest.5.1")
    winHttpRequset.Open "GET", BASE_URL & "/file/" & filePath, False
    winHttpRequset.send

    getFileURL = winHttpRequset.ResponseText

End Function

Function DeleteVBAModulesAndUserForms()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ThisWorkbook.VBProject

    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            'Thisworkbook or worksheet module
            'We do nothing
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Function
