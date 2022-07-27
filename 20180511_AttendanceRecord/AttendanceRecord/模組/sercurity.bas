Attribute VB_Name = "sercurity"

Option Explicit

' 每次更新版本都需要修改
' 由於改成強制更新，所以要拉到 private，不讓 user 可以自己更改
Private Const CURRENT_VERSION = "3.1.1"

' 隨工作簿不同而需更改的參數:
Private Const SERVICE_NAME = "AttendanceRecord"

Private Const BASE_URL = "https://tammkyq18g.execute-api.ap-southeast-1.amazonaws.com"
Private ran As New UTILS_CLASS

Public Sub security()
'
' 1. 驗證是否是最新版本，如果是舊版本就下載新版本、刪除舊版的程式，並且不再需要密碼登入
' 2. 驗證雲端密碼，錯誤就關閉 excel
' 3. 驗證強密碼，錯誤就關閉 excel
'
' * 測試環境:
'       office 2016 in windows 10
'       mac 版本容易出現錯誤，不推薦在 mac 執行

    Dim inputPassword As String
    Dim strongPassword As String
    Dim latestVersion As String
    Dim newFilePath As String

    On Error GoTo errorHandler

    ' https://stackoverflow.com/questions/20361241/how-do-i-delete-module1-using-vba
    ' ThisWorkbook.VBProject.References.AddFromGuid _
    '     GUID:="{0002E157-0000-0000-C000-000000000046}", _
    '     Major:=5, Minor:=3

    latestVersion = getLatestVersion()

    If Not isGreaterOrEqualToRequireVersion(getVersion(), getLatestVersion()) Then
        newFilePath = Application.ActiveWorkbook.Path & Application.PathSeparator & SERVICE_NAME & " " & latestVersion & ".xlsm"

        downloadLatestFileToCurrentFolder (newFilePath)

        MsgBox "檢查到新版本, 將開啟新版程式 (舊版本將無法執行程式).", vbOKOnly

        Workbooks.Open Filename:=newFilePath

        ' 由於已經下載新程式，舊程式理論上無法執行了，所以不需要使用者再輸入雲端密碼
        Exit Sub
    End If

    If isInvalidPassword() Then
        MsgBox ActiveWorkbook.Name & " will be close."
        ThisWorkbook.Close SaveChanges:=False
    End If

    Exit Sub

' something error happen, for example
' - no internet connect
' - parse error
' - 存在相同檔案名稱，寫入檔案失敗
errorHandler:
    MsgBox (Err.Description)

    inputPassword = Trim(Application.InputBox("Please input STRONG passward.", "Error", Type:=2))

    strongPassword = "7U5uE+SMKg^@?qSJ"

    If inputPassword <> strongPassword Then
        MsgBox "Wrong STRONG password, " & ActiveWorkbook.Name & " will be close."
        ThisWorkbook.Close SaveChanges:=False
    End If

End Sub



Public Function getVersion() As String
    getVersion = CURRENT_VERSION
End Function

Public Function getLatestVersion() As String
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

Private Function getFileURL(filePath) As String
'

    Dim winHttpRequset As Object

    Set winHttpRequset = CreateObject("WinHttp.WinHttpRequest.5.1")
    winHttpRequset.Open "GET", BASE_URL & "/file/" & filePath, False
    winHttpRequset.send

    getFileURL = winHttpRequset.ResponseText

End Function


Private Function isInvalidPassword() As Boolean
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

    isInvalidPassword = True
    If winHttpRequset.Status = 200 Then
        isInvalidPassword = False
    Else
        Set res = ran.ParseJSON(winHttpRequset.ResponseText)
        MsgBox res("obj.message")
    End If

End Function

Private Function downloadLatestFileToCurrentFolder(newFilePath)
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

Private Function DeleteVBAModulesAndUserForms()
' not working in protected mode

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

Public Function isGreaterOrEqualToRequireVersion(currentVersion As String, requireVersion As String) As Boolean
'
' compare which version is latest.
'
' @since 1.0.0
' @param {string} [currentVersion] currentVersion.
' @param {string} [requireVersion] requireVersion.
' @return {boolean} [isGreaterOrEqualToRequireVersion] if currentVersion > requireVersion, return true.
'

    Dim arrCurrentVersion As Variant
    Dim arrRequireVersion As Variant
    Dim arrCurrentVersionLength As Integer
    Dim arrRequireVersionLength As Integer

    arrCurrentVersion = Split(currentVersion, ".")
    arrRequireVersion = Split(requireVersion, ".")

    arrCurrentVersionLength = UBound(arrCurrentVersion) - LBound(arrCurrentVersion) + 1
    arrRequireVersionLength = UBound(arrRequireVersion) - LBound(arrRequireVersion) + 1

    If arrCurrentVersionLength <> 3 Or arrRequireVersionLength <> 3 Then
        MsgBox "版本號格式錯誤", vbOKOnly
        isGreaterOrEqualToRequireVersion = False
        Exit Function
    End If

    If arrCurrentVersion(0) > arrRequireVersion(0) Then
        isGreaterOrEqualToRequireVersion = True
        Exit Function
    ElseIf arrCurrentVersion(0) < arrRequireVersion(0) Then
        isGreaterOrEqualToRequireVersion = False
        Exit Function
    End If

    If arrCurrentVersion(1) > arrRequireVersion(1) Then
        isGreaterOrEqualToRequireVersion = True
        Exit Function
    ElseIf arrCurrentVersion(1) < arrRequireVersion(1) Then
        isGreaterOrEqualToRequireVersion = False
        Exit Function
    End If

    If arrCurrentVersion(2) > arrRequireVersion(2) Then
        isGreaterOrEqualToRequireVersion = True
        Exit Function
    ElseIf arrCurrentVersion(2) < arrRequireVersion(2) Then
        isGreaterOrEqualToRequireVersion = False
        Exit Function
    End If

    isGreaterOrEqualToRequireVersion = True

End Function
