Attribute VB_Name = "sercurity"

Option Explicit

' �C����s�������ݭn�ק�
' �ѩ�令�j���s�A�ҥH�n�Ԩ� private�A���� user �i�H�ۤv���
Private Const CURRENT_VERSION = "3.1.1"

' �H�u�@ï���P�ӻݧ�諸�Ѽ�:
Private Const SERVICE_NAME = "AttendanceRecord"

Private Const BASE_URL = "https://tammkyq18g.execute-api.ap-southeast-1.amazonaws.com"
Private ran As New UTILS_CLASS

Public Sub security()
'
' 1. ���ҬO�_�O�̷s�����A�p�G�O�ª����N�U���s�����B�R���ª����{���A�åB���A�ݭn�K�X�n�J
' 2. ���Ҷ��ݱK�X�A���~�N���� excel
' 3. ���ұj�K�X�A���~�N���� excel
'
' * ��������:
'       office 2016 in windows 10
'       mac �����e���X�{���~�A�����˦b mac ����

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

        MsgBox "�ˬd��s����, �N�}�ҷs���{�� (�ª����N�L�k����{��).", vbOKOnly

        Workbooks.Open Filename:=newFilePath

        ' �ѩ�w�g�U���s�{���A�µ{���z�פW�L�k����F�A�ҥH���ݭn�ϥΪ̦A��J���ݱK�X
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
' - �s�b�ۦP�ɮצW�١A�g�J�ɮץ���
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
' ���ұK�X.
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
' �U���s����.
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
        MsgBox "�������榡���~", vbOKOnly
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
