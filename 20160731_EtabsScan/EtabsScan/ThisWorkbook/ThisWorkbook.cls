VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' @license Version v3.0.2
' Version.vb
'
' Copyright (c) 2016-present, skyran
'
' This source code is licensed under the MIT license found in the
' LICENSE file in the root directory of this source tree.


' �H�u�@ï���P�ӻݧ�諸�Ѽ�:
' VERSION_URL: �Ӥu�@ï version.txt
' DOWNLOAD_URL: �Ӥu�@ï �U���ɮצ�m
Private Const VERSION_URL = "https://github.com/skyran1278/VbaProject/raw/master/20160731_EtabsScan/EtabsScanVersion.txt"
Private Const DOWNLOAD_URL = "https://github.com/skyran1278/VbaProject/raw/master/20160731_EtabsScan/EtabsScan.xlsm"

Private ran As New UTILS_CLASS

Private Sub Workbook_Open()
'
' * �ت�: ���ұK�X�A�ˬd�{���̷s�����A�æ۰ʴ��ܧ�s
'
' * �H�u�@ï���P�ӻݧ�諸�Ѽ�:
'       VERSION_URL: �Ӥu�@ï version.txt
'       DOWNLOAD_URL: �Ӥu�@ï �U���ɮצ�m
'
' * ���n�B�q�`���|��ʼƭ�:
'       �u�@����m: ������T
'       �W��: Cells(2, 3)
'       �ثe������: Cells(3, 3)
'       �̷s������: Cells(4, 3)
'
' * ��������:
'       office 2016 in windows 10
'       Mac �����e���X�{���~�A�����˦b Mac ����

    ' Dim ws_version As Worksheet
    ' Set ws_version = ThisWorkbook.Worksheets("Release Notes")

    On Error GoTo ErrorHandler

    Call VerifyPassword
    Call CheckVersion

    Exit Sub

ErrorHandler:
    MsgBox (Err.Description)

    inputPwd = Trim(Application.InputBox("Please Input Strong Passward.", "No Internet Connect", Type:=2))

    strongPwd = "7U5uE+SMKg^@?qSJ"

    If inputPwd <> strongPwd Then

        MsgBox "Wrong Password"
        ThisWorkbook.Close SaveChanges:=False

    End If

End Sub


Sub VerifyPassword()
'
' ���ұK�X.
'
' @since 1.0.0
'

    Dim srvXmlHttp As Object
    Dim inputPwd As String
    Dim cloudPwd As String
    Dim passwordUrl As String

    ' passwordUrl: pwd.txt
    passwordUrl = "https://oxlynrtpd4.execute-api.ap-southeast-1.amazonaws.com/login"

    inputPwd = Trim(Application.InputBox("Please Input Passward.", "Verify User Identity", Type:=2))

    Set winHttpRequset = CreateObject("WinHttp.WinHttpRequest.5.1")

    winHttpRequset.Open "POST", passwordUrl, False
    winHttpRequset.send "{""password"":" + inputPwd + "}"

    Set res = ran.ParseJSON(winHttpRequset.ResponseText)

    If winHttpRequset.Status <> 200 Then

        MsgBox res("obj.message")
        ThisWorkbook.Close SaveChanges:=False

    End If

End Sub


Sub CheckVersion()
'
' ���Ҫ�����.
'
' @since 1.0.0
'

    ' ���{�ǥ]�t���ܼ�
    Dim srvXmlHttp As Object
    Dim shell As Object
    Dim ws_version As Worksheet
    Dim currentVersion As String
    Dim latestVersion As String
    Dim releaseNotes As String

    Set srvXmlHttp = CreateObject("MSXML2.serverXMLHTTP")

    Set ws_version = ThisWorkbook.Worksheets("Release Notes")

    srvXmlHttp.Open "GET", VERSION_URL, False

    srvXmlHttp.send

    latestVersionAndReleaseNotes = srvXmlHttp.ResponseText

    currentVersion = ws_version.Cells(3, 3)

    ' �Ϥ��������M��s����
    latestVersionAndReleaseNotes = Split(latestVersionAndReleaseNotes, Chr(10) & "===" & Chr(10))
    latestVersion = latestVersionAndReleaseNotes(0)
    releaseNotes = latestVersionAndReleaseNotes(1)

    If CompareVersion(currentVersion, latestVersion) Then

        If MsgBox(releaseNotes, vbYesNo, "Download Latest Version From Browser") = vbYes Then

            Set shell = CreateObject("Wscript.Shell")
            shell.Run (DOWNLOAD_URL)
            MsgBox "Please close this file and use new file from browser.", vbOKOnly

        End If

    End If

End Sub


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

    ElseIf arrLatestVersion(0) < arrCurrentVersion(0) Then
        CompareVersion = False

    ' if  arrLatestVersion(0) == arrCurrentVersion(0)
    ElseIf arrLatestVersion(1) > arrCurrentVersion(1) Then
        CompareVersion = True

    ElseIf arrLatestVersion(1) < arrCurrentVersion(1) Then
        CompareVersion = False

    ' if  arrLatestVersion(1) == arrCurrentVersion(1)
    ElseIf arrLatestVersion(2) > arrCurrentVersion(2) Then
        CompareVersion = True

    ElseIf arrLatestVersion(2) < arrCurrentVersion(2) Then
        CompareVersion = False

    ' all equal
    Else
        CompareVersion = False

    End If

End Function

