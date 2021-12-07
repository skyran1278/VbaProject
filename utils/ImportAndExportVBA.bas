' https://www.rondebruin.nl/win/s9/win002.htm
'
' Import Export VBA Code
' �Ȥ䴩 �������O�Ҳ�, ���, �Ҳ�, ThisWorkbook
' ���䴩�ӧO�u�@���ץX
'
' �e������
' In the VBE Editor set a reference to "Microsoft Visual Basic For Applications Extensibility 5.3" and to "Microsoft Scripting Runtime" and then save the file.
' �����w���� > �H���s�� VBA �M�ת���ҫ�
'
' SOP
' �Ы� > Export > �ק� > Import
'
Option Explicit

Public Sub exportModules()
    Dim shouldExport As Boolean
    Dim sourceBook As Excel.Workbook
    Dim exportFolder As String
    Dim vbaFilePath As String
    Dim vbaComponent As VBIDE.VBComponent

    exportFolder = useBookNameAsFolder()
    createFolder (exportFolder)

    On Error Resume Next
    Kill exportFolder & "*"
    On Error GoTo 0

    Set sourceBook = Application.Workbooks(ActiveWorkbook.Name)

    If sourceBook.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected, not possible to export the code."
        Exit Sub
    End If

    For Each vbaComponent In sourceBook.VBProject.VBComponents
        shouldExport = False

        Select Case vbaComponent.Type
            Case vbext_ct_ClassModule
                shouldExport = True
                createFolder (exportFolder & "�������O�Ҳ�\")
                vbaFilePath = exportFolder & "�������O�Ҳ�\" & vbaComponent.Name & ".cls"
            Case vbext_ct_MSForm
                shouldExport = True
                createFolder (exportFolder & "���\")
                vbaFilePath = exportFolder & "���\" & vbaComponent.Name & ".frm"
            Case vbext_ct_StdModule
                shouldExport = True
                createFolder (exportFolder & "�Ҳ�\")
                vbaFilePath = exportFolder & "�Ҳ�\" & vbaComponent.Name & ".bas"
            Case vbext_ct_Document
                If vbaComponent.CodeModule.CountOfLines > 0 And vbaComponent.Name = "ThisWorkbook" Then
                    shouldExport = True
                    createFolder (exportFolder & "Microsoft Excel ����\")
                    vbaFilePath = exportFolder & "Microsoft Excel ����\" & vbaComponent.Name & ".cls"
                End If
        End Select

        If shouldExport Then
            vbaComponent.Export vbaFilePath
        End If
    Next vbaComponent

    MsgBox "Export Done"

End Sub

Public Sub importModules()
    Dim targetBook As Excel.Workbook
    Dim importFolder As String
    Dim folder As Variant
    Dim fileSystemObject As Object
    Dim vbaFile As Scripting.File

    Set fileSystemObject = New Scripting.fileSystemObject

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook, not possible to import in this workbook."
        Exit Sub
    End If

    Set targetBook = Application.Workbooks(ActiveWorkbook.Name)

    If targetBook.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected, not possible to export the code."
        Exit Sub
    End If

    importFolder = useBookNameAsFolder()

    Call deleteModules

    For Each folder In fileSystemObject.GetFolder(importFolder).SubFolders
        For Each vbaFile In fileSystemObject.GetFolder(folder).Files
            If vbaFile.Name = "ThisWorkbook.cls" Then
                With targetBook.VBProject.VBComponents("ThisWorkbook").CodeModule
                    .AddFromFile vbaFile.Path
                    If .Find("VERSION 1.0 CLASS", 1, 1, -1, -1) Then
                        ' "VERSION 1.0 CLASS"
                        ' BEGIN
                        '   MultiUse = -1  'True
                        ' END
                    .DeleteLines 1, 4
                    End If
                End With
            ElseIf (fileSystemObject.GetExtensionName(vbaFile.Name) = "cls") Or (fileSystemObject.GetExtensionName(vbaFile.Name) = "frm") Or (fileSystemObject.GetExtensionName(vbaFile.Name) = "bas") Then
                targetBook.VBProject.VBComponents.Import vbaFile.Path
            End If
        Next vbaFile
    Next folder

    MsgBox "Import Done"

End Sub

Function useBookNameAsFolder() As String
    Dim fileSystemObject As Object
    Dim workbookFolder As String
    Dim workbookName As String

    Set fileSystemObject = New Scripting.fileSystemObject

    workbookFolder = Application.ActiveWorkbook.Path
    workbookName = fileSystemObject.GetBaseName(ActiveWorkbook.Name)

    useBookNameAsFolder = workbookFolder & "\" & workbookName & "\"

End Function


Function createFolder(folder As String)
    Dim fileSystemObject As Object
    Set fileSystemObject = New Scripting.fileSystemObject

    If fileSystemObject.FolderExists(folder) = False Then
        MkDir folder
    End If

End Function

Function deleteModules()
    Dim vbaProject As VBIDE.VBProject
    Dim vbaComponent As VBIDE.VBComponent

    Set vbaProject = ActiveWorkbook.VBProject

    For Each vbaComponent In vbaProject.VBComponents
        If vbaComponent.Type = vbext_ct_Document Then
            If vbaComponent.Name = "ThisWorkbook" Then
                vbaComponent.CodeModule.DeleteLines 1, vbaComponent.CodeModule.CountOfLines
            End If
        Else
            vbaProject.VBComponents.Remove vbaComponent
        End If
    Next vbaComponent
End Function

