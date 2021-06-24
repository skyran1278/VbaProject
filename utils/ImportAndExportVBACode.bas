''' https://www.rondebruin.nl/win/s9/win002.htm
' In the VBE Editor set a reference to "Microsoft Visual Basic For Applications Extensibility 5.3" and to "Microsoft Scripting Runtime" and then save the file.
Option Explicit

Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If

    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)

    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If

    szExportPath = FolderWithVBAProjectFiles & "\"

    For Each cmpComponent In wkbSource.VBProject.VBComponents

        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                ' bExport = False

                ' if want to export ThisWorkbook
                szFileName = szFileName & "\ThisWorkbook" & ".cls"
                bExport = ExportThisWorkbook(cmpComponent)

        End Select

        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName

        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent

        End If

    Next cmpComponent

    MsgBox "Export done"
End Sub


Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents
    Dim szThisWorkbookPath As String

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)

    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"

    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents

    ' if want to import this workbook
    ImportThisWorkbook wkbTarget, szImportPath

    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files

        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If

    Next objFile

    MsgBox "Import Done"
End Sub


Function ExportThisWorkbook(cmpComponent As VBIDE.VBComponent) As Boolean
    Dim objFSO As Scripting.FileSystemObject
    Dim WorkbookPath As String
    Dim WorkbookName As String
    Dim thisWorkbookDir As String

    ExportThisWorkbook = False
    If cmpComponent.CodeModule.CountOfLines > 0 And cmpComponent.Type = vbext_ct_Document And cmpComponent.Name = "ThisWorkbook" Then
        Set objFSO = New Scripting.FileSystemObject

        WorkbookPath = Application.ActiveWorkbook.Path
        WorkbookName = objFSO.GetBaseName(ActiveWorkbook.Name)

        If Right(WorkbookPath, 1) <> "\" Then
            WorkbookPath = WorkbookPath & "\"
        End If

        thisWorkbookDir = WorkbookPath & WorkbookName & "\ThisWorkbook"

        If objFSO.FolderExists(thisWorkbookDir) = False Then
            On Error Resume Next
            MkDir thisWorkbookDir
            On Error GoTo 0
        End If

        ExportThisWorkbook = True
    End If

End Function

Function ImportThisWorkbook(wkbTarget As Excel.Workbook, szImportPath As String) As String
    Dim objFSO As Scripting.FileSystemObject
    Dim szThisWorkbookPath As String

    Set objFSO = New Scripting.FileSystemObject
    szThisWorkbookPath = szImportPath & "ThisWorkbook\ThisWorkbook.cls"

    If objFSO.FileExists(szThisWorkbookPath) Then
        ' cmpComponents.Import szThisWorkbookPath
        With wkbTarget.VBProject.VBComponents("ThisWorkbook").CodeModule
            .DeleteLines 1, .CountOfLines
            .AddFromFile szThisWorkbookPath

            If .Find("VERSION 1.0 CLASS", 1, 1, -1, -1) Then
                ' "VERSION 1.0 CLASS"
                ' BEGIN
                '   MultiUse = -1  'True
                ' END
                .DeleteLines 1, 4
            End If
        End With
    End If
End Function

Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim WorkbookPath As String
    Dim WorkbookName As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    WorkbookPath = Application.ActiveWorkbook.Path
    WorkbookName = FSO.GetBaseName(ActiveWorkbook.Name)

    If Right(WorkbookPath, 1) <> "\" Then
        WorkbookPath = WorkbookPath & "\"
    End If

    If FSO.FolderExists(WorkbookPath & WorkbookName) = False Then
        On Error Resume Next
        MkDir WorkbookPath & WorkbookName
        On Error GoTo 0
    End If

    If FSO.FolderExists(WorkbookPath & WorkbookName) = True Then
        FolderWithVBAProjectFiles = WorkbookPath & WorkbookName
    Else
        FolderWithVBAProjectFiles = "Error"
    End If

End Function

Function DeleteVBAModulesAndUserForms()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent

        Set VBProj = ActiveWorkbook.VBProject

        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function
