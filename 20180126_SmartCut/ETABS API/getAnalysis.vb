Sub Main()

    'set the following flag to True to attach to an existing instance of the program
    'otherwise a new instance of the program will be started
    Dim AttachToInstance As Boolean
    AttachToInstance = False

    'set the following flag to True to manually specify the path to ETABS.exe
    'this allows for a connection to a version of ETABS other than the latest installation
    'otherwise the latest installed version of ETABS will be launched
    Dim SpecifyPath As Boolean
    SpecifyPath = False

    'if the above flag is set to True, specify the path to ETABS below
    Dim ProgramPath As String
    'ProgramPath = "C:\Program Files (x86)\Computers and Structures\ETABS 2016\ETABS.exe"

    'full path to the model
    'set it to the desired path of your model
    Dim ModelDirectory As String
    ModelDirectory = "D:\Paul\ETABS"
    If Len(Dir(ModelDirectory, vbDirectory)) = 0 Then
        MkDir ModelDirectory
    End If

    Dim ModelName As String
    ModelName = "20181224 IDA v1.2.EDB"

    Dim ModelPath As String
    ModelPath = ModelDirectory & Application.PathSeparator & ModelName

    'create API helper object
    Dim myHelper As cHelper
    Set myHelper = New Helper

    'dimension the ETABS Object as cOAPI type
    Dim myETABSObject As cOAPI
    Set myETABSObject = Nothing

    'use ret to check return values of API calls
    Dim ret As Long

    If AttachToInstance Then
        'attach to a running instance of ETABS
        'get the active ETABS object
        Set myETABSObject = GetObject(, "CSI.ETABS.API.ETABSObject")

    Else
        If SpecifyPath Then
        'create an instance of the ETABS object from the specified path
        Set myETABSObject = myHelper.CreateObject(ProgramPath)
        Else
        'create an instance of the ETABS object from the latest installed ETABS
        Set myETABSObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
        End If
        'start ETABS application
        myETABSObject.ApplicationStart

    End If

    'get a reference to cSapModel to access all OAPI classes and functions
    Dim mySapModel As ETABS2016.cSapModel
    Set mySapModel = myETABSObject.SapModel

    'initialize model
    ret = mySapModel.InitializeNewModel()

    'open an existing file - If no file exists, run the Save example first.
    ret = mySapModel.File.OpenFile(ModelPath)

    'run analysis
    ret = mySapModel.Analyze.RunAnalysis()

    'close ETABS
    myETABSObject.ApplicationExit(False)

    'clean up variables
    Set mySapModel = Nothing
    Set myETABSObject = Nothing

    If ret = 0 Then
        MsgBox "API script completed successfully."
    Else
        MsgBox "API script FAILED to complete."
    End If

    Exit Sub

    ErrHandler:
        MsgBox "Cannot run API script: " & Err.Description

End Sub



