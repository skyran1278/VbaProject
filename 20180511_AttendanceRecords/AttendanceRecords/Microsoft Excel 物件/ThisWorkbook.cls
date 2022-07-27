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
' ¨Ì¿à UTILS_CLASS.cls, version.bas
Option Explicit

Private Sub Workbook_Open()
    Call security
End Sub
