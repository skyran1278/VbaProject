Attribute VB_Name = "version"

' 每次更新版本都需要修改
' 由於改成強制更新，所以要拉到 private，不讓 user 可以自己更改
Private Const CURRENT_VERSION = "3.0.1"

Public Function getVersion() As String
    getVersion = CURRENT_VERSION
End Function
