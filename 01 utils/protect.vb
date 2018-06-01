
Private Sub func()
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'

    Set srvXmlHttp = CreateObject("MSXML2.serverXMLHTTP")

    srvXmlHttp.Open "GET", "https://github.com/skyran1278/VbaProject/raw/master/20170413_BeamZValue/z-value.txt", False

    input_pwd = Trim(Application.InputBox("Please Input Passward.", "Verify User Identity", type:=2))

    srvXmlHttp.send

    cloud_pwd = srvXmlHttp.ResponseText
    ' 消除空白行
    cloud_pwd = Trim(Replace(cloud_pwd, Chr(10), ""))

    If input_pwd = cloud_pwd Then

        MsgBox "Sign In Success"

    Else

        MsgBox "Wrong Password"
        ThisWorkbook.Close SaveChanges:=False

    End If

End Sub


