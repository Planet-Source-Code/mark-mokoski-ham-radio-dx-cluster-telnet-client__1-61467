Attribute VB_Name = "modLogin"
    Option Explicit

    '********************************************************************************
    '
    '   monLogin
    '   Auto Login Telnet session with stored User Name and Password (optional)
    '
    '   Mark Mokoski
    '   18-MAY-2004
    '
    '********************************************************************************

Public Sub Autologon(infotext)

        Select Case (infotext)
            'Send user logon to winsock
            Case "LOGIN:"
                frmTelnet.WinsockClient.SendData (TelnetLogon & vbCrLf)
                Login = False
                'DXwindow.SetFocus
                'Send user password to winsock
            Case "PASSWORD:"
                frmTelnet.WinsockClient.SendData (TelnetLogon & vbCrLf)
                Passwrd = False
                'DXwindow.SetFocus
        End Select


End Sub
