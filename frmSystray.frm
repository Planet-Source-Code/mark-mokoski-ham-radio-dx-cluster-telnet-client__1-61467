VERSION 5.00
Begin VB.Form frmSystray 
   Caption         =   "DX Telnet Client - Systray"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3585
   Icon            =   "frmSystray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmSystray.frx":030A
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Private Sub Form_Load()

    Call SystrayOn(frmSystray, "DX Cluster Telnet Client - Not Connected")
    DXspotCaption = "Not Connected"
    frmTelnet.mnuShowSpots.Caption = "DX Cluster Telnet Client - " + DXspotCaption
    DXwindow.Caption = "DX Cluster Telnet Client - " + DXspotCaption
    
End Sub

Private Sub Form_Terminate()

    Call SystrayOff(frmSystray)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SystrayOff(frmSystray)

    'Force terminate program
    End

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

    Static lngMsg            As Long
    Dim blnflag              As Boolean
    Dim lngResult            As Long
    



    lngMsg = X / Screen.TwipsPerPixelX

        If blnflag = False Then
            blnflag = True
        
                Select Case lngMsg
                    Case WM_RBUTTONCLK      'to popup on right-click
                        Call SetForegroundWindow(Me.HWND)
                        Call RemoveBalloon(Me)
                        PopupMenu frmTelnet.mnuRestore

                    Case WM_LBUTTONDBLCLK   'open on left-dblclick
                        'Call SystrayOff(frmsystray)
                        Call SetForegroundWindow(Me.HWND)
                        Call RemoveBalloon(Me)
                        frmTelnet.WindowState = vbNormal
                        frmTelnet.Show
            
                        If DXwindow.Visible = False Then
                            'if form is hiden, show window and update grid
                            DXwindow.WindowState = vbNormal
                            DXwindow.Show
                            Call ViewDXHeardList
                            Call ViewWWVHeardList
                        Else
                            'If form loaded and in up or in task bar, dont update grid
                            DXwindow.WindowState = vbNormal
                            DXwindow.Show
                        End If
            
                    'If Telnet connected to cluster, set DX window ith focus

                        If frmTelnet.WinsockClient.State > 0 Then
                            DXwindow.SetFocus
                        Else
                            frmTelnet.SetFocus
                        End If

                End Select
        
            blnflag = False
        
        End If
    
End Sub
