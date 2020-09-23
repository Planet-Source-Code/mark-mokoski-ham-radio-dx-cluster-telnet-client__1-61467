VERSION 5.00
Begin VB.Form frmTCPIP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DX Cluster Telnet Host Configure"
   ClientHeight    =   1785
   ClientLeft      =   2580
   ClientTop       =   1800
   ClientWidth     =   4815
   Icon            =   "frmTCPIP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1785
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Trace"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdOKCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   2520
      MouseIcon       =   "frmTCPIP.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frmTCPIP.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdOKCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   960
      MouseIcon       =   "frmTCPIP.frx":0A56
      MousePointer    =   99  'Custom
      Picture         =   "frmTCPIP.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Text            =   "23"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtRemoteName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "127.0.0.0"
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Port Number"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP Address or Host Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmTCPIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Private Sub Check1_Click()

        If Check1.Value = 1 Then
            frmTelnet.TraceTelnet = True
            frmTelnet.Tracevt100 = True
        Else
            frmTelnet.TraceTelnet = False
            frmTelnet.Tracevt100 = False
        End If
    
End Sub

Private Sub cmdOKCancel_Click(Index As Integer)

    Dim sHostName                   As String
    Dim txtRemoteAddress            As String

    On Error Resume Next

        Select Case Index
            Case 0
            
                frmTelnet.WinsockClient.Close
                frmTelnet.WinsockClient.LocalPort = 0
               
   
                If SocketsInitialize() Then
   
                    'pass the host name to the function
                    sHostName = txtRemoteName.Text
                    txtRemoteAddress = GetIPFromHostName(sHostName)
      
                    SocketsCleanup
      
                    'Set new info as system param's
                    DXtelnethost = txtRemoteName
                    DXtelnetport = txtPort
                Else
   
                    MsgBox "Windows Sockets for 32 bit Windows " & _
                    "is not successfully responding."
   
                End If
   

            frmTelnet.RemoteIPAd = txtRemoteAddress
            frmTelnet.RemotePort = txtPort
            frmTelnet.WinsockClient.RemotePort = txtPort
            frmTelnet.WinsockClient.RemoteHost = txtRemoteAddress

                If Err > 0 Then
                    MsgBox Error
                Else
                    Unload Me
                End If

            Case 1
                Unload Me
        End Select

End Sub

Private Sub Form_Load()

    '   txtRemoteName = frmTelnet.RemoteIPAd
    '   txtPort = frmTelnet.RemotePort
    txtRemoteName = DXtelnethost
    txtPort = DXtelnetport

    Check1.Value = -(frmTelnet.TraceTelnet)

End Sub
