VERSION 5.00
Begin VB.Form frmAnnounce 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DX Telnet Client - Send Announcement"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmAnnounce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ANNtype 
      Height          =   315
      Left            =   2160
      MouseIcon       =   "frmAnnounce.frx":0742
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Text            =   "ALL NODES"
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton AnnounceCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      MouseIcon       =   "frmAnnounce.frx":0A4C
      MousePointer    =   99  'Custom
      Picture         =   "frmAnnounce.frx":0D56
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton AnnounceOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Send Announcement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      MouseIcon       =   "frmAnnounce.frx":1060
      MousePointer    =   99  'Custom
      Picture         =   "frmAnnounce.frx":136A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox AnnounceText 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Send Announcement to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Type announcement in text box below"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmAnnounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Private Sub AnnounceCancel_Click()

    Unload frmAnnounce

End Sub

Private Sub AnnounceOK_Click()

    Dim TypeText

    'If there is a comment, add it to announce out string, if null, end sub

        If AnnounceText.Text <> "" Then
            frmTelnet.Visible = True
            frmTelnet.WindowState = vbNormal
            frmTelnet.SetFocus
    
                Select Case ANNtype.Text
    
                    Case "ALL NODES/USERS"
                        TypeText = "announce/all"
                    Case Else
                        TypeText = "announce"
            
                End Select
    
            frmTelnet.WinsockClient.SendData (vbCrLf & TypeText & " " & AnnounceText.Text & vbCrLf)
    
        End If

    Unload frmAnnounce

End Sub

Private Sub AnnounceText_Change()

        If AnnounceText.Text <> "" Then
            AnnounceOK.BackColor = &HC0C0C0
            AnnounceOK.Enabled = True
        Else
            AnnounceOK.BackColor = &H8000000F
            AnnounceOK.Enabled = False

        End If

        If Len(AnnounceText.Text) >= 50 Then
            AnnounceText.Text = Mid$(AnnounceText.Text, 1, 50)
            AnnounceText.SelStart = Len(AnnounceText.Text) + 1
        End If

End Sub

Private Sub Form_Load()
    
    AnnounceOK.BackColor = &H8000000F
    AnnounceOK.Enabled = False
    
    ANNtype.AddItem "ALL NODES/USERS", 0
    ANNtype.AddItem "LOCAL NODE/USERS", 1
    ANNtype.Text = "ALL NODES/USERS"

End Sub
