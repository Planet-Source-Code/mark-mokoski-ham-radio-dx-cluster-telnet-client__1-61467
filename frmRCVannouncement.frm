VERSION 5.00
Begin VB.Form frmRCVannouncement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DX Cluster Telnet Client - Announcement Received"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmRCVannouncement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer AnncTimer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   300
      Top             =   1680
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
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
      Height          =   1095
      Left            =   4440
      MouseIcon       =   "frmRCVannouncement.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frmRCVannouncement.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Send Announcement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1560
      MouseIcon       =   "frmRCVannouncement.frx":0A56
      MousePointer    =   99  'Custom
      Picture         =   "frmRCVannouncement.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Frame RCVframe 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.Label AccounceText 
         Caption         =   "Label1"
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
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmRCVannouncement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Private Sub AnncTimer_Timer()

    'Window active for one minute
    AnncTimer.Enabled = False
    Unload Me

End Sub

Private Sub Command1_Click()

    'Open Send Announcement window
    frmAnnounce.Visible = True

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Start one minute timeout timer
    AnncTimer.Interval = 60000  'one minute timeout
    AnncTimer.Enabled = True

End Sub
