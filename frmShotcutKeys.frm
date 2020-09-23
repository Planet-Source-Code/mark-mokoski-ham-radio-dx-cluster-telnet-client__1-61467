VERSION 5.00
Begin VB.Form frmShortcutKeys 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DX Cluster Client - Shortcut Keys"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmShotcutKeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ShortcutOK 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "OK"
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
      Left            =   2040
      MouseIcon       =   "frmShotcutKeys.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frmShotcutKeys.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Caption         =   "Shortcuts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.TextBox Text1 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmShotcutKeys.frx":0A56
         Top             =   240
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmShortcutKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Private Sub ShortcutOK_Click()

    Unload frmShortcutKeys

End Sub
