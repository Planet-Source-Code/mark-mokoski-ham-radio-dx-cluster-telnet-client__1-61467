VERSION 5.00
Begin VB.Form frmDXspot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DX Spot"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmDXspot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timeout 
      Interval        =   5000
      Left            =   2040
      Top             =   1920
   End
   Begin VB.CommandButton DXuseRadio 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Use Radio Frequency"
      Height          =   975
      Left            =   3240
      MouseIcon       =   "frmDXspot.frx":1272
      MousePointer    =   99  'Custom
      Picture         =   "frmDXspot.frx":157C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton DXspotCancel 
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
      Left            =   2640
      MouseIcon       =   "frmDXspot.frx":1886
      MousePointer    =   99  'Custom
      Picture         =   "frmDXspot.frx":1B90
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton DXspotOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Send DX Spot"
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
      MouseIcon       =   "frmDXspot.frx":1E9A
      MousePointer    =   99  'Custom
      Picture         =   "frmDXspot.frx":21A4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox DXcomment 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox DXfreq 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox DXstation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Comment"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "DX Frequency (KHz)"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "DX Station"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDXspot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '*******************************************************************************
    '
    '   Send DX Spot Form
    '
    '   Get data from form and send DX Spot to Cluster
    '
    '   Mark Mokoski
    '   21-MAY-2004
    '
    '*******************************************************************************



Private Sub DXfreq_Change()

        If DXstation.Text <> "" And DXfreq.Text <> "" And Val(DXfreq.Text) >= 1800 Then
            DXspotOK.BackColor = &HC0C0C0
            DXspotOK.Enabled = True
        Else
            DXspotOK.BackColor = &H8000000F
            DXspotOK.Enabled = False

        End If

End Sub

Private Sub DXspotCancel_Click()

    Unload frmDXspot

End Sub

Private Sub DXspotOK_Click()

    'Test for null feilds

        If DXstation.Text <> "" And DXfreq.Text <> "" And Val(DXfreq.Text) >= 1800 Then
            frmTelnet.Visible = True
            frmTelnet.WindowState = vbNormal
            frmTelnet.SetFocus

                If DXcomment.Text <> "" Then
                    'If there is a comment, add it to dx out string
                    frmTelnet.WinsockClient.SendData (vbCrLf & "dx" & " " & DXstation.Text & " " & DXfreq.Text & " " & DXcomment.Text & vbCrLf)
                Else
                    'If there is NO comment, send this string
                    frmTelnet.WinsockClient.SendData (vbCrLf & "dx" & " " & DXstation.Text & " " & DXfreq.Text & vbCrLf)
                End If
    
        End If

    Unload frmDXspot

End Sub

Private Sub DXstation_Change()

        If DXstation.Text <> "" And DXfreq.Text <> "" Then
            DXspotOK.BackColor = &HC0C0C0
            DXspotOK.Enabled = True
        Else
            DXspotOK.BackColor = &H8000000F
            DXspotOK.Enabled = False

        End If

    DXstation.SelStart = Len(DXstation.Text) + 1
    DXstation.Text = UCase(DXstation.Text)

End Sub

Private Sub DXuseRadio_Click()

    'Get Spot Frequency from Radio

    'Hold off sending the Data until the last Message is sent

        If frmTelnet.ComRadio.PortOpen = True Then

                Do Until frmTelnet.ComRadio.PortOpen = False
                    DoEvents
                Loop

        End If

    'Enable Time Out Timer (5 sec)
    Timeout.Enabled = True
    'Get Frequency from function
    DXfreq = modRadioComm.RadioFreq

End Sub

Private Sub Form_Load()

        If RadioModel = "None" Then
            DXuseRadio.Enabled = False
            DXuseRadio.BackColor = &H8000000F
        Else
            DXuseRadio.Enabled = True
        End If

    DXspotOK.BackColor = &H8000000F
    DXspotOK.Enabled = False

End Sub

Private Sub Form_Terminate()

    'Close COM port just in case

        If frmTelnet.ComRadio.PortOpen = True Then
            frmTelnet.ComRadio.PortOpen = False
        End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Close COM port just in case

        If frmTelnet.ComRadio.PortOpen = True Then
            frmTelnet.ComRadio.PortOpen = False
        End If

End Sub

Private Sub Timeout_Timer()

    '5 second timeout timer for getting data from radio

End Sub
