VERSION 5.00
Begin VB.Form frmDxAlert 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DX Alert"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmDxAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ShowDX 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Show DX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      MouseIcon       =   "frmDxAlert.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmDxAlert.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Left            =   5520
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1800
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      MouseIcon       =   "frmDxAlert.frx":091E
      MousePointer    =   99  'Custom
      Picture         =   "frmDxAlert.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton SetRadio 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set Radio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      MouseIcon       =   "frmDxAlert.frx":0F32
      MousePointer    =   99  'Custom
      Picture         =   "frmDxAlert.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "DX Alert !"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label DXlabel 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5775
   End
End
Attribute VB_Name = "frmDxAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Option Explicit

    Private DXfreq               As String
    Private DXfreqVal            As Long
    Private X                    As Integer

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Color change time in milliseconds
    Timer1.Interval = 200
    Timer1.Enabled = True
    'Alert window timeout in milliseconds
    Timer2.Interval = 60000  '1 minute
    Timer2.Enabled = True

    Static X

    X = 1
    'Disable SetRadio button if type"NONE"

        If UCase(RadioModel) = "NONE" Then
            SetRadio.Enabled = False
            SetRadio.Visible = True
            SetRadio.BackColor = &H8000000F
        Else
            SetRadio.Enabled = True
            SetRadio.Visible = True
        End If

    frmDxAlert.Visible = True
    frmDxAlert.SetFocus


End Sub

Public Sub AlertDX(CallDX, freqDX)

    DXfreq = freqDX
    DXlabel.Caption = CallDX & " on " & freqDX & " KHz"
    Timer2.Enabled = False  'Reset the timer if window already open
    Timer2.Enabled = True   'Restart timer for new DX Alert
    
End Sub

Private Sub SetRadio_Click()

    'Compute frequency in Hz
    DXfreqVal = Val(DXfreq) * 1000
    'Send DX frequency to radio
    
    'Hold off sending the Data until the last Message is sent

        If frmTelnet.ComRadio.PortOpen = True Then

                Do Until frmTelnet.ComRadio.PortOpen = False
                    DoEvents
                Loop

        End If
    
    Call modRadioComm.SendRadio(DXfreqVal)
    Unload Me
    
End Sub

Private Sub ShowDX_Click()

    frmTelnet.mnuRdxswindow_Click
    Command2_Click

End Sub

Private Sub Timer1_Timer()

        If X = 7 Then X = 1

        Select Case X
            Case 1
                Label2.ForeColor = vbGreen
            Case 2
                Label2.ForeColor = vbYellow
            Case 3
                Label2.ForeColor = vbBlue
            Case 4
                Label2.ForeColor = vbMagenta
            Case 5
                Label2.ForeColor = vbCyan
            Case 6
                Label2.ForeColor = vbWhite
        End Select

    X = X + 1

End Sub

Private Sub Timer2_Timer()

    'After 1 minute time out, close Alert form
    Unload Me

End Sub
