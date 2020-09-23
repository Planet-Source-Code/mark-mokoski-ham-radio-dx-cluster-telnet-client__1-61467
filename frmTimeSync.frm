VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTimeSync 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DX Telnet Client - Time Synchronizer"
   ClientHeight    =   1560
   ClientLeft      =   4785
   ClientTop       =   2595
   ClientWidth     =   4335
   Icon            =   "frmTimeSync.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer SyncTimeOut 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   720
      Top             =   1800
   End
   Begin VB.CommandButton MinSysTray 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Min to Tray"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      MouseIcon       =   "frmTimeSync.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frmTimeSync.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click to Minimize Time Sync to System Tray"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Timer tmrTimer 
      Interval        =   60000
      Left            =   120
      Top             =   1800
   End
   Begin MSWinsockLib.Winsock WinSock 
      Left            =   1320
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   37
   End
   Begin VB.CommandButton cmdSynchronizeNow 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Synchronize Now"
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
      Left            =   120
      MouseIcon       =   "frmTimeSync.frx":0B8E
      MousePointer    =   99  'Custom
      Picture         =   "frmTimeSync.frx":0E98
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click to Synchronize Computer Clock Now"
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
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
      ToolTipText     =   "Status"
      Top             =   120
      Width           =   4095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuSyncNow 
         Caption         =   "Sync Time Now"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Time Sync Window"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTSclose 
         Caption         =   "Close this Menu"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEndSync 
         Caption         =   "End Time Sync"
      End
   End
End
Attribute VB_Name = "frmTimeSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

        Private Type SystemTime
            intYear                   As Integer
            intMonth                  As Integer
            intWeekDay                As Integer
            intDay                    As Integer
            intHour                   As Integer
            intMinute                 As Integer
            intSecond                 As Integer
            intMillisecond            As Integer
        End Type

    Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SystemTime) As Long

    Private intCounter                As Integer
    Private intCloseMethod            As Integer
    Private sngDelay                  As Single
    Private strTime                   As String

Private Sub cmdMinimize_Click()

        If Me.WindowState = vbMinimized Then
            frmTimeSync.Hide
        End If

End Sub

Private Sub cmdSynchronizeNow_Click()

    Dim intFileNumber            As Integer
    Dim strServer                As String

    'Set connection time out timer
    SyncTimeOut.Enabled = True
    
    intFileNumber = FreeFile
    'strServer = "ntps1-0.cs.tu-berlin.de"
    strServer = "time-a.nist.gov"
    On Error Resume Next

        If Right$(App.Path, 1) = "\" Then
            Open App.Path & "settings.txt" For Input Lock Write As #intFileNumber
            Line Input #intFileNumber, strServer
            Close #intFileNumber
        Else
            Open App.Path & "\settings.txt" For Input Lock Write As #intFileNumber
            Line Input #intFileNumber, strServer
            Close #intFileNumber
        End If

    On Error GoTo SyncERROR
    cmdSynchronizeNow.Enabled = False
    cmdSynchronizeNow.BackColor = &H8000000F
    lblStatus.Caption = "Synchronizing..."
    WinSock.Close
    strTime = Empty
    WinSock.RemoteHost = strServer
    WinSock.Connect

        Do Until WinSock.State = sckConnected Or SyncTimeOut.Enabled = False
            DoEvents
        Loop

        If WinSock.State = sckConnected Then
            SyncTimeOut.Enabled = False
            Exit Sub
        End If
    
SyncERROR:
    cmdSynchronizeNow.Enabled = True
    lblStatus.Caption = "Synchronizing ERROR"
    SyncTimeOut.Enabled = False
    WinSock.Close

End Sub

Private Sub Form_Load()

    ' Place this in the Form Load event of the form you want to disable the 'X':

    Dim hSysMenu                 As Long

    hSysMenu = GetSystemMenu(HWND, False)
    RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
    
        If App.PrevInstance Then
            intCloseMethod = 1
            Unload Me
        End If

    Call SystrayOn(frmTimeSync, "DX Telnet Client - Time Sync")
    frmTelnet.mnuTimeSync.Enabled = True
    cmdSynchronizeNow_Click

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

    Static lngMsg                As Long
    Dim blnflag                  As Boolean
    Dim lngResult                As Long






    lngMsg = X / Screen.TwipsPerPixelX

        If blnflag = False Then
            blnflag = True
        
                Select Case lngMsg
                    Case WM_RBUTTONCLK      'to popup on right-click
                        Call SetForegroundWindow(Me.HWND)
                        Call RemoveBalloon(frmSystray)
                        PopupMenu frmTimeSync.mnuFile

                    Case WM_LBUTTONDBLCLK   'open on left-dblclick
                        'Call SystrayOff(frmsystray)
                        Call SetForegroundWindow(Me.HWND)
                        frmTimeSync.WindowState = vbNormal
                        frmTimeSync.Show
            
                End Select
        
            blnflag = False
        
        End If
    
End Sub


Private Sub Form_Resize()

        If Me.WindowState = vbMinimized Then
            frmTimeSync.Hide
        End If

End Sub

Private Sub Form_Terminate()

    Call SystrayOff(frmTimeSync)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SystrayOff(frmTimeSync)

End Sub

Private Sub SynchronizeClock(strTimeString As String)

    Dim datTime                  As Date
    Dim dblTime                  As Double
    Dim lngTime                  As Long
    Dim udtSystemDate            As SystemTime

    strTimeString = Trim$(strTimeString)

        If Len(strTimeString) <> 4 Then
            cmdSynchronizeNow.Enabled = True
            cmdSynchronizeNow.BackColor = &HC0C0C0
            lblStatus.Caption = "Error synchronizing!"
            Exit Sub
        End If

    dblTime = Asc(Left$(strTimeString, 1)) * 256 ^ 3 + Asc(Mid$(strTimeString, 2, 1)) * 256 ^ 2 + Asc(Mid$(strTimeString, 3, 1)) * 256 ^ 1 + Asc(Right$(strTimeString, 1))
    lngTime = dblTime - 2840140800#
    datTime = DateAdd("s", CDbl(lngTime + CLng(sngDelay)), #1/1/1990#)
    udtSystemDate.intYear = Year(datTime)
    udtSystemDate.intMonth = Month(datTime)
    udtSystemDate.intDay = Day(datTime)
    udtSystemDate.intHour = Hour(datTime)
    udtSystemDate.intMinute = Minute(datTime)
    udtSystemDate.intSecond = Second(datTime)
    Call SetSystemTime(udtSystemDate)
    cmdSynchronizeNow.Enabled = True
    cmdSynchronizeNow.BackColor = &HC0C0C0
    lblStatus.Caption = "Last update: " & UTCdate & " at " & UTCtime & " UTC "

End Sub

Private Sub MinSysTray_Click()

    'Hide Time Sync Window and keep icon in Systray visible
    frmTimeSync.Hide

End Sub

Private Sub mnuEndSync_Click()

    Unload frmTimeSync

End Sub

Private Sub mnuRestore_Click()

    frmTimeSync.WindowState = vbNormal
    frmTimeSync.Visible = True

End Sub

Private Sub mnuSyncNow_Click()

    cmdSynchronizeNow_Click

End Sub

Public Sub mnuTSclose_Click()

End Sub

Private Sub SyncTimeOut_Timer()

    SyncTimeOut.Enabled = False

End Sub

Private Sub tmrTimer_Timer()

    intCounter = intCounter + 1

        If intCounter = 60 Then
            intCounter = 0
            cmdSynchronizeNow_Click
        End If

End Sub

Private Sub WinSock_Close()

    On Error Resume Next

        Do Until WinSock.State = sckClosed
            WinSock.Close
            DoEvents
        Loop

    sngDelay = ((Timer - sngDelay) / 2)
    Call SynchronizeClock(strTime)
    On Error GoTo 0

End Sub

Private Sub WinSock_Connect()
    
    DoEvents
    sngDelay = Timer

End Sub

Private Sub WinSock_DataArrival(ByVal bytesTotal As Long)

    Dim strData                  As String

    WinSock.GetData strData, vbString
    strTime = strTime & strData

End Sub
