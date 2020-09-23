VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmTelnet 
   BackColor       =   &H80000017&
   Caption         =   "DX Cluster Telnet Client"
   ClientHeight    =   7155
   ClientLeft      =   -210
   ClientTop       =   3015
   ClientWidth     =   9675
   FillColor       =   &H00800000&
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmTelnet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer RealTime 
      Interval        =   250
      Left            =   7080
      Top             =   600
   End
   Begin MSWinsockLib.Winsock WinsockClient 
      Left            =   6120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer cursor_timer 
      Interval        =   600
      Left            =   6600
      Top             =   600
   End
   Begin ComctlLib.StatusBar stbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   6660
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2469
            MinWidth        =   2469
            Picture         =   "frmTelnet.frx":030A
            Text            =   "Closed"
            TextSave        =   "Closed"
            Key             =   "Connection"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Connection Status"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "frmTelnet.frx":0624
            Text            =   "Shift+F2 to Connect to Cluster"
            TextSave        =   "Shift+F2 to Connect to Cluster"
            Key             =   "TelnetHost"
            Object.Tag             =   ""
            Object.ToolTipText     =   "DX Cluster Host"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6703
            MinWidth        =   6703
            Picture         =   "frmTelnet.frx":093E
            Text            =   "DX Cluster Telnet Client Ready !"
            TextSave        =   "DX Cluster Telnet Client Ready !"
            Key             =   "Status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Status Message"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Object.Width           =   2099
            MinWidth        =   2099
            Picture         =   "frmTelnet.frx":0C58
            Text            =   "To Tray"
            TextSave        =   "To Tray"
            Key             =   "Min"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click to Minimize Telnet Window to System Tray"
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmTelnet.frx":0F72
   End
   Begin MSCommLib.MSComm ComRadio 
      Left            =   7560
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RTSEnable       =   -1  'True
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect to DX Cluster"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Connection to DX Cluster"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinAll 
         Caption         =   "Minimize All Windows to System Tray"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "End Program"
      End
   End
   Begin VB.Menu mOpen 
      Caption         =   "&Connect"
   End
   Begin VB.Menu mClose 
      Caption         =   "Di&sconnect"
   End
   Begin VB.Menu mnuShow 
      Caption         =   "S&how"
      Begin VB.Menu mnuShowDX 
         Caption         =   "Show DX"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu mnuShowWWV 
         Caption         =   "Show WWV"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnuShowUsers 
         Caption         =   "Show Users"
         Shortcut        =   ^{F9}
      End
   End
   Begin VB.Menu mnuSend 
      Caption         =   "S&end"
      Begin VB.Menu mnuDXspot 
         Caption         =   "DX Spot"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuWWVspot 
         Caption         =   "WWV Spot"
         Enabled         =   0   'False
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuMessage 
         Caption         =   "Anouncement"
         Shortcut        =   +{F9}
      End
   End
   Begin VB.Menu mnuProperties 
      Caption         =   "&Properties"
   End
   Begin VB.Menu mnuDXspots 
      Caption         =   "&DX Spots"
   End
   Begin VB.Menu mSettings 
      Caption         =   "Telnet &Settings"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuDXhelp 
         Caption         =   "DX Telnet Client Help"
      End
      Begin VB.Menu mnuShortcut 
         Caption         =   "Shortc&ut Keys"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About DX Telnet Client"
      End
   End
   Begin VB.Menu mnuRestore 
      Caption         =   "Restore"
      Visible         =   0   'False
      Begin VB.Menu mnuShowSpots 
         Caption         =   "No DX Spots in List"
      End
      Begin VB.Menu mnuRsep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRrestore 
         Caption         =   "Restore DX Telnet Windows"
      End
      Begin VB.Menu mnuRmin 
         Caption         =   "Minimize DX Telnet Windows to Systray"
      End
      Begin VB.Menu mnuRsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRtelnet 
         Caption         =   "Show Telnet Client Window"
      End
      Begin VB.Menu mnuRdxswindow 
         Caption         =   "Show DX Spots Window"
      End
      Begin VB.Menu mnuTimeSync 
         Caption         =   "Show Time Sync Window"
      End
      Begin VB.Menu mnuRsep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRdxspot 
         Caption         =   "Send a DX Spot"
      End
      Begin VB.Menu mnuRannounce 
         Caption         =   "Send Announcement"
      End
      Begin VB.Menu mnuRsep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRproperties 
         Caption         =   "Show Properties Window"
      End
      Begin VB.Menu mnuRshortcut 
         Caption         =   "Show Shortcut Keys"
      End
      Begin VB.Menu mnuRsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRclose 
         Caption         =   "Close this Menu"
      End
      Begin VB.Menu mnuRsep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRshutdown 
         Caption         =   "Shutdown DX Telnet Client"
      End
   End
End
Attribute VB_Name = "frmTelnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Option Explicit

    Const GO_NORM = 0

    Const GO_ESC1 = 1
    Const GO_ESC2 = 2
    Const GO_ESC3 = 3
    Const GO_ESC4 = 4
    Const GO_ESC5 = 5

    Const GO_IAC1 = 6
    Const GO_IAC2 = 7
    Const GO_IAC3 = 8
    Const GO_IAC4 = 9
    Const GO_IAC5 = 10
    Const GO_IAC6 = 11


    Const SUSP = 237
    Const ABORT = 238      'Abort
    Const SE = 240         'End of Subnegotiation
    Const NOP = 241
    Const DM = 242         'Data Mark
    Const BREAK = 243      'BREAK
    Const IP = 244         'Interrupt Process
    Const AO = 245         'Abort Output
    Const AYT = 246        'Are you there
    Const EC = 247         'Erase character
    Const EL = 248         'Erase Line
    Const GOAHEAD = 249    'Go Ahead
    Const SB = 250         'What follows is subnegotiation
    Const WILLTEL = 251
    Const WONTTEL = 252
    Const DOTEL = 253
    Const DONTTEL = 254
    Const IAC = 255

    Const BINARY = 0
    Const ECHO = 1
    Const RECONNECT = 2
    Const SGA = 3
    Const AMSN = 4
    Const STATUS = 5
    Const TIMING = 6
    Const RCTAN = 7
    Const OLW = 8
    Const OPS = 9
    Const OCRD = 10
    Const OHTS = 11
    Const OHTD = 12
    Const OFFD = 13
    Const OVTS = 14
    Const OVTD = 15
    Const OLFD = 16
    Const XASCII = 17
    Const LOGOUT = 18
    Const BYTEM = 19
    Const DET = 20
    Const SUPDUP = 21
    Const SUPDUPOUT = 22
    Const SENDLOC = 23
    Const TERMTYPE = 24
    Const EOR = 25
    Const TACACSUID = 26
    Const OUTPUTMARK = 27
    Const TERMLOCNUM = 28
    Const REGIME3270 = 29
    Const X3PAD = 30
    Const NAWS = 31
    Const TERMSPEED = 32
    Const TFLOWCNTRL = 33
    Const LINEMODE = 34
    Const DISPLOC = 35
    Const ENVIRON = 36
    Const AUTHENTICATION = 37
    Const UNKNOWN39 = 39
    Const EXTENDED_OPTIONS_LIST = 255
    Const RANDOM_LOSE = 256



    '------------------------------------------------------------
    Private Declare Function OSWinHelp% Lib "USER32" Alias "WinHelpA" (ByVal HWND&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
    'Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    'Private Operating       As Boolean
    'Private Connected       As Boolean
    Public Receiving                 As Boolean

    Private parsedata(10)            As Integer
    Private ppno                     As Integer


    Private control_on               As Boolean


    Public RemoteIPAd                As String
    Public RemotePort                As Integer

    Public TraceTelnet               As Boolean
    Public Tracevt100                As Boolean

    Private sw_ugoahead              As Boolean
    Private sw_igoahead              As Boolean
    Private sw_echo                  As Boolean
    Private sw_linemode              As Boolean
    Private sw_termsent              As Boolean
    Private substate                 As Boolean


Private Sub cursor_timer_Timer()


        If Not Receiving Then
            ' Debug.Print "Timer"
            term_DriveCursor
        End If

End Sub


Private Sub Form_GotFocus()

    Call RemoveBalloon(frmSystray)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim CH            As String
    
    CH = Chr$(0)
    
    'Translate keycodes to VT100 escape sequences
  
        Select Case KeyCode
            Case vbKeyControl
                control_on = True
            Case vbKeyEnd
                CH = Chr$(27) + "[K"
            Case vbKeyHome
                CH = Chr$(27) + "[H"
            Case vbKeyLeft
                CH = Chr$(27) + "[D"
            Case vbKeyUp
                CH = Chr$(27) + "[A"
            Case vbKeyRight
                CH = Chr$(27) + "[C"
            Case vbKeyDown
                CH = Chr$(27) + "[B"
            Case vbKeyF1
                CH = Chr$(27) + "OP"
            Case vbKeyF2
                CH = Chr$(27) + "OQ"
            Case vbKeyF3
                CH = Chr$(27) + "OR"
            Case vbKeyF4
                CH = Chr$(27) + "OS"
            Case Else

                If control_on And KeyCode > 63 Then
                    CH = Chr$(KeyCode - 64)
                End If

        End Select

        If CH > Chr$(0) And Connected Then
            WinsockClient.SendData CH

                If TraceTelnet Then Debug.Print CH
  
        End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Dim CH            As String
    
        If Connected Then
            CH = Chr$(KeyAscii)

                If control_on Then

                        If KeyAscii > 63 Then
                            CH = Chr$(KeyAscii - 64)
                        Else
                            CH = Chr$(0)
                        End If

                End If
        
                If CH > Chr$(0) Then

                        If CH = Chr$(13) Then
                            CH = CH & Chr$(10)
                        End If

                    WinsockClient.SendData CH
                End If

        End If


End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

        Select Case KeyCode
            Case vbKeyControl
                control_on = False
        End Select

End Sub

Private Sub Form_Load()

    ' Place this in the Form Load event of the form you want to disable the 'X':

    Dim hSysMenu            As Long

    hSysMenu = GetSystemMenu(HWND, False)
    RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND


    RemoteIPAd = DXtelnethost
    RemotePort = DXtelnetport
    Call vt100.term_init
    mnuClose.Enabled = False
    
End Sub

Private Sub Form_Paint()

    term_redrawscreen

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

        With WinsockClient
            .Close                            ' Clear any errors...
            .RemoteHost = "0.0.0.0"
            .RemotePort = 0
        End With

    Operating = False
    Connected = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Delete temp spot files on exit
    On Error Resume Next
    Unload frmTimeSync
    Unload frmAbout
    Unload frmDxAlert
    Unload frmDXspot
    Unload frmProperties
    Unload frmShortcutKeys
    Unload frmTCPIP
    Unload frmSystray

    'WinsockClient_Close
    Kill App.Path & "\DXheard.lst"
    Kill App.Path & "\WWVheard.lst"

    End                                  ' End program forcefully

End Sub


Public Sub mClose_Click()

    frmTelnet.WinsockClient.SendData vbCrLf & "bye" & vbCrLf

    '    WinsockClient_Close
    
    '    End

End Sub

Private Sub mExit_Click()

    'Delete temp spot files on exit
    On Error Resume Next

    'Close winsock connection if open

        If WinsockClient.State > 0 Then
            WinsockClient_Close
        End If

    Kill App.Path & "\DXheard.lst"
    Kill App.Path & "\WWVheard.lst"

    'unload systray icon
    Call SystrayOff(frmSystray)
    Call SystrayOff(frmTimeSync)

    End

End Sub

Private Sub mnuAbout_Click()

    frmAbout.Visible = True

End Sub

Private Sub mnuClose_Click()

        If WinsockClient.State = 0 Then Exit Sub    'Connection already closed
    mClose_Click

End Sub

Private Sub mnuConnect_Click()

        If WinsockClient.State = 7 Then Exit Sub    'Connection already open
    mOpen_Click

End Sub

Public Sub mnuDXhelp_Click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:luis@yahoo.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus
    ShellExecute HWND, vbNullString, App.Path & "\html\DXClusterTelnet.htm", vbNullString, vbNullString, vbNormalFocus
   
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)

End Sub

Public Sub mnuDXspot_Click()

    frmDXspot.Show

End Sub

Private Sub mnuDXspots_Click()

        If DXwindow.Visible = False Then 'Test to see if window already open
    
            DXwindow.WindowState = vbNormal
            DXwindow.Show
            Call ViewDXHeardList
            Call ViewWWVHeardList
        Else
        
            'If window is minimized, restore to normal window size

                If DXwindow.WindowState = vbMinimized Then
                    DXwindow.WindowState = vbNormal
                    DXwindow.Show
                End If
        
            DXwindow.SetFocus
    
        End If

End Sub

Public Sub mnuMessage_Click()

    frmAnnounce.Show

End Sub

Private Sub mnuMinAll_Click()

    mnuRmin_Click

End Sub

Public Sub mnuProperties_Click()

        If frmProperties.Visible = False Then 'Test to see if window already open
        
            frmProperties.Visible = True
            frmProperties.SetFocus
        Else
            frmProperties.SetFocus
            'If window is minimized, restore to normal window size

                If frmProperties.WindowState = vbMinimized Then DXwindow.WindowState = vbNormal
        
        End If

End Sub

Private Sub mnuWWVspots_Click()

        If DXwindow.Visible = False Then 'Test to see if window already open
    
            DXwindow.WindowState = vbNormal
            DXwindow.Show
            Call ViewDXHeardList
            Call ViewWWVHeardList
        Else
            'If window is minimized, restore to normal window size

                If DXwindow.WindowState = vbMinimized Then
                    DXwindow.WindowState = vbNormal
                    DXwindow.Show
                End If
        
            DXwindow.SetFocus
        
        End If

End Sub

Private Sub mnuRannounce_Click()

    mnuMessage_Click

End Sub

Public Sub mnuRclose_Click()

    'If Telnet connected to cluster, set DX window with focus

        If WinsockClient.State > 0 Then

                If DXwindow.Visible = True Then

                        If frmAnnounce.Visible = False And _
                            frmDXspot.Visible = False Then
                            DXwindow.SetFocus
                        End If

                Else

                        If frmTelnet.Visible = True Then frmTelnet.SetFocus
                End If

        End If

    'If frmTelnet.Visible = True Then frmTelnet.SetFocus
    'If DXwindow.Visible = True Then DXwindow.SetFocus
    'if frmAnnouce or frmDXspot visible, exit sub

End Sub

Private Sub mnuRdxspot_Click()

    mnuDXspot_Click

End Sub

Public Sub mnuRdxswindow_Click()

    mnuDXspots_Click

End Sub




Public Sub mnuRmin_Click()

    'Hide open windows to systray
    'If DXwindow.Visible = True Then DXwindow.WindowState = vbMinimized

        If DXwindow.Visible = True Then Unload DXwindow
    'If frmTelnet.Visible = True Then frmTelnet.WindowState = vbMinimized

        If frmTelnet.Visible = True Then frmTelnet.Hide

        If frmProperties.Visible = True Then Unload frmProperties

        If frmAbout.Visible = True Then Unload frmAbout

        If frmDXspot.Visible = True Then Unload frmDXspot

        If frmShortcutKeys.Visible = True Then Unload frmShortcutKeys

        If frmTCPIP.Visible = True Then Unload frmTCPIP

        If frmTimeSync.Visible = True Then frmTimeSync.Hide

End Sub

Private Sub mnuRproperties_Click()

    mnuProperties_Click

End Sub

Private Sub mnuRrestore_Click()

    Call SetForegroundWindow(Me.HWND)
    frmTelnet.WindowState = vbNormal
    frmTelnet.Show
            
        If DXwindow.Visible = False Then
            'if form is hiden, show window and update grid
            Load DXwindow
            DXwindow.WindowState = vbNormal
            DXwindow.Show
            Call ViewDXHeardList
            Call ViewWWVHeardList
            'If program was running and ther is a current DX Spot caption, use it, if not, stock caption
            DXwindow.Caption = "DX Cluster Telnet Client - " + DXspotCaption

        Else
            'If form loaded and in up or in task bar, dont update grid
            DXwindow.WindowState = vbNormal
            DXwindow.Show
            'If program was running and ther is a current DX Spot caption, use it, if not, stock caption
            DXwindow.Caption = "DX Cluster Telnet Client - " + DXspotCaption

        End If
        
    'If Telnet connected to cluster, set DX window ith focus

        If WinsockClient.State > 0 Then
            DXwindow.SetFocus
        Else
            frmTelnet.SetFocus
        End If
        
End Sub

Private Sub mnuRshortcut_Click()

    mnuShortcut_Click

End Sub

Private Sub mnuRshutdown_Click()

    mExit_Click

End Sub

Private Sub mnuRtelnet_Click()

    Call SetForegroundWindow(Me.HWND)
    frmTelnet.WindowState = vbNormal
    frmTelnet.Show
    frmTelnet.SetFocus

End Sub

Private Sub mnuShortcut_Click()

    frmShortcutKeys.Visible = True

End Sub

Public Sub mnuShowDX_Click()

    frmTelnet.WinsockClient.SendData (vbCrLf & "sh/dx" & vbCrLf)

End Sub

Private Sub mnuShowSpots_Click()

    mnuDXspots_Click

End Sub

Public Sub mnuShowUsers_Click()

    frmTelnet.WinsockClient.SendData (vbCrLf & "sh/users" & vbCrLf)
    'frmTelnet.WinsockClient.SendData (vbCrLf & "who" & vbCrLf)

End Sub

Public Sub mnuShowWWV_Click()

    frmTelnet.WinsockClient.SendData (vbCrLf & "sh/wwv" & vbCrLf)

End Sub

Private Sub mnuSpotDX_Click()

End Sub

Private Sub mnuTimeSync_Click()

        If frmTimeSync.Visible = True Then
            frmTimeSync.SetFocus
        Else
            frmTimeSync.WindowState = vbNormal
            frmTimeSync.Visible = True
            frmTimeSync.SetFocus
        End If

End Sub

Public Sub mOpen_Click()

    On Error Resume Next                                   ' Handle errors...
    '------------------------------------------------------------

        If Not Operating Then
            Operating = True

                If TraceTelnet Then Debug.Print Int(Timer) & " - [DoConnect] : " & vbCrLf

                With WinsockClient

                        If .State <> 0 Then
                            .Close
                            .RemotePort = 0
                            .LocalPort = 0
                            DoEvents

                                Do
                                    DoEvents
                                Loop Until .State = 0

                        End If
            
                    .RemoteHost = RemoteIPAd
                    .RemotePort = RemotePort
                    .Connect ' Attempt new connection
                    term_init
                    frmTelnet.stbStatusBar.Panels(3).Text = "Connecting to Remote Host"
                    frmTelnet.stbStatusBar.Panels(3).Picture = LoadResPicture(121, 1)
            
                End With
        
            mOpen.Enabled = False
            mnuConnect.Enabled = False
            DXwindow.mnuConnect.Enabled = False
            DXwindow.mnuFconnect.Enabled = False
            mClose.Enabled = True
            mnuClose.Enabled = True
            DXwindow.mnuDisconnect = True
            DXwindow.mnuFdisconnect = True
            mSettings.Enabled = False
            mnuShow.Enabled = True
            DXwindow.mnuShow = True
            mnuSend.Enabled = True
            DXwindow.mnuSend = True
            mnuRdxspot.Enabled = True
            mnuRannounce.Enabled = True
        
        End If
        
End Sub

Private Sub mSettings_Click()

    frmTCPIP.Show vbModal, frmTelnet

    'Show Telnet setup tab from properties form

    'frmProperties.PropertiesTab.Tab = 4
    '    If frmProperties.Visible = False Then 'Test to see if window already open
    '
    '        frmProperties.Visible = True
    '        frmProperties.SetFocus
    '    Else
    '        frmProperties.SetFocus
    '        'If window is minimized, restore to normal window size
    '        If frmProperties.WindowState = vbMinimized Then DXwindow.WindowState = vbNormal
    '
    '    End If
    '

End Sub

Private Sub RealTime_Timer()

    'Display time in local and UTC

    'See if DXwindow open

        If DXwindow.Visible = True Then
            'If visable, put UTC time in window
            DXwindow.DXtime_UTC.Caption = UTCtime
            DXwindow.DXtime_local.Caption = Format(Time$, "HH:MM:SS")
        End If
    
    'Display Time in Main Window
    '    LocalTime.Caption = Format(Date, "long date") & "   " & Time$ & "  " & "Local"
    '    ZuluTime.Caption = Format(UTCdate, "long date") & "   " & UTCtime & "  " & "UTC"
     
    'See if it is the top of the hour. If so' purge DX and WWV list of staile records

        If Mid$(Time$, 4, 2) = "00" And PurgeSpots = True Then
            'Set false to keep from repeating for every second of first minute
            PurgeSpots = False
            Call modExpireDX.ExpireDX
        End If
    
    'Set PurgeSpotes = True to setup for next hour

        If Mid$(Time$, 4, 2) = "01" Then PurgeSpots = True

End Sub


Private Sub stbStatusBar_PanelClick(ByVal Panel As Panel)

        Select Case Panel.Key
            Case "Min"   'mnuRmin_Click

                If frmTelnet.Visible = True Then frmTelnet.Hide

                If frmProperties.Visible = True Then Unload frmProperties

                If frmAbout.Visible = True Then Unload frmAbout

                If frmDXspot.Visible = True Then Unload frmDXspot

                If frmShortcutKeys.Visible = True Then Unload frmShortcutKeys

                If frmTCPIP.Visible = True Then Unload frmTCPIP

                If frmTimeSync.Visible = True Then frmTimeSync.Hide
    
            Case Else

        End Select

End Sub

Private Sub WinsockClient_Close()
        
    frmTelnet.Visible = True
    frmTelnet.SetFocus
        
    On Error Resume Next

    frmTelnet.stbStatusBar.Panels(1).Text = "Closed"
    frmTelnet.stbStatusBar.Panels(1).Picture = LoadResPicture(119, 1)
    frmTelnet.stbStatusBar.Panels(2).Text = "Disconnected from DX Cluster"
    frmTelnet.stbStatusBar.Panels(3).Text = "Connection Reset"
    frmTelnet.stbStatusBar.Panels(3).Picture = LoadResPicture(121, 1)
        
        If TraceTelnet Then Debug.Print Int(Timer) & " - [Closed  ] : Connection Reset By Peer "

        With WinsockClient
            .Close                                     ' Clear any errors...
            .RemotePort = 0
            .LocalPort = 0
        End With

    Operating = False
    Connected = False

        If LoginAuto = 1 Then
            Passwrd = True
            Login = True
        End If
        
    Call term_eraseSCREEN
    Call term_eraseBUFFER
        
    'Display message box for disconnect

    Dim DisMsg
        
    DisMsg = MsgBox("Disconnected from DX CLuster" & vbCrLf & vbCrLf & "at " & UTCtime & " " & UTCdate & " UTC", vbOKOnly, "Telnet disconnected")
    frmTelnet.stbStatusBar.Panels(2).Text = "Shift+F2 to Connect to Cluster"
    frmTelnet.stbStatusBar.Panels(3).Text = "DX Cluster Telnet Client Ready !"
    frmTelnet.stbStatusBar.Panels(3).Picture = LoadResPicture(122, 1)
       
    mOpen.Enabled = True
    mnuConnect.Enabled = True
    DXwindow.mnuConnect = True
    DXwindow.mnuFconnect = True
    mSettings.Enabled = True
    mClose.Enabled = False
    mnuClose.Enabled = False
    DXwindow.mnuDisconnect = False
    DXwindow.mnuFdisconnect = False
    mnuShow.Enabled = False
    DXwindow.mnuShow = False
    mnuSend.Enabled = False
    DXwindow.mnuSend = False
    mnuRdxspot.Enabled = False
    mnuRannounce.Enabled = False
        
    'Set Form caption to "DX Cluster Telnet Client - " and the last DX spot
    DXspotCaption = "Not Connected"
    'Set frmTelnet menu item mnuShowSpots caption to DX spot info
    frmTelnet.mnuShowSpots.Caption = "DX Cluster Telnet Client - " + DXspotCaption
    'Set Systray Incon Caption
    Call ChangeSystrayToolTip(frmSystray, "DX Cluster Telnet Client - " + DXspotCaption)
            
        If DXwindow.Visible = True Then

            'Set Form caption to "DX Cluster Telnet Client - " and the last DX spot
            DXwindow.Caption = "DX Cluster Telnet Client - " + DXspotCaption
            
        End If

End Sub

Private Sub WinsockClient_Connect()

    Dim ConnectedText            As String

    frmTelnet.Visible = True
    frmTelnet.SetFocus

    On Error GoTo ERR_Connect

    '------------------------------------------------------------
        
        If TraceTelnet Then Debug.Print Int(Timer) & " - [Connect] : " & _
    "[" & WinsockClient.RemoteHost & "] " & _
    "[" & WinsockClient.RemoteHostIP & "] " & _
    "[" & CStr(WinsockClient.RemotePort) & "]"  ' Display connection info
        
         
    sw_ugoahead = True
    sw_igoahead = False
    sw_echo = True
    sw_linemode = False
    sw_termsent = False
    substate = False
         
    '         ConnectString = Chr$(IAC) & Chr$(DOTEL) & Chr$(ECHO) _
    '                       & Chr$(IAC) & Chr$(DOTEL) & Chr$(SGA) _
    '                       & Chr$(IAC) & Chr$(WILLTEL) & Chr$(NAWS) _
    '                       & Chr$(IAC) & Chr$(WILLTEL) & Chr$(TERMTYPE) _
    '                       & Chr$(IAC) & Chr$(WILLTEL) & Chr$(TERMSPEED)
    '
    '
    '        WinsockClient.SendData ConnectString
        
        If TraceTelnet Then Debug.Print "SENT: DOTEL  ECHO SGA"

        If TraceTelnet Then Debug.Print "SENT: WILL NAWS TERMTYPE TERMSPEED"
        
    Connected = True
    ConnectedText = DXtelnethost + " on port" + Str(DXtelnetport)
    frmTelnet.stbStatusBar.Panels(1).Text = "Connected"
    frmTelnet.stbStatusBar.Panels(1).Picture = LoadResPicture(118, 1)
    frmTelnet.stbStatusBar.Panels(2).Text = ConnectedText
    frmTelnet.stbStatusBar.Panels(3).Text = "Connection Accepted By Remote Host"
    frmTelnet.stbStatusBar.Panels(3).Picture = LoadResPicture(123, 1)
        
    'Set Form caption to "DX Cluster Telnet Client - " and the last DX spot
    DXspotCaption = "Connected"
    'Set frmTelnet menu item mnuShowSpots caption to DX spot info
    frmTelnet.mnuShowSpots.Caption = "DX Cluster Telnet Client - " + DXspotCaption
    'Set Systray Incon Caption
    Call ChangeSystrayToolTip(frmSystray, "DX Cluster Telnet Client - " + DXspotCaption)
            
        If DXwindow.Visible = True Then

            'Set Form caption to "DX Cluster Telnet Client - " and the last DX spot
            DXwindow.Caption = "DX Cluster Telnet Client - " + DXspotCaption
        
        End If

    Exit Sub

ERR_Connect:                 'Error Trap

    Call ERR_Handler
        
End Sub

Private Sub WinsockClient_DataArrival(ByVal bytesTotal As Long)
    

    On Error GoTo ERR_DataArrival

    Dim CH()                     As Byte
    Dim i                        As Integer
    Static cmd                   As Byte

    '------------------------------------------------------------
        
    'If Not Receiving Then
    '    Receiving = True
    '    term_CaretControl True
    'Else
    'Exit Sub
    'End If
    
        If (bytesTotal > 0) Then  ' If there is any data...
        
        
            WinsockClient.GetData CH, vbByte + vbArray, bytesTotal
        
            ' CH = Buf

                For i = 0 To bytesTotal - 1

                        Select Case cmd
                            Case GO_NORM
                                cmd = term_process_char(CH(i))
                            Case GO_IAC1
                                cmd = iac1(CH(i))
                            Case GO_IAC2
                                cmd = iac2(CH(i))
                            Case GO_IAC3
                                cmd = iac3(CH(i))
                            Case GO_IAC4
                                cmd = iac4(CH(i))
                            Case GO_IAC5
                                cmd = iac5(CH(i))
                            Case GO_IAC6
                                cmd = iac6(CH(i))
                            Case Else

                                If TraceTelnet Then Debug.Print "Invalid 'next (" + Str$(cmd) + ")' processing routine in cmd loop"
                        End Select

                    DoEvents
                Next i

        End If
    
    term_CaretControl False
    Receiving = False
    Erase CH
    Erase CH
    Erase CH
    Exit Sub
    
ERR_DataArrival:                 'Error Trap

    term_CaretControl False
    Receiving = False
    
    Call ERR_Handler
    
End Sub



Private Function iac1(CH As Byte) As Integer

    ' Debug.Print "IAC : ";
    iac1 = GO_NORM

        Select Case CH
            Case DOTEL
                iac1 = GO_IAC2
            Case DONTTEL
                iac1 = GO_IAC6
            Case WILLTEL
                iac1 = GO_IAC3
            Case WONTTEL
                iac1 = GO_IAC4
            Case SB
                iac1 = GO_IAC5
                ppno = 0
            Case SE
                ' End of negotiation string, string is in parsedata()

                Select Case parsedata(0)
                    Case TERMTYPE

                        If parsedata(1) = 1 Then

                                If TraceTelnet Then Debug.Print "SENT: SB TERMTYPE VT100"
                            WinsockClient.SendData Chr$(IAC) & Chr$(SB) & Chr$(TERMTYPE) & "DEC-VT100" & Chr$(0) & Chr$(IAC) & Chr$(SE)
                        End If

                    Case TERMSPEED

                        If parsedata(1) = 1 Then
                            ' Debug.Print "TERMSPEED"

                                If TraceTelnet Then Debug.Print "SENT: SB TERMSPEED 38400"
                            WinsockClient.SendData Chr$(IAC) & Chr$(WILLTEL) & Chr$(CH)
                            WinsockClient.SendData Chr$(IAC) & Chr$(SB) _
                            & Chr$(TERMSPEED) & Chr$(0) _
                            & "9600,9600" _
                            & Chr$(IAC) & Chr$(SE)
                            'Original termspeed = & "57600,57600"
                        End If

                End Select

        End Select

End Function

Private Function iac2(CH As Byte) As Integer

    'DO Processing Respond with WILL or WONT

        If TraceTelnet Then Debug.Print "                                                                   RECEIVED DO : ";
    iac2 = GO_NORM

        Select Case CH
            Case BINARY

                If TraceTelnet Then Debug.Print "BINARY"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(BINARY)

                If TraceTelnet Then Debug.Print "SENT: WONT BINARY"
            Case ECHO

                If TraceTelnet Then Debug.Print "ECHO"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(ECHO)

                If TraceTelnet Then Debug.Print "SENT: WONT ECHO"
            Case NAWS

                If TraceTelnet Then Debug.Print "WINDOW SIZE"
            WinsockClient.SendData Chr$(IAC) & Chr$(SB) & Chr$(NAWS) & Chr$(0) & Chr$(80) & Chr$(0) & Chr$(24) & Chr$(IAC) & Chr$(SE)

                If TraceTelnet Then Debug.Print "SENT: SB WINDOW SIZE 80x24"
            Case SGA

                If TraceTelnet Then Debug.Print "SGA"

                If Not sw_igoahead Then

                        If TraceTelnet Then Debug.Print "SENT: WILL SGA"
                    WinsockClient.SendData Chr$(IAC) & Chr$(WILLTEL) & Chr$(SGA)
                    sw_igoahead = True
                Else

                        If TraceTelnet Then Debug.Print "DID NOT RESPOND"
                End If

            Case TERMTYPE

                If TraceTelnet Then Debug.Print "TERMTYPE"

                If Not sw_termsent Then

                        If TraceTelnet Then Debug.Print "SENT: WILL TERMTYPE"
                    sw_termsent = True
                    WinsockClient.SendData Chr$(IAC) & Chr$(WILLTEL) & Chr$(TERMTYPE)

                        If TraceTelnet Then Debug.Print "SENT: SB TERMTYPE VT100"
                    WinsockClient.SendData Chr$(IAC) & Chr$(SB) & Chr$(TERMTYPE) & _
                    Chr$(0) & "VT100" & Chr$(IAC) & Chr$(SE)
                Else

                        If TraceTelnet Then Debug.Print "DID NOT RESPOND"
                End If
 
            Case TERMSPEED

                If TraceTelnet Then Debug.Print "TERMSPEED"

                If TraceTelnet Then Debug.Print "SENT: WILL TERMSPEED"
            WinsockClient.SendData Chr$(IAC) & Chr$(WILLTEL) & Chr$(TERMSPEED)
      
                If TraceTelnet Then Debug.Print "SENT: SB TERMSPEED 57600"
            WinsockClient.SendData Chr$(IAC) & Chr$(SB) & Chr$(TERMSPEED) & Chr$(0)
            WinsockClient.SendData "57600,57600"
            WinsockClient.SendData Chr$(IAC) & Chr$(SE)
      
            Case TFLOWCNTRL

                If TraceTelnet Then Debug.Print "TFLOWCNTRL"

                If TraceTelnet Then Debug.Print "SENT: WONT FLOWCONTROL"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
            Case LINEMODE

                If TraceTelnet Then Debug.Print "LINEMODE"

                If TraceTelnet Then Debug.Print "SENT: WONT LINEMODE"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
            Case STATUS

                If TraceTelnet Then Debug.Print "STATUS"

                If TraceTelnet Then Debug.Print "SENT: WONT STATUS"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
            Case TIMING

                If TraceTelnet Then Debug.Print "TIMING"

                If TraceTelnet Then Debug.Print "SENT: WONT TIMING"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
            Case DISPLOC

                If TraceTelnet Then Debug.Print "DISPLOC"

                If TraceTelnet Then Debug.Print "SENT: WONT DISPLOC"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
    
            Case ENVIRON

                If TraceTelnet Then Debug.Print "ENVIRON"

                If TraceTelnet Then Debug.Print "SENT: WONT ENVIRON"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
    
            Case UNKNOWN39

                If TraceTelnet Then Debug.Print "UNKNOWN39"

                If TraceTelnet Then Debug.Print "SENT: WONT " & Asc(CH)
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
    
            Case AUTHENTICATION

                If TraceTelnet Then Debug.Print "AUTHENTICATION"

                If TraceTelnet Then Debug.Print "SENT: WILL "; AUTHENTICATION; ""
            WinsockClient.SendData Chr$(IAC) & Chr$(WILLTEL) & Chr$(CH)
      
                If TraceTelnet Then Debug.Print "SENT: SB AUTHENTICATION"
            WinsockClient.SendData Chr$(IAC) & _
            Chr$(SB) & _
            Chr$(AUTHENTICATION) & _
            Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & _
            Chr$(IAC) & _
            Chr$(SE)
            Case Else

                If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)

                If TraceTelnet Then Debug.Print "SENT: WONT UNKNOWN CMD " & CH
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
        End Select

End Function

Private Function iac3(CH As Byte) As Integer

    ' WILL Processing - Respond with DO or DONT
      
        If TraceTelnet Then Debug.Print "                                                                   RECEIVED WILL : ";
    
    iac3 = GO_NORM
    
        Select Case CH
            Case ECHO

                If TraceTelnet Then Debug.Print "ECHO"

                If Not sw_echo Then
                    sw_echo = True
                    WinsockClient.SendData Chr$(IAC) & Chr$(DOTEL) & Chr$(ECHO)

                        If TraceTelnet Then Debug.Print "SENT: DO ECHO"
                End If

            Case SGA

                If TraceTelnet Then Debug.Print "SGA"

                If Not sw_ugoahead Then
                    sw_ugoahead = True
                    WinsockClient.SendData Chr$(IAC) & Chr$(DOTEL) & Chr$(SGA)

                        If TraceTelnet Then Debug.Print "SENT: DOTEL SGA"
                End If
        
            Case TERMSPEED

                If TraceTelnet Then Debug.Print "TERMSPEED"

                If TraceTelnet Then Debug.Print "SENT: DONT TERMSPEED"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
          
            Case TFLOWCNTRL

                If TraceTelnet Then Debug.Print "TFLOWCNTRL"

                If TraceTelnet Then Debug.Print "SENT: DONT FLOWCONTROL"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
          
            Case LINEMODE

                If TraceTelnet Then Debug.Print "LINEMODE"

                If TraceTelnet Then Debug.Print "SENT: DONT LINEMODE"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
          
            Case STATUS

                If TraceTelnet Then Debug.Print "STATUS"

                If TraceTelnet Then Debug.Print "SENT: DONT STATUS"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
          
            Case TIMING

                If TraceTelnet Then Debug.Print "TIMING"

                If TraceTelnet Then Debug.Print "SENT: DONT TIMING"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
          
            Case DISPLOC

                If TraceTelnet Then Debug.Print "DISPLOC"

                If TraceTelnet Then Debug.Print "SENT: WONT DISPLOC"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
        
            Case ENVIRON

                If TraceTelnet Then Debug.Print "ENVIRON"

                If TraceTelnet Then Debug.Print "SENT: WONT ENVIRON"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
        
            Case UNKNOWN39

                If TraceTelnet Then Debug.Print "UNKNOWN39"

                If TraceTelnet Then Debug.Print "SENT: WONT " & Asc(CH)
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
        
        
            Case Else

                If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)

                If TraceTelnet Then Debug.Print "SENT: WONT UNKNOWN CMD " & Asc(CH)
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
        End Select
    
End Function

Private Function iac4(CH As Byte) As Integer

    ' WONT Processing
  
        If TraceTelnet Then Debug.Print "                                                                   RECEIVED WONT : ";

    iac4 = GO_NORM

        Select Case CH
    
            Case ECHO

                If TraceTelnet Then Debug.Print "ECHO"

                If sw_echo = True Then
                    WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(ECHO)

                        If TraceTelnet Then Debug.Print "SENT: DONTEL ECHO"
                    sw_echo = False
                End If
      
            Case SGA

                If TraceTelnet Then Debug.Print "SGA"

                If TraceTelnet Then Debug.Print "SENT: DONT SGA"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(SGA)
            sw_igoahead = False
    
            Case TERMSPEED

                If TraceTelnet Then Debug.Print "TERMSPEED"

                If TraceTelnet Then Debug.Print "SENT: DONT TERMSPEED"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
            Case TFLOWCNTRL

                If TraceTelnet Then Debug.Print "FLOWCONTROL"

                If TraceTelnet Then Debug.Print "SENT: DONT FLOWCONTROL"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
            Case LINEMODE

                If TraceTelnet Then Debug.Print "LINEMODE"

                If TraceTelnet Then Debug.Print "SENT: DONT LINEMODE"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
            Case STATUS

                If TraceTelnet Then Debug.Print "STATUS"

                If TraceTelnet Then Debug.Print "SENT: DONT STATUS"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
            Case TIMING

                If TraceTelnet Then Debug.Print "TIMING"

                If TraceTelnet Then Debug.Print "SENT: DONT TIMING"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
            Case DISPLOC

                If TraceTelnet Then Debug.Print "DISPLOC"

                If TraceTelnet Then Debug.Print "SENT: DONT DISPLOC"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
            Case ENVIRON

                If TraceTelnet Then Debug.Print "ENVIRON"

                If TraceTelnet Then Debug.Print "SENT: DONT ENVIRON"
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
            Case UNKNOWN39

                If TraceTelnet Then Debug.Print "UNKNOWN39"

                If TraceTelnet Then Debug.Print "SENT: DONT " & Asc(CH)
            WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
            Case Else

                If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)

                If TraceTelnet Then Debug.Print "SENT: DONT UNKNOWN CMD " & Asc(CH)
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
        End Select

End Function

Private Function iac5(CH As Byte) As Integer

    Dim ich                      As Integer

    ' Collect parms after SB and until another IAC

  
    ich = CH

        If ich = IAC Then
            iac5 = GO_IAC1
            Exit Function
        End If
    
        If TraceTelnet Then Debug.Print "                                                                   RECEIVED : ";

        If TraceTelnet Then Debug.Print "SB("; ppno; ") = " & ich
    
    parsedata(ppno) = ich
    ppno = ppno + 1
    
    iac5 = GO_IAC5

End Function


Private Function iac6(CH As Byte) As Integer

    'DONT Processing

 
    iac6 = GO_NORM
        

        Select Case CH
            Case SE

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED SE"

                If TraceTelnet Then Debug.Print "SENT: SE_ACK " & CH

            Case ECHO

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "ECHO"

                If Not sw_echo Then
                    sw_echo = True
                    WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(ECHO)

                        If TraceTelnet Then Debug.Print "SENT: WONT ECHO"
                End If

            Case SGA

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "SGA"

                If Not sw_ugoahead Then
                    sw_ugoahead = True
                    WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(SGA)

                        If TraceTelnet Then Debug.Print "SENT: WONT SGA"
                End If
    
            Case TERMSPEED

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "TERMSPEED"

                If TraceTelnet Then Debug.Print "SENT: WONT TERMSPEED"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
            Case TFLOWCNTRL

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "TFLOWCNTRL"

                If TraceTelnet Then Debug.Print "SENT: WONT FLOWCONTROL"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
            Case LINEMODE

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "LINEMODE"

                If TraceTelnet Then Debug.Print "SENT: WONT LINEMODE"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
            Case STATUS

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "STATUS"

                If TraceTelnet Then Debug.Print "SENT: WONT STATUS"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
            Case TIMING

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "TIMING"

                If TraceTelnet Then Debug.Print "SENT: WONT TIMING"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
            Case DISPLOC

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "DISPLOC"

                If TraceTelnet Then Debug.Print "SENT: WONT DISPLOC"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
    
            Case ENVIRON

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "ENVIRON"

                If TraceTelnet Then Debug.Print "SENT: WONT ENVIRON"
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
    
            Case UNKNOWN39

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "UNKNOWN39"

                If TraceTelnet Then Debug.Print "SENT: WONT " & Asc(CH)
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
        
            Case Else

                If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";

                If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)

                If TraceTelnet Then Debug.Print "SENT: WONT UNKNOWN CMD " & Asc(CH)
            WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
        End Select

End Function


Private Sub WinsockClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
       
    frmTelnet.Visible = True
    frmTelnet.SetFocus
    
    Call WinsockClient_Close

   
    frmTelnet.stbStatusBar.Panels(1).Text = "Closed"
    frmTelnet.stbStatusBar.Panels(1).Picture = LoadResPicture(119, 1)
    frmTelnet.stbStatusBar.Panels(2).Text = "ERROR " + Str(Number)
    frmTelnet.stbStatusBar.Panels(3).Text = Description
    frmTelnet.stbStatusBar.Panels(3).Picture = LoadResPicture(120, 1)

End Sub

Private Sub Form_Resize()

    'If Me.WindowState = vbMinimized Then
    '    frmTelnet.Hide
    'End If

End Sub

Private Sub Form_Terminate()

    'Delete temp spot files on exit
    On Error Resume Next
    
    frmTelnet.Visible = True
    frmTelnet.SetFocus
    
    Unload frmTimeSync
    Unload frmAbout
    Unload frmDxAlert
    Unload frmDXspot
    Unload frmProperties
    Unload frmShortcutKeys
    Unload frmTCPIP
    Unload frmSystray
    
    'WinsockClient_Close
    Kill App.Path & "\DXheard.lst"
    Kill App.Path & "\WWVheard.lst"
    
    End                                  ' End program forcefully
    
End Sub

Public Sub ERR_Handler()
    
    frmTelnet.Visible = True
    frmTelnet.SetFocus
    
    Call WinsockClient_Close
     
    frmTelnet.stbStatusBar.Panels(1).Text = "Closed"
    frmTelnet.stbStatusBar.Panels(2).Text = "ERROR"
    frmTelnet.stbStatusBar.Panels(3).Text = "Connection Closed by Unkown Error"
    frmTelnet.stbStatusBar.Panels(3).Picture = LoadResPicture(120, 1)

End Sub
