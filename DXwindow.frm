VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form DXwindow 
   AutoRedraw      =   -1  'True
   Caption         =   "DX Cluster Telnet Client - No DX Spots in List"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DXwindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9510
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
      Left            =   8040
      MouseIcon       =   "DXwindow.frx":014A
      MousePointer    =   99  'Custom
      Picture         =   "DXwindow.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Click to Minimize DX Spots Window to System Tray"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Timer SpotTimer 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   8760
      Top             =   0
   End
   Begin VB.Frame DXFrame 
      Caption         =   "DX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MouseIcon       =   "DXwindow.frx":0896
      MousePointer    =   99  'Custom
      TabIndex        =   14
      ToolTipText     =   "Click to tune radio to DX frequency"
      Top             =   5280
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Label DXspotText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DX Spot Text"
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
         Height          =   360
         Left            =   1080
         MouseIcon       =   "DXwindow.frx":0BA0
         MousePointer    =   99  'Custom
         TabIndex        =   16
         ToolTipText     =   "Click to tune radio to DX frequency"
         Top             =   315
         Visible         =   0   'False
         Width           =   8115
      End
      Begin VB.Label DXspotlabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WWV"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   435
         Left            =   90
         MouseIcon       =   "DXwindow.frx":0EAA
         MousePointer    =   99  'Custom
         TabIndex        =   15
         ToolTipText     =   "Click to tune radio to DX frequency"
         Top             =   195
         Visible         =   0   'False
         Width           =   990
      End
   End
   Begin VB.CommandButton ClearWWV 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear WWV"
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
      Left            =   5160
      MouseIcon       =   "DXwindow.frx":11B4
      MousePointer    =   99  'Custom
      Picture         =   "DXwindow.frx":14BE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click to clear the WWV report list"
      Top             =   6120
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid WWVGrid 
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid DXGrid 
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Click on DX spot to tune radio to DX frequency"
      Top             =   360
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5318
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      ScrollBars      =   2
      MousePointer    =   99
      MouseIcon       =   "DXwindow.frx":1900
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   3375
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "UTC"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   2175
         TabIndex        =   7
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   555
         TabIndex        =   6
         Top             =   240
         Width           =   705
      End
      Begin VB.Label DXtime_UTC 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   1695
         TabIndex        =   5
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label DXtime_local 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   600
         Width           =   1605
      End
   End
   Begin VB.CommandButton ClearDX 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear DX "
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
      Left            =   3720
      MouseIcon       =   "DXwindow.frx":1C1A
      MousePointer    =   99  'Custom
      Picture         =   "DXwindow.frx":1F24
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click to clear the DX spot list"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton CancelDX 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   6600
      MouseIcon       =   "DXwindow.frx":2366
      MousePointer    =   99  'Custom
      Picture         =   "DXwindow.frx":2670
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label WWVcount 
      Caption         =   "WWV Reports"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label DXcount 
      Caption         =   "DX Spots"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "WWV Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3450
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "DX Spots"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   90
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFconnect 
         Caption         =   "Connect to DX Cluster"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuFdisconnect 
         Caption         =   "Disconnect from DX Cluster"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinAll 
         Caption         =   "Minimize All Windows to System Tray"
      End
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connect"
   End
   Begin VB.Menu mnuDisconnect 
      Caption         =   "&Disconnect"
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show"
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
      End
      Begin VB.Menu mnuAnnounce 
         Caption         =   "Announcement"
         Shortcut        =   +{F9}
      End
   End
   Begin VB.Menu mnuProperties 
      Caption         =   "&Properties"
   End
   Begin VB.Menu mnuTelnet 
      Caption         =   "&Telnet Window"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuDXhelp 
         Caption         =   "DX Telnet Client Help"
      End
      Begin VB.Menu mnuShortcut 
         Caption         =   "Shortcut Keys"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About DX Telnet Client"
      End
   End
End
Attribute VB_Name = "DXwindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim FmCall(32767)
    Dim DXCall(32767)
    Dim DXfreq(32767)
    Dim DXcomment(32767)
    Dim DXtime(32767)
    Dim DXhour            As Integer
    Dim DXlastHour(32767)
    Dim DXspoted(32767)

Private Sub CancelDX_Click()

    'Close the DX Spot Window
    Unload DXwindow

End Sub

Private Sub ClearDX_Click()

    'Clear DX Spot List
    'Clearh the DXHeard list and clear the DXHeard.lst file
    On Error Resume Next
    'Clear the DXHeard table

        For x = 1 To (DXwindow.DXGrid.Rows - 2)
            DXwindow.DXGrid.RemoveItem 1
            DoEvents
        Next x

    'Delete the DXHeard.lst file
    Kill App.Path & "\" & "DXHeard.lst"
    DXwindow.DXcount.Caption = "( 0 current DX spots )"
    'Restore stock caption to DXwindow
    DXwindow.Caption = "DX Cluster Telnet Client - No DX Spots in List"
    'Restore Tray Pop Up Menu info to stock caption
    frmTelnet.mnuShowSpots.Caption = "No DX Spots in List"
    'Restore Systray Tool Tip text to Stock caption
    Call ChangeSystrayToolTip(frmSystray, "DX Cluster Telnet Client - No DX Spots in List")

End Sub

Private Sub ClearWWV_Click()

    'Clear WWV Spot List
    'Clearh the WWVHeard list and clear the DXHeard.lst file
    On Error Resume Next
    'Clear the WWVHeard table

        For x = 1 To (DXwindow.WWVGrid.Rows - 2)
            DXwindow.WWVGrid.RemoveItem 1
            DoEvents
        Next x

    'Delete the WWVHeard.lst file
    Kill App.Path & "\" & "WWVHeard.lst"
    DXwindow.WWVcount.Caption = "( 0 current WWV reports )"

End Sub

Private Sub DXFrame_DragDrop(Source As Control, x As Single, y As Single)

    'Set selection to last spot recieved
    DXGrid.Row = 1
    DXGrid_Click

End Sub

Private Sub DXGrid_Click()

    'Set focus to col 2 (Frequency)
    DXGrid.Col = 2
    'Get Frequency string
    DXfreqStr = DXGrid.Clip

        If DXfreqStr = "" Then Exit Sub 'Test for null string
    'Compute frequency in Hz
    DXfreqVal = Val(DXfreqStr) * 1000
    'Send DX frequency to radio
    DXfreqVal = Val(DXfreqStr) * 1000
    
    'Hold off sending the Data until the last Message is sent

        If frmTelnet.ComRadio.PortOpen = True Then

                Do Until frmTelnet.ComRadio.PortOpen = False
                    DoEvents
                Loop

        End If
    
    Call modRadioComm.SendRadio(DXfreqVal)
    
End Sub

Private Sub DXGrid_EnterCell()

        If UCase(RadioModel) <> "NONE" Then

            'Highlight the current row

            Dim DXwatchMatch            As Boolean
            Dim CallDX                  As String

            DXwindow.DXGrid.Col = 1
            CallDX = DXwindow.DXGrid.Clip

                If CallDX = "" Then Exit Sub 'Test for null string
            DXwatchMatch = False
            DXfile = FreeFile
            On Error Resume Next
            Open App.Path & "\" & "DXwatch.lst" For Input As DXfile

                Do Until EOF(DXfile) = True
                    Input #DXfile, watchCall

                        If Mid(CallDX, 1, Len(watchCall)) = watchCall Then DXwatchMatch = True
                    DoEvents
                Loop

            Close DXfile
            'If this was a DX Watch match, Highlight to Yellow on Magenta cells

                If DXwatchMatch = True Then

                        For r = 0 To 4
                            DXwindow.DXGrid.Col = r
                            DXwindow.DXGrid.CellBackColor = vbMagenta
                            DXwindow.DXGrid.CellForeColor = vbYellow
                            DXwindow.DXGrid.CellFontBold = True
                        Next r

                    'If not a DX Watch match, Highlight to Black on Yellow cells
                Else

                        For r = 0 To 4
                            DXwindow.DXGrid.Col = r
                            DXwindow.DXGrid.CellBackColor = vbYellow
                            DXwindow.DXGrid.CellForeColor = vbBlack
                            DXwindow.DXGrid.CellFontBold = False
                        Next r

                End If

            DXwatchMatch = False

        End If

End Sub

Private Sub DXspotlabel_Click()

    'Set selection to last spot recieved
    DXGrid.Row = 1
    DXGrid_Click

End Sub

Private Sub DXspotText_Click()

    'Set selection to last spot recieved
    DXGrid.Row = 1
    DXGrid_Click

End Sub

Private Sub Form_GotFocus()

    Call RemoveBalloon(frmSystray)

End Sub

Private Sub Form_Load()

    'Set time(s) on form on load. Problem with tick interval not updating during form load
    DXwindow.DXtime_UTC.Caption = UTCtime
    DXwindow.DXtime_local.Caption = Format(Time$, "HH:MM:SS")

    'Set ToolTip for radio frequency control

        If RadioModel = "None" Then
            'If no radio selected, turn off ToolTip t set radio frequency
            DXFrame.ToolTipText = ""
            DXSpotLabel.ToolTipText = ""
            DXspotText.ToolTipText = ""
            DXGrid.ToolTipText = ""
            'Set mouse pointer to default
            DXFrame.MousePointer = 0
            DXSpotLabel.MousePointer = 0
            DXspotText.MousePointer = 0
            DXGrid.MousePointer = 0
        Else
            'If a rado is selected, turn on ToolTip for frequency control
            DXFrame.ToolTipText = "Click to tune radio to DX frequency"
            DXSpotLabel.ToolTipText = "Click to tune radio to DX frequency"
            DXspotText.ToolTipText = "Click to tune radio to DX frequency"
            DXGrid.ToolTipText = "Click on DX spot to tune radio to DX frequency"
            'Set mouse pointer to Icon
            DXFrame.MousePointer = 99
            DXSpotLabel.MousePointer = 99
            DXspotText.MousePointer = 99
            DXGrid.MousePointer = 99
        End If

    DXFrame.BackColor = DXwindow.BackColor
    
        If Connected = False Then
    
            DXwindow.mnuConnect = True
            DXwindow.mnuFconnect = True
            DXwindow.mnuDisconnect = False
            DXwindow.mnuFdisconnect = False
            DXwindow.mnuShow = False
            DXwindow.mnuSend = False
        Else
            DXwindow.mnuConnect = False
            DXwindow.mnuFconnect = False
            DXwindow.mnuDisconnect = True
            DXwindow.mnuFdisconnect = True
            DXwindow.mnuShow = True
            DXwindow.mnuSend = True
        End If
    
    'If program was running and ther is a current DX Spot caption, use it, if not, stock caption
    DXwindow.Caption = "DX Cluster Telnet Client - " + DXspotCaption

End Sub


Private Sub Form_Resize()

    Dim x                   As Long
    Dim y                   As Long

    Dim FontSize            As Long

    ResizeFormFor Me 'Resize the form
  
        If Me.WindowState = vbMinimized Then
            frmTelnet.SetFocus
            frmTelnet.WindowState = vbMinimized
            
        Else
            'If program was running and ther is a current DX Spot caption, use it, if not, stock caption
            DXwindow.Caption = "DX Cluster Telnet Client - " + DXspotCaption
            
            'Get the Font size afte the resize event
            DXGrid.Col = 0
            DXGrid.Row = 0
            FontSize = DXGrid.Font.Size
            'Change all the cells in the Gris to the new Font Size
            'Select the Row

                For x = 1 To (DXGrid.Rows - 2)
                    DXGrid.Row = x
                    'Select the Colom and reset Font Size

                        For y = 0 To 4
                            DXGrid.Col = y
                            DXGrid.CellFontSize = FontSize
                            DoEvents
                        Next y

                    DoEvents
                Next x
           
        End If

    Me.Refresh

End Sub


Private Sub Form_Terminate()

    'Load DX list into array
    i = 0
    'On (Err = 53) GoTo NoDXFile 'File not found error
    On Error GoTo NoDXFile 'File not found error
    DXfile = FreeFile
    Open (App.Path + "\" + "DXHeard.lst") For Input As DXfile
     
        Do Until EOF(DXfile) = True
            i = i + 1
            Input #DXfile, FmCall(i), DXCall(i), DXfreq(i), DXcomment(i), DXtime(i), DXlastHour(i), DXspoted(i)
            DoEvents
            MaxRecord = i
            DoEvents
        Loop
        
    Close DXfile
        
    'Write new DXlist records
    
    Kill (App.Path + "\" + "DXHeard.lst")
    DXfile2 = FreeFile
    Open (App.Path + "\" + "DXHeard.lst") For Output As DXfile
    
    g = 1

        Do Until g = MaxRecord + 1
            DXGrid.Row = i
            DXGrid.Col = 0
        
                Select Case DXGrid.CellBackColor
                    Case vbWhite
                        DXspoted(g) = 0
                    Case vbRed
                        DXspoted(g) = 1
                    Case vbYellow
                        DXspoted(g) = 2
                    Case vbMagenta
                        DXspoted(g) = 3
                End Select
    
            Write #DXfile, FmCall(g), DXCall(g), DXfreq(g), DXcomment(g), DXtime(g), DXlastHour(g), DXspoted(g)
            DoEvents
            g = g + 1
            i = i - 1
        Loop
    
    Close DXfile2

NoDXFile:

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Form_Terminate

End Sub

Private Sub MinSysTray_Click()

    'frmTelnet.mnuRmin_Click
    
        If frmProperties.Visible = True Then Unload frmProperties

        If frmAbout.Visible = True Then Unload frmAbout

        If frmDXspot.Visible = True Then Unload frmDXspot

        If frmShortcutKeys.Visible = True Then Unload frmShortcutKeys

        If frmTCPIP.Visible = True Then Unload frmTCPIP

        If frmTimeSync.Visible = True Then frmTimeSync.Hide
    Unload DXwindow

End Sub

Private Sub mnuAbout_Click()

    frmAbout.Visible = True

End Sub

Private Sub mnuAnnounce_Click()

    frmAnnounce.Show

End Sub

Private Sub mnuConnect_Click()

    frmTelnet.Visible = True
    frmTelnet.WindowState = vbNormal
    frmTelnet.SetFocus
    frmTelnet.mOpen_Click

End Sub

Private Sub mnuDisconnect_Click()

    frmTelnet.Visible = True
    frmTelnet.WindowState = vbNormal
    frmTelnet.SetFocus
    frmTelnet.mClose_Click

End Sub

Private Sub mnuDXhelp_Click()

    frmTelnet.mnuDXhelp_Click

End Sub

Private Sub mnuDXspot_Click()

    frmDXspot.Show

End Sub

Private Sub mnuFconnect_Click()

    mnuConnect_Click

End Sub

Private Sub mnuFdisconnect_Click()

    mnuDisconnect_Click

End Sub

Private Sub mnuMinAll_Click()

    frmTelnet.mnuRmin_Click

End Sub

Private Sub mnuProperties_Click()

    frmTelnet.mnuProperties_Click

End Sub

Private Sub mnuShortcut_Click()

    frmShortcutKeys.Visible = True

End Sub

Private Sub mnuShowDX_Click()

    frmTelnet.Visible = True
    frmTelnet.WindowState = vbNormal
    frmTelnet.SetFocus
    frmTelnet.mnuShowDX_Click

End Sub

Private Sub mnuShowUsers_Click()

    frmTelnet.Visible = True
    frmTelnet.WindowState = vbNormal
    frmTelnet.SetFocus
    frmTelnet.mnuShowUsers_Click

End Sub

Private Sub mnuShowWWV_Click()

    frmTelnet.Visible = True
    frmTelnet.WindowState = vbNormal
    frmTelnet.SetFocus
    frmTelnet.mnuShowWWV_Click

End Sub

Private Sub mnuTelnet_Click()

    frmTelnet.Visible = True
    frmTelnet.WindowState = vbNormal
    frmTelnet.SetFocus

End Sub

Private Sub SpotTimer_Timer()

    'On 30 second timeout, clear DX spot from DXFrame

        If DXwindow.Visible = False Then Exit Sub

    SpotTimer.Enabled = False
    DXFrame.Visible = False
    
End Sub
