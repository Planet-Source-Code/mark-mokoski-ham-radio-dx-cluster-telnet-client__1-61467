Attribute VB_Name = "DXtelnetMain"
    Option Explicit

    '***********************************************************************************
    '
    '   WA1ZEK Telnet DX Cluster Client
    '   11 May 2004
    '
    '
    '   "WA1ZEK Telnet DX Cluster Client"
    '   Copyright 2002
    '   Mark Mokoski, WA1ZEK / VK2IFH
    '
    '   Version info:
    '
    '   11-MAY-2004
    '   0.1.0   First development, basic monitor and Heard List windows
    '
    '   18-MAY-2004
    '   0.1.1   Added Telnet Host and Band properties with FlexGrid selection
    '           Intergated other existing modules into app
    '
    '************************************************************************************

    'Shell out API for HTML files, Mial and Web Browser
    Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

    'Disable "X" (close window) on form (Used on frmTelnet so you cant close app bt acident)
    Public Declare Function GetSystemMenu Lib "USER32" (ByVal HWND As Long, ByVal bRevert As Long) As Long


    Public Declare Function RemoveMenu Lib "USER32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Const SC_CLOSE = &HF060&
    Public Const MF_BYCOMMAND = &H0&
    
    Public Declare Function GetForegroundWindow Lib "USER32" () As Long

    'Declare Public Variables

    'Radio CommPort parameters
    Public RadioModel                     As String
    Public RadioVFO                       As String
    Public RadioComPort                   As String
    Public RadioComBits                   As String
    Public RadioComSpeed                  As String
    Public RadioComStop                   As String
    Public RadioComParity                 As String
    Public RadioComFlow                   As String
    Public RCV_Data_Len                   As Integer

    'Telnet parameters
    Public DXtelnethost                   As String
    Public DXtelnetport                   As Integer
    Public Connected                      As Boolean
    Public Operating                      As Boolean
    'Time Sync
    Public TimeSync                       As Integer

    'Telnet incomming text buffer
    Public TelnetDXtext                   As String
    Public Login                          As Boolean
    Public Passwrd                        As Boolean
    Public TelnetLogon                    As String
    Public TelnetPswrd                    As String
    Public LoginAuto                      As String



    'Output string to radio (Radiotype dependent)
    Public RadioOutString                 As String
    Public RcvFreq                        As String
    Public RawRadio_Responce              As String

    'Alert sounds for DX and WWV spots
    Public DXsound                        As String
    Public DXwatch                        As String
    Public WWVsound                       As String
    Public WAVdxspot                      As String
    Public WAVdxwatch                     As String
    Public WAVwwv                         As String
    Public ANNsound                       As String
    Public WAVannounce                    As String
    Public Sound_OK                       As Integer
    Public PurgeSpots                     As Boolean

    'Colors and Fonts
    Public InfoColor                      As Long
    Public InfoFont                       As String
    Public InfoFont_Size                  As Integer
    Public InfoFont_Bold                  As Boolean
    Public InfoFont_Italic                As Boolean
    Public MyColor                        As Long
    Public MyFont                         As String
    Public MyFont_Size                    As Integer
    Public MyFont_Bold                    As Boolean
    Public MyFont_Italic                  As Boolean
    Public DataColor                      As Long
    Public DataFont                       As String
    Public DataFont_size                  As Integer
    Public DataFont_Bold                  As Boolean
    Public DataFont_Italic                As Boolean
    Public SysColor                       As Long
    Public SysFont                        As String
    Public SysFont_Size                   As Integer
    Public SysFont_Bold                   As Boolean
    Public SysFont_Italic                 As Boolean
    Public Back_Color                     As Long

    'Product Registration
    Public FirstName                      As String
    Public LastName                       As String
    Public CallSign                       As String
    Public ProductID                      As String
    Public RegNumber                      As String

    'DX and WWV Spot expire times in hours
    Public Expire_DX                      As Integer
    Public Expire_WWV                     As Integer
    Public Use_Local_Clock                As Boolean
    Public BalloonToolTip                 As Boolean
    Public ShowAnnouncements              As Boolean

    'DX Watch
    Public DXwatchEnabled                 As Boolean
    Public DXarchiveEnabled               As Boolean
    Public AlertColor                     As Integer

    'DX Spots Window Tool Tip Caption
    Public DXspotCaption                  As String
    Public NoConnectionCaption            As String




Sub Main()

    ' ***************************************************************************
    ' * Test to see if App is allready running
    ' * If App is running, terminate copy
    ' ***************************************************************************

        If App.PrevInstance Then
            MsgBox "WA1ZEK Telnet DX Cluster Client application is already running." & vbCrLf & _
            "Only one instance (copy) of program this can be running" & vbCrLf & _
            "for proper operation.", vbCritical, "Application ERROR"
            End
        Else
            '  MsgBox "This is the first instance of your application."
        End If

    'Remove files from previous session
    On Error Resume Next
    Kill (App.Path & "\DXheard.lst")
    Kill (App.Path & "\WWVheard.lst")

    'Load settings saved in registry

    'Radio settings
    RadioModel = GetSetting("WA1ZEK", "DXtelnet", "Radio", "None")
    RadioVFO = GetSetting("WA1ZEK", "DXtelnet", "VFO", "VFO B")
    RadioComPort = GetSetting("WA1ZEK", "DXtelnet", "ComPort", "COM 1")
    RadioComBits = GetSetting("WA1ZEK", "DXtelnet", "ComBits", "8")
    RadioComSpeed = GetSetting("WA1ZEK", "DXtelnet", "ComSpeed", "9600")
    RadioComStop = GetSetting("WA1ZEK", "DXtelnet", "ComStop", "1")
    RadioComParity = GetSetting("WA1ZEK", "DXtelnet", "ComParity", "None")
    RadioComFlow = GetSetting("WA1ZEK", "DXtelnet", "ComFlow", "XON/XOFF")

    'Telnet Settings
    DXtelnethost = GetSetting("WA1ZEK", "DXtelnet", "DXtelnethost", "dxc.kb1h.com")
    DXtelnetport = GetSetting("WA1ZEK", "DXtelnet", "DXtelnetport", "23")
    TelnetLogon = GetSetting("WA1ZEK", "DXtelnet", "TelnetLogon", "")
    TelnetPswrd = GetSetting("WA1ZEK", "DXtelnet", "TelnetPswrd", "")
    LoginAuto = GetSetting("WA1ZEK", "DXtelnet", "LoginAuto", "0")

    'Time Sync
    TimeSync = GetSetting("WA1ZEK", "DXtelnet", "TimeSync", "0")

    'Alert sounds
    DXsound = GetSetting("WA1ZEK", "DXtelnet", "DXalert", "OFF")
    DXwatch = GetSetting("WA1ZEK", "DXtelnet", "DXwatch", "OFF")
    WWVsound = GetSetting("WA1ZEK", "DXtelnet", "WWValert", "OFF")
    WAVdxspot = GetSetting("WA1ZEK", "DXtelnet", "WAVdxspot", "")
    WAVdxwatch = GetSetting("WA1ZEK", "DXtelnet", "WAvdxwatch", "")
    WAVwwv = GetSetting("WA1ZEK", "DXtelnet", "WAVwwv", "")
    ANNsound = GetSetting("WA1ZEK", "DXtelnet", "ANNsound", "OFF")
    WAVannounce = GetSetting("WA1ZEK", "DXtelnet", "WAVannounce", "")
    Sound_OK = GetSetting("WA1ZEK", "DXtelnet", "Sound_OK", "0")

    'Color and Font settings
    InfoColor = GetSetting("WA1ZEK", "DXtelnet", "InfoColor", vbBlue)
    InfoFont = GetSetting("WA1ZEK", "DXtelnet", "InfoFont", "MS Sans Serif")
    InfoFont_Size = GetSetting("WA1ZEK", "DXtelnet", "InfoFont_size", 10)
    InfoFont_Bold = GetSetting("WA1ZEK", "DXtelnet", "InfoFont_bold", False)
    InfoFont_Italic = GetSetting("WA1ZEK", "DXtelnet", "InfoFont_italic", False)
    MyColor = GetSetting("WA1ZEK", "DXtelnet", "MyColor", vbRed)
    MyFont = GetSetting("WA1ZEK", "DXtelnet", "MyFont", "MS Sans Serif")
    MyFont_Size = GetSetting("WA1ZEK", "DXtelnet", "MyFont_size", 10)
    MyFont_Bold = GetSetting("WA1ZEK", "DXtelnet", "MyFont_bold", False)
    MyFont_Italic = GetSetting("WA1ZEK", "DXtelnet", "MyFont_italic", False)
    DataColor = GetSetting("WA1ZEK", "DXtelnet", "DataColor", vbBlack)
    DataFont = GetSetting("WA1ZEK", "DXtelnet", "DataFont", "MS Sans Serif")
    DataFont_size = GetSetting("WA1ZEK", "DXtelnet", "DataFont_size", 10)
    DataFont_Bold = GetSetting("WA1ZEK", "DXtelnet", "DataFont_bold", False)
    DataFont_Italic = GetSetting("WA1ZEK", "DXtelnet", "DataFont_italic", False)
    SysColor = GetSetting("WA1ZEK", "DXtelnet", "SysColor", vbGreen)
    SysFont = GetSetting("WA1ZEK", "DXtelnet", "SysFont", "MS Sans Serif")
    SysFont_Size = GetSetting("WA1ZEK", "DXtelnet", "SysFont_size", 10)
    SysFont_Bold = GetSetting("WA1ZEK", "DXtelnet", "SysFont_bold", False)
    SysFont_Italic = GetSetting("WA1ZEK", "DXtelnet", "SysFont_italic", False)
    Back_Color = GetSetting("WA1ZEK", "DXtelnet", "Back_Color", vbWhite)

    'DX and WWV Expire times in hours
    Expire_DX = GetSetting("WA1ZEK", "DXtelnet", "Expire_DX", 4)
    Expire_WWV = GetSetting("WA1ZEK", "DXtelnet", "Expire_WWV", 12)
    DXwatchEnabled = GetSetting("WA1ZEK", "DXtelnet", "DXwatchEnabled", True)
    DXarchiveEnabled = GetSetting("WA1ZEK", "DXtelnet", "DXarchiveEnabled", False)
    Use_Local_Clock = GetSetting("WA1ZEK", "DXtelnet", "Use_Local_Clock", False)
    BalloonToolTip = GetSetting("WA1ZEK", "DXtelnet", "BalloonToolTip", False)
    ShowAnnouncements = GetSetting("WA1ZEK", "DXtelnet", "ShowAnnouncements", False)

    AlertColor = 1
    PurgeSpots = True

    'Auto Telnet Login flags

        If LoginAuto = 0 Then
            Login = False
            Passwrd = False
        Else
            Login = True
            Passwrd = True
        End If

    'Load Systray form and hide
    Load frmSystray
    frmSystray.Hide

    'Load Time Sync form

        If TimeSync = 1 Then
            Load frmTimeSync
            frmTimeSync.Hide
            frmTelnet.mnuTimeSync.Enabled = True
        Else
            frmTelnet.mnuTimeSync.Enabled = False
        End If

    'Load main Telnet form
    Load frmTelnet
    frmTelnet.Visible = True

    Load DXwindow
    Call ViewDXHeardList
    Call ViewWWVHeardList
    DXwindow.Visible = True
    frmTelnet.SetFocus

    'Pass the Host name to the function to get IP number from Host Name

    Dim sHostName
    Dim txtRemoteAddress

    sHostName = DXtelnethost
    txtRemoteAddress = GetIPFromHostName(sHostName)

        With frmTelnet

            .RemoteIPAd = txtRemoteAddress
            .RemotePort = DXtelnetport
            .WinsockClient.RemotePort = DXtelnetport
            .WinsockClient.RemoteHost = txtRemoteAddress
    
            .mOpen.Enabled = True
            '.mnuConnect.Enabled = True
            .mSettings.Enabled = True
            .mClose.Enabled = False
            '.mnuClose.Enabled = False
            .mnuAbout.Enabled = True
            .mnuShow.Enabled = False
            .mnuSend.Enabled = False
            .mnuRdxspot.Enabled = False
            .mnuRannounce.Enabled = False
    
        End With

        With DXwindow

            .mnuConnect = True
            .mnuDisconnect = False
            .mnuShow = False
            .mnuSend = False
        End With


End Sub
