VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WA1ZEK DX Cluster Telnet Client - Properties"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8310
   Begin VB.CommandButton CommandCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      MouseIcon       =   "frmProperties.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmProperties.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      MouseIcon       =   "frmProperties.frx":091E
      MousePointer    =   99  'Custom
      Picture         =   "frmProperties.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin TabDlg.SSTab PropertiesTab 
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabHeight       =   520
      MouseIcon       =   "frmProperties.frx":0F32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Bands / Modes"
      TabPicture(0)   =   "frmProperties.frx":124C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text3"
      Tab(0).Control(1)=   "BandGrid"
      Tab(0).Control(2)=   "DeleteRange"
      Tab(0).Control(3)=   "AddRange"
      Tab(0).Control(4)=   "ModeType"
      Tab(0).Control(5)=   "EndKHz"
      Tab(0).Control(6)=   "StartKHz"
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(9)=   "Label3"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Radio Control"
      TabPicture(1)   =   "frmProperties.frx":1268
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text4"
      Tab(1).Control(1)=   "RadioFrame"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Sounds"
      TabPicture(2)   =   "frmProperties.frx":1284
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SoundFileFrame"
      Tab(2).Control(1)=   "SoundFrame"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "DX Set Up"
      TabPicture(3)   =   "frmProperties.frx":12A0
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Text2"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Text7"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Telnet Settings"
      TabPicture(4)   =   "frmProperties.frx":12BC
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TelnetFrame"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "User Information"
      TabPicture(5)   =   "frmProperties.frx":12D8
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame2"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   89
         Text            =   "frmProperties.frx":12F4
         Top             =   4440
         Width           =   5055
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid BandGrid 
         Height          =   1935
         Left            =   -73680
         TabIndex        =   82
         Top             =   2400
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483643
         AllowBigSelection=   0   'False
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         MousePointer    =   99
         MouseIcon       =   "frmProperties.frx":13B6
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.CommandButton DeleteRange 
         Caption         =   "Delete Range"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -71040
         MaskColor       =   &H8000000F&
         MouseIcon       =   "frmProperties.frx":16D0
         MousePointer    =   99  'Custom
         Picture         =   "frmProperties.frx":19DA
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton AddRange 
         Caption         =   "Add Range"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -72840
         MaskColor       =   &H8000000F&
         MouseIcon       =   "frmProperties.frx":1E1C
         MousePointer    =   99  'Custom
         Picture         =   "frmProperties.frx":2126
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox ModeType 
         Height          =   315
         ItemData        =   "frmProperties.frx":2568
         Left            =   -69600
         List            =   "frmProperties.frx":256A
         MouseIcon       =   "frmProperties.frx":256C
         MousePointer    =   99  'Custom
         TabIndex        =   76
         Text            =   "CW"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox EndKHz 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -71640
         TabIndex        =   75
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox StartKHz 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73680
         TabIndex        =   74
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Frame TelnetFrame 
         Caption         =   "Telnet Settings"
         Height          =   4455
         Left            =   -74640
         TabIndex        =   66
         Top             =   720
         Width           =   6975
         Begin VB.CheckBox SetTime 
            Caption         =   "Sync CPU Time every Hour"
            Height          =   195
            Left            =   3840
            MouseIcon       =   "frmProperties.frx":2876
            MousePointer    =   99  'Custom
            TabIndex        =   90
            Top             =   960
            Width           =   2295
         End
         Begin VB.CheckBox AutoLogin 
            Caption         =   "Use Auto Logon"
            Height          =   195
            Left            =   1320
            MouseIcon       =   "frmProperties.frx":2B80
            MousePointer    =   99  'Custom
            TabIndex        =   88
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox TelnetPassword 
            Height          =   285
            Left            =   4200
            TabIndex        =   87
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox TelnetUser 
            Height          =   285
            Left            =   1320
            TabIndex        =   85
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton RemoveTelnet 
            Caption         =   "Remove from Telnet List"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3840
            MouseIcon       =   "frmProperties.frx":2E8A
            MousePointer    =   99  'Custom
            Picture         =   "frmProperties.frx":3194
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000080&
            Height          =   820
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   73
            Text            =   "frmProperties.frx":35D6
            Top             =   3600
            Width           =   6255
         End
         Begin VB.CommandButton AddTelnet 
            Caption         =   "Add to Telnet List"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1320
            MouseIcon       =   "frmProperties.frx":36DE
            MousePointer    =   99  'Custom
            Picture         =   "frmProperties.frx":39E8
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   1200
            Width           =   2175
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid TelnetGrid 
            Height          =   1335
            Left            =   360
            TabIndex        =   71
            Top             =   2160
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   2355
            _Version        =   393216
            FixedCols       =   0
            BackColorSel    =   16711680
            ForeColorSel    =   -2147483643
            AllowBigSelection=   0   'False
            FocusRect       =   0
            SelectionMode   =   1
            MousePointer    =   99
            RowSizingMode   =   1
            MouseIcon       =   "frmProperties.frx":3E2A
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.TextBox TelnetPort 
            Height          =   285
            Left            =   5400
            TabIndex        =   70
            Text            =   "23"
            Top             =   280
            Width           =   735
         End
         Begin VB.TextBox TelnetHost 
            Height          =   285
            Left            =   1320
            TabIndex        =   68
            Text            =   "dxc.kb1h.com"
            Top             =   280
            Width           =   3495
         End
         Begin VB.Label LabelPassword 
            Caption         =   "Password"
            Height          =   255
            Left            =   3360
            TabIndex        =   86
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Labeluser 
            Alignment       =   2  'Center
            Caption         =   "User Name"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Caption         =   "Port"
            Height          =   255
            Left            =   4920
            TabIndex        =   69
            Top             =   320
            Width           =   375
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Caption         =   "Telnet Host"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   320
            Width           =   975
         End
      End
      Begin VB.Frame SoundFileFrame 
         Caption         =   "Sound File Select"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1695
         Left            =   -74760
         TabIndex        =   54
         Top             =   3360
         Width           =   7215
         Begin VB.CommandButton RemoveWave 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Remove Wave"
            Height          =   855
            Left            =   1530
            MouseIcon       =   "frmProperties.frx":4144
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   720
            Width           =   1300
         End
         Begin VB.CommandButton FindWAV 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Select File"
            Height          =   855
            Left            =   120
            MouseIcon       =   "frmProperties.frx":444E
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   720
            Width           =   1300
         End
         Begin VB.CommandButton CancelWAV 
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
            Left            =   5790
            MouseIcon       =   "frmProperties.frx":4758
            MousePointer    =   99  'Custom
            Picture         =   "frmProperties.frx":4A62
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   720
            Width           =   1300
         End
         Begin VB.TextBox WAVfile 
            Height          =   315
            Left            =   120
            TabIndex        =   60
            Text            =   "WAVfile"
            Top             =   240
            Width           =   6975
         End
         Begin VB.CommandButton testWAV 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Play Wave"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2970
            MouseIcon       =   "frmProperties.frx":4D6C
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   720
            Width           =   1300
         End
         Begin VB.CommandButton okWAV 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4400
            MouseIcon       =   "frmProperties.frx":5076
            MousePointer    =   99  'Custom
            Picture         =   "frmProperties.frx":5380
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   720
            Width           =   1300
         End
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000080&
         Height          =   1335
         Left            =   4440
         MultiLine       =   -1  'True
         TabIndex        =   51
         Text            =   "frmProperties.frx":568A
         Top             =   3840
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000080&
         Height          =   1335
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   50
         Text            =   "frmProperties.frx":5741
         Top             =   3840
         Width           =   3975
      End
      Begin VB.Frame Frame3 
         Caption         =   "DX and WWV Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2895
         Left            =   4440
         TabIndex        =   42
         Top             =   840
         Width           =   3015
         Begin VB.CheckBox AnnounceOK 
            Caption         =   " Show Announcements"
            Height          =   375
            Left            =   360
            TabIndex        =   96
            Top             =   2160
            Width           =   2535
         End
         Begin VB.CheckBox BalloonOK 
            Caption         =   " Enable Balloon Tool Tip"
            Height          =   255
            Left            =   360
            TabIndex        =   95
            Top             =   2520
            Width           =   2535
         End
         Begin VB.CheckBox CPUclock 
            Caption         =   " Use Local Clock for Spot Time"
            Height          =   255
            Left            =   360
            TabIndex        =   94
            Top             =   1920
            Width           =   2535
         End
         Begin VB.CheckBox DXarchive 
            Caption         =   " Archive DX Spots to File"
            Height          =   255
            Left            =   360
            MouseIcon       =   "frmProperties.frx":5814
            MousePointer    =   99  'Custom
            TabIndex        =   65
            Top             =   1610
            Width           =   2535
         End
         Begin VB.ComboBox WWV_Expire 
            Height          =   315
            ItemData        =   "frmProperties.frx":5B1E
            Left            =   1320
            List            =   "frmProperties.frx":5B75
            MouseIcon       =   "frmProperties.frx":5BCC
            MousePointer    =   99  'Custom
            TabIndex        =   47
            Text            =   "12"
            Top             =   1200
            Width           =   615
         End
         Begin VB.ComboBox DX_Expire 
            Height          =   315
            ItemData        =   "frmProperties.frx":5ED6
            Left            =   1320
            List            =   "frmProperties.frx":5F01
            MouseIcon       =   "frmProperties.frx":5F2C
            MousePointer    =   99  'Custom
            TabIndex        =   46
            Text            =   "4"
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Hours"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2040
            TabIndex        =   49
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Hours"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2040
            TabIndex        =   48
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "WWV Reports"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "DX Spots"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Choose how long DX Spots and WWV reports are kept"
            Height          =   495
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Product Registration"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3015
         Left            =   -74400
         TabIndex        =   33
         Top             =   1080
         Width           =   6375
         Begin VB.TextBox RegNumText 
            Height          =   285
            Left            =   1920
            TabIndex        =   41
            Top             =   2400
            Width           =   4215
         End
         Begin VB.TextBox ProdIDtext 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            TabIndex        =   40
            Text            =   "CM010-2004-XXXXXX"
            Top             =   1920
            Width           =   4215
         End
         Begin VB.TextBox CallText 
            Height          =   285
            Left            =   1920
            TabIndex        =   39
            Top             =   960
            Width           =   4215
         End
         Begin VB.TextBox NameText 
            Height          =   285
            Left            =   1920
            TabIndex        =   38
            Top             =   480
            Width           =   4215
         End
         Begin VB.Label Label6 
            Caption         =   "Registration Number"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   360
            TabIndex        =   37
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Product ID"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   360
            TabIndex        =   36
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label CallLabel 
            Caption         =   "Call"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   360
            TabIndex        =   35
            Top             =   960
            Width           =   375
         End
         Begin VB.Label NameLabel 
            Caption         =   "Name"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "DX Watch List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2895
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   4095
         Begin VB.CheckBox WatchEnabled 
            Caption         =   "DX Alert Enabled"
            Height          =   315
            Left            =   2400
            MouseIcon       =   "frmProperties.frx":6236
            MousePointer    =   99  'Custom
            TabIndex        =   64
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox DXlistText 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   2520
            Width           =   1575
         End
         Begin VB.CommandButton DXdelete 
            Caption         =   "Delete DX"
            Enabled         =   0   'False
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
            Left            =   2520
            Picture         =   "frmProperties.frx":6540
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton DXAdd 
            Caption         =   "Add DX"
            Enabled         =   0   'False
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
            Left            =   2520
            Picture         =   "frmProperties.frx":6982
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   600
            Width           =   1335
         End
         Begin VB.ListBox DXList 
            Enabled         =   0   'False
            Height          =   2595
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000080&
         Height          =   1455
         Left            =   -74040
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "frmProperties.frx":6DC4
         Top             =   3600
         Width           =   5535
      End
      Begin VB.Frame RadioFrame 
         Caption         =   "Radio Control"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2415
         Left            =   -74040
         TabIndex        =   8
         Top             =   960
         Width           =   5655
         Begin VB.ComboBox CommPort 
            Height          =   315
            Left            =   1320
            MouseIcon       =   "frmProperties.frx":6F1A
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Text            =   "COM 1"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.ComboBox CommBits 
            Height          =   315
            Left            =   3960
            MouseIcon       =   "frmProperties.frx":7224
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Text            =   "8"
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox RadioType 
            Height          =   315
            Left            =   1320
            MouseIcon       =   "frmProperties.frx":752E
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Text            =   "None"
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox CommStop 
            Height          =   315
            Left            =   3960
            MouseIcon       =   "frmProperties.frx":7838
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Text            =   "1"
            Top             =   960
            Width           =   1455
         End
         Begin VB.ComboBox CommParity 
            Height          =   315
            Left            =   3960
            MouseIcon       =   "frmProperties.frx":7B42
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Text            =   "None"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.ComboBox CommFlow 
            Height          =   315
            Left            =   3960
            MouseIcon       =   "frmProperties.frx":7E4C
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Text            =   "XON/XOFF"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.ComboBox CommSpeed 
            Height          =   315
            Left            =   1320
            MouseIcon       =   "frmProperties.frx":8156
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Text            =   "9600"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.ComboBox VFORadio 
            Height          =   315
            Left            =   1320
            MouseIcon       =   "frmProperties.frx":8460
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Text            =   "VFO A"
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label RadioModelLabel 
            Caption         =   "Radio Model"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label CommPortLabel 
            Caption         =   "COM Port"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label CommBitsLabel 
            Caption         =   "Data Bits"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   3000
            TabIndex        =   22
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label CommStopLabel 
            Caption         =   "Stop Bits"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   3000
            TabIndex        =   21
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label CommParityLabel 
            Caption         =   "Parity"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   3000
            TabIndex        =   20
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label CommFlowLabel 
            Caption         =   "Flow Control"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   3000
            TabIndex        =   19
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label CommSpeedLabel 
            Caption         =   "Speed"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label RadioVFOLabel 
            Caption         =   "Radio VFO"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame SoundFrame 
         Caption         =   "Alert Sounds"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2655
         Left            =   -74760
         TabIndex        =   3
         Top             =   720
         Width           =   7215
         Begin VB.TextBox annWAV 
            Height          =   285
            Left            =   2760
            TabIndex        =   92
            Text            =   "AnnounceWav"
            Top             =   1800
            Width           =   4335
         End
         Begin VB.ComboBox ANNonoff 
            Height          =   315
            Left            =   1440
            MouseIcon       =   "frmProperties.frx":876A
            MousePointer    =   99  'Custom
            TabIndex        =   91
            Text            =   "OFF"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox wwvWAV 
            Height          =   285
            Left            =   2760
            TabIndex        =   57
            Text            =   "wwvWav"
            Top             =   1320
            Width           =   4335
         End
         Begin VB.TextBox watchWAV 
            Height          =   285
            Left            =   2760
            TabIndex        =   56
            Text            =   "watchWav"
            Top             =   840
            Width           =   4335
         End
         Begin VB.TextBox spotWAV 
            Height          =   285
            Left            =   2760
            TabIndex        =   55
            Text            =   "spotWav"
            Top             =   360
            Width           =   4335
         End
         Begin VB.CheckBox SoundOK 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1440
            MouseIcon       =   "frmProperties.frx":8A74
            MousePointer    =   99  'Custom
            TabIndex        =   52
            Top             =   2280
            Width           =   255
         End
         Begin VB.ComboBox DX_watch 
            Height          =   315
            Left            =   1440
            MouseIcon       =   "frmProperties.frx":8D7E
            MousePointer    =   99  'Custom
            TabIndex        =   32
            Text            =   "OFF"
            Top             =   840
            Width           =   1215
         End
         Begin VB.ComboBox DXonoff 
            Height          =   315
            Left            =   1440
            MouseIcon       =   "frmProperties.frx":9088
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Text            =   "OFF"
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox WWVonoff 
            Height          =   315
            Left            =   1440
            MouseIcon       =   "frmProperties.frx":9392
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Text            =   "OFF"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Announcement"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label SoundOKLabel 
            Caption         =   "Use .WAV Sounds"
            Height          =   255
            Left            =   1680
            MouseIcon       =   "frmProperties.frx":969C
            MousePointer    =   99  'Custom
            TabIndex        =   53
            Top             =   2310
            Width           =   1455
         End
         Begin VB.Label DXwatchlabel 
            Caption         =   "DX Watch"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   840
            Width           =   975
         End
         Begin VB.Label DXSpotLabel 
            Caption         =   "DX Spot "
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label WWVAlertLabel 
            Caption         =   "WWV Alert"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   1215
         End
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Mode"
         Height          =   255
         Left            =   -69600
         TabIndex        =   81
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Stop Frequency (KHz)"
         Height          =   255
         Left            =   -71640
         TabIndex        =   80
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Start Frequency (KHz)"
         Height          =   255
         Left            =   -73680
         TabIndex        =   79
         Top             =   840
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog PropDialog 
      Left            =   240
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    '*************************************************************************************
    '
    'Proprerties window
    '
    'Get and set registry values and operating paramiters for AX25 Monitor
    '
    '*************************************************************************************

    'API cal to see if soundcard is present
    Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
    'API call to get UserName from system
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

    'Some module variables
    Dim WAVselect                As Integer
    '0 = DX spot wave
    '1 = DX watch wave
    '2 = WWV wave
    Dim WAVfileselect            As String
    'To return wave file spec to caling sub
    Dim TempCall                 As String
    Dim TempDXCall               As String

 
Private Sub CallDX_Change()

        If CallDX.Text <> "" Then
            'Force uppercase
            CallDX.Text = UCase(CallDX.Text)
            CallDX.SelStart = Len(CallDX.Text) + 1
        End If

End Sub

Private Sub CallDX_Click()

    CallDX.SelStart = 0
    CallDX.SelLength = Len(CallDX.Text) + 1

End Sub

Private Sub CallDX_DblClick()

    CallDX.SelStart = Len(CallDX.Text) + 1

End Sub

Private Sub CallDX_GotFocus()

    TempDXCall = CallDX.Text

End Sub

Private Sub CallDX_LostFocus()

        If CallDX.Text = CallTerminal.Text Then

            MsgBox "Packet Terminal and DX Terminal calls can not be the same." & vbCrLf & _
            "Choose a different SSID ('Dash' number) for each.", vbCritical, "Application ERROR"
            CallDX.Text = TempDXCall
        End If

End Sub

Private Sub CallTerminal_Change()

        If CallTerminal.Text <> "" Then
            'Force uppercase
            CallTerminal.Text = UCase(CallTerminal.Text)
            CallTerminal.SelStart = Len(CallTerminal.Text) + 1
        End If

End Sub

Private Sub CallTerminal_Click()

    CallTerminal.SelStart = 0
    CallTerminal.SelLength = Len(CallTerminal.Text) + 1

End Sub

Private Sub CallTerminal_DblClick()

    CallTerminal.SelStart = Len(CallTerminal.Text) + 1

End Sub

Private Sub CallTerminal_GotFocus()

    TempCall = CallTerminal.Text

End Sub

Private Sub CallTerminal_LostFocus()

        If CallTerminal.Text = CallDX.Text Then
            MsgBox "Packet Terminal and DX Terminal calls can not be the same." & vbCrLf & _
            "Choose a different SSID ('Dash' number) for each.", vbCritical, "Application ERROR"
            CallTerminal.Text = TempCall
        End If
   
End Sub

Private Sub AddRange_Click()

    'Set format for frequency to XXXX.XX KHz
    LowerKHz = Format(StartKHz.Text, "#######0.00")
    UpperKHz = Format(EndKHz.Text, "#######0.00")
    'Add Band/Mode to list

        If StartKHz.Text <> "" Or EndKHz.Text <> "" Then
            BandGrid.AddItem LowerKHz & vbTab & UpperKHz & vbTab & ModeType.Text, 1
        End If

    'Clear input text after add
    StartKHz.Text = ""
    EndKHz.Text = ""

    'Sort the grid, assending order by StartKHz
    BandGrid.Col = 0
    BandGrid.Sort = 3

    'Get rid of empty first line after sort
    BandGrid.Row = 1
    BandGrid.Col = 0

        If BandGrid.Text = "" Then
            BandGrid.RemoveItem (BandGrid.Row)
        End If

    'Disable Add Button
    AddRange.Enabled = False

End Sub

Private Sub AddTelnet_Click()

    'Get number of rows, then add new entry to bottom of list
    GridRows = (TelnetGrid.Rows - 1)
    TelnetGrid.AddItem TelnetHost.Text & vbTab & TelnetPort.Text & vbTab & TelnetUser.Text & vbTab & TelnetPassword.Text, GridRows

End Sub

Private Sub AddTelnet_GotFocus()

    RemoveTelnet.Enabled = False

End Sub

Private Sub ANNonoff_Change()

        If SoundOK.Value = 0 Then Exit Sub

        Select Case ANNonoff
            Case "ON"
                annWAV.BackColor = &H80000005
                annWAV.Enabled = True
            Case "OFF"
                annWAV.BackColor = &H80000004
                annWAV.Enabled = False
        End Select

End Sub

Private Sub ANNonoff_Click()

    ANNonoff_Change

End Sub

Private Sub AnnounceOK_Click()

        If AnnounceOK.Value = 0 Then
            ShowAnnouncements = False
        Else
            ShowAnnouncements = True
        End If

End Sub

Private Sub annWAV_Click()

    'IF Select wave open, save current file and then open new selection
    'Open wave select for new selection
    SoundFileFrame.Caption = "Select Announce Alert Sound"
    WAVselect = 3
    WAVfile.Text = annWAV.Text
    Call EnableWaveFind

End Sub

Private Sub AutoLogin_Click()

        If AutoLogin.Value = 1 Then
            TelnetUser.Enabled = True
            TelnetUser.BackColor = &H80000005
            TelnetUser.Visible = True
            Labeluser.Visible = True
            TelnetUser.Text = TelnetLogon
            TelnetPassword.Enabled = True
            TelnetPassword.BackColor = &H80000005
            TelnetPassword.Visible = True
            LabelPassword.Visible = True
            TelnetPassword.Text = TelnetPswrd
    
            'Don't change AutoLogin state while connected

                If frmTelnet.WinsockClient.State > 0 Then
                    Passwrd = False
                    Login = False
                Else
                    Passwrd = True
                    Login = True
                End If
    
        Else
            TelnetUser.Enabled = False
            TelnetUser.BackColor = &H8000000F
            TelnetUser.Visible = False
            Labeluser.Visible = False
            TelnetPassword.Enabled = False
            TelnetPassword.BackColor = &H8000000F
            TelnetPassword.Visible = False
            LabelPassword.Visible = False
            Passwrd = False
            Login = False

        End If

End Sub



Private Sub BalloonOK_Click()

        If BalloonOK.Value = 0 Then
            BalloonToolTip = False
        Else
            BalloonToolTip = True
        End If

End Sub

Private Sub BandGrid_Click()

    DeleteRange.Enabled = True
    AddRange.Enabled = False

End Sub

Private Sub BandGrid_LostFocus()

    DeleteRange.Enabled = False

End Sub

Private Sub CallText_Change()

    CallText.SelStart = Len(CallText.Text) + 1
    CallText.Text = UCase(CallText.Text)

End Sub

Private Sub CallText_Click()

    CallText.SelStart = 0
    CallText.SelLength = Len(CallText.Text) + 1

End Sub

Private Sub CallText_DblClick()

    CallText.SelStart = Len(CallText.Text) + 1

End Sub


Private Sub CallText_GotFocus()

    AddRange.Enabled = False

End Sub

Private Sub CancelWAV_Click()

    'SoundFileFrame.Visible = False
    Call DisableWaveFind

End Sub


Private Sub CommandCancel_Click()

    Unload frmProperties

End Sub



Private Sub CommandOK_Click()

    'Save setting
    'Radio settings
    RadioModel = RadioType.Text
    SaveSetting "WA1ZEK", "DXtelnet", "Radio", RadioType.Text
    RadioVFO = VFORadio.Text
    SaveSetting "WA1ZEK", "DXtelnet", "VFO", VFORadio.Text
    RadioComPort = CommPort.Text
    SaveSetting "WA1ZEK", "DXtelnet", "ComPort", CommPort.Text
    RadioComSpeed = CommSpeed.Text
    SaveSetting "WA1ZEK", "DXtelnet", "ComSpeed", CommSpeed.Text
    RadioComBits = CommBits.Text
    SaveSetting "WA1ZEK", "DXtelnet", "Combits", CommBits.Text
    RadioComStop = CommStop.Text
    SaveSetting "WA1ZEK", "DXtelnet", "ComStop", CommStop.Text
    RadioComParity = CommParity.Text
    SaveSetting "WA1ZEK", "DXtelnet", "ComParity", CommParity.Text
    RadioComFlow = CommFlow.Text
    SaveSetting "WA1ZEK", "DXtelnet", "ComFlow", CommFlow.Text
    'Telnet Settings
    DXtelnethost = TelnetHost.Text
    SaveSetting "WA1ZEK", "DXtelnet", "DXtelnethost", DXtelnethost
    DXtelnetport = TelnetPort.Text
    SaveSetting "WA1ZEK", "DXtelnet", "DXtelnetport", DXtelnetport
    TelnetLogon = TelnetUser.Text
    SaveSetting "WA1ZEK", "DXtelnet", "TelnetLogon", TelnetLogon
    TelnetPswrd = TelnetPassword.Text
    SaveSetting "WA1ZEK", "DXtelnet", "TelnetPswrd", TelnetPswrd
    LoginAuto = AutoLogin.Value
    SaveSetting "WA1ZEK", "DXtelnet", "LoginAuto", LoginAuto

    'Alert sounds
    DXsound = DXonoff.Text
    SaveSetting "WA1ZEK", "DXtelnet", "DXalert", DXonoff.Text
    DXwatch = DX_watch.Text
    SaveSetting "WA1ZEK", "DXtelnet", "DXwatch", DX_watch.Text
    WWVsound = WWVonoff.Text
    SaveSetting "WA1ZEK", "DXtelnet", "WWValert", WWVonoff.Text
    WAVdxspot = spotWAV.Text
    SaveSetting "WA1ZEK", "DXtelnet", "WAVdxspot", spotWAV.Text
    WAVdxwatch = watchWAV.Text
    SaveSetting "WA1ZEK", "DXtelnet", "WAVdxwatch", watchWAV.Text
    WAVwwv = wwvWAV.Text
    SaveSetting "WA1ZEK", "DXtelnet", "WAVwwv", wwvWAV.Text
    ANNsound = ANNonoff.Text
    SaveSetting "WA1ZEK", "DXtelnet", "ANNsound", ANNonoff.Text
    WAVannounce = annWAV.Text
    SaveSetting "WA1ZEK", "DXtelnet", "WAVannounce", annWAV.Text
    Sound_OK = SoundOK.Value
    SaveSetting "WA1ZEK", "DXtelnet", "Sound_OK", SoundOK.Value

    'Registration values
    username = NameText.Text
    SaveSetting "WA1ZEK", "DXtelnet", "UserName", NameText.Text
    UserCall = CallText.Text
    SaveSetting "WA1ZEK", "DXtelnet", "UserCall", CallText.Text
    RegNum = RegNumText.Text
    SaveSetting "WA1ZEK", "DXtelnet", "RegNum", RegNumText.Text


    'DX Watch and Spot expire settings
    'Timeouts
    Expire_DX = Val(DX_Expire.Text)
    SaveSetting "WA1ZEK", "DXtelnet", "Expire_DX", Expire_DX
    Expire_WWV = Val(WWV_Expire.Text)
    SaveSetting "WA1ZEK", "DXtelnet", "Expire_WWV", Expire_WWV
    SaveSetting "WA1ZEK", "DXtelnet", "Use_Local_Clock", CPUclock.Value
    SaveSetting "WA1ZEK", "DXtelnet", "BalloonToolTip", BalloonOK.Value
    SaveSetting "WA1ZEK", "DXtelnet", "ShowAnnouncements", AnnounceOK.Value

    'Save Watch List to file
    On Error Resume Next
    Kill App.Path & "\" & "DXwatch.lst"
    watchlist = FreeFile
    Open App.Path & "\" & "DXwatch.lst" For Output As watchlist

        For i = 0 To ((DXList.ListCount) - 1)
            Write #watchlist, DXList.List(i)
        Next i

    Close watchlist

        If WatchEnabled.Value = 1 Then
            DXwatchEnabled = True
        Else
            DXwatchEnabled = False
        End If

    SaveSetting "WA1ZEK", "DXtelnet", "DXwatchEnabled", DXwatchEnabled
    DXarchiveEnabled = DXarchive.Value
    SaveSetting "WA1ZEK", "DXtelnet", "DXarchiveEnabled", DXarchiveEnabled

    'Save Telnet Host settings
    Call modTelnetList.TelnetListSave
    'Save Band-Mode Settings
    Call modBandList.BandListSave

    'Pass the Host name to the function to get IP number from Host Name
    sHostName = DXtelnethost
    txtRemoteAddress = GetIPFromHostName(sHostName)

    frmTelnet.RemoteIPAd = txtRemoteAddress
    frmTelnet.RemotePort = DXtelnetport
    frmTelnet.WinsockClient.RemotePort = DXtelnetport
    frmTelnet.WinsockClient.RemoteHost = txtRemoteAddress

    'Time Sync
    TimeSync = SetTime.Value
    SaveSetting "WA1ZEK", "DXtelnet", "TimeSync", SetTime.Value
    'Start or stop frmTimeSync on checkbox change

        If TimeSync = 1 Then
            Load frmTimeSync
            frmTelnet.mnuTimeSync.Enabled = True
        Else
            Unload frmTimeSync
            frmTelnet.mnuTimeSync.Enabled = False
        End If
    
    'Unload form
    Unload frmProperties

End Sub

Private Sub CPUclock_Click()

        If CPUclock.Value = 1 Then
            'Use local clock for spot incomming time
            Use_Local_Clock = True
        Else
            'Use Cluster time for incomming spot
            Use_Local_Clock = False
        End If
    
End Sub

Private Sub DeleteRange_Click()

    'Delete a grid entry

        If BandGrid.Rows = 2 Then   'Can't remove the last row, so add one
            BandGrid.AddItem "" & vbTab & "" & vbTab & "", (BandGrid.Rows)
            BandGrid.RemoveItem (BandGrid.Row)
        Else
            BandGrid.RemoveItem (BandGrid.Row)
        End If

    'Sort the grid, assending order by StartKHz
    BandGrid.Col = 0
    BandGrid.Sort = 3

    'Disable Delete button
    DeleteRange.Enabled = False

End Sub

Private Sub DeleteRange_GotFocus()

    AddRange.Enabled = False

End Sub

Private Sub DX_Expire_GotFocus()

    DXdelete.Enabled = False

End Sub

Private Sub DX_watch_Change()

        If SoundOK.Value = 0 Then Exit Sub

        Select Case DX_watch
            Case "ON"
                watchWAV.BackColor = &H80000005
                watchWAV.Enabled = True
            Case "OFF"
                watchWAV.BackColor = &H80000004
                watchWAV.Enabled = False
        End Select

End Sub

Private Sub DX_watch_Click()

    DX_watch_Change

End Sub

Private Sub DXAdd_Click()

    'Add item to DX watch list
    DXList.AddItem UCase(Mid(DXlistText.Text, 1, Len(DXlistText)))
    'Clear text in input box (DXlistText)
    DXlistText.Text = ""
    'Restore focus to text input (for more adds)
    DXlistText.SetFocus
    DXlistText.SelStart = Len(DXlistText) + 1

End Sub

Private Sub DXarchive_GotFocus()

    DXdelete.Enabled = False

End Sub

Private Sub DXdelete_Click()

    'Remove item from list
    ListI = DXList.ListIndex
    DXList.RemoveItem ListI
    'Remove focus of control, shift to text input
    DXdelete.Enabled = False
    
End Sub

Private Sub DXList_Click()

    'Change Delete button enabled when item selected
    DXdelete.Enabled = True
    DXdelete.SetFocus

End Sub

Private Sub DXList_GotFocus()

    DXdelete.Enabled = True

End Sub

Private Sub DXlistText_Change()

    'Change Add button enabled based on text in DXlistText

        If DXlistText <> "" Then
            DXAdd.Enabled = True
            'Force uppercase
            DXlistText.Text = UCase(DXlistText.Text)
            DXlistText.SelStart = Len(DXlistText) + 1
        Else
            DXAdd.Enabled = False
        End If

    DXdelete.Enabled = False

End Sub

Private Sub DXlistText_GotFocus()

        If DXlistText.Text <> "" Then
            DXAdd.Enabled = True
        End If

    DXdelete.Enabled = False

End Sub

Private Sub DXlistText_KeyUp(KeyCode As Integer, Shift As Integer)

    'Test for <CR> to entertext into DXlist box

        If KeyCode = 13 Then
            DXlistText = (Mid(DXlistText.Text, 1, (Len(DXlistText) - 2)))
            DXAdd_Click
        End If

End Sub

Private Sub DXlistText_LostFocus()

    'DXlistText.Text = ""
    'DXAdd.Enabled = False

End Sub

Private Sub DXonoff_Change()

        If SoundOK.Value = 0 Then Exit Sub

        Select Case DXonoff
            Case "ON"
                spotWAV.BackColor = &H80000005
                spotWAV.Enabled = True
            Case "OFF"
                spotWAV.BackColor = &H80000004
                spotWAV.Enabled = False
        End Select

End Sub

Private Sub DXonoff_Click()

    DXonoff_Change

End Sub

Private Sub FindWAV_Click()

    'PropDialog.Path = App.Path & "\wave"
    PropDialog.InitDir = App.Path & "\wav files\"
    PropDialog.ShowOpen
    WAVfile.Text = PropDialog.FileName

End Sub

Private Sub Font1_Click()

        With PropDialog
            .FontName = Font1.Text
            .FontBold = InfoFont_Bold
            .FontItalic = InfoFont_Italic
            .FontSize = Int(InfoFont_Size)
            .ShowFont
        End With

    Font1.Font = PropDialog.FontName
    Font1.Font.Bold = PropDialog.FontBold
    Font1.Font.Italic = PropDialog.FontItalic
    Font1.Text = PropDialog.FontName
    Font1.FontSize = PropDialog.FontSize    'Only one size supported, change all font sizes
    Font2.FontSize = PropDialog.FontSize
    Font3.FontSize = PropDialog.FontSize
    Font4.FontSize = PropDialog.FontSize

End Sub

Private Sub Font2_Click()

        With PropDialog
            .FontName = Font2.Text
            .FontBold = MyFont_Bold
            .FontItalic = MyFont_Italic
            .FontSize = Int(MyFont_Size)
            .ShowFont
        End With

    Font2.Font = PropDialog.FontName
    Font2.Font.Bold = PropDialog.FontBold
    Font2.Font.Italic = PropDialog.FontItalic
    Font2.Text = PropDialog.FontName
    Font1.FontSize = PropDialog.FontSize    'Only one size supported, change all font sizes
    Font2.FontSize = PropDialog.FontSize
    Font3.FontSize = PropDialog.FontSize
    Font4.FontSize = PropDialog.FontSize

End Sub

Private Sub Font3_Click()

        With PropDialog
            .FontName = Font3.Text
            .FontBold = DataFont_Bold
            .FontItalic = DataFont_Italic
            .FontSize = Int(DataFont_size)
            .ShowFont
        End With

    Font3.Font = PropDialog.FontName
    Font3.Font.Bold = PropDialog.FontBold
    Font3.Font.Italic = PropDialog.FontItalic
    Font3.Text = PropDialog.FontName
    Font1.FontSize = PropDialog.FontSize    'Only one size supported, change all font sizes
    Font2.FontSize = PropDialog.FontSize
    Font3.FontSize = PropDialog.FontSize
    Font4.FontSize = PropDialog.FontSize

End Sub

Private Sub Font4_Click()

        With PropDialog
            .FontName = Font4.Text
            .FontBold = SysFont_Bold
            .FontItalic = SysFont_Italic
            .FontSize = Int(SysFont_Size)
            .ShowFont
        End With

    Font4.Font = PropDialog.FontName
    Font4.Font.Bold = PropDialog.FontBold
    Font4.Font.Italic = PropDialog.FontItalic
    Font4.Text = PropDialog.FontName
    Font1.FontSize = PropDialog.FontSize    'Only one size supported, change all font sizes
    Font2.FontSize = PropDialog.FontSize
    Font3.FontSize = PropDialog.FontSize
    Font4.FontSize = PropDialog.FontSize

End Sub

Private Sub Form_Load()

    'Disable Band - Mode controls if no radio

        If RadioModel = "None" Then
            StartKHz.Enabled = False
            StartKHz.BackColor = &H8000000F
            EndKHz.Enabled = False
            EndKHz.BackColor = &H8000000F
            ModeType.Enabled = False
            ModeType.BackColor = &H8000000F
            AddRange.Enabled = False
            DeleteRange.Enabled = False
            BandGrid.Enabled = False
            BandGrid.BackColor = &H8000000F
        Else
            'Enable Band-Mode Tab
            StartKHz.Enabled = True
            StartKHz.BackColor = &H80000005
            EndKHz.Enabled = True
            EndKHz.BackColor = &H80000005
            ModeType.Enabled = True
            ModeType.BackColor = &H80000005
            AddRange.Enabled = False
            DeleteRange.Enabled = False
            BandGrid.Enabled = True
            BandGrid.BackColor = &H80000005
        End If

    'Set up dropdown lists
    'Supported Radios
    RadioType.AddItem "None", 0
    RadioType.AddItem "TS-2000S", 1
    RadioType.AddItem "TS-2000X", 2
    RadioType.AddItem "TS-870S", 3
    RadioType.AddItem "TS-570S(G)", 4
    RadioType.AddItem "TS-570D(G)", 5
    RadioType.AddItem "TS-950S(DX)", 6
    RadioType.AddItem "TS-850S", 7
    RadioType.AddItem "TS-450S", 8
    RadioType.AddItem "TS-480S(AT/H)" '9
    'RadioType.AddItem "TS-940S", 9
    'RadioType.AddItem "TS-440S", 10

    'Radio VFO's
    VFORadio.AddItem "VFO A", 0
    VFORadio.AddItem "VFO B", 1

    'Radio Comm ports
    CommPort.AddItem "COM 1", 0
    CommPort.AddItem "COM 2", 1
    CommPort.AddItem "COM 3", 2
    CommPort.AddItem "COM 4", 3
    CommPort.AddItem "COM 5", 4
    CommPort.AddItem "COM 6", 5
    CommPort.AddItem "COM 7", 6
    CommPort.AddItem "COM 8", 7

    'Radio Comm port speed
    CommSpeed.AddItem "300", 0
    CommSpeed.AddItem "1200", 1
    CommSpeed.AddItem "2400", 2
    CommSpeed.AddItem "4800", 3
    CommSpeed.AddItem "9600", 4
    CommSpeed.AddItem "19200", 5
    CommSpeed.AddItem "38400", 6
    CommSpeed.AddItem "57600", 7
    CommSpeed.AddItem "115200", 8

    'Comm Databits
    CommBits.AddItem "8", 0
    CommBits.AddItem "7", 1

    'CommStop bits
    CommStop.AddItem "1", 0
    CommStop.AddItem "2", 1

    'Comm Parity bits
    CommParity.AddItem "None", 0
    CommParity.AddItem "Even", 1
    CommParity.AddItem "Odd", 2
    CommParity.AddItem "Mark", 3
    CommParity.AddItem "Space", 4

    'Comm Flow control
    CommFlow.AddItem "None", 0
    CommFlow.AddItem "XON/XOFF", 1
    CommFlow.AddItem "RTS/CTS", 2

    'DX Alert
    DXonoff.AddItem "OFF", 0
    DXonoff.AddItem "ON", 1
    spotWAV.Text = WAVdxspot

    'WWV Alert
    WWVonoff.AddItem "OFF", 0
    WWVonoff.AddItem "ON", 1
    wwvWAV.Text = WAVwwv

    'DX Watch
    DX_watch.AddItem "OFF", 0
    DX_watch.AddItem "ON", 1
    watchWAV.Text = WAVdxwatch

    'Announcement Alert
    ANNonoff.AddItem "OFF", 0
    ANNonoff.AddItem "ON", 1
    annWAV.Text = WAVannounce

    'Band Limits
    ModeType.AddItem "CW", 0
    ModeType.AddItem "LSB", 1
    ModeType.AddItem "USB", 2
    ModeType.AddItem "FM", 3
    ModeType.AddItem "AM", 4


    'Test for soundcard
    rtn = waveOutGetNumDevs() 'check for a sound card

        If rtn > 0 Then 'if returned is greater than 0 then a sound card exists
            SoundOK.Value = Sound_OK
            SoundOK.Visible = True
   
                If SoundOK.Value = 1 Then
                    SoundFileFrame.Visible = True
                    Call DisableWaveFind
                Else
                    Call DisableWaveFind
                End If
   
        Else 'otherwise no sound card found
            SoundOK.Value = 0
            SoundOK.Visible = True
            SoundOK.Enabled = False
            SoundOK.BackColor = &H8000000F
            Call DisableWaveFind
            spotWAV.Visible = True
            watchWAV.Visible = True
            wwvWAV.Visible = True
            annWAV.Visible = True

        End If

    'Settings from registry

    'Radio Settings
    RadioType.Text = RadioModel
    RadioType_Click 'Disable the rest of the radio settings if "None" as radio type
    VFORadio.Text = RadioVFO
    CommPort.Text = RadioComPort
    CommSpeed.Text = RadioComSpeed
    CommBits.Text = RadioComBits
    CommStop.Text = RadioComStop
    CommParity.Text = RadioComParity
    CommFlow.Text = RadioComFlow

    'Telnet Settings
    TelnetHost.Text = DXtelnethost
    TelnetPort.Text = DXtelnetport
    TelnetUser.Text = TelnetLogon
    TelnetPassword = TelnetPswrd
    AutoLogin.Value = LoginAuto

    'Time Sync
    SetTime.Value = TimeSync

    'Alert sounds
    DXonoff.Text = DXsound
    DXonoff_Change
    DX_watch.Text = DXwatch
    DX_watch_Change
    WWVonoff.Text = WWVsound
    WWVonoff_Change
    ANNonoff.Text = ANNsound
    annWAV.Text = WAVannounce
    ANNonoff_Change

        If SoundOK.Value = 0 Then SoundOK_Click

    'Registration Info, Product ID = CM010-2004-XXXXXX
    ProdIDtext = "CM010-2004-" & Format(App.Major, "00") & Format(App.Minor, "00") & Format(App.Revision, "00")

    'Get UserName from registy
    username = GetSetting("WA1ZEK", "DXtelnet", "UserName", "None")

        If username = "None" Then
            'Get Username from API call

            Dim UserNameSYS            As String * 255

            Call GetUserName(UserNameSYS, 255)
            username = UserNameSYS
        End If

    NameText.Text = Trim(username)

    'Get UserCall from Registry
    UserCall = GetSetting("WA1ZEK", "DXtelnet", "UserCall", "None")
    CallText.Text = UCase(UserCall)

    'Set Telnet Login to Ham Call
    TelnetUser.Text = CallText.Text

    'Get Program registration number from Registry
    RegNum = GetSetting("WA1ZEK", "DXtelnet", "RegNum", "")
    RegNumText.Text = RegNum

    'DX and WWV Exipre times in hours
    DX_Expire.Text = Str(Expire_DX)
    WWV_Expire.Text = Str(Expire_WWV)

        If Use_Local_Clock = True Then
            CPUclock.Value = 1
        Else
            CPUclock.Value = 0
        End If

        If BalloonToolTip = False Then
            BalloonOK.Value = 0
        Else
            BalloonOK.Value = 1
        End If

        If ShowAnnouncements = False Then
            AnnounceOK.Value = 0
        Else
            AnnounceOK.Value = 1
        End If

    'Load DX Watch list into list box from file
    'If no file, exit


    watchlist = FreeFile
    On Error GoTo nofile
    Open App.Path & "\" & "DXwatch.lst" For Input As watchlist

        Do Until EOF(watchlist) = True
            Input #watchlist, dxdata
            DXList.AddItem dxdata
            DoEvents
        Loop
    
    Close watchlist
nofile:

        If DXwatchEnabled = True Then
            WatchEnabled.Value = 1
            DXList.Enabled = True
            DXlistText.Enabled = True
        Else
            WatchEnabled.Value = 0
            DXList.Enabled = False
            DXdelete.Enabled = False
            DXAdd.Enabled = False
            DXlistText.Enabled = False
        End If

        If DXarchiveEnabled = True Then
            DXarchive.Value = 1
        Else
            DXarchive.Value = 0
        End If

    'Load Button Pictures
    FindWAV.Picture = LoadResPicture(102, 0)
    RemoveWave.Picture = LoadResPicture(104, 0)
    testWAV.Picture = LoadResPicture(103, 0)
    'Load Telnet List Grid
    Call modTelnetList.TelnetListWrite
    'Load Band List Grid
    Call modBandList.BandListWrite
    AddTelnet.Enabled = False

        If AutoLogin.Value = 1 Then
            TelnetUser.Enabled = True
            TelnetUser.BackColor = &H80000005
            TelnetUser.Visible = True
            Labeluser.Visible = True
            TelnetUser.Text = TelnetLogon
            TelnetPassword.Enabled = True
            TelnetPassword.BackColor = &H80000005
            TelnetPassword.Visible = True
            LabelPassword.Visible = True
            TelnetPassword.Text = TelnetPswrd
            'Don't change AutoLogin state while connected

                If frmTelnet.WinsockClient.State > 0 Then
                    Passwrd = False
                    Login = False
                Else
                    Passwrd = True
                    Login = True
                End If

        Else
            TelnetUser.Enabled = False
            TelnetUser.BackColor = &H8000000F
            TelnetUser.Visible = False
            Labeluser.Visible = False
            TelnetPassword.Enabled = False
            TelnetPassword.BackColor = &H8000000F
            TelnetPassword.Visible = False
            LabelPassword.Visible = False
            Passwrd = False
            Login = False

        End If

End Sub


Private Sub SoundOKLabel_Click()

        If SoundOK.Value = 1 Then
            SoundOK.Value = 0
        Else
            SoundOK.Value = 1
        End If

End Sub

Private Sub NameText_Click()

    NameText.SelStart = 0
    NameText.SelLength = Len(NameText.Text) + 1

End Sub

Private Sub NameText_DblClick()

    NameText.SelStart = Len(NameText.Text) + 1

End Sub

Private Sub okWAV_Click()

        Select Case WAVselect
    
            Case 0
                spotWAV.Text = WAVfile.Text
            Case 1
                watchWAV.Text = WAVfile.Text
            Case 2
                wwvWAV.Text = WAVfile.Text
            Case 3
                annWAV.Text = WAVfile.Text
        End Select

    'SoundFileFrame.Visible = False
    Call DisableWaveFind

End Sub

Private Sub PEip_Click()

    PEip.SelStart = 0
    PEip.SelLength = Len(PEip.Text) + 1

End Sub

Private Sub PEip_DblClick()

    PEip.SelStart = Len(PEip.Text) + 1

End Sub

Private Sub PEport_Click()

    PEport.SelStart = 0
    PEport.SelLength = Len(PEport.Text) + 1

End Sub

Private Sub PEport_DblClick()

    PEport.SelStart = Len(PEport.Text) + 1

End Sub

Private Sub PropertiesTab_GotFocus()

        If PropertiesTab.Tab = 0 Then
            StartKHz.SetFocus
        End If

End Sub

Private Sub PropertiesTab_LostFocus()

    AddTelnet.Enabled = False

End Sub

Private Sub RadioType_Click()

    'If no Radio (None) selected diable the rest of the radio selections

        If RadioType.Text = "None" Then
            VFORadio.Enabled = False
            CommPort.Enabled = False
            CommSpeed.Enabled = False
            CommBits.Enabled = False
            CommStop.Enabled = False
            CommParity.Enabled = False
            CommFlow.Enabled = False
            'Disable Band-Mode Tab
            StartKHz.Enabled = False
            StartKHz.BackColor = &H8000000F
            EndKHz.Enabled = False
            EndKHz.BackColor = &H8000000F
            ModeType.Enabled = False
            ModeType.BackColor = &H8000000F
            AddRange.Enabled = False
            DeleteRange.Enabled = False
            BandGrid.Enabled = False
            BandGrid.BackColor = &H8000000F

        Else
            VFORadio.Enabled = True
            CommPort.Enabled = True
            CommSpeed.Enabled = True
            CommBits.Enabled = True
            CommStop.Enabled = True
            CommParity.Enabled = True
            CommFlow.Enabled = True
            'Enable Band-Mode Tab
            StartKHz.Enabled = True
            StartKHz.BackColor = &H80000005
            EndKHz.Enabled = True
            EndKHz.BackColor = &H80000005
            ModeType.Enabled = True
            ModeType.BackColor = &H80000005
            AddRange.Enabled = False
            DeleteRange.Enabled = False
            BandGrid.Enabled = True
            BandGrid.BackColor = &H80000005

        End If

    'Change ToolTip and Mouse Pointer for DXwindow Radio control feature

        If RadioType.Text = "None" Then
            'If no radio selected, turn off ToolTip t set radio frequency
            DXwindow.DXFrame.ToolTipText = ""
            DXwindow.DXSpotLabel.ToolTipText = ""
            DXwindow.DXspotText.ToolTipText = ""
            DXwindow.DXGrid.ToolTipText = ""
            'Set mouse pointer to default
            DXwindow.DXFrame.MousePointer = 0
            DXwindow.DXSpotLabel.MousePointer = 0
            DXwindow.DXspotText.MousePointer = 0
            DXwindow.DXGrid.MousePointer = 0
        Else
            'If a rado is selected, turn on ToolTip for frequency control
            DXwindow.DXFrame.ToolTipText = "Click to tune radio to DX frequency"
            DXwindow.DXSpotLabel.ToolTipText = "Click to tune radio to DX frequency"
            DXwindow.DXspotText.ToolTipText = "Click to tune radio to DX frequency"
            DXwindow.DXGrid.ToolTipText = "Click on DX spot to tune radio to DX frequency"
            'Set mouse pointer to Icon
            DXwindow.DXFrame.MousePointer = 99
            DXwindow.DXSpotLabel.MousePointer = 99
            DXwindow.DXspotText.MousePointer = 99
            DXwindow.DXGrid.MousePointer = 99
        End If

    
End Sub

Private Sub RegNumText_Click()

    RegNumText.SelStart = 0
    RegNumText.SelLength = Len(RegNumText.Text) + 1

End Sub

Private Sub RegNumText_DblClick()

    RegNumText.SelStart = Len(RegNumText.Text) + 1

End Sub

Private Sub RemoveTelnet_Click()

    'Remove Host and Port fron list

    'Test to see if there is any text on row
    TelnetGrid.Col = 0
    HostText = TelnetGrid.Text

        If HostText = "" Then Exit Sub
    TelnetGrid.RemoveItem (TelnetGrid.Row)
    RemoveTelnet.Enabled = False

End Sub

Private Sub RemoveWave_Click()

    WAVfile.Text = ""

End Sub



Private Sub SoundOK_Click()

        If SoundOK.Value = 1 Then
            spotWAV.Enabled = True
            spotWAV.BackColor = &H80000005
            watchWAV.Enabled = True
            watchWAV.BackColor = &H80000005
            wwvWAV.Enabled = True
            wwvWAV.BackColor = &H80000005
            annWAV.Enabled = True
            annWAV.BackColor = &H80000005
    
        Else
            SoundFileFrame.Visible = False
            spotWAV.Enabled = False
            spotWAV.BackColor = &H8000000F
            watchWAV.Enabled = False
            watchWAV.BackColor = &H8000000F
            wwvWAV.Enabled = False
            wwvWAV.BackColor = &H8000000F
            annWAV.Enabled = False
            annWAV.BackColor = &H8000000F
            Call DisableWaveFind
        End If

End Sub


Private Sub spotWAV_Click()

    'IF Select wave open, save current file and then open new selection
    'If SoundFileFrame.Visible = True Then okWAV_Click
    'Open wave select for new selection
    SoundFileFrame.Caption = "Select DX Spot Sound"
    WAVselect = 0
    Call EnableWaveFind
    WAVfile.Text = spotWAV.Text
    
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub StartKHz_Change()

        If StartKHz.Text = "" Then
            AddRange.Enabled = False
        Else
            AddRange.Enabled = True
        End If

End Sub

Private Sub TelnetGrid_Click()

    RemoveTelnet.Enabled = True

End Sub

Private Sub TelnetGrid_DblClick()

    'Set double clicked row info as new host
    RemoveTelnet.Enabled = False

    'Test to see if ther is any text in row
    TelnetGrid.Col = 0
    HostText = TelnetGrid.Text

        If HostText = "" Then Exit Sub
    TelnetGrid.Col = 0
    TelnetHost.Text = TelnetGrid.Text
    TelnetGrid.Col = 1
    TelnetPort.Text = TelnetGrid.Text
    TelnetGrid.Col = 2
    TelnetUser.Text = TelnetGrid.Text
    TelnetGrid.Col = 3
    TelnetPassword.Text = TelnetGrid.Text
       
    AddTelnet.Enabled = False

End Sub

Private Sub TelnetGrid_GotFocus()

    AddTelnet.Enabled = False

End Sub

Private Sub TelnetHost_Change()

        If TelnetHost.Text = "" Then Exit Sub
    AddTelnet.Enabled = True

End Sub

Private Sub TelnetHost_GotFocus()

    RemoveTelnet.Enabled = False

End Sub

Private Sub TelnetPassword_Change()

    AddTelnet.Enabled = True

End Sub

Private Sub TelnetPassword_GotFocus()

    RemoveTelnet.Enabled = False

End Sub

Private Sub TelnetPort_Change()

        If TelnetHost.Text = "" Then Exit Sub
    AddTelnet.Enabled = True

End Sub

Private Sub TelnetPort_GotFocus()

    RemoveTelnet.Enabled = False

End Sub

Private Sub TelnetUser_Change()

    TelnetUser.Text = UCase(TelnetUser.Text)
    TelnetUser.SelStart = Len(TelnetUser.Text) + 1
    AddTelnet.Enabled = True

End Sub

Private Sub TelnetUser_GotFocus()

    RemoveTelnet.Enabled = False

End Sub

Private Sub testWAV_Click()

    'Check for "NULL" file name

        If WAVfile.Text = "" Then Exit Sub
    'Play the WAVE file
    PlaySound (WAVfile.Text)

End Sub


Private Sub WatchEnabled_Click()

        If WatchEnabled = 1 Then
            DXwatchEnabled = True
            DXList.Enabled = True
            DXlistText.Enabled = True
        Else
            DXwatchEnabled = False
            DXList.Enabled = False
            DXdelete.Enabled = False
            DXAdd.Enabled = False
            DXlistText.Enabled = False
        End If

End Sub

Private Sub WatchEnabled_GotFocus()

    DXdelete.Enabled = False

End Sub

Private Sub watchWAV_Click()

    'IF Select wave open, save current file and then open new selection
    'If SoundFileFrame.Visible = True Then okWAV_Click
    'Open wave select for new selection
    SoundFileFrame.Caption = "Select DX Watch Sound"
    WAVselect = 1
    Call EnableWaveFind
    WAVfile.Text = watchWAV.Text

End Sub

Private Sub WWV_Expire_GotFocus()

    DXdelete.Enabled = False

End Sub

Private Sub WWVonoff_Change()

        If SoundOK.Value = 0 Then Exit Sub

        Select Case WWVonoff
            Case "ON"
                wwvWAV.BackColor = &H80000005
                wwvWAV.Enabled = True
            Case "OFF"
                wwvWAV.BackColor = &H80000004
                wwvWAV.Enabled = False
        End Select

End Sub

Private Sub WWVonoff_Click()

    WWVonoff_Change

End Sub

Private Sub wwvWAV_Click()

    'IF Select wave open, save current file and then open new selection
    'Open wave select for new selection
    SoundFileFrame.Caption = "Select WWV Alert Sound"
    WAVselect = 2
    WAVfile.Text = wwvWAV.Text
    Call EnableWaveFind
    
End Sub

Private Sub DisableWaveFind()

    WAVfile.Enabled = False
    WAVfile.BackColor = &H8000000F
    FindWAV.Enabled = False
    FindWAV.BackColor = &H8000000F
    RemoveWave.Enabled = False
    RemoveWave.BackColor = &H8000000F
    testWAV.Enabled = False
    testWAV.BackColor = &H8000000F
    okWAV.Enabled = False
    okWAV.BackColor = &H8000000F
    CancelWAV.Enabled = False
    CancelWAV.BackColor = &H8000000F
    SoundFileFrame.Caption = ""
    WAVfile.Text = ""
    SoundFileFrame.Enabled = False
    SoundFileFrame.Visible = True

End Sub

Private Sub EnableWaveFind()

    WAVfile.Enabled = True
    WAVfile.BackColor = &H80000005
    FindWAV.Enabled = True
    FindWAV.BackColor = &HC0C0C0
    RemoveWave.Enabled = True
    RemoveWave.BackColor = &HC0C0C0
    testWAV.Enabled = True
    testWAV.BackColor = &HC0C0C0
    okWAV.Enabled = True
    okWAV.BackColor = &HC0C0C0
    CancelWAV.Enabled = True
    CancelWAV.BackColor = &HC0C0C0
    SoundFileFrame.Enabled = True
    SoundFileFrame.Visible = True

End Sub
