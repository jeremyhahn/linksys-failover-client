VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Linksys Failover Client"
   ClientHeight    =   6405
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10365
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   30
      TabIndex        =   0
      Top             =   480
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   10398
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "UPnP Primary Configuration"
      TabPicture(0)   =   "frmMain.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "PrimaryApply"
      Tab(0).Control(2)=   "PrimaryCancel"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(4)=   "PriDNSport"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "UPnP Failover Configuration"
      TabPicture(1)   =   "frmMain.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label36"
      Tab(1).Control(1)=   "Label40"
      Tab(1).Control(2)=   "Label41"
      Tab(1).Control(3)=   "Label35"
      Tab(1).Control(4)=   "Label28"
      Tab(1).Control(5)=   "Label27"
      Tab(1).Control(6)=   "Label26"
      Tab(1).Control(7)=   "Label25"
      Tab(1).Control(8)=   "Label24"
      Tab(1).Control(9)=   "secondaryCancel"
      Tab(1).Control(10)=   "SecondaryApply"
      Tab(1).Control(11)=   "Frame2"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Activity Log"
      TabPicture(2)   =   "frmMain.frx":0902
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "txtLogDisplay"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.TextBox PriDNSport 
         Height          =   285
         Left            =   -69870
         TabIndex        =   18
         Text            =   "53"
         Top             =   2970
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Failover Configuration"
         Height          =   3615
         Left            =   -74760
         TabIndex        =   67
         Top             =   1440
         Width           =   9855
         Begin VB.TextBox secHTTPport 
            Height          =   285
            Left            =   4890
            TabIndex        =   100
            Text            =   "80"
            Top             =   1095
            Width           =   615
         End
         Begin VB.CheckBox secPOP3enabled 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9165
            TabIndex        =   14
            Top             =   2955
            Width           =   255
         End
         Begin VB.TextBox SecPOP3ip 
            Height          =   285
            Left            =   7395
            TabIndex        =   13
            Top             =   2865
            Width           =   570
         End
         Begin VB.TextBox secPOP3port 
            Height          =   285
            Left            =   4905
            TabIndex        =   12
            Text            =   "110"
            Top             =   2865
            Width           =   615
         End
         Begin VB.CheckBox secSMTPenabled 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9165
            TabIndex        =   11
            Top             =   2475
            Width           =   255
         End
         Begin VB.TextBox SecSMTPip 
            Height          =   285
            Left            =   7395
            TabIndex        =   10
            Top             =   2415
            Width           =   570
         End
         Begin VB.TextBox secSMTPport 
            Height          =   285
            Left            =   4905
            TabIndex        =   9
            Text            =   "25"
            Top             =   2415
            Width           =   615
         End
         Begin VB.CheckBox secFTPenabled 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9165
            TabIndex        =   8
            Top             =   2025
            Width           =   255
         End
         Begin VB.TextBox SecFTPip 
            Height          =   285
            Left            =   7395
            TabIndex        =   7
            Top             =   1965
            Width           =   570
         End
         Begin VB.TextBox secFTPport 
            Height          =   285
            Left            =   4905
            TabIndex        =   6
            Text            =   "21"
            Top             =   1965
            Width           =   615
         End
         Begin VB.CheckBox secDNSenabled 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9165
            TabIndex        =   5
            Top             =   1575
            Width           =   255
         End
         Begin VB.TextBox SecDNSip 
            Height          =   285
            Left            =   7395
            TabIndex        =   4
            Top             =   1515
            Width           =   570
         End
         Begin VB.TextBox secDNSport 
            Height          =   285
            Left            =   4890
            TabIndex        =   3
            Text            =   "53"
            Top             =   1530
            Width           =   615
         End
         Begin VB.CheckBox secHTTPenabled 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9165
            TabIndex        =   2
            Top             =   1155
            Width           =   255
         End
         Begin VB.TextBox SecHTTPip 
            Height          =   285
            Left            =   7395
            TabIndex        =   1
            Top             =   1095
            Width           =   570
         End
         Begin VB.Label Label11 
            Caption         =   "TCP"
            Height          =   255
            Index           =   11
            Left            =   3465
            TabIndex        =   93
            Top             =   2955
            Width           =   375
         End
         Begin VB.Label Label55 
            Caption         =   "110"
            Height          =   195
            Left            =   1935
            TabIndex        =   92
            Top             =   2985
            Width           =   405
         End
         Begin VB.Label Label11 
            Caption         =   "TCP"
            Height          =   255
            Index           =   10
            Left            =   3465
            TabIndex        =   91
            Top             =   2505
            Width           =   375
         End
         Begin VB.Label Label54 
            Caption         =   "25"
            Height          =   255
            Left            =   1935
            TabIndex        =   90
            Top             =   2505
            Width           =   285
         End
         Begin VB.Label Label11 
            Caption         =   "TCP"
            Height          =   255
            Index           =   9
            Left            =   3465
            TabIndex        =   89
            Top             =   2055
            Width           =   375
         End
         Begin VB.Label Label53 
            Caption         =   "21"
            Height          =   255
            Left            =   1935
            TabIndex        =   88
            Top             =   2055
            Width           =   255
         End
         Begin VB.Label lblLANIP 
            Caption         =   "%LANIP%"
            Height          =   285
            Index           =   14
            Left            =   6615
            TabIndex        =   87
            Top             =   2925
            Width           =   765
         End
         Begin VB.Label lblLANIP 
            Caption         =   "%LANIP%"
            Height          =   285
            Index           =   13
            Left            =   6615
            TabIndex        =   86
            Top             =   2445
            Width           =   765
         End
         Begin VB.Label lblLANIP 
            Caption         =   "%LANIP%"
            Height          =   285
            Index           =   12
            Left            =   6615
            TabIndex        =   85
            Top             =   1995
            Width           =   765
         End
         Begin VB.Label lblLANIP 
            Caption         =   "%LANIP%"
            Height          =   285
            Index           =   11
            Left            =   6615
            TabIndex        =   84
            Top             =   1545
            Width           =   765
         End
         Begin VB.Label lblLANIP 
            Caption         =   "%LANIP%"
            Height          =   285
            Index           =   10
            Left            =   6615
            TabIndex        =   83
            Top             =   1125
            Width           =   765
         End
         Begin VB.Label Label52 
            Caption         =   "53"
            Height          =   240
            Left            =   1935
            TabIndex        =   82
            Top             =   1590
            Width           =   375
         End
         Begin VB.Label Label51 
            Caption         =   "Enable"
            Height          =   240
            Left            =   9045
            TabIndex        =   81
            Top             =   600
            Width           =   600
         End
         Begin VB.Label Label50 
            Caption         =   "IP Address"
            Height          =   285
            Left            =   6885
            TabIndex        =   80
            Top             =   600
            Width           =   1050
         End
         Begin VB.Label Label49 
            Caption         =   "Internal Port"
            Height          =   285
            Left            =   4770
            TabIndex        =   79
            Top             =   600
            Width           =   1050
         End
         Begin VB.Label Label12 
            Caption         =   "UDP"
            Height          =   285
            Index           =   2
            Left            =   3465
            TabIndex        =   78
            Top             =   1575
            Width           =   465
         End
         Begin VB.Label Label11 
            Caption         =   "TCP"
            Height          =   255
            Index           =   8
            Left            =   3465
            TabIndex        =   77
            Top             =   1125
            Width           =   375
         End
         Begin VB.Label Label48 
            Caption         =   "Protocol"
            Height          =   255
            Left            =   3375
            TabIndex        =   76
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label47 
            Caption         =   "Ext. Port"
            Height          =   255
            Left            =   1815
            TabIndex        =   75
            Top             =   615
            Width           =   735
         End
         Begin VB.Label Label46 
            Caption         =   "80"
            Height          =   255
            Left            =   1935
            TabIndex        =   74
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label45 
            Caption         =   "POP3"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   73
            Top             =   3015
            Width           =   735
         End
         Begin VB.Label Label44 
            Caption         =   "SMTP"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   72
            Top             =   2535
            Width           =   615
         End
         Begin VB.Label Label43 
            Caption         =   "FTP"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   71
            Top             =   2055
            Width           =   495
         End
         Begin VB.Label Label42 
            Caption         =   "DNS"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   70
            Top             =   1575
            Width           =   495
         End
         Begin VB.Label Label39 
            Caption         =   "HTTP"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   69
            Top             =   1095
            Width           =   735
         End
         Begin VB.Label Label38 
            Caption         =   "Service"
            Height          =   255
            Left            =   375
            TabIndex        =   68
            Top             =   615
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Primary Router Configuration"
         Height          =   3615
         Left            =   -74760
         TabIndex        =   40
         Top             =   1440
         Width           =   9855
         Begin VB.TextBox PriHTTPport 
            Height          =   285
            Left            =   4890
            TabIndex        =   15
            Text            =   "80"
            Top             =   1095
            Width           =   615
         End
         Begin VB.TextBox PriHTTPip 
            Height          =   285
            Left            =   7395
            TabIndex        =   16
            Top             =   1095
            Width           =   570
         End
         Begin VB.CheckBox PriHTTPenabled 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9165
            TabIndex        =   17
            Top             =   1155
            Width           =   255
         End
         Begin VB.TextBox PriDNSip 
            Height          =   285
            Left            =   7395
            TabIndex        =   19
            Top             =   1515
            Width           =   570
         End
         Begin VB.CheckBox PriDNSenabled 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9165
            TabIndex        =   20
            Top             =   1575
            Width           =   255
         End
         Begin VB.TextBox PriFTPport 
            Height          =   285
            Left            =   4905
            TabIndex        =   21
            Text            =   "21"
            Top             =   1965
            Width           =   615
         End
         Begin VB.TextBox PriFTPip 
            Height          =   285
            Left            =   7395
            TabIndex        =   22
            Top             =   1965
            Width           =   570
         End
         Begin VB.CheckBox PriFTPenabled 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9165
            TabIndex        =   23
            Top             =   2025
            Width           =   255
         End
         Begin VB.TextBox PriSMTPport 
            Height          =   285
            Left            =   4905
            TabIndex        =   24
            Text            =   "25"
            Top             =   2415
            Width           =   615
         End
         Begin VB.TextBox PriSMTPip 
            Height          =   285
            Left            =   7395
            TabIndex        =   25
            Top             =   2415
            Width           =   570
         End
         Begin VB.CheckBox PriSMTPenabled 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9165
            TabIndex        =   26
            Top             =   2475
            Width           =   255
         End
         Begin VB.TextBox PriPOP3port 
            Height          =   285
            Left            =   4905
            TabIndex        =   27
            Text            =   "110"
            Top             =   2865
            Width           =   615
         End
         Begin VB.TextBox PriPOP3ip 
            Height          =   285
            Left            =   7395
            TabIndex        =   28
            Top             =   2865
            Width           =   570
         End
         Begin VB.CheckBox PriPOP3enabled 
            Caption         =   "Check1"
            Height          =   255
            Left            =   9165
            TabIndex        =   29
            Top             =   2955
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Service"
            Height          =   255
            Left            =   375
            TabIndex        =   66
            Top             =   615
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "HTTP"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   65
            Top             =   1095
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "DNS"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   64
            Top             =   1575
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "FTP"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   63
            Top             =   2055
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "SMTP"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   62
            Top             =   2535
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "POP3"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   61
            Top             =   3015
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "80"
            Height          =   255
            Left            =   1935
            TabIndex        =   60
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Ext. Port"
            Height          =   255
            Left            =   1815
            TabIndex        =   59
            Top             =   615
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Protocol"
            Height          =   255
            Left            =   3375
            TabIndex        =   58
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "TCP"
            Height          =   255
            Index           =   0
            Left            =   3465
            TabIndex        =   57
            Top             =   1125
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "UDP"
            Height          =   285
            Index           =   0
            Left            =   3465
            TabIndex        =   56
            Top             =   1575
            Width           =   465
         End
         Begin VB.Label Label13 
            Caption         =   "Internal Port"
            Height          =   285
            Left            =   4770
            TabIndex        =   55
            Top             =   600
            Width           =   1050
         End
         Begin VB.Label Label14 
            Caption         =   "IP Address"
            Height          =   285
            Left            =   6885
            TabIndex        =   54
            Top             =   600
            Width           =   1050
         End
         Begin VB.Label Label15 
            Caption         =   "Enable"
            Height          =   240
            Left            =   9045
            TabIndex        =   53
            Top             =   600
            Width           =   600
         End
         Begin VB.Label Label16 
            Caption         =   "53"
            Height          =   240
            Left            =   1935
            TabIndex        =   52
            Top             =   1590
            Width           =   375
         End
         Begin VB.Label lblLANIP 
            Caption         =   "%LANIP%"
            Height          =   285
            Index           =   0
            Left            =   6615
            TabIndex        =   51
            Top             =   1125
            Width           =   765
         End
         Begin VB.Label lblLANIP 
            Caption         =   "%LANIP%"
            Height          =   285
            Index           =   1
            Left            =   6615
            TabIndex        =   50
            Top             =   1545
            Width           =   765
         End
         Begin VB.Label lblLANIP 
            Caption         =   "%LANIP%"
            Height          =   285
            Index           =   2
            Left            =   6615
            TabIndex        =   49
            Top             =   1995
            Width           =   765
         End
         Begin VB.Label lblLANIP 
            Caption         =   "%LANIP%"
            Height          =   285
            Index           =   3
            Left            =   6615
            TabIndex        =   48
            Top             =   2445
            Width           =   765
         End
         Begin VB.Label lblLANIP 
            Caption         =   "%LANIP%"
            Height          =   285
            Index           =   4
            Left            =   6615
            TabIndex        =   47
            Top             =   2925
            Width           =   765
         End
         Begin VB.Label Label17 
            Caption         =   "21"
            Height          =   255
            Left            =   1935
            TabIndex        =   46
            Top             =   2055
            Width           =   255
         End
         Begin VB.Label Label11 
            Caption         =   "TCP"
            Height          =   255
            Index           =   1
            Left            =   3465
            TabIndex        =   45
            Top             =   2055
            Width           =   375
         End
         Begin VB.Label Label18 
            Caption         =   "25"
            Height          =   255
            Left            =   1935
            TabIndex        =   44
            Top             =   2505
            Width           =   285
         End
         Begin VB.Label Label11 
            Caption         =   "TCP"
            Height          =   255
            Index           =   2
            Left            =   3465
            TabIndex        =   43
            Top             =   2505
            Width           =   375
         End
         Begin VB.Label Label19 
            Caption         =   "110"
            Height          =   195
            Left            =   1935
            TabIndex        =   42
            Top             =   2985
            Width           =   405
         End
         Begin VB.Label Label11 
            Caption         =   "TCP"
            Height          =   255
            Index           =   3
            Left            =   3465
            TabIndex        =   41
            Top             =   2955
            Width           =   375
         End
      End
      Begin VB.ListBox txtLogDisplay 
         Height          =   5325
         Left            =   90
         TabIndex        =   39
         Top             =   420
         Width           =   10125
      End
      Begin VB.CommandButton SecondaryApply 
         Caption         =   "Apply"
         Height          =   345
         Left            =   -67320
         TabIndex        =   36
         Top             =   5340
         Width           =   1215
      End
      Begin VB.CommandButton secondaryCancel 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   -66090
         TabIndex        =   35
         Top             =   5340
         Width           =   1155
      End
      Begin VB.CommandButton PrimaryCancel 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   -66090
         TabIndex        =   32
         Top             =   5340
         Width           =   1155
      End
      Begin VB.CommandButton PrimaryApply 
         Caption         =   "Apply"
         Height          =   345
         Left            =   -67320
         TabIndex        =   31
         Top             =   5340
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Enable"
         Height          =   240
         Left            =   -65940
         TabIndex        =   99
         Top             =   4590
         Width           =   600
      End
      Begin VB.Label Label25 
         Caption         =   "IP Address"
         Height          =   285
         Left            =   -68100
         TabIndex        =   98
         Top             =   4560
         Width           =   1050
      End
      Begin VB.Label Label26 
         Caption         =   "Internal Port"
         Height          =   285
         Left            =   -70230
         TabIndex        =   97
         Top             =   4575
         Width           =   1050
      End
      Begin VB.Label Label27 
         Caption         =   "Protocol"
         Height          =   255
         Left            =   -71625
         TabIndex        =   96
         Top             =   4590
         Width           =   615
      End
      Begin VB.Label Label28 
         Caption         =   "Ext. Port"
         Height          =   255
         Left            =   -73185
         TabIndex        =   95
         Top             =   4590
         Width           =   735
      End
      Begin VB.Label Label35 
         Caption         =   "Service"
         Height          =   255
         Left            =   -74610
         TabIndex        =   94
         Top             =   4590
         Width           =   1215
      End
      Begin VB.Label Label41 
         Caption         =   "(c) 2004 Pc Technics"
         Height          =   225
         Left            =   -74640
         TabIndex        =   38
         Top             =   7350
         Width           =   1905
      End
      Begin VB.Label Label40 
         Caption         =   "Freeware version 1.0"
         Height          =   225
         Left            =   -74640
         TabIndex        =   37
         Top             =   7140
         Width           =   1575
      End
      Begin VB.Label Label36 
         Caption         =   $"frmMain.frx":091E
         Height          =   495
         Left            =   -74280
         TabIndex        =   33
         Top             =   750
         Width           =   9135
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMain.frx":09D8
         Height          =   495
         Left            =   -74280
         TabIndex        =   30
         Top             =   750
         Width           =   9135
      End
   End
   Begin VB.Label Label37 
      BackColor       =   &H00FF0000&
      Caption         =   "Linksys Failover Client"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   90
      TabIndex        =   34
      Top             =   90
      Width           =   3915
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuShutdown 
         Caption         =   "Shutdown"
      End
      Begin VB.Menu menuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuLogging 
      Caption         =   "Logging"
      Begin VB.Menu menuRefreshLog 
         Caption         =   "Refresh Log"
      End
      Begin VB.Menu menuClearLog 
         Caption         =   "Clear Log"
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu menuEnableLogging 
         Caption         =   "Enable Logging"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu menuTools 
      Caption         =   "Tools"
      Begin VB.Menu menuUpdateNow 
         Caption         =   "Update Now"
      End
      Begin VB.Menu menuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu menuRouterPassword 
         Caption         =   "Router Password"
      End
      Begin VB.Menu menuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuOptinos 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Help"
      Begin VB.Menu menuCheck4Upgrade 
         Caption         =   "Check For Upgrade"
      End
      Begin VB.Menu menuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu menuAboutUs 
         Caption         =   "About"
      End
      Begin VB.Menu menuAbout 
         Caption         =   "Copyright"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()


Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = (Screen.Height / 2) - (Me.Height / 2)


POPULATE_LISTVIEW
 
 
 
    'Get Out Logging Option Preference From DB File
    VERIFY_SYSCONFIG
    Open App.Path & "/sysconfig.dat" For Input As #1
  
         Input #1, AdminPassword
         Input #1, LoggingPref
       
     Close #1
     
     If LoggingPref = "1" Then
       frmMain.menuEnableLogging.Checked = True
   Else
       frmMain.menuEnableLogging.Checked = False
 End If
     







    Open App.Path & "\IP.dat" For Input As #1

        Input #1, dbRouterIP, UpdateInterval

    Close #1
    
   
   
   
   
   

LANaddress = Split(dbRouterIP, ".")

octet1 = LANaddress(0)
octet2 = LANaddress(1)
octet3 = LANaddress(2)

NetID = octet1 & "." & octet2 & "." & octet3 & "."


lblLANIP(0).Caption = NetID
lblLANIP(1).Caption = NetID
lblLANIP(2).Caption = NetID
lblLANIP(3).Caption = NetID
lblLANIP(4).Caption = NetID
lblLANIP(10).Caption = NetID
lblLANIP(11).Caption = NetID
lblLANIP(12).Caption = NetID
lblLANIP(13).Caption = NetID
lblLANIP(14).Caption = NetID







    Open App.Path & "\PrimaryConfigs.dat" For Input As #2

      Input #2, dbPriHTTPport, dbPriHTTPip, dbPriHTTPenabled
      Input #2, dbPriDNSport, dbPriDNSip, dbPriDNSenabled
      Input #2, dbPriFTPport, dbPriFTPip, dbPriFTPenabled
      Input #2, dbPriSMTPport, dbPriSMTPip, dbPriSMTPenabled
      Input #2, dbPriPOP3port, dbPriPOP3ip, dbPriPOP3enabled


    Close #2



      PriHTTPport = dbPriHTTPport
      PriHTTPip = dbPriHTTPip
      PriHTTPenabled = dbPriHTTPenabled
      
      
      PriDNSport = dbPriDNSport
      PriDNSip = dbPriDNSip
      PriDNSenabled = dbPriDNSenabled
      
      
      PriFTPport = dbPriFTPport
      PriFTPip = dbPriFTPip
      PriFTPenabled = dbPriFTPenabled
      
      
      
      
      PriSMTPport = dbPriSMTPport
      PriSMTPip = dbPriSMTPip
      PriSMTPenabled = dbPriSMTPenabled
      
      
      
      PriPOP3port = dbPriPOP3port
      PriPOP3ip = dbPriPOP3ip
      PriPOP3enabled = dbPriPOP3enabled
      
      
   ' ----------------------------------------------------------------------
   '                          Failover Configs
   
   
   
   Open App.Path & "\FailoverConfigs.dat" For Input As #3

      Input #3, dbSecHTTPport, dbSecHTTPip, dbSecHTTPenabled
      Input #3, dbSecDNSport, dbSecDNSip, dbSecDNSenabled
      Input #3, dbSecFTPport, dbSecFTPip, dbSecFTPenabled
      Input #3, dbSecSMTPport, dbSecSMTPip, dbSecSMTPenabled
      Input #3, dbSecPOP3port, dbSecPOP3ip, dbSecPOP3enabled


   Close #3



      secHTTPport = dbSecHTTPport
      SecHTTPip = dbSecHTTPip
      secHTTPenabled = dbSecHTTPenabled
      
      
      secDNSport = dbSecDNSport
      SecDNSip = dbSecDNSip
      secDNSenabled = dbSecDNSenabled
      
      
      secFTPport = dbSecFTPport
      SecFTPip = dbSecFTPip
      secFTPenabled = dbSecFTPenabled
      
      
      
      
      secSMTPport = dbSecSMTPport
      SecSMTPip = dbSecSMTPip
      secSMTPenabled = dbSecSMTPenabled
      
      
      
      secPOP3port = dbSecPOP3port
      SecPOP3ip = dbSecPOP3ip
      secPOP3enabled = dbSecPOP3enabled
      
      
      
      
      PrimaryApply.Enabled = False
      SecondaryApply.Enabled = False


End Sub




Private Sub menuAbout_Click()
Load frmAbout
frmAbout.Show
End Sub



Private Sub menuAboutUs_Click()
Load frmAboutAuthor
frmAboutAuthor.Show
End Sub






Private Sub menuCheck4Upgrade_Click()

 Set XML_HTTP2 = CreateObject("MSXML2.ServerXMLHTTP.3.0")


 XML_HTTP2.Open "GET", "http://www.pc-technics.com/API/software_updates.php?IDENTITY=PT_SOFTWARE_BOT&Software=LFC"
 
 XML_HTTP2.Send ""

      
    If XML_HTTP2.Status <> 200 Then MsgBox "Web Server Error : " & XML_HTTP.Status, vbCritical, "Linksys Failover Client"

HTTP_RESPONSE = XML_HTTP2.ResponseText

  UpdateArray = Split(HTTP_RESPONSE, ",")
  
  DoUpdate = UpdateArray(0)
  DownloadURL = UpdateArray(1)


       If DoUpdate = "1" Then
 
           ConfirmDownload = MsgBox("There is a new update available. Would you like to download the new version?", vbYesNo, "Linksys Failover Client")
 
        If ConfirmDownload = vbYes Then
   
            Dim IE As Object
            Set IE = CreateObject("InternetExplorer.Application")
                IE.Navigate DownloadURL
            
            DoEvents
                IE.Visible = True
  
            Set IE = Nothing
    End If
    
    
    Else
    
    MsgBox "There are currently no upgrades available. Please check back later.", vbInformation, "Linksys Failover Client"
       
  
End If
  
  
  
  
  Set XML_HTTP2 = Nothing


End Sub







Private Sub menuClearLog_Click()
frmMain.txtLogDisplay.Clear


Open App.Path & "/LOG.txt" For Output As #1

   Write #1, "Log Cleared - " & Now
   
Close #1



 
POPULATE_LISTVIEW



End Sub




Private Sub menuEnableLogging_Click()


If frmMain.menuEnableLogging.Checked = True Then



    VERIFY_SYSCONFIG
    Open App.Path & "/sysconfig.dat" For Input As #1
  
         Input #1, AdminPassword
         Input #1, LoggingPref
       
     Close #1


    Open App.Path & "/sysconfig.dat" For Output As #1
 
      Write #1, AdminPassword
       Write #1, "0"
    
      Close #1
   
   frmMain.menuEnableLogging.Checked = False
  
Else
  
     VERIFY_SYSCONFIG
     Open App.Path & "/sysconfig.dat" For Input As #1
  
         Input #1, AdminPassword
         Input #1, LoggingPref
       
    Close #1


     Open App.Path & "/sysconfig.dat" For Output As #1
 
         Write #1, AdminPassword
         Write #1, "1"
    
    Close #1

   frmMain.menuEnableLogging.Checked = True

End If


End Sub

Private Sub menuExit_Click()

    If PrimaryApply.Enabled = True Then

       MsgBox "You have unsaved settings on the Primary Configuration tab. Please apply your settings before exiting.", vbInformation, "Linksys Failover Client"
        
       Exit Sub
End If


 




    If SecondaryApply.Enabled = True Then

       MsgBox "You have unsaved settings on the Failover Configuration tab. Please apply your settings before exiting.", vbYesNo, "Linksys Failover Client"

       Exit Sub
    
End If
 

 

             
             
             
Me.Hide
End Sub



Private Sub menuOptinos_Click()
Load frmOptions
frmOptions.Show
End Sub






Private Sub menuRefreshLog_Click()
frmMain.txtLogDisplay.Clear
POPULATE_LISTVIEW
End Sub

Private Sub menuRouterPassword_Click()
Load frmRouterPassword
frmRouterPassword.Show
End Sub

Private Sub menuShutdown_Click()
Unload frmService
Unload Me
End Sub



Private Sub menuUpdateNow_Click()
DO_GLOBAL_POLL

UpdateControls
End Sub

Private Sub PriDNSenabled_Click()
PrimaryApply.Enabled = True
End Sub

Private Sub PriDNSip_Change()
PrimaryApply.Enabled = True
End Sub

Private Sub PriDNSport_Change()
PrimaryApply.Enabled = True
End Sub

Private Sub PriFTPenabled_Click()
PrimaryApply.Enabled = True
End Sub

Private Sub PriFTPip_Change()
PrimaryApply.Enabled = True
End Sub

Private Sub PriFTPport_Change()
PrimaryApply.Enabled = True
End Sub

Private Sub PriHTTPenabled_Click()
PrimaryApply.Enabled = True
End Sub

Private Sub PriHTTPip_Change()
PrimaryApply.Enabled = True
End Sub









Private Sub PriHTTPport_Change()
PrimaryApply.Enabled = True
End Sub







Private Sub PrimaryApply_Click()
 
PrimaryApply.Enabled = False



If Not VerifyPort(PriHTTPport) Then
MsgBox "You have entered an invalid port number for the HTTP protocol. Valid port numbers are 1 - 65535.", vbCritical, "Linksys Failover Client"
PriHTTPport.SetFocus
PrimaryApply.Enabled = True
Exit Sub
End If



If Not VerifyPort(PriDNSport) Then
MsgBox "You have entered an invalid port number for the DNS protocol. Valid port numbers are 1 - 65535.", vbCritical, "Linksys Failover Client"
PriDNSport.SetFocus
PrimaryApply.Enabled = True
Exit Sub
End If

If Not VerifyPort(PriFTPport) Then
MsgBox "You have entered an invalid port number for the FTP protocol. Valid port numbers are 1 - 65535.", vbCritical, "Linksys Failover Client"
PriFTPport.SetFocus
PrimaryApply.Enabled = True
Exit Sub
End If

If Not VerifyPort(PriSMTPport) Then
MsgBox "You have entered an invalid port number for the HTTP protocol. Valid port numbers are 1 - 65535.", vbCritical, "Linksys Failover Client"
PriSMTPport.SetFocus
PrimaryApply.Enabled = True
Exit Sub
End If


If Not VerifyPort(PriPOP3port) Then
MsgBox "You have entered an invalid port number for the POP3 protocol. Valid port numbers are 1 - 65535.", vbCritical, "Linksys Failover Client"
PriPOP3port.SetFocus
PrimaryApply.Enabled = True
Exit Sub
End If





If Not VerifyIP(PriHTTPip) Then
MsgBox "You have entered an invalid IP range for the HTTP protocol. Valid hosts on a class C network are 1 - 254.", vbCritical, "Linksys Failover Client"
PriHTTPip.SetFocus
PrimaryApply.Enabled = True
Exit Sub
End If


If Not VerifyIP(PriDNSip) Then
MsgBox "You have entered an invalid IP range for the DNS protocol. Valid hosts on a class C network are 1 - 254.", vbCritical, "Linksys Failover Client"
PriDNSip.SetFocus
PrimaryApply.Enabled = True
Exit Sub
End If


If Not VerifyIP(PriFTPip) Then
MsgBox "You have entered an invalid IP range for the FTP protocol. Valid hosts on a class C network are 1 - 254.", vbCritical, "Linksys Failover Client"
PriFTPip.SetFocus
PrimaryApply.Enabled = True
Exit Sub
End If


If Not VerifyIP(PriSMTPip) Then
MsgBox "You have entered an invalid IP range for the SMTP protocol. Valid hosts on a class C network are 1 - 254.", vbCritical, "Linksys Failover Client"
PriSMTPip.SetFocus
PrimaryApply.Enabled = True
Exit Sub
End If


If Not VerifyIP(PriPOP3ip) Then
MsgBox "You have entered an invalid IP range for the POP3 protocol. Valid hosts on a class C network are 1 - 254.", vbCritical, "Linksys Failover Client"
PriPOP3ip.SetFocus
PrimaryApply.Enabled = True
Exit Sub
End If




 Open App.Path & "\PrimaryConfigs.dat" For Output As #4
      
      Write #4, PriHTTPport, PriHTTPip, PriHTTPenabled
      Write #4, PriDNSport, PriDNSip, PriDNSenabled
      Write #4, PriFTPport, PriFTPip, PriFTPenabled
      Write #4, PriSMTPport, PriSMTPip, PriSMTPenabled
      Write #4, PriPOP3port, PriPOP3ip, PriPOP3enabled
    
Close #4


    
End Sub











Private Sub PrimaryCancel_Click()
Unload Me
End Sub

Private Sub PriPOP3enabled_Click()
PrimaryApply.Enabled = True
End Sub

Private Sub PriPOP3ip_Change()
PrimaryApply.Enabled = True
End Sub

Private Sub PriPOP3port_Change()
PrimaryApply.Enabled = True
End Sub

Private Sub PriSMTPenabled_Click()
PrimaryApply.Enabled = True
End Sub

Private Sub PriSMTPip_Change()
PrimaryApply.Enabled = True
End Sub

Private Sub PriSMTPport_Change()
PrimaryApply.Enabled = True
End Sub


Private Sub secDNSenabled_Click()
SecondaryApply.Enabled = True
End Sub

Private Sub secDNSip_Change()
SecondaryApply.Enabled = True
End Sub

Private Sub secDNSport_Change()
SecondaryApply.Enabled = True
End Sub

Private Sub secFTPenabled_Click()
SecondaryApply.Enabled = True
End Sub

Private Sub secFTPip_Change()
SecondaryApply.Enabled = True
End Sub

Private Sub secFTPport_Change()
SecondaryApply.Enabled = True
End Sub

Private Sub secHTTPenabled_Click()
SecondaryApply.Enabled = True
End Sub

Private Sub secHTTPip_Change()
SecondaryApply.Enabled = True
End Sub

Private Sub secHTTPport_Change()
SecondaryApply.Enabled = True
End Sub





















Private Sub SecondaryApply_Click()

SecondaryApply.Enabled = False



If Not VerifyPort(secHTTPport) Then
MsgBox "You have entered an invalid port number for the HTTP protocol. Valid port numbers are 1 - 65535.", vbCritical, "Linksys Failover Client"
secHTTPport.SetFocus
SecondaryApply.Enabled = True
Exit Sub
End If



If Not VerifyPort(secDNSport) Then
MsgBox "You have entered an invalid port number for the DNS protocol. Valid port numbers are 1 - 65535.", vbCritical, "Linksys Failover Client"
secDNSport.SetFocus
SecondaryApply.Enabled = True
Exit Sub
End If




If Not VerifyPort(secFTPport) Then
MsgBox "You have entered an invalid port number for the FTP protocol. Valid port numbers are 1 - 65535.", vbCritical, "Linksys Failover Client"
secFTPport.SetFocus
SecondaryApply.Enabled = True
Exit Sub
End If




If Not VerifyPort(secSMTPport) Then
MsgBox "You have entered an invalid port number for the HTTP protocol. Valid port numbers are 1 - 65535.", vbCritical, "Linksys Failover Client"
secSMTPport.SetFocus
SecondaryApply.Enabled = True
Exit Sub
End If


If Not VerifyPort(secPOP3port) Then
MsgBox "You have entered an invalid port number for the POP3 protocol. Valid port numbers are 1 - 65535.", vbCritical, "Linksys Failover Client"
secPOP3port.SetFocus
SecondaryApply.Enabled = True
Exit Sub
End If





If Not VerifyPort(SecHTTPip) Then
MsgBox "You have entered an invalid IP range for the HTTP protocol. Valid hosts on a class C network are 1 - 254.", vbCritical, "Linksys Failover Client"
SecHTTPip.SetFocus
SecondaryApply.Enabled = True
Exit Sub
End If





If Not VerifyIP(SecDNSip) Then
MsgBox "You have entered an invalid IP range for the DNS protocol. Valid hosts on a class C network are 1 - 254.", vbCritical, "Linksys Failover Client"
SecDNSip.SetFocus
SecondaryApply.Enabled = True
Exit Sub
End If


If Not VerifyIP(SecFTPip) Then
MsgBox "You have entered an invalid IP range for the FTP protocol. Valid hosts on a class C network are 1 - 254.", vbCritical, "Linksys Failover Client"
SecFTPip.SetFocus
SecondaryApply.Enabled = True
Exit Sub
End If




If Not VerifyIP(SecSMTPip) Then
MsgBox "You have entered an invalid IP range for the SMTP protocol. Valid hosts on a class C network are 1 - 254.", vbCritical, "Linksys Failover Client"
SecSMTPip.SetFocus
SecondaryApply.Enabled = True
Exit Sub
End If



If Not VerifyIP(SecPOP3ip) Then
MsgBox "You have entered an invalid IP range for the POP3 protocol. Valid hosts on a class C network are 1 - 254.", vbCritical, "Linksys Failover Client"
SecPOP3ip.SetFocus
SecondaryApply.Enabled = True
Exit Sub
End If






Open App.Path & "\FailoverConfigs.dat" For Output As #5
      
      Write #5, secHTTPport, SecHTTPip, secHTTPenabled
      Write #5, secDNSport, SecDNSip, secDNSenabled
      Write #5, secFTPport, SecFTPip, secFTPenabled
      Write #5, secSMTPport, SecSMTPip, secSMTPenabled
      Write #5, secPOP3port, SecPOP3ip, secPOP3enabled

Close #5



End Sub













Private Sub secondaryCancel_Click()
Unload Me
End Sub

Private Sub secPOP3enabled_Click()
SecondaryApply.Enabled = True
End Sub

Private Sub secPOP3ip_Change()
SecondaryApply.Enabled = True
End Sub

Private Sub secPOP3port_Change()
SecondaryApply.Enabled = True
End Sub

Private Sub secSMTPenabled_Click()
SecondaryApply.Enabled = True
End Sub

Private Sub secSMTPip_Change()
SecondaryApply.Enabled = True
End Sub

Private Sub secSMTPport_Change()
SecondaryApply.Enabled = True
End Sub
