VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Object = "{9D520B81-DCB4-11D3-8ED1-97B0B30DF77E}#1.0#0"; "MBTray.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Object = "{71A54968-8589-4C1A-9739-425B9222A82F}#1.0#0"; "Cryptx.ocx"
Begin VB.Form frmService 
   Caption         =   "LFC Update Engine"
   ClientHeight    =   1515
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4125
   Icon            =   "frmService.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   840
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin MBTray.Tray Tray1 
      Left            =   3660
      Top             =   1050
      _ExtentX        =   847
      _ExtentY        =   847
      Icon            =   "frmService.frx":08CA
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin NTService.NTService NTService1 
      Left            =   420
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "Linksys Failover Client"
      Interactive     =   -1  'True
      ServiceName     =   "LFC"
      StartMode       =   2
   End
   Begin CryptXCtl.CryptX Crypto 
      Left            =   0
      Top             =   1050
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "TrayMenu"
      Begin VB.Menu TrayAbout 
         Caption         =   "About"
      End
      Begin VB.Menu TrayCopyright 
         Caption         =   "Copyright"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu TrayToggle 
         Caption         =   "Show/Hide"
      End
      Begin VB.Menu TrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strBuffer As String
Dim cchBuffer As Integer
Dim strHeader As String

Dim dtNextTime As Date


'Set Up The NT Service Routines
Private Sub NTService1_Continue(Success As Boolean)
  subLogCommand "NT Service Continued ...", ""

  Success = True
End Sub





     Private Sub NTService1_Pause(Success As Boolean)
        subLogCommand "NT Service Paused ...", ""
        Success = True
     End Sub




           Private Sub NTService1_Start(Success As Boolean)
             subLogCommand "Service Started ...", ""
             frmMain.txtLogDisplay.AddItem ""
             frmMain.txtLogDisplay.AddItem "NT Service Started .."
             
             NTService1.ControlsAccepted = svcCtrlPauseContinue
             NTService1.StartService
             
             Me.Hide
             Success = True
           End Sub




                  
                  Private Sub NTService1_Stop()
                    subLogCommand "NT Service Stopped ...", ""
                    frmMain.txtLogDisplay.AddItem ""
                    frmMain.txtLogDisplay.AddItem "Service Stopped"
                  End Sub














Private Sub Form_Load()

Me.Visible = False

Tray1.Create
Tray1.ToolTip = "Linksys Failover Client v" & AppVersion & vbCrLf & _
                "(c) 2004 Pc Technics"
                
DO_GLOBAL_POLL
UpdateControls



   VERIFY_IP_SETTINGS
   Open App.Path & "/IP.dat" For Input As #1

        Input #1, RouterIP, UpdateInterval

   Close #1
    

                    'Set up the timer
                     Timer1.Interval = 500
                     Timer1.Enabled = True
                     dtNextTime = DateAdd("n", UpdateInterval, Now)

End Sub






Private Sub Timer1_Timer()

If Now >= dtNextTime Then
Timer1.Enabled = False



  VERIFY_IP_SETTINGS
  Open App.Path & "/IP.dat" For Input As #1

        Input #1, RouterIP, UpdateInterval

    Close #1


            
            Tray1.Destroy
            Tray1.Create





                'Run The Polling Sub Again
                 DO_GLOBAL_POLL
                 UpdateControls
                 




dtNextTime = DateAdd("n", UpdateInterval, Now)
Timer1.Enabled = True
End If

End Sub













' ###### Start System Tray ######
Private Sub Tray1_DblClick(ByVal Button As Long)
Load frmMain
frmMain.Show
End Sub









Private Sub TrayAbout_Click()
Load frmAboutAuthor
frmAboutAuthor.Show
End Sub






Private Sub TrayCopyright_Click()
Load frmAbout
frmAbout.Show
End Sub









Private Sub Tray1_MouseUp(ByVal Button As Long)
    If Button = vbRightButton Then PopupMenu TrayMenu
End Sub






Private Sub TrayExit_Click()
Unload frmMain
Unload Me
End Sub




Private Sub TrayToggle_Click()

     If frmMain.Visible Then
        Unload frmMain
   Else
        Load frmMain
        frmMain.Show
End If
End Sub

'###### End System Tray #####
