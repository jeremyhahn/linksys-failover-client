VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Router Address"
      Height          =   4575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4875
      Begin VB.TextBox txtUpdateInterval 
         Height          =   285
         Left            =   2550
         TabIndex        =   2
         Text            =   "1"
         Top             =   3150
         Width           =   405
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   3600
         TabIndex        =   6
         Top             =   4050
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Okay"
         Height          =   345
         Left            =   2790
         TabIndex        =   5
         Top             =   4050
         Width           =   795
      End
      Begin VB.TextBox txtRouterIP 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Text            =   "192.168.1.1"
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "minuite(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3030
         TabIndex        =   9
         Top             =   3210
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Every"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   8
         Top             =   3210
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "  What interval would you like the client to perform routine queries to the servers?"
         Height          =   465
         Left            =   300
         TabIndex        =   7
         Top             =   2430
         Width           =   4305
      End
      Begin VB.Label Label2 
         Caption         =   "Router Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1470
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "  The update engine needs to know the IP adress of your Linksys router."
         Height          =   465
         Left            =   300
         TabIndex        =   3
         Top             =   600
         Width           =   4065
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If (txtUpdateInterval = vbNullString) Then
MsgBox "You must enter an update interval", vbCritical, "Linksys Failover Client by Pc Technics"
Exit Sub
End If


If (txtRouterIP = vbNullString) Then
MsgBox "You must enter the IP address of your router.", vbCritical, "Linksys Failover Client by Pc Technics"
Exit Sub
End If


If (txtUpdateInterval < 1) Then
MsgBox "You must enter a positive number greater than 0 as your update interval.", vbCritical, "Linksys Failover Client by Pc Technics"
Exit Sub
End If



 Open App.Path & "/IP.dat" For Output As #1
      
      Write #1, txtRouterIP, txtUpdateInterval
      
  Close #1
      

Unload Me

Unload frmMain

Load frmMain
frmMain.Show


End Sub





Private Sub Command2_Click()
Unload Me
End Sub







Private Sub Form_Load()

Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = (Screen.Height / 2) - (Me.Height / 2)


    VERIFY_IP_SETTINGS
    'Get The Router IP
     Open App.Path & "\IP.dat" For Input As #2
          Input #2, RouterIP, UpdateInterval
      Close #2


txtRouterIP = RouterIP
txtUpdateInterval = UpdateInterval


End Sub
