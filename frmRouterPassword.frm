VERSION 5.00
Object = "{71A54968-8589-4C1A-9739-425B9222A82F}#1.0#0"; "Cryptx.ocx"
Begin VB.Form frmRouterPassword 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Router Password"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   FillColor       =   &H00FF0000&
   Icon            =   "frmRouterPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Okay"
      Default         =   -1  'True
      Height          =   330
      Left            =   2790
      TabIndex        =   3
      Top             =   1140
      Width           =   780
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   525
      Width           =   2325
   End
   Begin CryptXCtl.CryptX Crypto 
      Left            =   180
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF0000&
      Caption         =   "Confirm:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   600
      Width           =   720
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF0000&
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1080
   End
End
Attribute VB_Name = "frmRouterPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()

    Me.Hide

End Sub






Private Sub cmdOK_Click()
   

   
   Select Case txtPassword
   
   
   Case "admin"
       
       MsgBox "You should change your Linksys router password, as this is the default password which ships with ALL linksys routers. This password is not very secure."



   Case vbNullString
   
       MsgBox "Please change your BLANK router password. This program was not intended for insecure networks."
       
       txtPassword.Text = vbNullString
       txtConfirmPassword.Text = vbNullString
       txtPassword.SetFocus
       
         Exit Sub
       

    End Select
    
    
    
    If Not txtPassword = txtConfirmPassword Then
       MsgBox "Passwords do not match!", vbInformation, "Linksys Failover Client"
       Exit Sub
End If
    
    
    
    

    
    
    Open App.Path & "/sysconfig.dat" For Output As #1
    
     Crypto.Password = "PcTechnicsCryptoAPI"
     Password = Crypto.Encrypt(txtPassword)
     
     Write #1, Password
     Write #1, "1"

    Close #1
    

           MsgBox "Password has been set.", vbInformation, "Linksys Failover Client"

                Me.Hide



    
    
End Sub
