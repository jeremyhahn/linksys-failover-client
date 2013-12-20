VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copyright 2004 Pc Technics"
   ClientHeight    =   3375
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2329.485
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Linksys Failover Client v1.1"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1020
      TabIndex        =   3
      Top             =   60
      Width           =   3765
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "All Rights Reserved"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2130
      TabIndex        =   2
      Top             =   3060
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   56.343
      X2              =   5281.227
      Y1              =   1470.164
      Y2              =   1470.164
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright 2004 Pc Technics                                            Designed By: Jeremy Hahn"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1020
      TabIndex        =   0
      Top             =   2280
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   56.343
      X2              =   5267.141
      Y1              =   1470.164
      Y2              =   1470.164
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAbout.frx":08CA
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1605
      Left            =   180
      TabIndex        =   1
      Top             =   450
      Width           =   5430
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = (Screen.Height / 2) - (Me.Height / 2)
End Sub
