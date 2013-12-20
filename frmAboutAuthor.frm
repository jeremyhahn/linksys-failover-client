VERSION 5.00
Begin VB.Form frmAboutAuthor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About :: Pc Technics"
   ClientHeight    =   3375
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAboutAuthor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2329.485
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "All Rights Reserved"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   3030
      Width           =   1665
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   56.343
      X2              =   5281.227
      Y1              =   1263.099
      Y2              =   1263.099
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAboutAuthor.frx":08CA
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
      Height          =   1410
      Left            =   930
      TabIndex        =   0
      Top             =   180
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   56.343
      X2              =   5267.141
      Y1              =   1263.099
      Y2              =   1263.099
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  For custom versions, suggestions, bugreports, or general feedback, please see our website at www.pc-technics.com. "
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
      Height          =   675
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   3870
   End
End
Attribute VB_Name = "frmAboutAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = (Screen.Height / 2) - (Me.Height / 2)
End Sub
