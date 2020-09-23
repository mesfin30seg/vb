VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2610
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4575
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1542.074
   ScaleMode       =   0  'User
   ScaleWidth      =   4295.677
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1890
      TabIndex        =   1
      Top             =   255
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   1800
      Width           =   1020
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1005
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   360
      Picture         =   "frmSettings.frx":0442
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   345
      TabIndex        =   0
      Top             =   360
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   330
      TabIndex        =   2
      Top             =   1140
      Width           =   1335
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If txtPassword = "blood" And txtUserName = "admin" Then
        LoginSucceeded = True
        txtUserName = ""
        txtPassword = ""
        Unload Me
        frm_Menu.Show vbModal
        Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtUserName = ""
        txtPassword = ""
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        End If
        End
End Sub

Private Sub Form_Load()
frmSettings.Caption = "Admin Password"
End Sub
