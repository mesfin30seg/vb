VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBlood_BankSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4365
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBlood_BankSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar CBBSProgressBar 
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   8265
      Begin VB.Timer progressTimer 
         Interval        =   1000
         Left            =   5160
         Top             =   3240
      End
      Begin VB.Timer loadTimer 
         Interval        =   1000
         Left            =   6000
         Top             =   3240
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   240
         Shape           =   2  'Oval
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Please Wai:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblProductName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   5
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Copyright (C) IBRUQ: Product is copyrighted in the year 2008"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   2280
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   " Warning : Use of pirated copy is illegal."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   2
         Top             =   2640
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmBlood_BankSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Windows Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub progressTimer_Timer()
    i = Rnd() * 20
    If CBBSProgressBar.Value < 100 Then
        If CBBSProgressBar.Value + i < 100 Then
            CBBSProgressBar.Value = CBBSProgressBar.Value + i
        Else
            CBBSProgressBar.Value = 100
        End If
    Else
    Unload Me
   frmSettings.Show
    'MDI_BBMS.Show
    End If
    End Sub

   
