VERSION 5.00
Begin VB.Form frm_Menu 
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   12270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_Backup 
      Caption         =   "Backup Database"
      Height          =   2175
      Left            =   9840
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   615
      Left            =   8640
      MouseIcon       =   "frmBacground.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmBacground.frx":030A
      TabIndex        =   3
      ToolTipText     =   "Donor Form"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDonor_List 
      Caption         =   "Donor &List"
      Height          =   615
      Left            =   7200
      MouseIcon       =   "frmBacground.frx":074C
      MousePointer    =   99  'Custom
      Picture         =   "frmBacground.frx":0A56
      TabIndex        =   2
      ToolTipText     =   "View Booking"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecipient_Record 
      Caption         =   "&Recipient Form"
      Height          =   615
      Left            =   5760
      MouseIcon       =   "frmBacground.frx":0E98
      MousePointer    =   99  'Custom
      Picture         =   "frmBacground.frx":11A2
      TabIndex        =   1
      ToolTipText     =   "View Flight Schedules"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDonor_Record 
      Caption         =   "&Donor Form"
      Height          =   615
      Left            =   4320
      MouseIcon       =   "frmBacground.frx":15E4
      MousePointer    =   99  'Custom
      Picture         =   "frmBacground.frx":18EE
      TabIndex        =   0
      ToolTipText     =   "Recipient Form"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MAIN MENU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   555
      Left            =   5880
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   10200
      MouseIcon       =   "frmBacground.frx":1D30
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Exit Database"
      Top             =   6960
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   4230
      Left            =   4320
      Picture         =   "frmBacground.frx":203A
      Top             =   2040
      Width           =   5295
   End
End
Attribute VB_Name = "frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Backup_Click()
frmBackUp.Show vbModal
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub cmdDonor_List_Click()
frmREFERENCE_RECORD.Show vbModal
End Sub

Private Sub cmdDonor_Record_Click()
frmDONOR_RECORD.Show vbModal

End Sub

Private Sub cmdRecipient_Record_Click()
frmRECIPIENT_RECORD.Show vbModal
End Sub

Private Sub cmdRpt_Donor_Click()
drp_Donor.Show vbModal
End Sub

Private Sub cmdRptReci_Click()
drp_Recipient.Show vbModal
End Sub

Private Sub cmdSettings_Click()
'frmSettings.Show vbModal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Exit Blood Bank?", vbYesNo + vbQuestion, "Exit") = vbYes Then
End
Else
Cancel = True
End If
End Sub

Private Sub lblExit_Click()
If MsgBox("Exit Blood Bank?", vbYesNo + vbQuestion, "Exit") = vbYes Then
End
Else
Cancel = True
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.FontSize = 14
lblExit.ForeColor = &HFF00&
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.FontSize = 24
lblExit.ForeColor = &HFF&
End Sub
Private Sub Form_Load()
frm_Menu.Caption = "Main Menu"
End Sub



