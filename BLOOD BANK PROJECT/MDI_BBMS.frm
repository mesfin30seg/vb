VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDI_BBMS 
   BackColor       =   &H8000000C&
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5745
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "5:51 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "9/1/2008"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4710
            Text            =   "Blood Bank Software Running"
            TextSave        =   "Blood Bank Software Running"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MDI_BBMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Activate()
frmSettings.Show vbModal, MDI_BBMS
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Exit Blood Bank?", vbYesNo + vbQuestion, "Exit") = vbYes Then
End
Else
Cancel = True
End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
drp_Recipient.Hide
drp_Donor.Hide
frmDONOR_RECORD.Hide
frmREFERENCE_RECORD.Hide
frmREFERENCE_RECORD.Hide
End Sub

Private Sub mnuClose_Click()
On Error Resume Next
Unload Me.ActiveForm
End Sub

Private Sub mnuDonor_Click()
frmDONOR_RECORD.Show
frmRECIPIENT_RECORD.Hide
frmREFERENCE_RECORD.Hide
drp_Donor.Hide
drp_Recipient.Hide
End Sub

Private Sub mnuDRP_Click()
drp_Donor.Show
drp_Recipient.Hide
frmRECIPIENT_RECORD.Hide
frmDONOR_RECORD.Hide
frmREFERENCE_RECORD.Hide
End Sub

Private Sub mnuExit_Click()
If MsgBox("Exit Blood Bank?", vbYesNo + vbQuestion, "Exit") = vbYes Then
End
Else
Cancel = True
End If
End Sub

Private Sub mnuRecipient_Click()
frmRECIPIENT_RECORD.Show
frmDONOR_RECORD.Hide
frmREFERENCE_RECORD.Hide
drp_Recipient.Hide
drp_Donor.Hide
End Sub

Private Sub mnuReference_Click()
frmRECIPIENT_RECORD.Hide
frmDONOR_RECORD.Hide
frmREFERENCE_RECORD.Show
drp_Recipient.Hide
drp_Donor.Hide
End Sub

Private Sub mnuTRP_Click()
drp_Recipient.Show
drp_Donor.Hide
frmDONOR_RECORD.Hide
frmREFERENCE_RECORD.Hide
frmREFERENCE_RECORD.Hide
End Sub


