VERSION 5.00
Begin VB.Form frmBackUp 
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timPrgBar 
      Left            =   6240
      Top             =   3480
   End
   Begin VB.TextBox txtfileName 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtpath 
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Restore DataBase"
      Height          =   1095
      Left            =   3840
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin VB.DirListBox drilistbackup 
      Height          =   1665
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.DriveListBox driBackup 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdBackUp 
      Caption         =   "BackUp DataBase"
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   3120
      Width           =   2175
   End
End
Attribute VB_Name = "frmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BLOOD_BANKDB As Database
Dim rstDonor As Recordset
Dim rstReference As Recordset
Dim rstCheck_No As Recordset
Dim VarPre_No As String
Dim VarGenerate_No As String
Dim VarCheck_No As String
Dim FileSystemObject As Object
Dim strFilename As String

Private Sub cmdBackUp_Click()
On Error GoTo er
strFilename = "" + txtpath.Text + "\" + txtfileName.Text + ".mdb"
Set FileSystemObject = CreateObject("Scripting.filesystemobject")
FileSystemObject.copyfile App.Path & "\Blood_Bank.mdb", strFilename
Exit Sub
er:
MsgBox "Invalid Path", vbCritical, "Invali Path setting! "
End Sub

Private Sub driBackup_Change()
Dim d, fs As Object
    
    'Set the constrctions to created objectes
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(driBackup.Drive))
    
    'Set the contents of the selected drives
    If d.IsReady Then
    drilistbackup.Path = driBackup.Drive
        drilistbackup.SetFocus
        Else
        MsgBox "The Drive Is Not Ready!", vbExclamation, "Drive Not Ready!"
    End If
End Sub

Private Sub drilistbackup_Click()
 txtpath.Text = "" & drilistbackup.Path
End Sub

Private Sub Form_Load()
frmBackUp.Caption = "Back Up Files"
cmdBackUp.Caption = "Click to BackUp"

Set BLOOD_BANKDB = OpenDatabase(App.Path & "\BLOOD_BANK.mdb", False, False)
Set rstDonor = BLOOD_BANKDB.OpenRecordset("DONOR_RECORD")
Set rstReference = BLOOD_BANKDB.OpenRecordset("DONOR_REFERENCE_DATA")
Set rstCheck_No = BLOOD_BANKDB.OpenRecordset("Used_No")

txtfileName.Text = FormatDateTime(Now, vbLongDate)

End Sub
