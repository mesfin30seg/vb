VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecipient 
   BackColor       =   &H80000011&
   Caption         =   "Bleeding Process (Transfusion)"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Caption         =   "Cross Matching"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   3015
      Left            =   2520
      TabIndex        =   42
      Top             =   4320
      Width           =   5655
      Begin VB.CommandButton cmdMatch 
         BackColor       =   &H8000000C&
         Caption         =   "&Match"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4440
         TabIndex        =   55
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txt_BloodGroup 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2520
         TabIndex        =   53
         Top             =   360
         Width           =   1695
      End
      Begin VB.ListBox lstBagNo 
         Appearance      =   0  'Flat
         Height          =   420
         ItemData        =   "frmRecipient.frx":0000
         Left            =   2520
         List            =   "frmRecipient.frx":0002
         TabIndex        =   46
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lbl_Rhesus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   54
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000B&
         Caption         =   "Enter Blood Group"
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
         Left            =   240
         TabIndex        =   52
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000A&
         Caption         =   "Rhesus:"
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
         Left            =   240
         TabIndex        =   49
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblExp_Date 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   48
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblDonor_No 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   47
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000B&
         Caption         =   "Epiration Date:"
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
         Left            =   240
         TabIndex        =   45
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000B&
         Caption         =   "Donor's Number:"
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
         Left            =   240
         TabIndex        =   44
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000B&
         Caption         =   "Bag Number:"
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
         Left            =   240
         TabIndex        =   43
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Screening Officer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   3015
      Left            =   8400
      TabIndex        =   32
      Top             =   4320
      Width           =   4695
      Begin VB.TextBox txtScreen_Off 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   36
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtDesig 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   35
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtRemark 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   1320
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtp_ScreenDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   33
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643841
         CurrentDate     =   39655
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000B&
         Caption         =   "Officer's Name:"
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
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000B&
         Caption         =   "Designation:"
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
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackColor       =   &H8000000B&
         Caption         =   "Remark"
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
         Left            =   240
         TabIndex        =   38
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label23 
         BackColor       =   &H8000000B&
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   2160
         Width           =   975
      End
   End
   Begin VB.Frame fraNon_ViralScreen 
      BackColor       =   &H8000000A&
      Caption         =   "None Viral Screening"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   8400
      TabIndex        =   25
      Top             =   720
      Width           =   4695
      Begin VB.ComboBox cbo_Rhesus 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   51
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtHBE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         Left            =   2520
         TabIndex        =   28
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtPCV 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         Left            =   2520
         TabIndex        =   27
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox cboBlood_Group 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   2520
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000B&
         Caption         =   "Rhesus:"
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
         Left            =   240
         TabIndex        =   50
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000A&
         Caption         =   "Hemoglobin Estimate(%):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000A&
         Caption         =   "Blood Pack Cell Volume:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         Caption         =   "Blood Group:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   13440
      Picture         =   "frmRecipient.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Save New Record"
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton cmdNew 
      Height          =   375
      Left            =   13440
      Picture         =   "frmRecipient.frx":0106
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Add Record"
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton cmdSave_Edit 
      Height          =   375
      Left            =   13440
      Picture         =   "frmRecipient.frx":0638
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Edit Record"
      Top             =   6720
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   375
      Left            =   13440
      Picture         =   "frmRecipient.frx":073A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Find Record"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   13440
      Picture         =   "frmRecipient.frx":083C
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Delete Record"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton cmdLastRecord 
      Caption         =   ">|"
      Height          =   375
      Left            =   12720
      TabIndex        =   18
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   12360
      TabIndex        =   17
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton cmdFirstRecord 
      Caption         =   "|<"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   7560
      Width           =   375
   End
   Begin VB.Frame fra_Recipient 
      BackColor       =   &H8000000B&
      Caption         =   "Recipient Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   3495
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   5655
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643841
         CurrentDate     =   39654
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   11
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ComboBox cboSex 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   10
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txtPAddress 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtFullName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblGenerate_No 
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblPrefix_No 
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000B&
         Caption         =   "Full Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000B&
         Caption         =   "Recipient's Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000B&
         Caption         =   "Permanent Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000B&
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000B&
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000B&
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblRecipientNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   13320
      TabIndex        =   41
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   3240
      TabIndex        =   24
      Top             =   7560
      Width           =   9135
   End
End
Attribute VB_Name = "frmRecipient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BLOOD_BANKDB As Database
Dim rstRecipient As Recordset
Dim rstDonor As Recordset
Dim rstSold As Recordset

Dim VarRecipientNumber As String
Dim VarPre_No As String
Dim VarGenerate_No As String

Dim Var_BBN As String
Dim Var_BloodGroup As String
Dim Var_Rhesus As String

Private Sub cmdDelete_Click()
On Error GoTo ErrorTrap:
If MsgBox("Delete Recipient?", vbQuestion + vbYesNo, "Delete") = vbNo Then
Exit Sub
End If
With rstRecipient
.Delete
Clearform
End With
Exit Sub
ErrorTrap:
 MsgBox Err.Description, vbInformation, "Error"
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
Dim StrSearch As String
StrSearch = InputBox("Enter Recipient's Number:", "Find Donor")
On Error Resume Next
With rstRecipient
.Index = "Recipient's Number"
.Seek "=", StrSearch
If .NoMatch Then
MsgBox "No Record", vbInformation, "Find Recipient"
Exit Sub
Else
LoadData
End If
End With
End Sub
Private Sub LoadData()
On Error GoTo ErrorTrap:
With rstRecipient
txtFullName = ![Full Name]
lblRecipientNumber = ![Recipient's Number]
txtPAddress = ![Permanent Address]
cboSex = !Sex
txtAge = !Age
dtpDate = !Date
lblGenerate_No = ![Prefix No]
cboBlood_Group = ![Blood Group]
cbo_Rhesus = ![Rhesus]
txtPCV = !PCV
txtHBE = ![Hemoglobin Est]
End With
Exit Sub
ErrorTrap:
End Sub

Private Sub Clearform()
On Error Resume Next
txtFullName = ""
lblRecipientNumber = ""
txtPAddress = ""
cboSex = ""
txtAge = ""
dtpDate = ""
txtBBN = ""
cboBlood_Group = ""
cbo_Rhesus = ""
txtPCV = ""
txtHBE = ""
End Sub

Private Sub cmdFirstRecord_Click()
On Error Resume Next
With rstRecipient
.MoveFirst
LoadData
End With
End Sub

Private Sub cmdLastRecord_Click()
On Error Resume Next
With rstRecipient
.MoveLast
LoadData
End With
End Sub

Private Sub cmdMatch_Click()
On Error Resume Next
With rstDonor
.Delete
End With
'With rstSold
'.MoveLast
'.AddNew
'![Full Name] = lblName
'![Donor's Number] = lblDonorNo
'lblAddress = ![Permanent Address]
'!Age = lblAge
'!Date = lblCollectionDate
'![Bag Number] = lblExpDate
'![Blood Group] = lblBloodGroup
'!Rhesus = lblRhesus
'![Bag Number] = lblBBN
'![Hemoglobin Estimate] = lblHEST
'![Blood Pack Cell Volume ] = lblBPCV
'!HIV = lblHIV
'![Hepatitis B Surface Antigen] = lblHBSA
'![Hepatitis C Virus] = lblHCV
'!Syphilis = lblSyphilis
'![Physical Exam] = lblPhyExam
'![Officer's Name] = lblScreenOff
'!Designition = lblRank
'!Remark = lblRemark
'![Screening Date] = lblScreenDate
'.Update
'.Bookmark = .LastModified

'End With
'End With
'Exit Sub
End Sub

Private Sub cmdNew_Click()
On Error Resume Next

Clearform

VarPre_No = "RC/"
lblPrefix_No = VarPre_No

With rstRecipient
.MoveLast
If .BOF And .EOF Then
VarGenerate_No = Format(1, "000")
lblGenerate_No = VarGenerate_No
Else
lblGenerate_No = Format(CDbl(![Prefix No]) + 1, "000")
End If
End With
lblRecipientNumber = lblPrefix_No & lblGenerate_No
txtFullName.SetFocus
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
With rstRecipient
.MoveNext
If .EOF Then
 .MoveLast
End If
LoadData
End With

End Sub

Private Sub cmdPrevious_Click()
On Error Resume Next
With rstRecipient
.MovePrevious
If .BOF Then
 .MoveFirst
End If
LoadData
End With
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorTrap:
With rstRecipient

'.MoveLast
.AddNew
![Full Name] = txtFullName
![Recipient's Number] = lblRecipientNumber
![Permanent Address] = txtPAddress
!Sex = cboSex
!Age = txtAge
!Date = dtpDate
![Prefix No] = lblGenerate_No
![Blood Group] = cboBlood_Group
![Rhesus] = cbo_Rhesus
!PCV = txtPCV
![Hemoglobin Est] = txtHBE

.Update
.Bookmark = .LastModified
Clearform
End With
Exit Sub
ErrorTrap:
MsgBox Err.Description, vbInformation, "Error"

End Sub

Private Sub cmdSave_Edit_Click()
On Error Resume Next
With rstRecipient
.Edit
![Full Name] = txtFullName
![Recipient's Number] = lblRecipientNumber
![Permanent Address] = txtPAddress
!Sex = cboSex
!Age = txtAge
!Date = dtpDate
![Prefix No] = lblGenerate_No
![Blood Group] = cboBlood_Group
![Rhesus] = cbo_Rhesus
!PCV = txtPCV
![Hemoglobin Est] = txtHBE
.Update
.Bookmark = .LastModified
Clearform
End With
End Sub

Private Sub Command1_Click()
'Var_BloodGroup = txt_BloodGroup
'Set rstDonor = BLOOD_BANKDB.OpenRecordset(" Select * from Donor where [Blood Group] = '" & Var_BloodGroup & "';", dbOpenDynaset)
'With rstDonor
'lstBagNo.Clear
'While Not .EOF
'lstBagNo.AddItem ![Bag Number]
'.MoveNext
'Wend
'End With
End Sub

Private Sub Form_Load()

Set BLOOD_BANKDB = OpenDatabase(App.Path & "\BLOOD_BANK.mdb", False, False)
Set rstRecipient = BLOOD_BANKDB.OpenRecordset("Recipient")
Set rstDonor = BLOOD_BANKDB.OpenRecordset("Donor")
Set rstSold = BLOOD_BANKDB.OpenRecordset("Sold")

cboSex.AddItem "Female"
cboSex.AddItem "Male"

cboBlood_Group.AddItem "A"
cboBlood_Group.AddItem "B"
cboBlood_Group.AddItem "AB"
cboBlood_Group.AddItem "O"

cbo_Rhesus.AddItem "Positive"
cbo_Rhesus.AddItem "Negative"

End Sub

Private Sub lstBagNo_Click()
lblDonor_No = ""
lblExp_Date = ""
lbl_Rhesus = ""
Var_BBN = lstBagNo.Text
Set rstDonor = BLOOD_BANKDB.OpenRecordset(" Select * from Donor where [Blood Group]= '" & Var_BloodGroup & "' And [Bag Number] = '" & Var_BBN & "';", dbOpenDynaset)
With rstDonor
lblDonor_No = ![Donor's Number]
lblExp_Date = ![Exp Date]
lbl_Rhesus = ![Rhesus]
End With
End Sub

Private Sub txt_BloodGroup_Change()
Var_BloodGroup = txt_BloodGroup
Set rstDonor = BLOOD_BANKDB.OpenRecordset(" Select * from Donor where [Blood Group] = '" & Var_BloodGroup & "';", dbOpenDynaset)
With rstDonor
lstBagNo.Clear
While Not .EOF
lstBagNo.AddItem ![Bag Number]
.MoveNext
Wend
End With

End Sub







