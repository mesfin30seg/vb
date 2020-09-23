VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDONOR_RECORD 
   Caption         =   "Blood Donation Details"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLastRecord 
      Caption         =   ">|"
      Height          =   375
      Left            =   9120
      TabIndex        =   63
      ToolTipText     =   "Goto Last Record"
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   8760
      TabIndex        =   62
      ToolTipText     =   "Goto Next Record"
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   375
      Left            =   4440
      TabIndex        =   61
      ToolTipText     =   "Goto Previous Record"
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton cmdFirstRecord 
      Caption         =   "|<"
      Height          =   375
      Left            =   4080
      TabIndex        =   60
      ToolTipText     =   "Go to First Record"
      Top             =   8760
      Width           =   375
   End
   Begin VB.TextBox txtRemark 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   52
      Top             =   7920
      Width           =   1935
   End
   Begin VB.TextBox txt_Rank 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      TabIndex        =   51
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox txtScreen_Off 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   50
      Top             =   7440
      Width           =   2175
   End
   Begin VB.TextBox txtHBE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   9120
      TabIndex        =   49
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox txtPCV 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   5520
      TabIndex        =   48
      Top             =   6480
      Width           =   735
   End
   Begin VB.ComboBox cbo_Rhesus 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   47
      Top             =   6480
      Width           =   1455
   End
   Begin VB.ComboBox cboBlood_Group 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   9120
      TabIndex        =   46
      Top             =   6000
      Width           =   975
   End
   Begin VB.ComboBox cboPhy_Exam 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   9120
      TabIndex        =   45
      Top             =   5520
      Width           =   1815
   End
   Begin VB.ComboBox cbo_HCV 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   43
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ComboBox cbo_HBSA 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   5520
      TabIndex        =   42
      Top             =   5520
      Width           =   1455
   End
   Begin VB.ComboBox cbo_Syph 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   41
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ComboBox cbo_HIV 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   1560
      TabIndex        =   40
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtAdverse_RXN 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Top             =   3840
      Width           =   2055
   End
   Begin VB.ComboBox cboEver_Do 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   37
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtPAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   2880
      Width           =   2415
   End
   Begin MSMask.MaskEdBox mskDonation_Date 
      Height          =   255
      Left            =   8040
      TabIndex        =   32
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtBBN 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   31
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cboDonorType 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   30
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   29
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtParentName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   28
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ComboBox cboSex 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   27
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtFullName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   26
      Top             =   960
      Width           =   2895
   End
   Begin MSMask.MaskEdBox mskExpir_Date 
      Height          =   255
      Left            =   12360
      TabIndex        =   34
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskPhone_No 
      Height          =   255
      Left            =   2640
      TabIndex        =   35
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskLast_Do_Date 
      Height          =   255
      Left            =   2640
      TabIndex        =   38
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskSceen_Date 
      Height          =   255
      Left            =   7800
      TabIndex        =   53
      Top             =   8040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   10200
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Donor.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Donor.frx":0112
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Donor.frx":0224
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Donor.frx":0336
            Key             =   "Button"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Donor.frx":0448
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Record"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find Donor"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Button"
            Object.ToolTipText     =   "Edit Record"
            ImageKey        =   "Button"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Record"
            ImageKey        =   "Delete"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label33 
      Caption         =   "%"
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
      Left            =   6240
      TabIndex        =   66
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label32 
      Caption         =   "%"
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
      Left            =   9840
      TabIndex        =   65
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label31 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   4800
      TabIndex        =   64
      Top             =   8760
      Width           =   3975
   End
   Begin VB.Label lblGenerate_No 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   10080
      TabIndex        =   59
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblPrefix_No 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   10080
      TabIndex        =   58
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label30 
      Caption         =   "PARTICULARS OF SCREENING OFFICER"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3720
      TabIndex        =   56
      Top             =   6960
      Width           =   5175
   End
   Begin VB.Label Label29 
      Caption         =   "DONOR'S VIRAL AND NONE VIRAL SCREENING"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   55
      Top             =   5040
      Width           =   6135
   End
   Begin VB.Label Label28 
      Caption         =   "DONOR'S PARTICULARS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4560
      TabIndex        =   54
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label15 
      Caption         =   "Physical Examination:"
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
      Left            =   7080
      TabIndex        =   44
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Previous Donation:"
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
      Left            =   6120
      TabIndex        =   36
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblDonorNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label26 
      Caption         =   "Remarks:"
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
      Left            =   2040
      TabIndex        =   24
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "Screening Date:"
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
      Left            =   6240
      TabIndex        =   23
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label Label24 
      Caption         =   "Rank:"
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
      Left            =   6240
      TabIndex        =   22
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "Screening Officer:"
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
      Left            =   2040
      TabIndex        =   21
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Hemoglobin Estimate:"
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
      Left            =   7080
      TabIndex        =   20
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label12 
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
      Left            =   3240
      TabIndex        =   19
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label14 
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
      Left            =   7080
      TabIndex        =   18
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label11 
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
      Left            =   600
      TabIndex        =   17
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "Syphilis:"
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
      TabIndex        =   16
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "Hepatitis C Virus:"
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
      Left            =   3240
      TabIndex        =   15
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label18 
      Caption         =   "Hepatitis B Surface Anti:"
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
      Left            =   3240
      TabIndex        =   14
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label19 
      Caption         =   "HIV Status:"
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
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Blood Bag Number:"
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
      Left            =   600
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Donation Date:"
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
      Left            =   6120
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label5 
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
      Left            =   600
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Left            =   600
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Left            =   600
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Left            =   6120
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      Caption         =   "Expiration Date:"
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
      Left            =   12480
      TabIndex        =   5
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "Donor Type:"
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
      Left            =   6120
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "Donor Phone No:"
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
      TabIndex        =   3
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label22 
      Caption         =   "Last Donation Date:"
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
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label23 
      Caption         =   "Parent/Spouse Name:"
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
      Left            =   6120
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label27 
      Caption         =   "Adverse Reaction:"
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
      Left            =   6120
      TabIndex        =   0
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "frmDONOR_RECORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BLOOD_BANKDB As Database
Dim rstDonor As Recordset
Dim rstReference As Recordset
Dim rstCheck_No As Recordset
Dim rstVS As Recordset
Dim VarPre_No As String
Dim VarGenerate_No As String
Dim VarCheck_No As String
Private Sub NewRec()
On Error Resume Next
Clearform
GetID
txtFullName.SetFocus
End Sub
Private Sub GetID()
On Error GoTo ErrorTrap
VarPre_No = "DN/"
lblPrefix_No = VarPre_No
VarGenerate_No = VarCheck_No
With rstCheck_No
On Error Resume Next
.MoveLast
VarGenerate_No = Format(1, "000")
lblGenerate_No = VarGenerate_No
lblGenerate_No = Format(CDbl(![Available_No]) + 1, "000")
lblDonorNo = lblPrefix_No & lblGenerate_No
End With
Exit Sub
ErrorTrap:
 MsgBox Err.Description, vbInformation, "Error"
End Sub
Private Sub Clearform()
On Error Resume Next
txtFullName = ""
lblDonorNo = ""
txtPAddress = ""
cboSex = ""
txtAge = ""
txtAdverse_RXN = ""
txtBBN = ""
lblDonorNo = ""
cboPhy_Exam = ""
cboBlood_Group = ""
txtPCV = ""
txtHBE = ""
cboEver_Do = ""
cbo_HIV = ""
cbo_HBSA = ""
cbo_HCV = ""
cbo_Syph = ""
cbo_Rhesus = ""
txtScreen_Off = ""
txt_Rank = ""
txtRemark = ""
mskSceen_Date = ""
mskDonation_Date = ""
mskLast_Do_Date = ""
cboDonorType = ""
mskPhone_No = ""
txtParentName = ""
txtReaction = ""
End Sub
Private Sub EditRec()
On Error Resume Next
With rstDonor
.Edit
![DONOR NAME] = txtFullName
![DONOR NUMBER] = lblDonorNo
!SEX = cboSex
![PARENT/SPOUSE NAME] = txtParentName
!AGE = txtAge
![DONOR TYPE] = cboDonorType
![BLOOD BAG NUMBER] = txtBBN
![DONATION DATE] = mskDonation_Date
![PERMANENT ADDRESS] = txtPAddress
![DONOR PHONE NUMBER] = mskPhone_No
![EVER DONATED?] = cboEver_Do
![LAST DONATION DATE] = mskLast_Do_Date
![ADVERSE REACTION] = txtAdverse_RXN
![HIV STATUS] = cbo_HIV
![SYPHILIS] = cbo_Syph
!HBSA = cbo_HBSA
!HCV = cbo_HCV
![PHYSICAL EXAM] = cboPhy_Exam
![BLOOD GROUP] = cboBlood_Group
!RHESUS = cbo_Rhesus
!BPCV = txtPCV
![HEMOGLOBIN EST] = txtHBE
![SCREENING OFFICER] = txtScreen_Off
!RANK = txt_Rank
!REMARK = txtRemark
![SCREENING DATE] = mskSceen_Date
.Update
.Bookmark = .LastModified
End With
Clearform
End Sub
Private Sub SaveRec()
On Error GoTo ErrorTrap:
With rstDonor
.AddNew
![DONOR NAME] = txtFullName
![DONOR NUMBER] = lblDonorNo
!SEX = cboSex
![PARENT/SPOUSE NAME] = txtParentName
!AGE = txtAge
![DONOR TYPE] = cboDonorType
![BLOOD BAG NUMBER] = txtBBN
![DONATION DATE] = mskDonation_Date
![PERMANENT ADDRESS] = txtPAddress
![DONOR PHONE NUMBER] = mskPhone_No
![EVER DONATED?] = cboEver_Do
![LAST DONATION DATE] = mskLast_Do_Date
![ADVERSE REACTION] = txtAdverse_RXN
![HIV STATUS] = cbo_HIV
![SYPHILIS] = cbo_Syph
!HBSA = cbo_HBSA
!HCV = cbo_HCV
![PHYSICAL EXAM] = cboPhy_Exam
![BLOOD GROUP] = cboBlood_Group
!RHESUS = cbo_Rhesus
!BPCV = txtPCV
![HEMOGLOBIN EST] = txtHBE
![SCREENING OFFICER] = txtScreen_Off
!RANK = txt_Rank
!REMARK = txtRemark
![SCREENING DATE] = mskSceen_Date
.Update
.Bookmark = .LastModified
Clearform
End With
With rstCheck_No
.AddNew
![Available_No] = lblGenerate_No
.Update
.Bookmark = .LastModified
End With
Exit Sub
ErrorTrap:
  MsgBox Err.Description, vbInformation, "Error"
 End Sub
Private Sub FindRecord()
On Error Resume Next
Dim StrSearch As String
StrSearch = InputBox("Enter Donor's Number:", "Find Donor")
On Error Resume Next
With rstDonor
.Index = "DONOR NUMBER"
.Seek "=", StrSearch
If .NoMatch Then
MsgBox "No Record", vbInformation, "Find Donor"
Exit Sub
Else
GetRecord
End If
End With
txtFullName.SetFocus
End Sub
Private Sub GetRecord()
On Error Resume Next
With rstDonor
txtFullName = ![DONOR NAME]
lblDonorNo = ![DONOR NUMBER]
cboSex = !SEX
txtParentName = ![PARENT/SPOUSE NAME]
txtAge = !AGE
cboDonorType = ![DONOR TYPE]
txtBBN = ![BLOOD BAG NUMBER]
mskDonation_Date = ![DONATION DATE]
txtPAddress = ![PERMANENT ADDRESS]
mskPhone_No = ![DONOR PHONE NUMBER]
cboEver_Do = ![EVER DONATED?]
mskLast_Do_Date = ![LAST DONATION DATE]
txtAdverse_RXN = ![ADVERSE REACTION]
cbo_HIV = ![HIV STATUS]
cbo_Syph = ![SYPHILIS]
cbo_HBSA = !HBSA
cbo_HCV = !HCV
cboPhy_Exam = ![PHYSICAL EXAM]
cboBlood_Group = ![BLOOD GROUP]
cbo_Rhesus = !RHESUS
txtPCV = !BPCV
txtHBE = ![HEMOGLOBIN EST]
txtScreen_Off = ![SCREENING OFFICER]
txt_Rank = !RANK
txtRemark = !REMARK
mskSceen_Date = ![SCREENING DATE]
End With
End Sub
Private Sub Delete()
On Error GoTo ErrorTrap:
If MsgBox("Delete Donor Record?", vbQuestion + vbYesNo, "Delete") = vbNo Then
Exit Sub
End If
With rstDonor
.Delete
Clearform
End With
Exit Sub
ErrorTrap:
 MsgBox Err.Description, vbInformation, "Error"
End Sub


Private Sub Form_Load()
Set BLOOD_BANKDB = OpenDatabase(App.Path & "\BLOOD_BANK.mdb", False, False)
Set rstDonor = BLOOD_BANKDB.OpenRecordset("DONOR_RECORD")
Set rstReference = BLOOD_BANKDB.OpenRecordset("DONOR_REFERENCE_DATA")
Set rstCheck_No = BLOOD_BANKDB.OpenRecordset("Used_No")

cboSex.AddItem "Female"
cboSex.AddItem "Male"
cboPhy_Exam.AddItem "Normal"
cboPhy_Exam.AddItem "Not Normal"
cboBlood_Group.AddItem "A"
cboBlood_Group.AddItem "B"
cboBlood_Group.AddItem "AB"
cboBlood_Group.AddItem "O"
cbo_Rhesus.AddItem "Positive"
cbo_Rhesus.AddItem "Negative"
cbo_HIV.AddItem "Positive"
cbo_HIV.AddItem "Negative"
cbo_HBSA.AddItem "Positive"
cbo_HBSA.AddItem "Negative"
cbo_HCV.AddItem "Positive"
cbo_HCV.AddItem "Negative"
cbo_Syph.AddItem "Positive"
cbo_Syph.AddItem "Negative"
cboDonorType.AddItem "Commercial"
cboDonorType.AddItem "Replacement"
cboDonorType.AddItem "Voluntary"
cboEver_Do.AddItem "No"
cboEver_Do.AddItem "Yes"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        
        Case "New"
                NewRec
        
                    Case "Save"
                            SaveRec
        
                                Case "Find"
                                    FindRecord
            
                    Case "Button"
                            EditRec
        Case "Delete"
            Delete
    End Select
End Sub

Private Sub cmdFirstRecord_Click()
On Error Resume Next
With rstDonor
.MoveFirst
GetRecord
End With
End Sub

Private Sub cmdLastRecord_Click()
On Error Resume Next
With rstDonor
.MoveLast
GetRecord
End With
End Sub
Private Sub cmdNext_Click()
On Error Resume Next
With rstDonor
.MoveNext
If .EOF Then
 .MoveLast
End If
GetRecord
End With
End Sub

Private Sub cmdPrevious_Click()
On Error Resume Next
With rstDonor
.MovePrevious
If .BOF Then
 .MoveFirst
End If
GetRecord
End With
End Sub
