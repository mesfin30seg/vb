VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRECIPIENT_RECORD 
   Caption         =   "None Viral Screen"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   14115
   WindowState     =   2  'Maximized
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
      Height          =   495
      Left            =   8160
      TabIndex        =   67
      Top             =   6000
      Width           =   735
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   14115
      _ExtentX        =   24897
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
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Button"
            Object.ToolTipText     =   "Button"
            ImageKey        =   "Button"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
      EndProperty
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
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   7920
      Width           =   2055
   End
   Begin VB.TextBox txt_Rank 
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
      Height          =   285
      Left            =   7800
      TabIndex        =   34
      Top             =   7440
      Width           =   1455
   End
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
      Height          =   285
      Left            =   3000
      TabIndex        =   33
      Top             =   7440
      Width           =   1935
   End
   Begin VB.ListBox lstBagNo 
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
      Height          =   615
      ItemData        =   "frmNon_Viral_Screen.frx":0000
      Left            =   3120
      List            =   "frmNon_Viral_Screen.frx":0002
      TabIndex        =   30
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txt_BloodGroup 
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
      Height          =   285
      Left            =   3120
      TabIndex        =   29
      Top             =   4920
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
      Height          =   285
      Left            =   7800
      TabIndex        =   27
      Top             =   3600
      Width           =   1335
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
      Height          =   315
      Left            =   3480
      TabIndex        =   26
      Top             =   3600
      Width           =   1335
   End
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
      Left            =   7800
      TabIndex        =   25
      Top             =   3120
      Width           =   1335
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
      Left            =   3480
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskReceive_Date 
      Height          =   255
      Left            =   7800
      TabIndex        =   23
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
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
      Height          =   285
      Left            =   3480
      TabIndex        =   22
      Top             =   2640
      Width           =   1335
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
      Left            =   7800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   1680
      Width           =   2175
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
      Left            =   3480
      TabIndex        =   20
      Top             =   2040
      Width           =   1335
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
      Height          =   285
      Left            =   7800
      TabIndex        =   19
      Top             =   1200
      Width           =   2895
   End
   Begin MSMask.MaskEdBox mskScreen_Date 
      Height          =   255
      Left            =   7800
      TabIndex        =   37
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Left            =   12840
      Top             =   6240
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
            Picture         =   "frmNon_Viral_Screen.frx":0004
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNon_Viral_Screen.frx":0116
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNon_Viral_Screen.frx":0228
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNon_Viral_Screen.frx":033A
            Key             =   "Button"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNon_Viral_Screen.frx":044C
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblrxn 
      BackColor       =   &H80000009&
      Caption         =   "Label18"
      Height          =   255
      Left            =   12600
      TabIndex        =   66
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lbldonorphone 
      BackColor       =   &H80000009&
      Caption         =   "Label17"
      Height          =   255
      Left            =   12600
      TabIndex        =   65
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblbpcv 
      BackColor       =   &H80000009&
      Caption         =   "Label17"
      Height          =   375
      Left            =   12600
      TabIndex        =   64
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblhemoest 
      BackColor       =   &H80000009&
      Caption         =   "Label18"
      Height          =   255
      Left            =   12600
      TabIndex        =   63
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblscrenoff 
      BackColor       =   &H80000009&
      Caption         =   "Label19"
      Height          =   255
      Left            =   12600
      TabIndex        =   62
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblrank 
      BackColor       =   &H80000009&
      Caption         =   "Label24"
      Height          =   255
      Left            =   12600
      TabIndex        =   61
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblremark 
      BackColor       =   &H80000009&
      Caption         =   "Label25"
      Height          =   255
      Left            =   12600
      TabIndex        =   60
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblscrendate 
      BackColor       =   &H80000009&
      Caption         =   "Label26"
      Height          =   255
      Left            =   12600
      TabIndex        =   59
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lbllastdonordate 
      BackColor       =   &H80000009&
      Caption         =   "Label27"
      Height          =   255
      Left            =   12600
      TabIndex        =   58
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblDonorname 
      BackColor       =   &H80000009&
      Caption         =   "Label18"
      Height          =   375
      Left            =   11280
      TabIndex        =   57
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblsex 
      BackColor       =   &H80000009&
      Caption         =   "Label19"
      Height          =   375
      Left            =   11280
      TabIndex        =   56
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblparent 
      BackColor       =   &H80000009&
      Caption         =   "Label24"
      Height          =   375
      Left            =   11280
      TabIndex        =   55
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblage 
      BackColor       =   &H80000009&
      Caption         =   "Label25"
      Height          =   375
      Left            =   11280
      TabIndex        =   54
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lbldonortype 
      BackColor       =   &H80000009&
      Caption         =   "Label26"
      Height          =   255
      Left            =   11280
      TabIndex        =   53
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblBBN 
      BackColor       =   &H80000009&
      Caption         =   "Label27"
      Height          =   255
      Left            =   11280
      TabIndex        =   52
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblphysicalexam 
      BackColor       =   &H80000009&
      Caption         =   "Label36"
      Height          =   375
      Left            =   12600
      TabIndex        =   51
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblbloodgroup 
      BackColor       =   &H80000009&
      Caption         =   "Label37"
      Height          =   375
      Left            =   12600
      TabIndex        =   50
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblhcv 
      BackColor       =   &H80000009&
      Caption         =   "Label 35"
      Height          =   375
      Left            =   12600
      TabIndex        =   49
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbldonordate 
      BackColor       =   &H80000009&
      Caption         =   "Label38"
      Height          =   255
      Left            =   11280
      TabIndex        =   48
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblpaddress 
      BackColor       =   &H80000009&
      Caption         =   "Label39"
      Height          =   255
      Left            =   11280
      TabIndex        =   47
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lbleverdonate 
      BackColor       =   &H80000009&
      Caption         =   "Label41"
      Height          =   255
      Left            =   11280
      TabIndex        =   46
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblhiv 
      BackColor       =   &H80000009&
      Caption         =   "Label42"
      Height          =   255
      Left            =   11280
      TabIndex        =   45
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblsyphilis 
      BackColor       =   &H80000009&
      Caption         =   "Label43"
      Height          =   255
      Left            =   11280
      TabIndex        =   44
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblhbsa 
      BackColor       =   &H80000009&
      Caption         =   "Label44"
      Height          =   255
      Left            =   11280
      TabIndex        =   43
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label43 
      Caption         =   "Label43"
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label30 
      BackColor       =   &H80000009&
      Caption         =   "PARTICULARS OF SCREENING OFFICER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   3360
      TabIndex        =   41
      Top             =   6840
      Width           =   6015
   End
   Begin VB.Label Label29 
      BackColor       =   &H80000009&
      Caption         =   "RECIPIENT CROSS MATCHING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   40
      Top             =   4320
      Width           =   4695
   End
   Begin VB.Label Label28 
      BackColor       =   &H80000009&
      Caption         =   "RECIPIENT PARTICULAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3960
      TabIndex        =   39
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000B&
      Caption         =   "Screen Date:"
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
      Left            =   5760
      TabIndex        =   36
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label lblExp_Date 
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
      Height          =   255
      Left            =   7800
      TabIndex        =   32
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lbl_Rhesus 
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
      Height          =   255
      Left            =   7800
      TabIndex        =   31
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblDonor_No 
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
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   6240
      Width           =   1695
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
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   1200
      Width           =   1815
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
      Left            =   1200
      TabIndex        =   17
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000B&
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
      Left            =   5760
      TabIndex        =   16
      Top             =   7440
      Width           =   1215
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
      Left            =   1200
      TabIndex        =   15
      Top             =   7440
      Width           =   1455
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
      Left            =   1200
      TabIndex        =   14
      Top             =   5520
      Width           =   1335
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
      Left            =   1200
      TabIndex        =   13
      Top             =   6240
      Width           =   1575
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
      Left            =   5760
      TabIndex        =   12
      Top             =   5520
      Width           =   1575
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
      Left            =   5760
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
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
      Left            =   1200
      TabIndex        =   10
      Top             =   4920
      Width           =   1695
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
      Left            =   1200
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
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
      Left            =   1200
      TabIndex        =   8
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
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
      Left            =   5760
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
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
      Left            =   5760
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000B&
      Caption         =   "Date Received:"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
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
      Left            =   2640
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
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
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
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
      Left            =   5760
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
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
      Left            =   5760
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmRECIPIENT_RECORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BLOOD_BANKDB As Database
Dim rstRecipient As Recordset
Dim rstDonor As Recordset
Dim rstReference As Recordset
Dim DeRef As ConnectionDesigner
Dim VarRecipientNumber As String
Dim VarPre_No As String
Dim VarGenerate_No As String

Dim Var_BBN As String
Dim Var_BloodGroup As String
Dim Var_Rhesus As String

Private Sub cmdMatch_Click()
On Error Resume Next
If MsgBox("Is Cross Matching Ok? Please Check!!", vbQuestion + vbYesNo, "Cross Match") = vbNo Then
Exit Sub
End If
With rstReference
'.MoveLast
.AddNew
![Full Name] = lblDonorname
![Donor's Number] = lblDonor_No
![PERMANENT ADDRESS] = lblpaddress
!Date = lbldonordate
![Exp Date] = lblExp_Date
![Bag Number] = lstBagNo.Text
![PHYSICAL EXAM] = lblphysicalexam
![BLOOD GROUP] = lblBloodGroup
!RHESUS = lbl_Rhesus
![Blood Pack Cell Volume] = Label25
![Hemoglobin Estimate] = Label26
!HIV = Label27
![Hepatitis B Surface Antigen] = Label28
![Hepatitis C Virus] = Label29
![SYPHILIS] = Label30
![Officer's Name] = Label31
!Designition = Label32
!REMARK = Label33
![SCREENING DATE] = Label34
![DONOR TYPE] = Label38
![Donor Phone] = Label39
![Parent Name] = Label40
![LAST DONATION DATE] = Label41
![ADVERSE REACTION] = Label42
!SEX = Label43
!AGE = Label44
![EVER DONATES] = lbleverdonate
.Update
.Bookmark = .LastModified
End With
With rstDonor
.Delete
End With
End Sub

Private Sub lstBagNo_Click()
On Error Resume Next
lblDonor_No = ""
lblExp_Date = ""
lbl_Rhesus = ""
Var_BBN = lstBagNo.Text
Set rstDonor = BLOOD_BANKDB.OpenRecordset(" Select * from DONOR_RECORD where [BLOOD GROUP]= '" & Var_BloodGroup & "' And [BLOOD BAG NUMBER] = '" & Var_BBN & "';", dbOpenDynaset)
With rstDonor
lblDonor_No = ![DONOR NUMBER]
lblExp_Date = ![EXPIRATION DATE]
lbl_Rhesus = ![RHESUS]
lblDonorname = ![DONOR NAME]
lblpaddress = ![PERMANENT ADDRESS]
lblsex = !SEX
lblage = !AGE
lbldonordate = ![DONATION DATE]
lblBBN = ![BLOOD BAG NUMBER]
lblphysicalexam = ![PHYSICAL EXAM]
lblBloodGroup = ![BLOOD GROUP]
lblBPCV = ![BPCV]
lblhemoest = ![HEMOGLOBIN EST]
lblHIV = ![HIV STATUS]
lblHBSA = !HBSA
lblHCV = !HCV
lblSyphilis = !SYPHILIS
lblscrenoff = ![SCREENING OFFICER]
lblRank = !RANK
lblRemark = !REMARK
lblscrendate = ![SCREENING DATE]
lbldonortype = ![DONOR TYPE]
lbldonorphone = ![DONOR PHONE NUMBER]
lblparent = ![PARENT/SPOUSE NAME]
lbllastdonordate = ![LAST DONATION DATE]
lblrxn = ![ADVERSE REACTION]
lbleverdonate = ![EVER DONATED?]
End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            NewRecord
        Case "Save"
            SaveRecord
        Case "Find"
            'ToDo: Add 'Find' button code.
            MsgBox "Add 'Find' button code."
        Case "Button"
            'ToDo: Add 'Button' button code.
            MsgBox "Add 'Button' button code."
        Case "Delete"
            'ToDo: Add 'Delete' button code.
            MsgBox "Add 'Delete' button code."
    End Select
End Sub

Private Sub NewRecord()
On Error Resume Next
'Clearform
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
Private Sub SaveRecord()
On Error GoTo ErrorTrap:
With rstRecipient
'.MoveLast
.AddNew
![Full Name] = txtFullName
![Recipient's Number] = lblRecipientNumber
![PERMANENT ADDRESS] = txtPAddress
!SEX = cboSex
!AGE = txtAge
!Date = mskReceive_Date
![Prefix No] = lblGenerate_No
![BLOOD GROUP] = cboBlood_Group
![RHESUS] = cbo_Rhesus
!PCV = txtPCV
![HEMOGLOBIN EST] = txtHBE
![Enter Group] = txt_BloodGroup
![Bag No] = lstBagNo.Text
![Donor's No] = lblDonor_No
![Exp Date] = lblExp_Date
![Rhesus 2] = lbl_Rhesus
![SCREENING OFFICER] = txtScreen_Off
!RANK = txtDesig
!REMARK = txtRemark
![SCREENING DATE] = mskScreen_Date
.Update
.Bookmark = .LastModified
'Clearform
End With
Exit Sub
ErrorTrap:
MsgBox Err.Description, vbInformation, "Error"
End Sub

Private Sub Form_Load()
Set BLOOD_BANKDB = OpenDatabase(App.Path & "\BLOOD_BANK.mdb", False, False)
Set rstRecipient = BLOOD_BANKDB.OpenRecordset("Recipient")
Set rstDonor = BLOOD_BANKDB.OpenRecordset("DONOR_RECORD")
Set rstReference = BLOOD_BANKDB.OpenRecordset("Reference Data")
cboSex.AddItem "Female"
cboSex.AddItem "Male"
cboBlood_Group.AddItem "A"
cboBlood_Group.AddItem "B"
cboBlood_Group.AddItem "AB"
cboBlood_Group.AddItem "O"
cbo_Rhesus.AddItem "Positive"
cbo_Rhesus.AddItem "Negative"
End Sub

Private Sub txt_BloodGroup_Change()
Var_BloodGroup = txt_BloodGroup
Set rstDonor = BLOOD_BANKDB.OpenRecordset(" Select * from DONOR_RECORD where [BLOOD GROUP] = '" & Var_BloodGroup & "';", dbOpenDynaset)
With rstDonor
lstBagNo.Clear
While Not .EOF
lstBagNo.AddItem ![BLOOD BAG NUMBER]
.MoveNext
Wend
End With
End Sub
