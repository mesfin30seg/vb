VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRECIPIENT_RECORD 
   Caption         =   "Blood Recipient Details"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Print All Report"
      Height          =   495
      Left            =   6840
      TabIndex        =   80
      Top             =   9360
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   120
      Top             =   7680
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BLOOD BANK PROJECT\BLOOD_BANK.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BLOOD BANK PROJECT\BLOOD_BANK.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select* from Recipient_Record"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BLOOD BANK PROJECT\BLOOD_BANK.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BLOOD BANK PROJECT\BLOOD_BANK.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Recipient_Record"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Individual Report"
      Height          =   495
      Left            =   3360
      TabIndex        =   79
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton cmdBloodValidity 
      BackColor       =   &H80000009&
      Caption         =   "&Blood Validity"
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
      Left            =   9000
      TabIndex        =   75
      ToolTipText     =   "Check for Validity"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdFirstRecord 
      Caption         =   "|<"
      Height          =   375
      Left            =   3240
      TabIndex        =   73
      ToolTipText     =   "Go to First Record"
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   375
      Left            =   3600
      TabIndex        =   72
      ToolTipText     =   "Goto Previous Record"
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   7920
      TabIndex        =   71
      ToolTipText     =   "Goto Next Record"
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton cmdLastRecord 
      Caption         =   ">|"
      Height          =   375
      Left            =   8280
      TabIndex        =   70
      ToolTipText     =   "Goto Last Record"
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton cmdMatch 
      BackColor       =   &H80000009&
      Caption         =   "&Cross Match"
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
      Left            =   9000
      TabIndex        =   67
      ToolTipText     =   "Click To Cross Match"
      Top             =   6000
      Width           =   1695
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
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
            Object.ToolTipText     =   "Add New Record"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Record"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find Record"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Button"
            Object.ToolTipText     =   "Save Edited "
            ImageKey        =   "Button"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Record"
            ImageKey        =   "Delete"
         EndProperty
      EndProperty
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
      Height          =   735
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   7800
      Width           =   2055
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
      Left            =   7200
      TabIndex        =   34
      Top             =   7320
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
      Left            =   3360
      TabIndex        =   33
      Top             =   7320
      Width           =   1935
   End
   Begin VB.ListBox lstBagNo 
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
      Height          =   525
      ItemData        =   "frm_Recipient.frx":0000
      Left            =   3360
      List            =   "frm_Recipient.frx":0002
      TabIndex        =   30
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txt_BloodGroup 
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
      Left            =   3360
      TabIndex        =   29
      Top             =   4680
      Width           =   1695
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
      Left            =   3360
      TabIndex        =   27
      Top             =   3480
      Width           =   1095
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
      Height          =   315
      Left            =   3360
      TabIndex        =   26
      Top             =   3000
      Width           =   1095
   End
   Begin VB.ComboBox cbo_Rhesus 
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
      Left            =   7440
      TabIndex        =   25
      Top             =   3120
      Width           =   1335
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
      Left            =   3360
      TabIndex        =   24
      Top             =   2520
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskReceive_Date 
      Height          =   255
      Left            =   7440
      TabIndex        =   23
      Top             =   2640
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
      Left            =   3360
      TabIndex        =   22
      Top             =   2040
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
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   1560
      Width           =   2175
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
      Left            =   3360
      TabIndex        =   20
      Top             =   1560
      Width           =   1335
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
      Height          =   285
      Left            =   7440
      TabIndex        =   19
      Top             =   1080
      Width           =   2895
   End
   Begin MSMask.MaskEdBox mskScreen_Date 
      Height          =   255
      Left            =   7200
      TabIndex        =   37
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
      Left            =   240
      Top             =   6600
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
            Picture         =   "frm_Recipient.frx":0004
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Recipient.frx":0116
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Recipient.frx":0228
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Recipient.frx":033A
            Key             =   "Button"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Recipient.frx":044C
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label17 
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
      Left            =   4440
      TabIndex        =   82
      Top             =   3480
      Width           =   255
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
      Left            =   4440
      TabIndex        =   81
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblDayV1 
      Height          =   375
      Left            =   -360
      TabIndex        =   78
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDayV2 
      Height          =   375
      Left            =   -360
      TabIndex        =   77
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDay 
      Height          =   375
      Left            =   -360
      TabIndex        =   76
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label31 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   3960
      TabIndex        =   74
      Top             =   8760
      Width           =   3975
   End
   Begin VB.Label lblPrefix_No 
      Height          =   255
      Left            =   360
      TabIndex        =   69
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblGenerate_No 
      Height          =   255
      Left            =   0
      TabIndex        =   68
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblrxn 
      BackColor       =   &H80000009&
      Caption         =   "Label18"
      Height          =   255
      Left            =   13320
      TabIndex        =   66
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbldonorphone 
      BackColor       =   &H80000009&
      Caption         =   "Label17"
      Height          =   255
      Left            =   13440
      TabIndex        =   65
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblbpcv 
      BackColor       =   &H80000009&
      Caption         =   "Label17"
      Height          =   255
      Left            =   13440
      TabIndex        =   64
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblhemoest 
      BackColor       =   &H80000009&
      Caption         =   "Label18"
      Height          =   255
      Left            =   13440
      TabIndex        =   63
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblscrenoff 
      BackColor       =   &H80000009&
      Caption         =   "Label19"
      Height          =   255
      Left            =   13320
      TabIndex        =   62
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblrank 
      BackColor       =   &H80000009&
      Caption         =   "Label24"
      Height          =   255
      Left            =   13320
      TabIndex        =   61
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblremark 
      BackColor       =   &H80000009&
      Caption         =   "Label25"
      Height          =   255
      Left            =   13440
      TabIndex        =   60
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblscrendate 
      BackColor       =   &H80000009&
      Caption         =   "Label26"
      Height          =   255
      Left            =   13320
      TabIndex        =   59
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbllastdonordate 
      BackColor       =   &H80000009&
      Caption         =   "Label27"
      Height          =   255
      Left            =   13320
      TabIndex        =   58
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblDonorname 
      BackColor       =   &H80000009&
      Caption         =   "Label18"
      Height          =   255
      Left            =   11760
      TabIndex        =   57
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblsex 
      BackColor       =   &H80000009&
      Caption         =   "Label19"
      Height          =   255
      Left            =   11760
      TabIndex        =   56
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblparent 
      BackColor       =   &H80000009&
      Caption         =   "Label24"
      Height          =   255
      Left            =   11760
      TabIndex        =   55
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblAge 
      BackColor       =   &H80000009&
      Caption         =   "Label25"
      Height          =   255
      Left            =   11880
      TabIndex        =   54
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbldonortype 
      BackColor       =   &H80000009&
      Caption         =   "Label26"
      Height          =   255
      Left            =   11880
      TabIndex        =   53
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblBBN 
      BackColor       =   &H80000009&
      Caption         =   "Label27"
      Height          =   255
      Left            =   11880
      TabIndex        =   52
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblphysicalexam 
      BackColor       =   &H80000009&
      Caption         =   "Label36"
      Height          =   255
      Left            =   13080
      TabIndex        =   51
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblbloodgroup 
      BackColor       =   &H80000009&
      Caption         =   "Label37"
      Height          =   255
      Left            =   13920
      TabIndex        =   50
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblhcv 
      BackColor       =   &H80000009&
      Caption         =   "Label 35"
      Height          =   255
      Left            =   13080
      TabIndex        =   49
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbldonordate 
      BackColor       =   &H80000009&
      Caption         =   "Label38"
      Height          =   255
      Left            =   11880
      TabIndex        =   48
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblpaddress 
      BackColor       =   &H80000009&
      Caption         =   "Label39"
      Height          =   255
      Left            =   11880
      TabIndex        =   47
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbleverdonate 
      BackColor       =   &H80000009&
      Caption         =   "Label41"
      Height          =   255
      Left            =   11880
      TabIndex        =   46
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblhiv 
      BackColor       =   &H80000009&
      Caption         =   "Label42"
      Height          =   255
      Left            =   11880
      TabIndex        =   45
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblsyphilis 
      BackColor       =   &H80000009&
      Caption         =   "Label43"
      Height          =   255
      Left            =   12000
      TabIndex        =   44
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblhbsa 
      BackColor       =   &H80000009&
      Caption         =   "Label44"
      Height          =   255
      Left            =   11880
      TabIndex        =   43
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label43 
      Caption         =   "Label43"
      Height          =   375
      Left            =   -3360
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label30 
      Caption         =   "PARTICULARS OF SCREENING OFFICER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Top             =   6720
      Width           =   5295
   End
   Begin VB.Label Label29 
      Caption         =   "RECIPIENT CROSS MATCHING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Label Label28 
      Caption         =   "RECIPIENT PARTICULARS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3960
      TabIndex        =   39
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label8 
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
      Left            =   5880
      TabIndex        =   36
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label lblExp_Date 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Left            =   12360
      TabIndex        =   32
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbl_Rhesus 
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
      Left            =   7440
      TabIndex        =   31
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label lblDonor_No 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label lblRecipientNumber 
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
      Left            =   3480
      TabIndex        =   18
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label22 
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
      Left            =   1800
      TabIndex        =   17
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label21 
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
      Left            =   5880
      TabIndex        =   16
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label20 
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
      Left            =   1800
      TabIndex        =   15
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Left            =   1080
      TabIndex        =   14
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label6 
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
      Left            =   1080
      TabIndex        =   13
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000009&
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
      Left            =   12000
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label16 
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
      Left            =   5640
      TabIndex        =   11
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label9 
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
      Left            =   1080
      TabIndex        =   10
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Left            =   1080
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label4 
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
      Left            =   1080
      TabIndex        =   8
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Left            =   1080
      TabIndex        =   7
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   5640
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label10 
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
      Left            =   5640
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label11 
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
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label12 
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
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label13 
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
      Left            =   5520
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label14 
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
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label15 
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
      Left            =   5520
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
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

Private Sub cmdBloodValidity_Click()
On Error Resume Next
Dim sDate As String
      Dim intNumDays As Integer
      Dim intNow As Integer
      Dim Enterdate As Integer
      Dim DayBalance As Integer
            sDate = lblExp_Date
           Enterdate = CInt(DateValue(sDate) - Now())
           intNumDays = Enterdate
           lblDayV1 = intNumDays * (-1)
           lblDayV2 = 35 - Val(lblDayV1)
           DayBalance = lblDayV2
           
        If lblDayV2 > 35 Then
            MsgBox "Wrong Donation Date, Please Check", vbInformation = vbOKOnly
                   Exit Sub
                    End If
        If lblDayV1 >= 35 Then
            MsgBox "Blood is Expired"
            With rstDonor
            .Delete
            Clearform
            End With
         Else
         MsgBox "Blood is Still Healthy:" & "Remaining Days is:" _
         & vbCrLf & DayBalance, vbInformation
         
   End If
  
End Sub

Private Sub cmdMatch_Click()
'On Error Resume Next
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
'![Exp Date] = lblExp_Date
![Bag Number] = lstBagNo.Text
![PHYSICAL EXAM] = lblphysicalexam
![BLOOD GROUP] = lblBloodGroup
!RHESUS = lbl_Rhesus
![Blood Pack Cell Volume] = lblBPCV
![Hemoglobin Estimate] = lblhemoest
!HIV = lblHIV
![Hepatitis B Surface Antigen] = lblHBSA
![Hepatitis C Virus] = lblHCV
![SYPHILIS] = lblSyphilis
![Officer's Name] = lblscrenoff
!RANK = lblRank
!REMARK = lblRemark
![SCREENING DATE] = lblscrendate
![DONOR TYPE] = lbldonortype
![Donor Phone] = lbldonorphone
![Parent Name] = lblParent
![LAST DONATION DATE] = lbllastdonordate
![ADVERSE REACTION] = lblrxn
!SEX = lblsex
!AGE = lblAge
![EVER DONATES] = lbleverdonate
.Update
.Bookmark = .LastModified
End With
With rstDonor
.Delete
End With
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
With rstRecipient
.MoveNext
LoadData
End With
End Sub

Private Sub cmdPrevious_Click()
On Error Resume Next
With rstRecipient
.MovePrevious
LoadData
End With
End Sub

Private Sub Command1_Click()
Dim Slip As String
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Set db = New ADODB.Connection
db.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data source=" & App.Path & "\BLOOD_BANK.mdb"
Slip = InputBox("Enter Donor's No. To Print")
Set rs = New ADODB.Recordset
rs.Open "Select * From RECIPIENT_RECORD Where [Recipient's Number]=" & "'" & Slip & "'", db, adOpenDynamic, adLockOptimistic
Set drp_Recipient.DataSource = rs
drp_Recipient.Refresh
drp_Recipient.Show vbModal
End Sub

Private Sub Command2_Click()
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Set db = New ADODB.Connection
db.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data source=" & App.Path & "\BLOOD_BANK.mdb"
Set rs = New ADODB.Recordset
rs.Open "Select * From RECIPIENT_RECORD order by [Recipient's Number]", db, adOpenStatic, adLockReadOnly
Set drp_Recipient.DataSource = rs
drp_Recipient.Refresh
drp_Recipient.Show vbModal
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
lblExp_Date = ![DONATION DATE]
lbl_Rhesus = ![RHESUS]
lblDonorname = ![DONOR NAME]
lblpaddress = ![PERMANENT ADDRESS]
lblsex = !SEX
lblAge = !AGE
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
lblParent = ![PARENT/SPOUSE NAME]
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
            cmdFind_Record
        Case "Button"
            cmdSave_Edit
            
        Case "Delete"
            Delete
    End Select
End Sub
Private Sub Delete()
On Error GoTo ErrorTrap:
If MsgBox("Delete Recipient Record?", vbQuestion + vbYesNo, "Delete Record") = vbNo Then
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

Private Sub NewRecord()
On Error Resume Next
Clearform
VarPre_No = "RC/"
lblPrefix_No = VarPre_No
With rstRecipient
.MoveLast
VarGenerate_No = Format(1, "000")
lblGenerate_No = VarGenerate_No
lblGenerate_No = Format(CDbl(![Prefix No]) + 1, "000")
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
!RANK = txt_Rank
!REMARK = txtRemark
![SCREENING DATE] = mskScreen_Date
.Update
.Bookmark = .LastModified
Clearform
End With
Exit Sub
ErrorTrap:
MsgBox Err.Description, vbInformation, "Error"
End Sub

Private Sub Form_Load()
Set BLOOD_BANKDB = OpenDatabase(App.Path & "\BLOOD_BANK.mdb", False, False)
Set rstRecipient = BLOOD_BANKDB.OpenRecordset("RECIPIENT_RECORD")
Set rstDonor = BLOOD_BANKDB.OpenRecordset("DONOR_RECORD")
Set rstReference = BLOOD_BANKDB.OpenRecordset("DONOR_REFERENCE_DATA")
cboSex.AddItem "Female"
cboSex.AddItem "Male"
cboBlood_Group.AddItem "A"
cboBlood_Group.AddItem "B"
cboBlood_Group.AddItem "AB"
cboBlood_Group.AddItem "O"
cbo_Rhesus.AddItem "Positive"
cbo_Rhesus.AddItem "Negative"
lblDay.Caption = Format(Date, "mm:dd:yy")
End Sub

Private Sub cmdFind_Record()
On Error Resume Next
Dim StrSearch As String
StrSearch = InputBox("Enter Recipient's Number:", "Find Donor")
lstBagNo.Visible = False
lblBag_No.Visible = True
With rstRecipient
.Index = "Recipient's Number"
.Seek "=", StrSearch
If .NoMatch Then
MsgBox "No Record", vbInformation, "Find Recipient"
Exit Sub
Else
On Error GoTo ErrorTrap:
With rstRecipient
txtFullName = ![Full Name]
lblRecipientNumber = ![Recipient's Number]
txtPAddress = ![PERMANENT ADDRESS]
cboSex = !SEX
txtAge = !AGE
mskReceive_Date = !Date
lblGenerate_No = ![Prefix No]
cboBlood_Group = ![BLOOD GROUP]
cbo_Rhesus = ![RHESUS]
txtPCV = !PCV
txtHBE = ![HEMOGLOBIN EST]
txt_BloodGroup = ![Enter Group]
lblBag_No = ![Bag No]
lblDonor_No = ![Donor's No]
lblExp_Date = ![Exp Date]
lbl_Rhesus = ![Rhesus 2]
txtScreen_Off = ![SCREENING OFFICER]
txt_Rank = !RANK
txtRemark = !REMARK
mskScreen_Date = ![SCREENING DATE]
End With
Exit Sub
ErrorTrap:
End If
End With
End Sub
Private Sub LoadData()
With rstRecipient
txtFullName = ![Full Name]
lblRecipientNumber = ![Recipient's Number]
txtPAddress = ![PERMANENT ADDRESS]
cboSex = !SEX
txtAge = !AGE
mskReceive_Date = !Date
lblGenerate_No = ![Prefix No]
cboBlood_Group = ![BLOOD GROUP]
cbo_Rhesus = ![RHESUS]
txtPCV = !PCV
txtHBE = ![HEMOGLOBIN EST]
txt_BloodGroup = ![Enter Group]
lblBag_No = ![Bag No]
lblDonor_No = ![Donor's No]
lblExp_Date = ![Exp Date]
lbl_Rhesus = ![Rhesus 2]
txtScreen_Off = ![SCREENING OFFICER]
txt_Rank = !RANK
txtRemark = !REMARK
mskScreen_Date = ![SCREENING DATE]
End With
End Sub

Private Sub Clearform()
On Error Resume Next
txtFullName = ""
lblRecipientNumber = ""
txtPAddress = ""
cboSex = ""
txtAge = ""
mskReceive_Date = ""
txtBBN = ""
cboBlood_Group = ""
cbo_Rhesus = ""
txtPCV = ""
txtHBE = ""
txt_BloodGroup = ""
lstBagNo.Text = ""
lblBag_No = ""
lblDonor_No = ""
lblExp_Date = ""
lbl_Rhesus = ""
txtScreen_Off = ""
txt_Rank = ""
txtRemark = ""
mskScreen_Date = ""
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
Private Sub cmdSave_Edit()
On Error GoTo ErrorTrap:
With rstRecipient
'.MoveLast
.Edit
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
![Bag No] = lblBag_No
![Donor's No] = lblDonor_No
![Exp Date] = lblExp_Date
![Rhesus 2] = lbl_Rhesus
![SCREENING OFFICER] = txtScreen_Off
!RANK = txt_Rank
!REMARK = txtRemark
![SCREENING DATE] = mskScreen_Date
.Update
.Bookmark = .LastModified
Clearform
End With
Exit Sub
ErrorTrap:
MsgBox Err.Description, vbInformation, "Error"
End Sub

'Private Sub txtScreen_Off_Change()
'If MsgBox("Have You Cross Matched?", vbInformation + vbYesNo, "Cross Match") = vbNo Then
'Exit Sub
'End If
'Else
'txtScreen_Off_Change
'End Sub

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
