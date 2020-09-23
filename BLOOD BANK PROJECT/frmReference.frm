VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmREFERENCE_RECORD 
   BackColor       =   &H80000009&
   Caption         =   "Records of Donor"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT INDIVIDUAL REPORT"
      Height          =   375
      Left            =   5760
      TabIndex        =   63
      Top             =   8400
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   10920
      Top             =   5520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "Select * from DONOR_REFERENCE_DATA"
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
      Height          =   330
      Left            =   10920
      Top             =   5880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "Select * from DONOR_REFERENCE_DATA"
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
      Caption         =   "PRINT ALL REPORT"
      Height          =   375
      Left            =   8400
      TabIndex        =   62
      Top             =   8400
      Width           =   2415
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find Donor Record"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Record"
            ImageKey        =   "Delete"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra_Donor 
      Appearance      =   0  'Flat
      Caption         =   "Transfused Blood Record"
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
      Height          =   7455
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   10455
      Begin VB.Frame fraMenu 
         Caption         =   "Donors/Recipients Personal Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   -7440
         TabIndex        =   5
         Top             =   0
         Width           =   6735
         Begin VB.CommandButton cmdRecipient 
            Caption         =   "Recipient's Data"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   1200
            TabIndex        =   7
            Top             =   3120
            Width           =   4575
         End
         Begin VB.CommandButton cmdDonor 
            Caption         =   "Donor's Data"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   1200
            TabIndex        =   6
            Top             =   1680
            Width           =   4575
         End
      End
      Begin VB.Label Label8 
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
         Left            =   8640
         TabIndex        =   65
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label Label4 
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
         Left            =   8640
         TabIndex        =   64
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label lblPreviousD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   60
         Top             =   5400
         Width           =   2295
      End
      Begin VB.Label Label36 
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
         Left            =   480
         TabIndex        =   59
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label lbl_Sex 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   58
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lbl_AGE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   57
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label33 
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
         Height          =   255
         Left            =   480
         TabIndex        =   56
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label32 
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
         Height          =   255
         Left            =   480
         TabIndex        =   55
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblAdverseRXN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   7920
         TabIndex        =   54
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblLastDon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   53
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblParent 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   52
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label lblDon_Phone_No 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   51
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label lblDonType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   50
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label23 
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
         Left            =   5640
         TabIndex        =   49
         Top             =   1080
         Width           =   1695
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
         Height          =   375
         Left            =   5640
         TabIndex        =   48
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "Parent or Husband Name:"
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
         Left            =   480
         TabIndex        =   47
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label20 
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
         Left            =   480
         TabIndex        =   46
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   45
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lblScreenDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   44
         Top             =   6840
         Width           =   1935
      End
      Begin VB.Label lblRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   43
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label lblRank 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   42
         Top             =   6840
         Width           =   2295
      End
      Begin VB.Label lblScreenOff 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   41
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Label lblName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   40
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblHEST 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   39
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label lblBPCV 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   38
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label lblAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   2520
         TabIndex        =   37
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lblRhesus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   36
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label lblBloodGroup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   35
         Top             =   4080
         Width           =   1935
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
         Left            =   5640
         TabIndex        =   34
         Top             =   4440
         Width           =   1695
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
         Left            =   5640
         TabIndex        =   33
         Top             =   3720
         Width           =   1935
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
         Left            =   5640
         TabIndex        =   32
         Top             =   4080
         Width           =   1575
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
         Left            =   5640
         TabIndex        =   31
         Top             =   4800
         Width           =   2175
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
         Left            =   5640
         TabIndex        =   30
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label lblPhyExam 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   29
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblSyphilis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   28
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label lblHCV 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   27
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label lblHBSA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   26
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "HIV:"
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
         TabIndex        =   25
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Hepatitis B Surface Antigen (HBSAg):"
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
         Left            =   5640
         TabIndex        =   24
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label17 
         Caption         =   "Hepatitis C Virus (HCV):"
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
         Left            =   5640
         TabIndex        =   23
         Top             =   2760
         Width           =   1455
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
         Left            =   5640
         TabIndex        =   22
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblHIV 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7920
         TabIndex        =   21
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblBBN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label lblCollectionDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label lblDonorNo 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   840
         Width           =   2295
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
         Left            =   480
         TabIndex        =   16
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Collection Date:"
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
         Left            =   480
         TabIndex        =   15
         Top             =   3120
         Width           =   1575
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
         Left            =   480
         TabIndex        =   14
         Top             =   2160
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
         Left            =   480
         TabIndex        =   13
         Top             =   840
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
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   1215
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
         Left            =   480
         TabIndex        =   11
         Top             =   6480
         Width           =   1575
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
         Left            =   480
         TabIndex        =   10
         Top             =   6840
         Width           =   735
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
         Left            =   5640
         TabIndex        =   9
         Top             =   6840
         Width           =   1455
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
         Left            =   5640
         TabIndex        =   8
         Top             =   6480
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   10440
         Y1              =   6360
         Y2              =   6360
      End
   End
   Begin VB.CommandButton cmdLastRecord 
      Caption         =   ">|"
      Height          =   375
      Left            =   10440
      TabIndex        =   3
      ToolTipText     =   "Goto Last Record"
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      ToolTipText     =   "Goto Next Record"
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Goto Previous Record"
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton cmdFirstRecord 
      Caption         =   "|<"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Go to First Record"
      Top             =   7920
      Width           =   375
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   7020
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":0000
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":0112
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRecord 
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   7920
      Width           =   9015
   End
End
Attribute VB_Name = "frmREFERENCE_RECORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BLOOD_BANKDB As Database
Dim rstReference As Recordset
Private Sub Command1_Click()
'On Error Resume Next
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Set db = New ADODB.Connection
db.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data source=" & App.Path & "\BLOOD_BANK.mdb"
Set rs = New ADODB.Recordset
rs.Open "Select * From DONOR_REFERENCE_DATA order by [Donor's Number]", db, adOpenStatic, adLockReadOnly
Set reportDonor.DataSource = rs
reportDonor.Refresh
reportDonor.Show vbModal
End Sub
Private Sub Command2_Click()
'On Error Resume Next
Dim Slip As String
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Set db = New ADODB.Connection
db.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data source=" & App.Path & "\BLOOD_BANK.mdb"
Slip = InputBox("Enter Donor's No. To Print")
Set rs = New ADODB.Recordset
rs.Open "Select * From DONOR_REFERENCE_DATA Where [Donor's Number]=" & "'" & Slip & "'", db, adOpenDynamic, adLockOptimistic
Set reportDonor.DataSource = rs
reportDonor.Refresh
reportDonor.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Find"
            Find_Record
        Case "Delete"
            DELETERECORD
    End Select
End Sub
Private Sub DELETERECORD()
On Error GoTo ErrorTrap:
If MsgBox("Delete Donor Record?", vbQuestion + vbYesNo, "Delete") = vbNo Then
Exit Sub
End If
With rstReference
.Delete
Clearform
End With
Exit Sub
ErrorTrap:
 MsgBox Err.Description, vbInformation, "Error"
 End Sub
 Private Sub Clearform()
On Error Resume Next
lblName = ""
 lblDonorNo = ""
 lbl_AGE = ""
 lbl_Sex = ""
 lblAddress = ""
 lblCollectionDate = ""
 'lblExpDate = ""
 lblBBN = ""
 lblPhyExam = ""
 lblbloodgroup = ""
 lblRhesus = ""
 lblbpcv = ""
 lblHEST = ""
 lblhiv = ""
 lblhbsa = ""
 lblhcv = ""
 lblsyphilis = ""
 lblScreenOff = ""
 lblrank = ""
 lblremark = ""
 lblScreenDate = ""
 lblDonType = ""
lblDon_Phone_No = ""
lblparent = ""
lblLastDon = ""
lblAdverseRXN = ""
lblPreviousD = ""
End Sub
Private Sub Find_Record()
'On Error Resume Next
Dim StrSearch As String
StrSearch = InputBox("Enter Donor's Number:", "Search Donor")
'Set rstDonor = BLOOD_BANKDB.OpenRecordset(" Select * from DONOR_REFERENCE_DATA where [Donor's Number] = '" & StrSearch & "'';", dbOpenDynaset)
On Error Resume Next
With rstReference
.Index = "Donor's Number"
.Seek "=", StrSearch
If .NoMatch Then
MsgBox "No Record", vbInformation, "Search Donor"
Exit Sub
Else
LoadData
End If
End With
End Sub
Private Sub LoadData()
'On Error GoTo ErrorTrap
With rstReference
 lblName = ![Full Name]
 lblDonorNo = ![Donor's Number]
 lbl_AGE = !AGE
 lbl_Sex = !SEX
 lblAddress = ![PERMANENT ADDRESS]
 lblCollectionDate = !Date
 'lblExpDate = ![Exp Date]
 lblBBN = ![Bag Number]
 lblPhyExam = ![PHYSICAL EXAM]
 lblbloodgroup = ![BLOOD GROUP]
 lblRhesus = !RHESUS
 lblbpcv = ![Blood Pack Cell Volume]
 lblHEST = ![Hemoglobin Estimate]
 lblhiv = !HIV
 lblhbsa = ![Hepatitis B Surface Antigen]
 lblhcv = ![Hepatitis C Virus]
 lblsyphilis = ![SYPHILIS]
 lblScreenOff = ![Officer's Name]
 lblrank = !RANK
 lblremark = !REMARK
 lblScreenDate = ![SCREENING DATE]
 lblDonType = ![DONOR TYPE]
lblDon_Phone_No = ![Donor Phone]
lblparent = ![Parent Name]
lblLastDon = ![LAST DONATION DATE]
lblAdverseRXN = ![ADVERSE REACTION]
lblPreviousD = ![EVER DONATES]
End With
Exit Sub
ErrorTrap:
End Sub

Private Sub cmdFirstRecord_Click()
On Error Resume Next
With rstReference
.MoveFirst
LoadData
End With
End Sub

Private Sub cmdLastRecord_Click()
On Error Resume Next
With rstReference
.MoveLast
LoadData
End With
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
With rstReference
.MoveNext
If .EOF Then
 .MoveLast
End If
LoadData
End With
End Sub

Private Sub cmdPrevious_Click()
On Error Resume Next
With rstReference
.MovePrevious
If .BOF Then
 .MoveFirst
End If
LoadData
End With
End Sub

Private Sub Form_Load()
On Error Resume Next
Set BLOOD_BANKDB = OpenDatabase(App.Path & "\BLOOD_BANK.mdb", False, False)
Set rstReference = BLOOD_BANKDB.OpenRecordset("DONOR_REFERENCE_DATA")
lblDay.Caption = Format(Date, "mm:dd:yy")
End Sub




