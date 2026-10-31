VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAllocation 
   Caption         =   "ALLOCATION FORM"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14475
   Icon            =   "frmAllocation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   14475
   StartUpPosition =   2  'CenterScreen
   Tag             =   "XY"
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrAutoAllocate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Tag             =   "XY"
      Top             =   0
   End
   Begin VB.Timer tmrAutoAllocateLoadFrame 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7080
      Tag             =   "XY"
      Top             =   0
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6600
      Tag             =   "XY"
      Top             =   0
   End
   Begin TabDlg.SSTab SStabFindRetail 
      Height          =   1815
      Left            =   15360
      TabIndex        =   113
      Tag             =   "XY"
      Top             =   6720
      Visible         =   0   'False
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   3201
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "FIND"
      TabPicture(0)   =   "frmAllocation.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdPopUpFind"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin lvButton.lvButtons_H cmdPopUpFind 
         Height          =   375
         Left            =   240
         TabIndex        =   115
         Tag             =   "XY"
         Top             =   1320
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "&FIND"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmAllocation.frx":57FE
         cBack           =   -2147483633
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   114
         Tag             =   "XY"
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtFindApplicationNo 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   118
            Tag             =   "XY"
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "ALLOCATION NO."
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Tag             =   "XY"
            Top             =   360
            Width           =   1695
         End
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   3120
         TabIndex        =   116
         Tag             =   "XY"
         Top             =   1320
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "&CANCEL"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmAllocation.frx":5D98
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame fraAllocation 
      Caption         =   "ALLOCATION"
      Height          =   3375
      Left            =   15120
      TabIndex        =   98
      Tag             =   "XY"
      Top             =   1800
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   240
         TabIndex        =   105
         Tag             =   "XY"
         Top             =   1680
         Width           =   6015
         Begin VB.TextBox txtAllocateAmnt 
            BackColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   110
            Tag             =   "XY"
            Top             =   480
            Width           =   3135
         End
         Begin VB.CheckBox chkAutoAllocate 
            Caption         =   "AUTO ALLOCATE"
            Height          =   195
            Left            =   0
            TabIndex        =   108
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lblAllocateAmnt 
            BackStyle       =   0  'Transparent
            Caption         =   "ALLOCATE AMOUNT"
            Height          =   255
            Left            =   240
            TabIndex        =   109
            Tag             =   "XY"
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.TextBox txtInvBal 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   104
         Tag             =   "XY"
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txtBillingAmnt 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   102
         Tag             =   "XY"
         Top             =   720
         Width           =   4215
      End
      Begin VB.ComboBox cboBillingNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmAllocation.frx":6332
         Left            =   2040
         List            =   "frmAllocation.frx":6334
         TabIndex        =   99
         Tag             =   "XY"
         Top             =   360
         Width           =   4215
      End
      Begin lvButton.lvButtons_H cmdPopUpAllocate 
         Height          =   375
         Left            =   240
         TabIndex        =   106
         Tag             =   "XY"
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&ALLOCATE"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmAllocation.frx":6336
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPopUpClose 
         Height          =   375
         Left            =   5160
         TabIndex        =   107
         Tag             =   "XY"
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&CLOSE"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmAllocation.frx":BB28
         cBack           =   -2147483633
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE BALLANCE"
         Height          =   255
         Left            =   240
         TabIndex        =   103
         Tag             =   "XY"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblBillingAmnt 
         BackStyle       =   0  'Transparent
         Caption         =   "BILLING AMOUNT"
         Height          =   255
         Left            =   240
         TabIndex        =   101
         Tag             =   "XY"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblBillingNo 
         BackStyle       =   0  'Transparent
         Caption         =   "BILLING NO."
         Height          =   255
         Left            =   240
         TabIndex        =   100
         Tag             =   "XY"
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraLedger 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   15120
      TabIndex        =   81
      Tag             =   "XY"
      Top             =   2280
      Visible         =   0   'False
      Width           =   14055
      Begin VB.TextBox txtldgrStatus 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   95
         Tag             =   "XY"
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox txtLdgrMnthlyDue 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   93
         Tag             =   "XY"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtLdgrTerms 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   91
         Tag             =   "XY"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtLdgrAcctAmnt 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   89
         Tag             =   "XY"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtLdgrTransDate 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   87
         Tag             =   "XY"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtLdgrRefNo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   85
         Tag             =   "XY"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtLdgrAcctType 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   83
         Tag             =   "XY"
         Top             =   600
         Width           =   2535
      End
      Begin lvButton.lvButtons_H cmdLedgerHide 
         Height          =   375
         Left            =   12960
         TabIndex        =   82
         Tag             =   "XYF"
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "HIDE"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmAllocation.frx":C0C2
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxLedger 
         Height          =   3255
         Left            =   4440
         TabIndex        =   97
         Tag             =   "XY"
         Top             =   600
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   12
         RowHeightMin    =   315
         BackColorBkg    =   -2147483633
         HighLight       =   0
         GridLines       =   2
         FormatString    =   "|<REF. TYPE |<REF. NO. |<REF. DATE |<DEBIT |<CREDIT |<OS BAL |<SEQ NO. |<JOB ORDER ID |<SP INVNT NO. |<JO ORDER NO |<ACCT. BAL."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label31 
         BackColor       =   &H00808080&
         Caption         =   "LEDGER TRANSACTION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Tag             =   "XY"
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   96
         Tag             =   "XY"
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "MONTHLY DUE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   94
         Tag             =   "XY"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "TERMS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   92
         Tag             =   "XY"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCT. AMOUNT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Tag             =   "XY"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "TRANS. DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Tag             =   "XY"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "REF. NO."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Tag             =   "XY"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT TYPE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Tag             =   "XY"
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Timer tmrLoadLedger 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   5640
      Tag             =   "XY"
      Top             =   0
   End
   Begin VB.TextBox txtFlxInvDtlsAllocateAmount 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   9960
      MaxLength       =   16
      TabIndex        =   74
      Tag             =   "XY"
      Text            =   "txtFlxInvDtlsAll"
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmrLoadAccount 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   6120
      Tag             =   "XY"
      Top             =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAccnt 
      Height          =   2655
      Left            =   15120
      TabIndex        =   73
      Tag             =   "XY"
      Top             =   6600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   14
      RowHeightMin    =   285
      BackColorBkg    =   8421504
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin VB.TextBox txtFlxCashAllocation 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8280
      MaxLength       =   16
      TabIndex        =   72
      Tag             =   "XY"
      Text            =   "txtFlxCashAlloca"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin lvButton.lvButtons_H cmdShowLedger 
      Height          =   375
      Left            =   120
      TabIndex        =   52
      Tag             =   "XY"
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "SHOW LEDGER"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmAllocation.frx":C65C
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAutoAlloc 
      Height          =   375
      Left            =   1920
      TabIndex        =   53
      Tag             =   "XY"
      Top             =   4200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "AUTO ALLOCATE"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmAllocation.frx":11E4E
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   15120
      TabIndex        =   62
      Tag             =   "XY"
      Text            =   "HoldBackColor"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   15120
      Locked          =   -1  'True
      TabIndex        =   61
      Tag             =   "XY"
      Text            =   "HoldGrayBackColor"
      Top             =   1140
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdRemoveInvoiceDetails 
      Height          =   375
      Left            =   12720
      TabIndex        =   54
      Tag             =   "XY"
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "REMOVE ITEMS"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmAllocation.frx":17640
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Tag             =   "XY"
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "&NEW"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmAllocation.frx":17BDA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Tag             =   "XY"
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "&EDIT"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmAllocation.frx":18174
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   15120
      TabIndex        =   59
      Tag             =   "XY"
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "&CANCEL"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmAllocation.frx":1870E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   15120
      TabIndex        =   60
      Tag             =   "XY"
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "&SAVE"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmAllocation.frx":18CA8
      cBack           =   -2147483633
   End
   Begin VB.Frame fraHoldBottom 
      Caption         =   "INVOICE DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   50
      Tag             =   "XY"
      Top             =   4680
      Width           =   14295
      Begin VB.TextBox txtPYMTREMAINING 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   9600
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   124
         Tag             =   "XY"
         Text            =   "0.00"
         Top             =   4950
         Width           =   1695
      End
      Begin VB.TextBox txtPYMTRCVD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   122
         Tag             =   "XY"
         Text            =   "0.00"
         Top             =   4950
         Width           =   1695
      End
      Begin VB.TextBox txtAccntType 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   75
         Tag             =   "XY"
         Top             =   480
         Width           =   1575
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   10920
         TabIndex        =   71
         Tag             =   "XY"
         Top             =   360
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   661
         Caption         =   "FIND"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmAllocation.frx":19242
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtBillAmnt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   66
         Tag             =   "XY"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   65
         Tag             =   "XY"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtAllocAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   64
         Tag             =   "XY"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtRefNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   63
         Tag             =   "XY"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtFind 
         Height          =   375
         Left            =   12000
         MaxLength       =   20
         TabIndex        =   58
         Tag             =   "XY"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtAllctedAmntBottom 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   57
         Tag             =   "XY"
         Text            =   "0.00"
         Top             =   4950
         Width           =   1695
      End
      Begin VB.TextBox txtTotalCnt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   56
         Tag             =   "XY"
         Text            =   "0"
         Top             =   4950
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid flxInvoiceDetails 
         Height          =   3975
         Left            =   120
         TabIndex        =   51
         Tag             =   "XY"
         Top             =   915
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   7011
         _Version        =   393216
         Cols            =   15
         RowHeightMin    =   315
         BackColorBkg    =   -2147483633
         HighLight       =   2
         GridLines       =   2
         FormatString    =   $"frmAllocation.frx":197DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPYMTREMAINING 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "REMAINING"
         Height          =   375
         Left            =   9600
         TabIndex        =   125
         Tag             =   "XY"
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label lblPYMTRCVD 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PYMT. RCVD."
         Height          =   375
         Left            =   7680
         TabIndex        =   123
         Tag             =   "XY"
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label lblTotalAllocate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ALLOCATE"
         Height          =   375
         Left            =   5520
         TabIndex        =   121
         Tag             =   "XY"
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label lblTotalCnt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COUNT"
         Height          =   375
         Left            =   960
         TabIndex        =   120
         Tag             =   "XY"
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCNT. TYPE"
         Height          =   255
         Left            =   7080
         TabIndex        =   76
         Tag             =   "XY"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "BILL AMOUNT"
         Height          =   255
         Left            =   5400
         TabIndex        =   70
         Tag             =   "XY"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCE"
         Height          =   255
         Left            =   3840
         TabIndex        =   69
         Tag             =   "XY"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "ALLOC. AMOUNT"
         Height          =   255
         Left            =   1920
         TabIndex        =   68
         Tag             =   "XY"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "REF. NO."
         Height          =   255
         Left            =   600
         TabIndex        =   67
         Tag             =   "XY"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Tag             =   "XY"
         Top             =   4950
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   600
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   6165
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   0
      TabCaption(0)   =   "GENERAL"
      TabPicture(0)   =   "frmAllocation.frx":1988F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCust"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPayType"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblAccntType"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboCustomer"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtAllocationNo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtAllctedDate"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAllctedAmntGenTab"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtTaxHeldPercentage"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtTaxHeldAmount"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtNetPayable"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboPayType"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboAccntType"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "CASH ALLOCATION DETAIL"
      TabPicture(1)   =   "frmAllocation.frx":198AB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtCashRcvd"
      Tab(1).Control(1)=   "cmdAddCashReceive"
      Tab(1).Control(2)=   "cmdClearCashAllocation"
      Tab(1).Control(3)=   "cmdRemoveCashAllocation"
      Tab(1).Control(4)=   "fraHoldTop"
      Tab(1).Control(5)=   "txtTTLCash"
      Tab(1).Control(6)=   "lblCashRcvd"
      Tab(1).Control(7)=   "Label11"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "CHECK ALLOCATION DETAIL"
      TabPicture(2)   =   "frmAllocation.frx":198C7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtPRDate"
      Tab(2).Control(1)=   "txtChckBal"
      Tab(2).Control(2)=   "txtNewChckBal"
      Tab(2).Control(3)=   "txtChckAmnt"
      Tab(2).Control(4)=   "txtChckDate"
      Tab(2).Control(5)=   "txtChckNo"
      Tab(2).Control(6)=   "txtChckStat"
      Tab(2).Control(7)=   "txtChckTyp"
      Tab(2).Control(8)=   "txtNmOfBnk"
      Tab(2).Control(9)=   "txtAccntName"
      Tab(2).Control(10)=   "txtPRNo"
      Tab(2).Control(11)=   "cmdFindCheck"
      Tab(2).Control(12)=   "lblCheckBal"
      Tab(2).Control(13)=   "Label21"
      Tab(2).Control(14)=   "Label20"
      Tab(2).Control(15)=   "Label19"
      Tab(2).Control(16)=   "Label18"
      Tab(2).Control(17)=   "Label17"
      Tab(2).Control(18)=   "Label16"
      Tab(2).Control(19)=   "Label15"
      Tab(2).Control(20)=   "Label14"
      Tab(2).Control(21)=   "Label13"
      Tab(2).Control(22)=   "Label12"
      Tab(2).ControlCount=   23
      Begin VB.TextBox txtCashRcvd 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -73440
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   420
         Width           =   2535
      End
      Begin lvButton.lvButtons_H cmdAddCashReceive 
         Height          =   375
         Left            =   -70800
         TabIndex        =   21
         Tag             =   "XY"
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&ADD"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmAllocation.frx":198E3
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtPRDate 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   119
         Tag             =   "XY"
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox cboAccntType 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmAllocation.frx":1F0D5
         Left            =   1800
         List            =   "frmAllocation.frx":1F0DF
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Tag             =   "XY"
         Top             =   840
         Width           =   2895
      End
      Begin VB.ComboBox cboPayType 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmAllocation.frx":1F0F9
         Left            =   1800
         List            =   "frmAllocation.frx":1F10F
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Tag             =   "XY"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtChckBal 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -63480
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   48
         Tag             =   "XY"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtNewChckBal 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   47
         Tag             =   "XY"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtChckAmnt 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   46
         Tag             =   "XY"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtChckDate 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   45
         Tag             =   "XY"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtChckNo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   44
         Tag             =   "XY"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtChckStat 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -73080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   43
         Tag             =   "XY"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtChckTyp 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -73080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   42
         Tag             =   "XY"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtNmOfBnk 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -73080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   41
         Tag             =   "XY"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtAccntName 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -73080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   40
         Tag             =   "XY"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtPRNo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -73080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   39
         Tag             =   "XY"
         Top             =   600
         Width           =   2535
      End
      Begin lvButton.lvButtons_H cmdClearCashAllocation 
         Height          =   375
         Left            =   -62880
         TabIndex        =   27
         Tag             =   "XY"
         Top             =   420
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "CLEAR ALLOCATION"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmAllocation.frx":1F168
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRemoveCashAllocation 
         Height          =   375
         Left            =   -64680
         TabIndex        =   26
         Tag             =   "XY"
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "REMOVE ITEMS"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmAllocation.frx":2495A
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Frame fraHoldTop 
         Caption         =   "AVAILABLE CASH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74880
         TabIndex        =   24
         Tag             =   "XY"
         Top             =   840
         Width           =   14055
         Begin MSFlexGridLib.MSFlexGrid flxCashAllocation 
            Height          =   1695
            Left            =   120
            TabIndex        =   25
            Tag             =   "XY"
            Top             =   240
            Width           =   13815
            _ExtentX        =   24368
            _ExtentY        =   2990
            _Version        =   393216
            Cols            =   8
            RowHeightMin    =   315
            BackColorBkg    =   -2147483633
            Enabled         =   0   'False
            HighLight       =   0
            GridLines       =   2
            AllowUserResizing=   1
            FormatString    =   "|<CASH NO |<CASH RECEIVED |<TOTAL DISPENSE |^BALANCE |^ALLOCATE |^NEW BALANCE |^DATE RECEIVED"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.TextBox txtTTLCash 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   22
         Tag             =   "XY"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtNetPayable 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   11
         Tag             =   "XY"
         Text            =   "0.00"
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox txtTaxHeldAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   10
         Tag             =   "XY"
         Text            =   "0"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtTaxHeldPercentage 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   9
         Tag             =   "XY"
         Text            =   "0"
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtAllctedAmntGenTab 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   8
         Tag             =   "XY"
         Text            =   "0.00"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtAllctedDate 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   7
         Tag             =   "XY"
         Text            =   "<AUTO>"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtAllocationNo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   6
         Tag             =   "XY"
         Text            =   "<AUTO>"
         Top             =   480
         Width           =   2295
      End
      Begin lvButton.lvButtons_H cmdFindCheck 
         Height          =   375
         Left            =   -65160
         TabIndex        =   49
         Tag             =   "XY"
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "frmAllocation.frx":24EF4
         ImgSize         =   40
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cboCustomer 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmAllocation.frx":2A6E6
         Left            =   1800
         List            =   "frmAllocation.frx":2A6E8
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "XY"
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblCashRcvd 
         BackStyle       =   0  'Transparent
         Caption         =   "CASH RECEIVED"
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label lblAccntType 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT TYPE"
         Height          =   255
         Left            =   360
         TabIndex        =   79
         Tag             =   "XY"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblPayType 
         BackStyle       =   0  'Transparent
         Caption         =   "PAYMENT TYPE"
         Height          =   255
         Left            =   360
         TabIndex        =   78
         Tag             =   "XY"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblCheckBal 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK BALANCE"
         Height          =   255
         Left            =   -64920
         TabIndex        =   38
         Tag             =   "XY"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "NEW CHECK  BALANCE"
         Height          =   375
         Left            =   -69360
         TabIndex        =   37
         Tag             =   "XY"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK AMOUNT"
         Height          =   255
         Left            =   -69360
         TabIndex        =   36
         Tag             =   "XY"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK DATE"
         Height          =   255
         Left            =   -69360
         TabIndex        =   35
         Tag             =   "XY"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK NUMBER"
         Height          =   255
         Left            =   -69360
         TabIndex        =   34
         Tag             =   "XY"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "P.R. DATE"
         Height          =   255
         Left            =   -69360
         TabIndex        =   33
         Tag             =   "XY"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK STATUS"
         Height          =   255
         Left            =   -74640
         TabIndex        =   32
         Tag             =   "XY"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK TYPE"
         Height          =   255
         Left            =   -74640
         TabIndex        =   31
         Tag             =   "XY"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME OF BANK"
         Height          =   255
         Left            =   -74640
         TabIndex        =   30
         Tag             =   "XY"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT NAME"
         Height          =   255
         Left            =   -74640
         TabIndex        =   29
         Tag             =   "XY"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "P.R. NUMBER"
         Height          =   255
         Left            =   -74640
         TabIndex        =   28
         Tag             =   "XY"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Tag             =   "XY"
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "NET PAYABLE"
         Height          =   255
         Left            =   6480
         TabIndex        =   18
         Tag             =   "XY"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "TAX w/HELD(AMOUNT)"
         Height          =   255
         Left            =   6480
         TabIndex        =   17
         Tag             =   "XY"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TAX w/HELD(%)"
         Height          =   255
         Left            =   6480
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ALLOCATED AMOUNT"
         Height          =   255
         Left            =   6480
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ALLOCATED DATE"
         Height          =   255
         Left            =   6480
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ALLOCATION NO."
         Height          =   255
         Left            =   6480
         TabIndex        =   13
         Tag             =   "XY"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblCust 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Tag             =   "XY"
         Top             =   480
         Width           =   1095
      End
   End
   Begin lvButton.lvButtons_H cmdRemoveZeroAlloc 
      Height          =   375
      Left            =   10680
      TabIndex        =   112
      Tag             =   "XY"
      Top             =   4200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "ZERO ALLOCATION"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmAllocation.frx":2A6EA
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   3
      Tag             =   "XY"
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "&FIND"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmAllocation.frx":2AC84
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "&PRINT"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmAllocation.frx":2B21E
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   12000
      Tag             =   "XY"
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14715
      m_MDI_Height    =   9720
   End
End
Attribute VB_Name = "frmAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OwnerForm As String
Public AllocationID As String

Dim ObjKeyTrapping As New clsKeyTrapping
Dim formAction As String
Dim DidRecordEntry As Boolean

Dim DtlsRow As Integer
Dim DtlsCol As Integer

Dim DidflxACInUsed As Boolean

Dim DidflxAccntInUsed As Boolean
Dim LoadflxAccount As Boolean
Dim flxAccntSelRow As Integer
Dim flxAccntTTLCol As Integer

Dim arrCustomer() As String
Dim arrCurBill() As String
Dim arrPayType() As String

Dim CUSTOMERID As Integer
Dim CBOCUSTOMER_TAG As String
Dim CboCustInUsed As Boolean

Dim SearchcboCust As Boolean
Dim SearchRefNo As Boolean

Dim TaxHeldPerValue As Double
Dim TaxHeldAmntValue As Double

Dim RefNo As String

Dim BillingNo As String

Dim flxInvoiceDetailsTTLCol As Integer
Dim flxFindRowNo As Integer
Dim flxFindColNo As Integer

Dim CHKSPRID As Integer

Dim PERCENT As String
Dim Amount As String

Dim pAppNo As Integer
Dim PayType As String

Dim sql As String
Dim sqlCash As String
Dim sqlCheck As String

Dim ActiveTaxWHeldObject As String

Dim MaxTAXWithHELD As Double

Dim ALLOCATIONSTATUS As String

Dim SkipCboPayType  As String

Sub EnableSelectedButton()
'cmdTrans Legend
'0 ->NEW
'1 ->EDIT
'2 ->Find
'3 - >Print
'4 ->Save
'5 ->Cancel
If Val(txtAllocationNo.Text) <= 0 Then
    cmdTrans(1).Enabled = False
    cmdTrans(3).Enabled = False
Else
    If ALLOCATIONSTATUS = "FOR PAYMENT" _
        Or ALLOCATIONSTATUS = "ALLOCATION NEW ENTRY" _
        Or ALLOCATIONSTATUS = "UPDATING ALLOCATION RECORD" Then
        
        cmdTrans(1).Enabled = True
    Else
        cmdTrans(1).Enabled = False
    End If
    cmdTrans(3).Enabled = True
End If
End Sub

Public Sub ShowRecord(pCHECKNO As String, pCUSTOMERNAME As String, pPRNO As String, pPRDATE As String, pUSABLEBAL As String, _
                        pBANK As String, pSTATUS As String, pCHECKTYPE As String, pCHECKDATE As String, pCHECKBAL As String, _
                        pCHKSPRID As Integer)
On Error GoTo er
    Dim NetPayValue As Double
    Dim CheckBalValue As Double
    
    NetPayValue = Val(Replace(txtNetPayable.Text, ",", ""))
    CheckBalValue = Val(Replace(pCHECKBAL, ",", ""))
    
    If NetPayValue > CheckBalValue Then
        MsgBox "Net payable is greater than the check balance!", vbExclamation, "Validation!"
        Exit Sub
    End If
    
    txtPRno.Text = pPRNO
    txtAccntName.Text = pCUSTOMERNAME
    txtNmOfBnk.Text = pBANK
    txtChckTyp.Text = pCHECKTYPE
    txtChckStat.Text = pSTATUS
    txtPRDate.Text = pPRDATE
    txtChckNo.Text = pCHECKNO
    txtChckDate.Text = pCHECKDATE
    txtChckAmnt.Text = pUSABLEBAL
    txtChckBal.Text = pCHECKBAL
    
    CHKSPRID = pCHKSPRID
    
    cboPayType.Locked = True
    cboPayType.BackColor = txtHoldGrayBackColor.BackColor
    
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->ShowRecord()"
End Sub

Sub SetCboListData()
Dim Qry As String
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    'CUSTOMERID, CUSTOMERNAME, ADDRESS, CTYPE
    Call modAllocation.sp_customer_list(SysBranchID, tmpRs)
    If tmpRs.State = 1 Then
        Call modTools.addToListObjectWithArray(cboCustomer, tmpRs, arrCustomer, "CUSTOMERNAME", 2)
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->SetCboListData()"
End Sub

Sub flxAccntColsChanger()
    With flxAccnt
        flxAccntTTLCol = 14
        .Clear
        .Rows = 2
        .Cols = flxAccntTTLCol
        '                0        1             2       3         4           5            6           7         8             9            10          11             12          13           14
        .FormatString = "   |^ACCOUNT TYPE |<STATUS |^REF. NO. |<REF. DATE |<DUE DATE |<REF. AMOUNT |<REF. BAL. |<BILL NO. |^BILLING AMOUNT. |<SPSORDERID |<DUEDATE NF |<MONTH DIFF. |<JOBORDERID |TRANSDATE_NF"
     
        .ColWidth(0) = 200
        .ColWidth(1) = 1600  'ACCOUNT_TYPE
        .ColWidth(2) = 2500  'STATUS
        .ColWidth(3) = 1600  'REF. NO.
        .ColWidth(4) = 2000  'REF. DATE
        .ColWidth(5) = 2000  'DUE DATE
        .ColWidth(6) = 2000  'ACCT. AMOUNT
        .ColWidth(7) = 1500  'ACCT. BAL.
        .ColWidth(8) = 1500  'BILL NO.
        .ColWidth(9) = 2000  'BILLING AMOUNT.
        .ColWidth(10) = 0    'SPSORDERID
        .ColWidth(11) = 0    'MONTHDIFF
        .ColWidth(12) = 0    'DUEDATE_NF
        .ColWidth(13) = 0    'JOBORDERID
        .ColWidth(14) = 0    'TRANSDATE_NF
    End With
End Sub

Sub flxInvoiceDetailsColsChanger()
    With flxInvoiceDetails
        flxInvoiceDetailsTTLCol = 16
        .Clear
        .Rows = 2
        .Cols = flxInvoiceDetailsTTLCol
       '                 0   1              2        3          4            5              6          7              8              9           10          11          12          13          14          15
       .FormatString = "   |^ACCOUNT TYPE |<STATUS |^REF. NO. |>REF. AMOUNT |>REF. BALANCE |>ALLOCATE |>NEW BALANCE |>BILL AMOUNT |^BILLING NO. |<REF. DATE |<DUE DATE |DUEDATE NF |MONTH DIFF. |JOBORDERID |TRANSDATE_NF "
       .ColWidth(0) = 200
       .ColWidth(1) = 2000 'ACCOUNT TYPE
       .ColWidth(2) = 2500 'STATUS
       .ColWidth(3) = 1600 'REF. NO.
       .ColWidth(4) = 1600 'REF. AMOUNT
       .ColWidth(5) = 1700 'REF. BALANCE
       .ColWidth(6) = 2100 'ALLOCATE
       .ColWidth(7) = 2100 'New BALANCE
       .ColWidth(8) = 1600 'BILL AMOUNT
       .ColWidth(9) = 1600 'BILLING NO.
       .ColWidth(10) = 2000 'REF. DATE
       .ColWidth(11) = 2000 'DUE DATE
       .ColWidth(12) = 0 'DUEDATE_NF
       .ColWidth(13) = 0 'MONTHDIFF
       .ColWidth(14) = 0 'JOBORDERID
       .ColWidth(15) = 0 'TRANSDATE_NF
    End With
End Sub

Sub flxCashAllocationColsChanger()
    With flxCashAllocation
        .Clear
        .Rows = 2
        .Cols = 8
        '                 0   1              2             3                4         5            6           7               8
       .FormatString = "   |<CASH NO |<CASH RECEIVED |<TOTAL DISPENSE |^BALANCE |^ALLOCATE |^NEW BALANCE |^DATE RECEIVED |<ACTUAL CASH ID"
       .ColWidth(0) = 200
       .ColWidth(1) = 1500 'CASH NO
       .ColWidth(2) = 2000 'CASH RECEIVED
       .ColWidth(3) = 2000 'TOTAL DISPENSE
       .ColWidth(4) = 2000 'BALANCE
       .ColWidth(5) = 2000 'ALLOCATE
       .ColWidth(6) = 2000 'NEW BALANCE
       .ColWidth(7) = 2500 'DATE RECEIVED
       .ColWidth(8) = 0
    End With
End Sub

Sub flxLedgerColsChanger()
On Error Resume Next
    With flxLedger
        .Clear
        .Rows = 2
        .Cols = 7
       '                0    1              2             3    4         5       6
       .FormatString = "  |<REF. TYPE |<REF. NO. |<REF. DATE |<DEBIT |<CREDIT |<OS BAL"
       .ColWidth(0) = 200
       .ColWidth(1) = 1500 'REF. TYPE
       .ColWidth(2) = 1500 'REF. NO.
       .ColWidth(3) = 1000 'REF. DATE
       .ColWidth(4) = 1300 'DEBIT
       .ColWidth(5) = 1300 'CREDIT
       .ColWidth(6) = 1300 'OS BAL
    End With
End Sub

Public Sub LoadCashRecord(pALLOCATIONID As Integer)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim DtlsRow As Integer
    
    If Trim(cboCustomer.Text) = "" Or cboCustomer.ListIndex < 0 Then
        MsgBox "Please select a customer!", vbExclamation, "Validation!"
        SSTab1.Tab = 0
        cboCustomer.SetFocus
        Exit Sub
    End If
    
    Call flxCashAllocationColsChanger
    
    DoEvents
    CUSTOMERID = arrCustomer(cboCustomer.ListIndex, 0)
    
    Call modAllocation.sp_cash_list(SysBranchID, CUSTOMERID, pALLOCATIONID, tmpRs)
    'CASHID, CASHRECEIVED, TOTALDISPENSE, BALANCE, ALLOCATEDAMOUNT, NEWBALANCE, DATECREATED
    
    If tmpRs.EOF = True Then
        MsgBox "No available cash record!", vbExclamation, "Validation!"
        Exit Sub
    End If
    
    tmpRs.MoveFirst
    
    DtlsRow = 1
    With tmpRs
        Do While Not tmpRs.EOF
            '                0    1               2               3          4          5          6                7             8
            'FormatString = " |<CASH NO |<CASH RECEIVED |<TOTAL DISPENSE |^BALANCE |^ALLOCATE |^NEW BALANCE |^DATE RECEIVED |ACTUAL CASH ID
            If DtlsRow = 1 Then
                flxCashAllocation.TextMatrix(DtlsRow, 0) = ""
                flxCashAllocation.TextMatrix(DtlsRow, 1) = !CASHID
                flxCashAllocation.TextMatrix(DtlsRow, 2) = Format(Trim(!CASHRECEIVED), "#,##0.00")
                flxCashAllocation.TextMatrix(DtlsRow, 3) = Format(Trim(!TOTALDISPENSE), "#,##0.00")
                flxCashAllocation.TextMatrix(DtlsRow, 4) = Format(Trim(!Balance), "#,##0.00")
                flxCashAllocation.TextMatrix(DtlsRow, 5) = Format(Trim(!AllocatedAmount), "#,##0.00")
                flxCashAllocation.TextMatrix(DtlsRow, 6) = Format(Trim(!NewBalance), "#,##0.00")
                flxCashAllocation.TextMatrix(DtlsRow, 7) = !DateCreated
                flxCashAllocation.TextMatrix(DtlsRow, 8) = !CASHID
            Else
                flxCashAllocation.AddItem ("") & _
                vbTab & !CASHID & _
                vbTab & Format(Trim(!CASHRECEIVED), "#,##0.00") & _
                vbTab & Format(Trim(!TOTALDISPENSE), "#,##0.00") & _
                vbTab & Format(Trim(!Balance), "#,##0.00") & _
                vbTab & Format(Trim(!AllocatedAmount), "#,##0.00") & _
                vbTab & Format(Trim(!NewBalance), "#,##0.00") & _
                vbTab & !DateCreated & _
                vbTab & !CASHID
            End If
            DtlsRow = DtlsRow + 1
            tmpRs.MoveNext
        Loop
    End With
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call ComputeTTLCashAllocation
    
    On Local Error Resume Next
    If pALLOCATIONID <= 0 Then txtCashRcvd.SetFocus
    
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->LoadCashRecord()"
End Sub

Private Sub cboCustomer_LostFocus()
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub
    
    If Trim(cboCustomer.Text) <> "" And cboCustomer.ListIndex < 0 Then
        Call cboCustomer_Click
    Else
        If Trim(cboCustomer.Text) = "" Then
            Call flxAccntColsChanger
        End If
    End If
Exit Sub
er:
End Sub

Private Sub chkAutoAllocate_Click()
On Error GoTo er
    If chkAutoAllocate.Value = 1 Then
        txtAllocateAmnt.Locked = False
        txtAllocateAmnt.BackColor = txtHoldWhiteBackColor.BackColor
        
        If txtBillingAmnt.Text <= txtInvBal.Text Then
            txtAllocateAmnt.Text = txtBillingAmnt.Text
        ElseIf txtInvBal.Text <= txtBillingAmnt.Text Then
            txtAllocateAmnt.Text = txtBillingAmnt.Text
        End If
    Else
    
        txtAllocateAmnt.Locked = True
        txtAllocateAmnt.BackColor = txtHoldGrayBackColor.BackColor
    End If

Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->chkAutoAllocate_Click()"
End Sub

Private Sub cboAccntType_Click()
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub
    
    If Trim(cboCustomer.Text) <> "" And cboCustomer.ListIndex < 0 Then
        CBOCUSTOMER_TAG = "SKIP"
        cboCustomer.ListIndex = modTools.getListIndex(cboCustomer, cboCustomer.Text)
        CBOCUSTOMER_TAG = ""
    End If
    
    If cboCustomer.ListIndex < 0 Or Trim(cboCustomer.Text) = "" Then
        MsgBox "Customer name is required!", vbExclamation, "Validation!"
        cboCustomer.SetFocus
        Exit Sub
    End If
    
    CUSTOMERID = arrCustomer(cboCustomer.ListIndex, 0)
    
    'Clear Invoice Details Flexgrid
    Call flxInvoiceDetailsColsChanger
    
    tmrLoadAccount.Enabled = True
    Call Unload(frm_OnProcess)
    frm_OnProcess.OwnerForm = "frmAllocation"
    frm_OnProcess.Show vbModal
Exit Sub
er:
    CBOCUSTOMER_TAG = ""
    tmrLoadAccount.Enabled = False
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    MsgBox Err.Description, vbCritical, "Error->cboAccntType_Click()"
End Sub

Private Sub cboBillingNo_Click()
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub
    Dim cbolistIndex As Integer

    cbolistIndex = modTools.getListIndex(cboBillingNo, cboBillingNo.Text)
    If cbolistIndex < 0 Then
        cboBillingNo.SetFocus
    Else
        cboBillingNo.ListIndex = cbolistIndex

        txtBillingAmnt.Text = arrCurBill(cboBillingNo.ListIndex, 1)
        txtInvBal.Text = arrCurBill(cboBillingNo.ListIndex, 3)
    End If
    
    Call chkAutoAllocate_Click
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cboBillingNo_Click()"
End Sub

Private Sub cboCustomer_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    
    If DidRecordEntry = False Then Exit Sub
    
    
    If SearchcboCust Then
        SearchcboCust = False
        Exit Sub
    End If

    With cboCustomer
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
            'Legend
            'arrCustomer(cboCustomer.ListIndex,0) = CUSTOMERID
            'arrCustomer(cboCustomer.ListIndex,1) = CUSTOMERNAME

            CUSTOMERID = arrCustomer(i, 0)

            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboCust = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With
        
Exit Sub
er:
    tmrLoadAccount.Enabled = False
    Call Unload(frm_OnProcess)
    MsgBox Err.Description, vbCritical, "Error ->cboCustomer_Change()"
End Sub

Private Sub cboCustomer_Click()
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub
    If CBOCUSTOMER_TAG = "SKIP" Then Exit Sub
    
    Call flxCashAllocationColsChanger
    Call flxInvoiceDetailsColsChanger
    
    If modTools.getListIndex(cboCustomer, cboCustomer.Text) < 0 Then
        cboCustomer.Text = ""
        MsgBox "The customer name is not found in the list!", vbExclamation, "Validation!"
        Exit Sub
    End If
    
    If cboCustomer.ListIndex < 0 Then
        SearchcboCust = True
        cboCustomer.ListIndex = modTools.getListIndex(cboCustomer, cboCustomer.Text)
        Exit Sub
    End If
    
    If cboCustomer.ListIndex < 0 Then
        cboCustomer.Text = ""
        MsgBox "This customer name does not exist!", vbExclamation, "Validation!"
        Exit Sub
    End If
    
    CUSTOMERID = arrCustomer(cboCustomer.ListIndex, 0)
    
    tmrLoadAccount.Enabled = True
    
On Local Error Resume Next
    Call Unload(frm_OnProcess)
    frm_OnProcess.OwnerForm = "frmAllocation"
    frm_OnProcess.Show vbModal
Exit Sub
er:
    tmrLoadAccount.Enabled = False
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    MsgBox Err.Description, vbCritical, "Error->cboCustomer_Click()"
End Sub


Private Sub cboCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboCust = True
    End If
End Sub

Private Sub cboPayType_Click()
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub
    If Trim(SkipCboPayType) = "Skip" Then Exit Sub
    If Trim(cboCustomer.Text) = "" Or cboCustomer.ListIndex < 0 Then
        MsgBox "Please select a customer!", vbExclamation, "Validation!"
        SSTab1.Tab = 0
        SkipCboPayType = "Skip"
        cboPayType.ListIndex = -1
        SkipCboPayType = ""
        cboCustomer.SetFocus
        Exit Sub
    End If
    
    cmdAddCashReceive.Enabled = True
    txtTaxHeldPercentage.Locked = False
    txtTaxHeldPercentage.BackColor = txtHoldWhiteBackColor.BackColor
        
    txtTaxHeldAmount.Locked = False
    txtTaxHeldAmount.BackColor = txtHoldWhiteBackColor.BackColor
        
    txtCashRcvd.Locked = False
    txtCashRcvd.BackColor = txtHoldWhiteBackColor.BackColor
    
    
    If cboPayType.Text = "" Then
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        
    ElseIf cboPayType.Text = "CASH PAYMENT" Or cboPayType.Text = "CREDIT CARD PAYMENT" Or cboPayType.Text = "ONLINE PAYMENT" Then
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True 'Cash payment
        SSTab1.TabEnabled(2) = False 'Check payment
        If flxCashAllocation.TextMatrix(1, 1) <> "" Then
            cboPayType.Locked = True
            cboPayType.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
    ElseIf cboPayType.Text = "CHECK PAYMENT" Then
        SSTab1.Tab = 2
        SSTab1.TabEnabled(2) = True     'Check Payment
        SSTab1.TabEnabled(1) = False    'Cash Payment
        If txtPRno.Text <> "" Or txtPRDate.Text <> "" Or txtAccntName.Text <> "" Or txtNmOfBnk.Text <> "" Or txtChckTyp.Text <> "" Then
           cboPayType.Locked = True
           cboPayType.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
    ElseIf cboPayType.Text = "ACKNOWLEDGEMENT" Then
        ActiveTaxWHeldObject = "Skip"
        txtTaxHeldPercentage.Text = "0.00"
        txtTaxHeldAmount.Text = "0.00"
        txtTaxHeldPercentage.Locked = True
        txtTaxHeldPercentage.BackColor = txtHoldGrayBackColor.BackColor
            
        txtTaxHeldAmount.Locked = True
        txtTaxHeldAmount.BackColor = txtHoldGrayBackColor.BackColor
        
        txtCashRcvd.Text = ""
        txtCashRcvd.Locked = True
        txtCashRcvd.BackColor = txtHoldGrayBackColor.BackColor
        
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True     'Cash Payment
        SSTab1.TabEnabled(2) = False    'Check Payment
        
        cmdAutoAlloc.Enabled = False
        
        'Load Cash Record
        Call LoadCashRecord(-1)
        cmdAddCashReceive.Enabled = False
        
        If flxCashAllocation.TextMatrix(1, 1) <> "" Then
            cboPayType.Locked = True
            cboPayType.BackColor = txtHoldGrayBackColor.BackColor
        End If
    End If
    
    If cboPayType.Text = "CHECK PAYMENT" Or cboPayType.Text = "CASH PAYMENT" Or cboPayType.Text = "CREDIT CARD PAYMENT" Or cboPayType.Text = "ONLINE PAYMENT" Then
        If formAction = "NEW" Then
            cmdAutoAlloc.Enabled = True
        Else
            cmdAutoAlloc.Enabled = False
        End If
    End If
    
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cboPayType_Click()"
End Sub

Private Sub cmdAddCashReceive_Click()
On Error GoTo er
    Dim NewBalance As Double, NetPayValue As Double, CashValue As Double
    
    If DidRecordEntry = False Then Exit Sub
    If cmdAddCashReceive.Enabled = False Then Exit Sub
    
    NetPayValue = Val(Replace(txtNetPayable.Text, ",", ""))
    CashValue = Val(Replace(txtCashRcvd.Text, ",", ""))
    
    If Trim(txtCashRcvd.Text) = "" Then
        MsgBox "Please input cash received!", vbExclamation, "VALIDATION."
        txtCashRcvd.SelStart = 0
        txtCashRcvd.SelLength = Len(txtCashRcvd.Text)
        txtCashRcvd.SetFocus
        Exit Sub
    End If
    
    If NetPayValue > CashValue Then
        MsgBox "Cash receive is lesser than net payable!", vbExclamation, "Unable to allocate!"
        
        txtCashRcvd.SelStart = 0
        txtCashRcvd.SelLength = Len(txtCashRcvd.Text)
        txtCashRcvd.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(Trim(txtCashRcvd.Text)) = False Then
        MsgBox "Invalid cash received! Numeric is required.", vbExclamation, "VALIDATION."
        
        txtCashRcvd.SelStart = 0
        txtCashRcvd.SelLength = Len(txtCashRcvd.Text)
        txtCashRcvd.SetFocus
        Exit Sub
    End If
    
    If modTools.getFlxRowIndex(flxCashAllocation, 1, "<AUTO>", 1) >= 1 Then
        MsgBox "Multiple cash received instance is not allowed!", vbExclamation, "VALIDATION."
        
        txtCashRcvd.SelStart = 0
        txtCashRcvd.SelLength = Len(txtCashRcvd.Text)
        txtCashRcvd.SetFocus
        Exit Sub
    End If

    Call flxCashAllocationColsChanger
    NewBalance = Val(Replace(txtCashRcvd.Text, ",", "")) - Val(Replace(txtNetPayable.Text, ",", ""))
    '                0      1           2               3           4           5           6               7               8
    'FormatString = " |<CASH NO |<CASH RECEIVED |<TOTAL DISPENSE |^BALANCE |^ALLOCATE |^NEW BALANCE |^DATE RECEIVED |ACTUAL CASH ID
    DtlsRow = 1
    If Trim(flxCashAllocation.TextMatrix(DtlsRow, 1)) = "" Then
        flxCashAllocation.TextMatrix(DtlsRow, 1) = "<AUTO>"                                 'CASH NO
        flxCashAllocation.TextMatrix(DtlsRow, 2) = Format(txtCashRcvd.Text, "#,##0.00")     'CASH RECEIVED
        flxCashAllocation.TextMatrix(DtlsRow, 3) = "0"                                      'TOTAl DISPENSE
        flxCashAllocation.TextMatrix(DtlsRow, 4) = Format(txtCashRcvd.Text, "#,##0.00")     'BALANCE
        flxCashAllocation.TextMatrix(DtlsRow, 5) = Format(txtNetPayable.Text, "#,##0.00")   'ALLOCATE
        flxCashAllocation.TextMatrix(DtlsRow, 6) = Format(NewBalance, "#,##0.00")           'NEW BALANCE
        flxCashAllocation.TextMatrix(DtlsRow, 7) = "<AUTO>"                                 'DATE RECEIVED
        flxCashAllocation.TextMatrix(DtlsRow, 8) = 0                                        'ACTUAL CASH ID
    Else
        flxCashAllocation.AddItem "" & _
         vbTab & "<AUTO>" & _
         vbTab & Format(txtCashRcvd.Text, "#,##0.00") & _
         vbTab & "0" & _
         vbTab & Format(txtCashRcvd.Text, "#,##0.00") & _
         vbTab & Format(txtNetPayable.Text, "#,##0.00") & _
         vbTab & Format(NewBalance, "#,##0.00") & _
         vbTab & "<AUTO>" & _
         vbTab & 0
    End If
    
    Call ComputeTTLCashAllocation
    txtCashRcvd.Text = ""
    
    cmdClearCashAllocation.Enabled = True
    
    cboPayType.Locked = True
    cboPayType.BackColor = txtHoldGrayBackColor.BackColor
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error ->cmdAddCashReceive_Click()"
End Sub

Public Function ComputeTTLCashAllocation() As Boolean
On Error GoTo er
    Dim intRow As Integer
    Dim Alloc As Double
    ComputeTTLCashAllocation = False
    Alloc = 0
    For intRow = 1 To flxCashAllocation.Rows - 1
        Alloc = Alloc + Val(Replace(flxCashAllocation.TextMatrix(intRow, 5), ",", ""))
    Next
    
    txtTTLCash.Text = Format(Alloc, "#,##0.00")
    If Alloc <= 0 Then
        cmdClearCashAllocation.Enabled = False
    Else
        If DidRecordEntry Then cmdClearCashAllocation.Enabled = True
    End If
    
    ComputeTTLCashAllocation = True
Exit Function
er:
    ComputeTTLCashAllocation = False
End Function

Public Function ComputeTTLAccountAllocAmount() As Boolean
On Error GoTo er
    Dim intRow As Integer
    Dim Alloc As Double
    
    ComputeTTLAccountAllocAmount = False
    For intRow = 1 To flxInvoiceDetails.Rows - 1
        Alloc = Alloc + Val(Replace(flxInvoiceDetails.TextMatrix(intRow, 6), ",", ""))
    Next
    txtAllctedAmntGenTab.Text = Format(Alloc, "#,##0.00")
    txtAllctedAmntBottom.Text = txtAllctedAmntGenTab.Text
    ComputeTTLAccountAllocAmount = True
Exit Function
er:
    ComputeTTLAccountAllocAmount = False
End Function

Private Sub cmdPopUpAllocate_Click()
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub
    If txtAllocateAmnt.Text = "" Then
        MsgBox "ALLOCATE AMOUNT IS REQUIRED!", vbInformation, "ALLOCATION"
        Exit Sub
    End If
    tmrAutoAllocate.Enabled = True
    Call Unload(frm_OnProcess)
    frm_OnProcess.OwnerForm = "frmAllocation"
    frm_OnProcess.Show vbModal
Exit Sub
er:
    tmrAutoAllocate.Enabled = False
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    MsgBox Err.Description, vbCritical, "Error->cmdAllocate_Click()"
End Sub

Private Sub cmdAutoAlloc_Click()
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub

    If cboCustomer.Text = "" Or CUSTOMERID <= 0 Then
        cboCustomer.SetFocus
        Exit Sub
    End If
    
    cmdRemoveZeroAlloc.Enabled = False
    
    tmrAutoAllocateLoadFrame.Enabled = True
    Call Unload(frm_OnProcess)
    frm_OnProcess.OwnerForm = "frmAllocation"
    frm_OnProcess.Show vbModal
 Exit Sub
er:
    tmrAutoAllocateLoadFrame.Enabled = False
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    MsgBox Err.Description, vbCritical, "Error -> cmdAutoAlloc_Click()"
End Sub

Private Sub cmdCancel_Click()
    SStabFindRetail.Visible = False
    txtFindApplicationNo.Text = ""
    DoEvents
    Call EnableFields
    DoEvents
    Call EnableSelectedButton
End Sub

Private Sub cmdClearCashAllocation_Click()
On Error GoTo er
    Dim Ans As VbMsgBoxResult
    Dim flxRow As Integer
    If cboPayType.Text = "ACKNOWLEDGEMENT" Then
        cmdClearCashAllocation.Enabled = False
        MsgBox "For payment type ACKNOWLEDGEMENT this action is not allowed!", vbExclamation, "Validation!"
        Exit Sub
    End If
    
    Ans = MsgBox("Are you do you want to clear the cash allocation?", vbYesNo + vbQuestion, "Clear Cash Allocation?")
    If Ans = vbNo Then Exit Sub
    
    '                  0   1              2             3                4         5            6           7               8
    '.FormatString = "   |<CASH NO |<CASH RECEIVED |<TOTAL DISPENSE |^BALANCE |^ALLOCATE |^NEW BALANCE |^DATE RECEIVED |<ACTUAL CASH ID"
    For flxRow = 1 To flxCashAllocation.Rows - 1
        flxCashAllocation.TextMatrix(flxRow, 5) = 0
    Next
     
    DoEvents
    Call ComputeTTLCashAllocation
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error -> cmdClearCashAllocation_Click()"
End Sub

Private Sub cmdPopUpClose_Click()
On Error Resume Next
    
    fraAllocation.Visible = False
    
    Call EnableFields
    
    cboBillingNo.Text = ""
    txtBillingAmnt.Text = ""
    txtInvBal.Text = ""
    txtAllocateAmnt.Text = ""
    

End Sub

Sub EnableFields()
On Error GoTo er
'cmdTrans Legend
'0 ->NEW
'1 ->EDIT
'2 ->Find
'3 ->PRINT
'4 ->Save
'5 ->Cancel
    If cmdRemoveInvoiceDetails.Enabled = False Then
        cmdRemoveInvoiceDetails.Enabled = True
    Else
        cmdRemoveInvoiceDetails.Enabled = False
    End If

    If Trim(flxInvoiceDetails.TextMatrix(1, 1)) <> "" Then
        cmdRemoveZeroAlloc.Enabled = True
    Else
        cmdRemoveZeroAlloc.Enabled = False
    End If
    
    cmdTrans(0).Enabled = True
    cmdTrans(2).Enabled = True
 
    SSTab1.Enabled = True
    cmdShowLedger.Enabled = True
    fraHoldBottom.Enabled = True
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->EnableFields()"
End Sub

Private Sub cmdFind_Click()
    flxAccnt.Visible = False
    Call modTools.flxFind(flxInvoiceDetails, txtFind, Me.Caption, flxInvoiceDetailsTTLCol, flxFindRowNo, flxFindColNo)
End Sub

Private Sub cmdFindCheck_Click()
    frmFindCheck.Show vbModal
End Sub

Sub FindAndLoadRecord()
    Dim tmpRs As New ADODB.Recordset
    Dim PRDATE As String
    Dim CheckID As String
    Dim ACCNTTYPE As String
    
    Call ClearEntryForm(False)
    DoEvents
    DoEvents
    If Trim(AllocationID) = "" Then Exit Sub
    
    DoEvents
    Call modAllocation.sp_allocation_list(SysBranchID, AllocationID, tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount <= 0 Then
            MsgBox "No record found!", vbExclamation, "Validation!"
            tmpRs.Close
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    SSTab1.Tab = 0
    tmpRs.MoveFirst
    With tmpRs
        CUSTOMERID = !CUSTOMERID
        cboCustomer.Text = !CustomerName
        CBOCUSTOMER_TAG = "SKIP"
        cboCustomer.ListIndex = modTools.getListIndex(cboCustomer, cboCustomer.Text)
        CBOCUSTOMER_TAG = ""
        
        ALLOCATIONSTATUS = !ALLOCATIONSTATUS
        ACCNTTYPE = !ACCOUNTTYPE
        PayType = !PayType
        txtAllocationNo.Text = !AllocationID
        txtAllctedDate.Text = !ALLOCATEDDATE
        txtAllctedAmntGenTab.Text = !ALLOC_AMOUNT
        
        PERCENT = !TAXWHELDPERCENT
        Amount = !TAXWHELDAMOUNT
        txtTaxHeldPercentage.Text = PERCENT
        txtTaxHeldAmount.Text = Amount
        
        txtPRno.Text = !PRNO
        txtPRDate.Text = !PRDATE
        txtAccntName.Text = !CHKACCTNAME
        txtNmOfBnk.Text = !BANK
        txtChckTyp.Text = !CHECKTYPE
        txtChckStat.Text = !ChkStatus
        
        CheckID = !CHKSPRID
        If CheckID = "" Then
            CheckID = 0
        End If
        CHKSPRID = CheckID
        
        txtChckNo.Text = !CHECKNO
        txtChckDate.Text = !CHECKDATE
        txtChckAmnt.Text = IIf(Trim(!CHECKAMOUNT) = "", "", Format(Val(!CHECKAMOUNT), "#,##0.00"))
        txtChckBal.Text = IIf(Trim(!CHECKBAL) = "", "", Format(Val(!CHECKBAL), "#,##0.00"))
        
        
        'ACCNTTYPE = "SERVICE"
        'ACCNTTYPE = "SPARE PARTS"
        cboAccntType.ListIndex = modTools.getListIndex(cboAccntType, ACCNTTYPE)
        
        
        If PayType = "CASH PAYMENT" Or PayType = "ACKNOWLEDGEMENT" Or PayType = "CREDIT CARD PAYMENT" Or PayType = "ONLINE PAYMENT" Then
            Call modTools.changeListIndex(cboPayType, PayType)
            Call LoadCashRecord(!AllocationID)
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
        ElseIf PayType = "CHECK PAYMENT" Then
            Call modTools.changeListIndex(cboPayType, PayType)
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = True
        End If
        
        DtlsRow = 1
        Do While Not tmpRs.EOF
            'ALLOCATIONID, PAYTYPE, CUSTOMERID, CUSTOMERNAME, ALLOCATEDDATE, TAXWHELDPERCENT, TAXWHELDAMOUNT,
            'ALLOCATIONSTATUS, ACCOUNTSTATUS, ACCOUNTTYPE, REFNO, REFAMOUNT, REFBALANCE, ALLOC_AMOUNT,
            'BALANCEDURINGALLOCATION, BILLAMOUNT, BILLINGNO, DUEDATE, PRNO, PRDATE, CHECKAMOUNT, USABLEBAL,
            'BANK,CHKSPRID, CHECKNO, TTLDEBIT, TTLCREDIT, CHECKBAL, CHKACCTNAME, CHECKTYPE, CHKSTATUS, TRANSDATE,
            'DUEDATE, DUEDATE_NF, MONTHDIFF, JOBORDERID,TRANSDATE_NF
             
             If DtlsRow = 1 Then
                flxInvoiceDetails.TextMatrix(DtlsRow, 1) = !ACCOUNTTYPE
                flxInvoiceDetails.TextMatrix(DtlsRow, 2) = !ACCOUNTSTATUS
                flxInvoiceDetails.TextMatrix(DtlsRow, 3) = !RefNo
                flxInvoiceDetails.TextMatrix(DtlsRow, 4) = Format(!REFAMOUNT, "#,##0.00")
                flxInvoiceDetails.TextMatrix(DtlsRow, 5) = Format(!REFBALANCE, "#,##0.00")
                flxInvoiceDetails.TextMatrix(DtlsRow, 6) = Format(!ALLOC_AMOUNT, "#,##0.00")
                flxInvoiceDetails.TextMatrix(DtlsRow, 7) = Format(!BALANCEDURINGALLOCATION, "#,##0.00")
                flxInvoiceDetails.TextMatrix(DtlsRow, 8) = Format(!BILLAMOUNT, "#,##0.00")
                flxInvoiceDetails.TextMatrix(DtlsRow, 9) = !BillingNo
                flxInvoiceDetails.TextMatrix(DtlsRow, 10) = !TRANSDATE
                flxInvoiceDetails.TextMatrix(DtlsRow, 11) = !DUEDATE
                flxInvoiceDetails.TextMatrix(DtlsRow, 12) = !DUEDATE_NF
                flxInvoiceDetails.TextMatrix(DtlsRow, 13) = !MONTHDIFF
                flxInvoiceDetails.TextMatrix(DtlsRow, 14) = !JOBORDERID
                flxInvoiceDetails.TextMatrix(DtlsRow, 15) = !TRANSDATE_NF
            Else
                flxInvoiceDetails.AddItem _
                    vbTab & !ACCOUNTTYPE & _
                    vbTab & !ACCOUNTSTATUS & _
                    vbTab & !RefNo & _
                    vbTab & Format(!REFAMOUNT, "#,##0.00") & _
                    vbTab & Format(!REFBALANCE, "#,##0.00") & _
                    vbTab & Format(!ALLOC_AMOUNT, "#,##0.00") & _
                    vbTab & Format(!BALANCEDURINGALLOCATION, "#,##0.00") & _
                    vbTab & Format(!BILLAMOUNT, "#,##0.00") & _
                    vbTab & !BillingNo & _
                    vbTab & !TRANSDATE & _
                    vbTab & !DUEDATE & _
                    vbTab & !DUEDATE_NF & _
                    vbTab & !MONTHDIFF & _
                    vbTab & !JOBORDERID & _
                    vbTab & !TRANSDATE_NF
            End If
           
            DtlsRow = DtlsRow + 1
            tmpRs.MoveNext
        Loop
    End With
    If tmpRs.State <> 0 Then tmpRs.Close
    
    'Count Total Invoice Details
    txtTotalCnt.Text = flxInvoiceDetails.Rows - 1
    
    
    Call ComputeTTLAccountAllocAmount
    Call ComputeNetPayable
    Call ComputeNewCheckBalance
    Call txtPYMTRCVD_Change
    
    DoEvents
    'Enabled Edit and Print Button
    Call EnableSelectedButton
End Sub

Private Sub cmdPopUpFind_Click()
On Error GoTo er
    If txtFindApplicationNo.Text = "" Then Exit Sub

    AllocationID = txtFindApplicationNo.Text
    
    'Call the sub routine to find and display allocation record
    Call FindAndLoadRecord
    
    DoEvents
    Call cmdCancel_Click
    txtFindApplicationNo.Text = ""
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cmdPopUpFind_Click()"
End Sub

Private Sub cmdLedgerHide_Click()
    fraLedger.Visible = False
    cmdShowLedger.Enabled = False
    
    txtLdgrAcctType.Text = ""
    txtLdgrRefNo.Text = ""
    txtLdgrTransDate.Text = ""
    txtLdgrAcctAmnt.Text = ""
    txtLdgrTerms.Text = ""
    txtLdgrMnthlyDue.Text = ""
    txtldgrStatus.Text = ""
    
    cmdTrans(0).Enabled = True
    cmdTrans(1).Enabled = True
    cmdTrans(2).Enabled = True
    cmdTrans(3).Enabled = True
    SSTab1.Enabled = True
    
    If DidRecordEntry Then
        cmdShowLedger.Enabled = True
        cmdRemoveZeroAlloc.Enabled = True
        cmdRemoveInvoiceDetails.Enabled = True
    End If
    
    fraHoldBottom.Enabled = True
End Sub

Private Sub cmdRemoveInvoiceDetails_Click()
On Error GoTo er
    Dim Ans As VbMsgBoxResult
    Dim flxIndex As Integer
    
    If cmdRemoveInvoiceDetails.Enabled = False Then
        Exit Sub
    End If
    cmdRemoveInvoiceDetails.Enabled = False
    
    If DidRecordEntry = False Then Exit Sub
    If flxInvoiceDetails.Row < 1 Then Exit Sub
    
    If flxInvoiceDetails.TextMatrix(flxInvoiceDetails.Row, 1) = "" Then Exit Sub
    flxIndex = flxInvoiceDetails.Row
    
    DoEvents
    
    Ans = MsgBox("Remove selected item, which account type=" & flxInvoiceDetails.TextMatrix(flxIndex, 1) & " and refno=" & flxInvoiceDetails.TextMatrix(flxIndex, 3) & "?", vbQuestion + vbYesNo, "REMOVE ITEM")
    If Ans = vbNo Then Exit Sub
    
    '                 0     1              2         3          4            5              6          7              8              9           10          11          12          13          14
    '.FormatString = "   |^ACCOUNT TYPE |<STATUS |^REF. NO. |<REF. AMOUNT |<REF. BALANCE |<ALLOCATE |<NEW BALANCE |<BILL AMOUNT |^BILLING NO. |<REF. DATE |<DUE DATE |DUEDATE NF |MONTH DIFF. |JOBORDERID "
    If flxIndex = 1 And flxInvoiceDetails.Rows <= 2 Then
        flxInvoiceDetails.TextMatrix(flxIndex, 1) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 2) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 3) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 4) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 5) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 6) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 7) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 8) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 9) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 10) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 11) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 12) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 13) = ""
        flxInvoiceDetails.TextMatrix(flxIndex, 14) = ""
    Else
        Call flxInvoiceDetails.RemoveItem(flxIndex)
    End If
    
    Call ComputeTTLAccountAllocAmount
    'Count Total Invoice Details
    txtTotalCnt.Text = flxInvoiceDetails.Rows - 1
    
    Call txtTaxHeldPercentage_Change
    
    Call ComputeNewCheckBalance
    cmdRemoveInvoiceDetails.Enabled = False
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error-> cmdRemoveCashAllocation_Click()"
End Sub

Private Sub cmdRemoveCashAllocation_Click()
 On Error GoTo er
    Dim Ans As VbMsgBoxResult
    
    If DidRecordEntry = False Then Exit Sub
    If cboPayType.Text = "ACKNOWLEDGEMENT" Then
        cmdRemoveCashAllocation.Enabled = False
        Exit Sub
    End If
    If cmdRemoveCashAllocation.Enabled = False Then Exit Sub
    
    Ans = MsgBox("Remove selected cash record?", vbYesNo + vbQuestion, "Remove?")
    If Ans = vbNo Then Exit Sub
    
    
    cmdRemoveCashAllocation.Enabled = False
    
    If flxCashAllocation.Rows = 2 Then
        flxCashAllocation.Rows = 1                  ' Delete the 1st active row
        cmdClearCashAllocation.Enabled = False
        
    ElseIf flxCashAllocation.RowSel >= 1 Then
        flxCashAllocation.RemoveItem (flxCashAllocation.RowSel) ' Delete  the active row
    End If
    
    If flxCashAllocation.Rows = 1 Then
        cboPayType.Locked = False
        cboPayType.BackColor = txtHoldWhiteBackColor.BackColor
    End If
        
    Call ComputeTTLCashAllocation

    If flxCashAllocation.Rows <= 1 Then
        cmdAddCashReceive.Enabled = True
    Else
        cmdAddCashReceive.Enabled = False
    End If
    
    txtCashRcvd.Locked = False
    txtCashRcvd.BackColor = txtHoldWhiteBackColor.BackColor
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error-> cmdRemoveCashAllocation_Click()"
End Sub

Public Function QueueCashAllocation() As Boolean
On Error GoTo er
    Dim cashreceive As Double
    Dim allocation As Double
    Dim flxRow As Integer
    Dim Qry As String
    
    QueueCashAllocation = False
    Qry = "DROP  TEMPORARY TABLE IF  EXISTS sp_mrcretails.TMP_CASH_ALLOCATION;"
    Call LYRcon.Execute(Qry)
     
    'create a temporary table
    Qry = "CREATE  TEMPORARY TABLE IF NOT EXISTS sp_mrcretails.TMP_CASH_ALLOCATION(  " & _
            "   CashID varchar(20),       " & _
            "   CASH_RECEIVED DECIMAL(24,2), " & _
            "   ALLOC_AMOUNT DECIMAL(24,2),    " & _
            "   ACTUAL_CASHID   INT         " & _
            ")"
    'InputBox Qry, Qry, Qry
    Call LYRcon.Execute(Qry)
    DoEvents
    Call LYRcon.Execute("Truncate sp_mrcretails.TMP_CASH_ALLOCATION;")
    
    DoEvents
    Qry = "INSERT INTO sp_mrcretails.TMP_CASH_ALLOCATION (CashID,CASH_RECEIVED,ALLOC_AMOUNT,ACTUAL_CASHID) Values "
    For flxRow = 1 To flxCashAllocation.Rows - 1
        If flxRow <> 1 Then
            Qry = Qry & ","
        End If
        
        cashreceive = Replace(Trim(flxCashAllocation.TextMatrix(flxRow, 2)), ",", "")
        allocation = Replace(Trim(flxCashAllocation.TextMatrix(flxRow, 5)), ",", "")
        
                                    'Cash ID                         'Cash receive      'Alloc. Amount          'ACTUAL CASH ID
        Qry = Qry & "('" & Trim(flxCashAllocation.TextMatrix(flxRow, 1)) & "'," & cashreceive & "," & allocation & "," & Trim(flxCashAllocation.TextMatrix(flxRow, 8)) & ")"
    Next
    'InputBox Qry, Qry, Qry
    Call LYRcon.Execute(Qry)
    
    QueueCashAllocation = True
Exit Function
er:
    QueueCashAllocation = False
End Function
 

Private Sub cmdShowLedger_Click()
On Error GoTo er
    If flxInvoiceDetails.TextMatrix(flxInvoiceDetails.Row, 1) = "" Then Exit Sub
    
    RefNo = flxInvoiceDetails.TextMatrix(flxInvoiceDetails.Row, 3)
    
    Call LoadLedger(CUSTOMERID, cboAccntType.Text, RefNo)
    
    DoEvents
    fraLedger.Visible = True
    
    cmdTrans(0).Enabled = False
    cmdTrans(1).Enabled = False
    cmdTrans(2).Enabled = False
    cmdTrans(3).Enabled = False
    SSTab1.Enabled = False
    cmdShowLedger.Enabled = False
    cmdRemoveZeroAlloc.Enabled = False
    fraHoldBottom.Enabled = False
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->cmdShowLedger_Click()"
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
'cmdTrans Legend
'0 ->NEW
'1 ->EDIT
'2 ->Find
'3 - >Print
'4 ->Save
'5 ->Cancel

    Dim level As String
    Dim message As String
    Dim Ans As VbMsgBoxResult

    If Index = 0 Then 'NEW
        formAction = "NEW"
        Call SetButtonVisibility(True)
        
        'Locked the Product Entry Object
        Call LockedEntryObject(True)
        Call ClearEntryForm(True)
        PayType = ""
        AllocationID = ""
        
        If cboCustomer.Text <> "" Then
            cboCustomer.Locked = True
            cboCustomer.BackColor = txtHoldGrayBackColor.BackColor
            cboPayType.Locked = True
            cboPayType.BackColor = txtHoldGrayBackColor.BackColor
        Else
            cboPayType.Locked = False
            cboPayType.BackColor = txtHoldWhiteBackColor.BackColor
        End If

 
        DidRecordEntry = True
        cmdAddCashReceive.Enabled = True
        flxCashAllocation.Enabled = True
        cboAccntType.Locked = False
        cboAccntType.BackColor = txtHoldWhiteBackColor.BackColor
             
        cmdRemoveInvoiceDetails.Enabled = False
        
        Me.Caption = "ALLOCATION FORM (NEW)"
    
    ElseIf Index = 1 Then 'EDIT
        formAction = "UPDATE"
        Call SetButtonVisibility(True)
        
        'Locked the Product Entry Object
        Call LockedEntryObject(True)
        DidRecordEntry = True
        flxCashAllocation.Enabled = True
        cboAccntType.Locked = False
        cboAccntType.BackColor = txtHoldWhiteBackColor.BackColor
        
        cmdRemoveInvoiceDetails.Enabled = False
        
        txtCashRcvd.Locked = True
        txtCashRcvd.BackColor = txtHoldGrayBackColor.BackColor
                        
        If cboPayType.ListIndex > 0 Then
            cboPayType.Locked = True
            cboPayType.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
        If cboAccntType.ListIndex >= 0 Then
            cboAccntType.Locked = True
            cboAccntType.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
        cboCustomer.Locked = True
        cboCustomer.BackColor = txtHoldGrayBackColor.BackColor
        
        cmdRemoveZeroAlloc.Enabled = True
        
        DoEvents
        tmrLoadAccount.Enabled = True
        Call Unload(frm_OnProcess)
        frm_OnProcess.OwnerForm = "frmAllocation"
        frm_OnProcess.Show vbModal
        
        Me.Caption = "ALLOCATION FORM (UPDATE)"
        
    ElseIf Index = 2 Then 'FIND Button
        
        SStabFindRetail.Visible = True
        
        cmdTrans(0).Enabled = False
        cmdTrans(1).Enabled = False
        cmdTrans(2).Enabled = False
        cmdTrans(3).Enabled = False
        SSTab1.Enabled = False
        cmdShowLedger.Enabled = False
        cmdAutoAlloc.Enabled = False
        cmdRemoveZeroAlloc.Enabled = False
        fraHoldBottom.Enabled = False
        txtFindApplicationNo.SetFocus
        
    ElseIf Index = 3 Then 'PRINT Button
        Dim tmpRs As New ADODB.Recordset
        Dim tmpRsCash As New ADODB.Recordset
        Dim tmpRsCheck As New ADODB.Recordset
        
        If Trim(txtAllocationNo.Text) = "" Then
            MsgBox "Allocation number is not found!", vbExclamation, "Validation!"
            Exit Sub
        End If
        
        
        Screen.MousePointer = vbHourglass
        Dim rptAllocForm As New rptAllocationMainForm
        
        rptAllocForm.DiscardSavedData
        
        DoEvents
        Call modAllocation.sp_allocation_printform(CInt(txtAllocationNo.Text), tmpRs)
        
        DoEvents
        Call modAllocation.sp_allocation_printform_cash_attachment(CInt(txtAllocationNo.Text), tmpRsCash)
        
        DoEvents
        Call modAllocation.sp_allocation_printform_check_attachment(CInt(txtAllocationNo.Text), tmpRsCheck)
        
        DoEvents
        'Main report
        With rptAllocForm
            .Database.SetDataSource tmpRs
            .ReadRecords
        End With
        
        'Subreport Cash Attachment
        With rptAllocForm.Subreport5.OpenSubreport
            .Database.SetDataSource tmpRsCash
            .ReadRecords
        End With
        
        'Subreport Check Attachment
        With rptAllocForm.Subreport6.OpenSubreport
            .Database.SetDataSource tmpRsCheck
            .ReadRecords
        End With
        
        PRNatt.CRViewer1.ReportSource = rptAllocForm
        PRNatt.CRViewer1.ViewReport
        PRNatt.CRViewer1.Zoom 150
        Screen.MousePointer = vbDefault
        
        DoEvents
        PRNatt.WindowState = vbMaximized
        PRNatt.Show vbModal
        
    ElseIf Index = 4 Then  'SAVE Button
        If flxInvoiceDetails.TextMatrix(flxInvoiceDetails.Row, 1) = "" Then
            MsgBox "No invoice detail!", vbExclamation, "Validation!"
            Exit Sub
        End If
        
        If Trim(cboCustomer.Text) = "" Then
            MsgBox "CUSTOMER IS REQUIRED!", vbExclamation, "Validation!"
            cboCustomer.SetFocus
            Exit Sub
        End If
        
        
        If ValidateEntry() = False Then Exit Sub
    
        DoEvents
        If formAction = "UPDATE" Then
            Ans = MsgBox("DO YOU WANT TO SAVE YOUR CHANGES?", vbQuestion + vbYesNo, Me.Caption)
            If Ans = vbNo Then Exit Sub
        Else
            Ans = MsgBox("PLEASE MAKE SURE THAT THE DATA IS CORRECT!" & _
                vbCrLf & _
                vbCrLf & "DO YOU WISH TO CONTINUE?", vbYesNo + vbQuestion, Me.Caption)
    
            If Ans = vbNo Then Exit Sub
        End If
        
        If cboPayType.Text = "CASH PAYMENT" Or cboPayType.Text = "ACKNOWLEDGEMENT" Or cboPayType.Text = "CREDIT CARD PAYMENT" Or cboPayType.Text = "ONLINE PAYMENT" Then
            'Insert data to the temporary tabe
            If QueueCashAllocation = False Then
                MsgBox "Temporary cash allocation values was not inserted!", vbCritical, Caption
                Exit Sub
            End If
        End If
        
        
        Dim Qry As String
        Dim HasValue As Boolean
        
        Dim SPINVNO As String
        Dim REF_NO As String
        Dim REF_AMOUNT As String
        Dim ALLOC_AMOUNT As String
        Dim Balance As String
        Dim DUEDATE_NF As String
        Dim SQNO As Integer
        Dim BillingNo As String
        Dim PERCENT As String
        Dim Amount As String
        Dim ALN_DATECREATED As String
        
        Dim i As Integer
        
        Qry = "Drop Temporary Table IF  Exists sp_mrcretails.TMP_ACCOUNT_ALLOCATION;"
        Call LYRcon.Execute(Qry)
         DoEvents
        Qry = " Create Temporary Table IF Not Exists sp_mrcretails.TMP_ACCOUNT_ALLOCATION( " & _
        "   SPINVNO VarChar(20),        " & _
        "   JOBORDERNO Varchar(20),     " & _
        "   REF_AMOUNT Decimal(24,2),   " & _
        "   ALLOC_AMOUNT Decimal(24,2), " & _
        "   BALANCE Decimal(24,2),      " & _
        "   BILLINGNO Varchar(20),      " & _
        "   MONTHDIFF Int,              " & _
        "   DUEDATE Date,               " & _
        "   SQNO Int,                   " & _
        "   JOBORDERID Int,             " & _
        "   TRANSDATE_NF   Date         " & _
        ");"
        
        Call LYRcon.Execute(Qry)
    
        DoEvents
        SQNO = 1
        
        HasValue = False
        Qry = "Insert Into sp_mrcretails.TMP_ACCOUNT_ALLOCATION(SPINVNO, JOBORDERNO, REF_AMOUNT, ALLOC_AMOUNT, BALANCE, BILLINGNO, MONTHDIFF, DUEDATE,SQNO,JOBORDERID,TRANSDATE_NF) Values"
    
        '               0      1            2        3            4               5            6          7              8              9           10          11         12         13         14         15
        '.FormatString = " |^ACCOUNT TYPE |<STATUS |^REF. NO. |<REF. AMOUNT |<REF. BALANCE |<ALLOCATE |<NEW BALANCE |<BILL AMOUNT |^BILLING NO. |<REF. DATE |<DUE DATE |DUEDATE_NF |MONTHDIFF |JOBORDERID |TRANSDATE_NF "
        For i = 1 To flxInvoiceDetails.Rows - 1
            
            DoEvents
            HasValue = True
            If i > 1 Then Qry = Qry & ","
                 
            REF_AMOUNT = Replace(Trim(flxInvoiceDetails.TextMatrix(i, 4)), ",", "")
            If Trim(REF_AMOUNT) = "" Then REF_AMOUNT = "0"
    
            ALLOC_AMOUNT = Replace(Trim(flxInvoiceDetails.TextMatrix(i, 6)), ",", "")
            If Trim(ALLOC_AMOUNT) = "" Then ALLOC_AMOUNT = "0"
    
            Balance = Replace(Trim(flxInvoiceDetails.TextMatrix(i, 7)), ",", "")
            If Trim(Balance) = "" Then Balance = "0"
    
            DUEDATE_NF = Format(Trim(flxInvoiceDetails.TextMatrix(i, 12)), "yyyy-mm-dd")
            
            If Trim(flxInvoiceDetails.TextMatrix(i, 1)) = "SPARE PARTS" Then
                SPINVNO = Trim(flxInvoiceDetails.TextMatrix(i, 3))
                REF_NO = ""
            ElseIf Trim(flxInvoiceDetails.TextMatrix(i, 1)) = "SERVICE" Then
                REF_NO = Trim(flxInvoiceDetails.TextMatrix(i, 3))
                SPINVNO = ""
            End If
            
            Qry = Qry & "('" & SPINVNO & "','" & REF_NO & "'," & REF_AMOUNT & "," & ALLOC_AMOUNT & "," & Balance & "," & _
                        "'" & Trim(flxInvoiceDetails.TextMatrix(i, 9)) & "'," & Trim(flxInvoiceDetails.TextMatrix(i, 13)) & ",'" & DUEDATE_NF & "'" & _
                        "," & SQNO & "," & Trim(flxInvoiceDetails.TextMatrix(i, 14)) & ",'" & Trim(flxInvoiceDetails.TextMatrix(i, 15)) & "')"
         Next
         
        Call LYRcon.Execute(Qry)
        'InputBox Qry, Qry, Qry
        
        AllocationID = ""
        ALN_DATECREATED = ""
        If formAction = "NEW" Then
            Call modAllocation.sp_allocation_create(SysBranchID, cboAccntType.Text, cboPayType.Text, CUSTOMERID, CHKSPRID, txtChckNo.Text, txtTaxHeldPercentage.Text, txtTaxHeldAmount.Text, UserID, AllocationID, ALN_DATECREATED, ALLOCATIONSTATUS, level, message)
        Else
            Call modAllocation.sp_allocation_update(SysBranchID, txtAllocationNo.Text, cboAccntType.Text, cboPayType.Text, CUSTOMERID, CHKSPRID, txtChckNo.Text, Replace(Trim(txtTaxHeldPercentage.Text), ",", ""), Replace(Trim(txtTaxHeldAmount.Text), ",", ""), UserID, level, message)
        End If
        
        If Trim(level) = "" Then
            Exit Sub
        ElseIf level = "Granted" Or level = "Warning" Then
            If level = "Granted" Then
                MsgBox message, vbInformation, "SAVE."
            Else
                MsgBox message, vbExclamation, "SAVE!"
            End If
        Else
            MsgBox message, vbCritical, "Error!"
            Exit Sub
        End If
        
        If formAction = "NEW" Then
            txtAllocationNo.Text = AllocationID
            txtAllctedDate.Text = ALN_DATECREATED
        End If
        
        DoEvents
        Call LockedEntryObject(False)
        Call SetButtonVisibility(False)
        
        cmdAddCashReceive.Enabled = False
        cmdClearCashAllocation.Enabled = False
        cmdRemoveCashAllocation.Enabled = False
        cmdRemoveZeroAlloc.Enabled = False

        cmdAutoAlloc.Enabled = False
        cmdRemoveInvoiceDetails.Enabled = False
 
        cmdShowLedger.Enabled = False
        
        cmdFindCheck.Enabled = False
 
 
        PayType = ""
        
        DidRecordEntry = False
        formAction = ""
        
        
        DoEvents
        'Enabled Edit and Print Button
        Call EnableSelectedButton
        
        Me.Caption = "ALLOCATION FORM"
        
    ElseIf Index = 5 Then 'CANCEL Button
        Ans = MsgBox("Cancel the currrent operation?", vbYesNo + vbQuestion, Me.Caption)
        If Ans = vbNo Then Exit Sub
        DidRecordEntry = False
        
        Call LockedEntryObject(False)
        Call SetButtonVisibility(False)
        
        If formAction = "UPDATE" Then
            tmrLoad.Enabled = True
        Else
            Call ClearEntryForm(False)
        End If
        
        cmdAddCashReceive.Enabled = False
        cmdClearCashAllocation.Enabled = False
        cmdRemoveInvoiceDetails.Enabled = False
        cmdRemoveZeroAlloc.Enabled = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        cmdAutoAlloc.Enabled = False
        cmdRemoveInvoiceDetails.Enabled = False
        cmdTrans(1).Enabled = False
        cmdTrans(3).Enabled = False
        PayType = ""
        
        Me.Caption = "ALLOCATION FORM"
        formAction = ""
    End If

Exit Sub
er:
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->cmdTrans_Click()"
End Sub

Sub LockedEntryObject(DataEntry As Boolean)
    DataEntry = Not DataEntry
    
    cboCustomer.Locked = DataEntry
    txtCashRcvd.Locked = DataEntry
    
    txtRefNo.Locked = DataEntry
    txtAllocAmount.Locked = DataEntry
    txtTaxHeldPercentage.Locked = DataEntry
    txtTaxHeldAmount.Locked = DataEntry
    
    cmdFindCheck.Enabled = Not DataEntry
    cmdShowLedger.Enabled = False
    cmdAutoAlloc.Enabled = False

    If cboCustomer.Locked Then
        cboCustomer.BackColor = txtHoldGrayBackColor.BackColor
        cboAccntType.BackColor = txtHoldGrayBackColor.BackColor
        cboPayType.BackColor = txtHoldGrayBackColor.BackColor
        txtTaxHeldPercentage.BackColor = txtHoldGrayBackColor.BackColor
        txtTaxHeldAmount.BackColor = txtHoldGrayBackColor.BackColor
        txtCashRcvd.BackColor = txtHoldGrayBackColor.BackColor
        
        txtRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtAllocAmount.BackColor = txtHoldGrayBackColor.BackColor
        
    Else
        cboCustomer.BackColor = txtHoldWhiteBackColor.BackColor
        cboAccntType.BackColor = txtHoldWhiteBackColor.BackColor
        cboPayType.BackColor = txtHoldWhiteBackColor.BackColor
        txtTaxHeldPercentage.BackColor = txtHoldWhiteBackColor.BackColor
        txtTaxHeldAmount.BackColor = txtHoldWhiteBackColor.BackColor
        txtCashRcvd.BackColor = txtHoldWhiteBackColor.BackColor
    
        txtRefNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtAllocAmount.BackColor = txtHoldWhiteBackColor.BackColor
        
    End If

End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->NEW
'1 ->EDIT
'2 ->Find
'3 ->Save
'4 ->Cancel
    
    Dim i As Integer
    
    'hide all button
    For i = 0 To 5
        cmdTrans(i).Visible = False
    Next i
    
    If DataEntry Then
        cmdTrans(4).Visible = True
        cmdTrans(5).Visible = True
        
        If formAction = "NEW" Then
            cmdAutoAlloc.Enabled = True
        Else
            cmdAutoAlloc.Enabled = False
        End If
    Else
        cmdAutoAlloc.Enabled = False
        For i = 0 To 3
            cmdTrans(i).Visible = True
        Next i
    End If
End Sub

Private Sub cmdRemoveZeroAlloc_Click()
On Error GoTo er
    Dim Ans As VbMsgBoxResult
    Dim TotaRows As Integer
    Dim hasRemove As Boolean
    Dim i As Integer
    
    If DidRecordEntry = False Then Exit Sub

    Ans = MsgBox("REMOVE ALL ZERO ALLOCATION?", vbQuestion + vbYesNo, "REMOVE ZERO ALLOCATION")
    If Ans = vbNo Then Exit Sub
    
    TotaRows = flxInvoiceDetails.Rows - 1
    '                 0   1              2        3          4            5              6
    '.FormatString = "   |^ACCOUNT TYPE |<STATUS |^REF. NO. |<REF. AMOUNT |<REF. BALANCE |<ALLOCATE "
    '7              8              9           10          11          12          13          14          15
    '|<NEW BALANCE |<BILL AMOUNT |^BILLING NO. |<REF. DATE |<DUE DATE |DUEDATE NF |MONTH DIFF. |JOBORDERID |TRANSDATE_NF "
    For i = TotaRows To 1 Step -1
        If Val(flxInvoiceDetails.TextMatrix(i, 6)) <= 0 Then
            hasRemove = True
            If i = 1 And flxInvoiceDetails.Rows <= 2 Then
                flxInvoiceDetails.TextMatrix(i, 1) = ""
                flxInvoiceDetails.TextMatrix(i, 2) = ""
                flxInvoiceDetails.TextMatrix(i, 3) = ""
                flxInvoiceDetails.TextMatrix(i, 4) = ""
                flxInvoiceDetails.TextMatrix(i, 5) = ""
                flxInvoiceDetails.TextMatrix(i, 6) = ""
                flxInvoiceDetails.TextMatrix(i, 7) = ""
                flxInvoiceDetails.TextMatrix(i, 8) = ""
                flxInvoiceDetails.TextMatrix(i, 9) = ""
                flxInvoiceDetails.TextMatrix(i, 10) = ""
                flxInvoiceDetails.TextMatrix(i, 11) = ""
                flxInvoiceDetails.TextMatrix(i, 12) = ""
                flxInvoiceDetails.TextMatrix(i, 13) = ""
                flxInvoiceDetails.TextMatrix(i, 14) = ""
                flxInvoiceDetails.TextMatrix(i, 15) = ""
            Else
                Call flxInvoiceDetails.RemoveItem(i)
            End If
            DoEvents
        End If
        DoEvents
    Next

    If hasRemove Then
        MsgBox "ALL RECORDS WITH ZERO ALLOCATION HAVE BEEN REMOVED.", vbInformation, "REMOVE ZERO ALLOCATION."
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error -> cmdRemoveZeroAlloc_Click()"
End Sub

Private Sub flxCashAllocation_Click()
    If flxCashAllocation.Rows > 1 Then
        If cboPayType.Text = "CASH PAYMENT" Or cboPayType.Text = "CREDIT CARD PAYMENT" Or cboPayType.Text = "ONLINE PAYMENT" Then
            cmdRemoveCashAllocation.Enabled = True
        End If
    End If
    txtFlxCashAllocation.Visible = False
End Sub

Private Sub flxCashAllocation_DblClick()
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub
    If flxCashAllocation.Rows = 1 Then Exit Sub
    If flxCashAllocation.Cols = 8 Then Exit Sub
    
    '                  0     1            2             3                4         5            6           7               8
    '.FormatString = "    |<CASH NO |<CASH RECEIVED |<TOTAL DISPENSE |^BALANCE |^ALLOCATE |^NEW BALANCE |^DATE RECEIVED |<ACTUAL CASH ID"
    DtlsCol = flxCashAllocation.Col
    If DtlsCol = 5 Then ' ALLOCATE
         DidflxACInUsed = True
         With txtFlxCashAllocation
            .Top = SSTab1.Top + fraHoldTop.Top + flxCashAllocation.CellTop + flxCashAllocation.Top
            .Left = SSTab1.Left + fraHoldTop.Left + flxCashAllocation.CellLeft + flxCashAllocation.Left
            .Width = flxCashAllocation.CellWidth
            .Height = flxCashAllocation.CellHeight
            .Text = Val(Replace(Trim(flxCashAllocation.TextMatrix(flxCashAllocation.Row, 5)), ",", ""))  'ALLOCATE
            .Visible = True
            .BorderStyle = 1
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
        
    ElseIf DtlsCol = 2 Then 'Cash Received
         DidflxACInUsed = True
         If cboPayType.Text = "CASH PAYMENT" Or cboPayType.Text = "CREDIT CARD PAYMENT" Or cboPayType.Text = "ONLINE PAYMENT" Then
             With txtFlxCashAllocation
                .Top = SSTab1.Top + fraHoldTop.Top + flxCashAllocation.CellTop + flxCashAllocation.Top
                .Left = SSTab1.Left + fraHoldTop.Left + flxCashAllocation.CellLeft + flxCashAllocation.Left
                .Width = flxCashAllocation.CellWidth
                .Height = flxCashAllocation.CellHeight
                .Text = Val(Replace(Trim(flxCashAllocation.TextMatrix(flxCashAllocation.Row, 2)), ",", ""))  'Cash Received
                .Visible = True
                .BorderStyle = 1
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        End If
    End If
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->flxCashAllocation_DblClick()"
End Sub

Private Sub flxCashAllocation_KeyPress(KeyAscii As Integer)
On Error GoTo er
    If KeyAscii = 13 Then Call flxCashAllocation_DblClick
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->flxCashAllocation_KeyPress()"
End Sub

Private Sub flxAccnt_DblClick()
On Error GoTo er

    If DidRecordEntry = False Then Exit Sub
    
    If cboPayType.Text = "ACKNOWLEDGEMENT" Then
        If flxAccnt.TextMatrix(flxAccnt.Row, 8) <> "" Then
            MsgBox "FOR PAYMENT TYPE ACKNOWLEDGEMENT," & vbCrLf & "CUSTOMER ACCOUNT WITH ATTACHED BILLING RECORD IS NOT ALLOWED!", vbExclamation, "Validation!"
            Exit Sub
        End If
    
    End If
    '                          0     1             2       3         4           5            6           7         8             9            10          11             12          13           14
    'flxAccnt.FormatString = "   |^ACCOUNT TYPE |<STATUS |^REF. NO. |<REF. DATE |<DUE DATE |<REF. AMOUNT |<REF. BAL. |<BILL NO. |^BILLING AMOUNT. |<SPSORDERID |<DUEDATE NF |<MONTH DIFF. |<JOBORDERID |TRANSDATE_NF"
    
    txtRefNo.Text = flxAccnt.TextMatrix(flxAccnt.Row, 3)
    txtAllocAmount.Text = "0"
    txtBalance.Text = flxAccnt.TextMatrix(flxAccnt.Row, 7)
    txtBillAmnt.Text = flxAccnt.TextMatrix(flxAccnt.Row, 9)
    txtAccntType.Text = flxAccnt.TextMatrix(flxAccnt.Row, 1)
    
    txtAllocAmount.SelStart = 0
    txtAllocAmount.SelLength = Len(txtAllocAmount.Text)
    txtAllocAmount.SetFocus
    
    flxAccnt.Visible = False
    txtAllocAmount.SelStart = 0
    txtAllocAmount.SelLength = Len(txtAllocAmount.Text)
    txtAllocAmount.SetFocus
    
    cmdShowLedger.Enabled = True
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->flxAccnt_DblClick()"
End Sub

Private Sub flxAccnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxAccnt_DblClick
End Sub

Private Sub flxInvoiceDetails_Click()
On Error GoTo er
    flxAccnt.Visible = False
    
    If flxInvoiceDetails.Row > 0 Then
        If flxInvoiceDetails.TextMatrix(flxInvoiceDetails.Row, 1) <> "" Then
            cmdShowLedger.Enabled = True
        End If
    End If
    
    If DidRecordEntry = False Then Exit Sub
    
    
    DoEvents
    If flxInvoiceDetails.TextMatrix(1, 1) = "" Then
       cmdRemoveInvoiceDetails.Enabled = False
       Exit Sub
    End If
    
    If flxInvoiceDetails.Rows = 1 Then Exit Sub
    
    txtFlxCashAllocation.Visible = False
    txtFlxInvDtlsAllocateAmount.Visible = False
    cmdRemoveInvoiceDetails.Enabled = True
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->flxInvoiceDetails_Click()"
End Sub

Private Sub flxInvoiceDetails_DblClick()
On Error GoTo er
    If DidRecordEntry = False Then
        Call cmdShowLedger_Click
        Exit Sub
    End If
    
    If txtRefNo.Locked Then Exit Sub
    If flxCashAllocation.Row <= 0 Then Exit Sub
    '                 0      1         2               3        4           5                6         7           8               9       10
    '.FormatString = "  |^STATUS |<ACCOUNT TYPE |^REF. NO. |<REF. AMOUNT |<REF. BALANCE |<ALLOCATE |<NEW BALANCE |<BILL AMOUNT |^BILLING NO. |<REF. DATE |<DUE DATE
    DtlsCol = flxInvoiceDetails.Col
    If DtlsCol = 6 Then ' ALLOCATE
       DidflxACInUsed = True
       With txtFlxInvDtlsAllocateAmount
            .Top = fraHoldBottom.Top + flxInvoiceDetails.CellTop + flxInvoiceDetails.Top
            .Left = fraHoldBottom.Left + flxInvoiceDetails.CellLeft + flxInvoiceDetails.Left
            .Width = flxInvoiceDetails.CellWidth
            .Height = flxInvoiceDetails.CellHeight
            .Text = Val(Replace(Trim(flxInvoiceDetails.TextMatrix(flxInvoiceDetails.Row, 6)), ",", ""))  'ALLOCATE
            .Visible = True
            .BorderStyle = 1
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
      End With
      
    Else
        Call cmdShowLedger_Click
        
        DoEvents
        txtFlxInvDtlsAllocateAmount.Visible = False
    End If
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->flxInvoiceDetails_DblClick()"
End Sub

Private Sub flxInvoiceDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxInvoiceDetails_DblClick
    End If
End Sub

Private Sub Form_Load()
'cmdTrans Legend
'0 ->NEW
'1 ->EDIT
'2 ->Find
'3 ->PRINT
'4 ->Save
'5 ->Cancel

SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False

cboAccntType.ListIndex = 0

MaxTAXWithHELD = 12

DoEvents

Call SetCboListData

Call flxCashAllocationColsChanger
Call flxInvoiceDetailsColsChanger

If OwnerForm = "frmAllocationSummary" Then
    If Trim(AllocationID) <> "" Then
        tmrLoad.Enabled = True
    End If
End If
End Sub

Sub LockedChckAllocDetail()
 
If formAction = "NEW" Then
    txtPRno.Locked = True
    txtPRDate.Locked = True
    txtAccntName.Locked = True
    txtNmOfBnk.Locked = True
    txtChckTyp.Locked = True
    txtChckStat.Locked = True
    txtChckNo.Locked = True
    txtChckDate.Locked = True
    txtChckAmnt.Locked = True
    txtNewChckBal.Locked = True
    txtChckBal.Locked = True
    
    txtPRno.BackColor = txtHoldGrayBackColor.BackColor
    txtPRDate.BackColor = txtHoldGrayBackColor.BackColor
    txtAccntName.BackColor = txtHoldGrayBackColor.BackColor
    txtNmOfBnk.BackColor = txtHoldGrayBackColor.BackColor
    txtChckTyp.BackColor = txtHoldGrayBackColor.BackColor
    txtChckStat.BackColor = txtHoldGrayBackColor.BackColor
    txtChckNo.BackColor = txtHoldGrayBackColor.BackColor
    txtChckDate.BackColor = txtHoldGrayBackColor.BackColor
    txtChckAmnt.BackColor = txtHoldGrayBackColor.BackColor
    txtNewChckBal.BackColor = txtHoldGrayBackColor.BackColor
    txtChckBal.BackColor = txtHoldGrayBackColor.BackColor
    
ElseIf formAction = "UPDATE" Then
    txtAccntName.Locked = True
    txtNmOfBnk.Locked = True
    txtChckTyp.Locked = True
    txtChckStat.Locked = True
    txtChckNo.Locked = True
    txtChckDate.Locked = True
    txtChckAmnt.Locked = True
    txtNewChckBal.Locked = True
    txtChckBal.Locked = True
    
    txtAccntName.BackColor = txtHoldGrayBackColor.BackColor
    txtNmOfBnk.BackColor = txtHoldGrayBackColor.BackColor
    txtChckTyp.BackColor = txtHoldGrayBackColor.BackColor
    txtChckStat.BackColor = txtHoldGrayBackColor.BackColor
    txtChckNo.BackColor = txtHoldGrayBackColor.BackColor
    txtChckDate.BackColor = txtHoldGrayBackColor.BackColor
    txtChckAmnt.BackColor = txtHoldGrayBackColor.BackColor
    txtNewChckBal.BackColor = txtHoldGrayBackColor.BackColor
    txtChckBal.BackColor = txtHoldGrayBackColor.BackColor

End If

End Sub

Sub ClearEntryForm(Focus1stObject As Boolean)
    cboCustomer.Text = ""
    cboPayType.ListIndex = -1
    txtAllocationNo.Text = "<AUTO>"
    txtAllctedDate.Text = "<AUTO>"
    txtAllctedAmntGenTab.Text = "0.00"
    txtTaxHeldPercentage.Text = "0"
    txtTaxHeldAmount.Text = "0"
    txtNetPayable.Text = "0.00"
    txtCashRcvd.Text = ""
    txtPRno.Text = ""
    txtAccntName.Text = ""
    txtNmOfBnk.Text = ""
    txtChckTyp.Text = ""
    txtChckStat.Text = ""
    txtPRDate.Text = ""
    txtChckNo.Text = ""
    txtChckDate.Text = ""
    txtChckAmnt.Text = ""
    txtNewChckBal.Text = ""
    txtChckBal.Text = ""
    txtFind.Text = ""
    txtAllctedAmntBottom.Text = ""
    txtTTLCash.Text = ""
    txtAllctedAmntGenTab.Text = "0.00"
    txtAllctedAmntBottom.Text = "0.00"
    txtTotalCnt.Text = "0"
    txtRefNo.Text = ""
    txtAllocAmount.Text = "0"
    txtBalance.Text = "0.00"
    txtAccntType.Text = ""
    
    txtPYMTRCVD.Text = "0.00"
    txtPYMTREMAINING.Text = "0.00"
    
    DoEvents
    Call flxInvoiceDetailsColsChanger
    DoEvents
    Call flxCashAllocationColsChanger
    DoEvents
    Call flxAccntColsChanger
    
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.Tab = 0
    If Focus1stObject = True Then
        If SSTab1.Tab = 0 Then cboCustomer.SetFocus
    End If
End Sub

Function ValidateEntry() As Boolean
    Dim i As Integer
    Dim strAmount As String
    ValidateEntry = False
    
    If Trim(cboCustomer.Text) = "" Then
        MsgBox "CUSTOMER NAME IS REQUIRED!", vbExclamation, "Validation!"
        cboCustomer.SetFocus
        Exit Function
    End If
    
    If Trim(cboAccntType.Text) = "" Then
        MsgBox "ACCOUNT TYPE IS REQUIRED!", vbExclamation, "Validation!"
        cboAccntType.SetFocus
        Exit Function
    End If
    
    If Trim(cboPayType.Text) = "" Then
        MsgBox "PAYMENT TYPE IS REQUIRED!", vbExclamation, "Validation!"
        cboPayType.SetFocus
        Exit Function
    End If
    
    If Trim(txtAllocationNo.Text) = "" Then
        MsgBox "APPLICATION NO. IS REQUIRED!", vbExclamation, "Validation!"
        txtAllocationNo.SetFocus
        Exit Function
    End If
       
    If Trim(txtAllctedDate.Text) = "" Then
        MsgBox "ALLOCATED DATE IS REQUIRED!", vbExclamation, "Validation!"
        txtAllctedDate.SetFocus
        Exit Function
    End If
    
    If Trim(txtTaxHeldPercentage.Text) = "" Then
        MsgBox "Tax w/HELD(%) IS REQUIRED!", vbExclamation, "Validation!"
        txtTaxHeldPercentage.SetFocus
        Exit Function
    End If
    
    If Trim(txtTaxHeldAmount.Text) = "" Then
        MsgBox "Tax w/HELD(AMOUNT) IS REQUIRED!", vbExclamation, "Validation!"
        txtTaxHeldAmount.SetFocus
        Exit Function
    End If
    
    If Trim(txtAllctedAmntGenTab.Text) = "" Then
        MsgBox "ALLOCATED AMOUNT IS REQUIRED!", vbExclamation, "Validation!"
        txtAllctedAmntGenTab.SetFocus
        Exit Function
    End If
    
    
    Call ComputeTTLAccountAllocAmount
    'ActiveTaxWHeldObject = "txtTaxHeldPercentage"
    'Call txtTaxHeldPercentage_Change
    'ActiveTaxWHeldObject = ""
    
    If cboPayType.Text = "CASH PAYMENT" Or cboPayType.Text = "ACKNOWLEDGEMENT" Or cboPayType.Text = "CREDIT CARD PAYMENT" Or cboPayType.Text = "ONLINE PAYMENT" Then
        If Trim(flxCashAllocation.TextMatrix(1, 1)) = "" Then
            MsgBox "Cash record is missing!", vbExclamation, "Validation!"
            SSTab1.Tab = 1
            flxCashAllocation.SetFocus
            Exit Function
        End If
        
        Call ComputeTTLCashAllocation
        If Val(Replace(txtNetPayable.Text, ",", "")) <> Val(Replace(txtTTLCash.Text, ",", "")) Then
            MsgBox "Invalid cash allocation!", vbExclamation, "Validation!"
            SSTab1.Tab = 1
            Exit Function
        End If
        
    ElseIf cboPayType.Text = "CHECK PAYMENT" Then
        If Trim(txtPRno.Text) = "" Then
            MsgBox "P.R. NUMBER IS REQUIRED!", vbExclamation, "Validation!"
            SSTab1.Tab = 2
            txtPRno.SetFocus
            Exit Function
        End If
    
        If Trim(txtPRDate.Text) = "" Then
            MsgBox "P.R. DATE IS REQUIRED!", vbExclamation, "Validation!"
            SSTab1.Tab = 2
            txtPRDate.SetFocus
            Exit Function
        End If
        
        If Trim(txtChckNo.Text) = "" Then
            MsgBox "CHECK NO. IS REQUIRED!", vbExclamation, "Validation!"
            SSTab1.Tab = 2
            txtChckNo.SetFocus
            Exit Function
        End If
        
 
        If Trim(txtNewChckBal.Text) = "" Then
            MsgBox "NEW CHECK BALANCE IS REQUIRED!", vbExclamation, "Validation!"
            SSTab1.Tab = 2
            txtNewChckBal.SetFocus
            Exit Function
        End If
        
        If Val(Replace(txtNewChckBal.Text, ",", "")) < 0 Then
            MsgBox "INSUFFICIENT CHECK BALANCE!", vbExclamation, "Validation!"
            SSTab1.Tab = 2
            txtNewChckBal.SetFocus
            Exit Function
        End If
    End If
    
    
    DoEvents
    For i = 1 To flxInvoiceDetails.Rows - 1 Step 1
        'For Negative Balance Validation
        strAmount = Replace(flxInvoiceDetails.TextMatrix(i, 6), ",", "")
        strAmount = Trim(strAmount)
        If strAmount = "" Then strAmount = "0.00"
        If Val(strAmount) <= 0 Then
           MsgBox "Please remove customer account with zero(0) allocation, please check!", vbExclamation, "Validation!"
           flxInvoiceDetails.TopRow = i
           flxInvoiceDetails.RowSel = i
           Exit Function
        End If
    Next i
    
    ValidateEntry = True
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    flxAccnt.Visible = False
End Sub

 
Private Sub Form_Resize()
On Error Resume Next
If Me.Top <= 0 Then
    Me.Top = 200
    Me.Left = 200
End If

If Me.Width <= 15000 Then Me.Width = 15000
If Me.Height <= 10890 Then Me.Height = 10890
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DidRecordEntry = True Then
        Cancel = 1
        MsgBox "PLEASE CANCEL THE CURRENT TRANSACTION!", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    OwnerForm = ""
    AllocationID = ""
    Set frmAllocation = Nothing
End Sub

Private Sub fraHoldBottom_Click()
    flxAccnt.Visible = False
End Sub


Private Sub RXResizer1_ResizeComplete()
On Error Resume Next
'cmdTrans Legend
'0 ->NEW
'1 ->EDIT
'2 ->Find
'3 ->Print
'4 ->Save
'5 ->Cancel
    
    'Align Save Button
    cmdTrans(4).Left = cmdTrans(0).Left
    cmdTrans(4).Top = cmdTrans(0).Top
    
    'Align Cancel Button
    cmdTrans(5).Left = Me.Width - cmdTrans(5).Width - 200
    cmdTrans(5).Top = cmdTrans(0).Top
    
    fraLedger.Left = SSTab1.Left
    fraLedger.Top = SSTab1.Top
    fraLedger.Width = SSTab1.Width
    
    lblCheckBal.Left = txtChckBal.Left - 1560
    lblCust.Left = cboCustomer.Left - 1440
    lblAccntType.Left = cboAccntType.Left - 1440
    lblPayType.Left = cboPayType.Left - 1440
    lblCashRcvd.Left = txtCashRcvd.Left - 1440
     
    
    fraAllocation.Left = (Me.Width / 2) - (fraAllocation.Width / 2)
    fraAllocation.Top = (Me.Height / 2) - (fraAllocation.Height / 2)
    
    SStabFindRetail.Left = (Me.Width / 2) - (SStabFindRetail.Width / 2)
    SStabFindRetail.Top = (Me.Height / 2) - (SStabFindRetail.Height / 2)
    
     
    '                                   0   1              2        3          4            5              6          7              8              9           10          11          12          13          14          15
    'flxInvoiceDetails.FormatString = "   |^ACCOUNT TYPE |<STATUS |^REF. NO. |>REF. AMOUNT |>REF. BALANCE |>ALLOCATE |>NEW BALANCE |>BILL AMOUNT |^BILLING NO. |<REF. DATE |<DUE DATE |DUEDATE NF |MONTH DIFF. |JOBORDERID |TRANSDATE_NF "
    flxInvoiceDetails.Col = 2
    txtTotalCnt.Left = flxInvoiceDetails.CellLeft + flxInvoiceDetails.Left
    txtTotalCnt.Width = flxInvoiceDetails.ColWidth(2) '2=>STATUS
    lblTotalCnt.Left = txtTotalCnt.Left
    lblTotalCnt.Width = txtTotalCnt.Width
    
    
    flxInvoiceDetails.Col = 6
    txtAllctedAmntBottom.Left = flxInvoiceDetails.CellLeft + flxInvoiceDetails.Left
    txtAllctedAmntBottom.Width = flxInvoiceDetails.ColWidth(6) '6=>ALLOCATE
    lblTotalAllocate.Left = txtAllctedAmntBottom.Left
    lblTotalAllocate.Width = txtAllctedAmntBottom.Width
    If Round(txtAllctedAmntBottom.Left / Me.Width, 2) > 0.48 Then
        flxInvoiceDetails.Col = 4
        txtAllctedAmntBottom.Left = flxInvoiceDetails.CellLeft + flxInvoiceDetails.Left
        txtAllctedAmntBottom.Width = flxInvoiceDetails.ColWidth(4) '4=>REF. AMOUNT
        lblTotalAllocate.Left = txtAllctedAmntBottom.Left
        lblTotalAllocate.Width = txtAllctedAmntBottom.Width
        
        
        flxInvoiceDetails.Col = 5
        txtPYMTRCVD.Left = flxInvoiceDetails.CellLeft + flxInvoiceDetails.Left
        txtPYMTRCVD.Width = flxInvoiceDetails.ColWidth(5) '5=>REF. BALANCE
        lblPYMTRCVD.Left = txtPYMTRCVD.Left
        lblPYMTRCVD.Width = txtPYMTRCVD.Width
        
        
        flxInvoiceDetails.Col = 6
        txtPYMTREMAINING.Left = flxInvoiceDetails.CellLeft + flxInvoiceDetails.Left
        txtPYMTREMAINING.Width = flxInvoiceDetails.ColWidth(6) '6=>ALLOCATE
        lblPYMTREMAINING.Left = txtPYMTREMAINING.Left
        lblPYMTREMAINING.Width = txtPYMTREMAINING.Width
    Else
        flxInvoiceDetails.Col = 7
        txtPYMTRCVD.Left = flxInvoiceDetails.CellLeft + flxInvoiceDetails.Left
        txtPYMTRCVD.Width = flxInvoiceDetails.ColWidth(7) '7=>NEW BALANCE
        lblPYMTRCVD.Left = txtPYMTRCVD.Left
        lblPYMTRCVD.Width = txtPYMTRCVD.Width
        
        
        flxInvoiceDetails.Col = 8
        txtPYMTREMAINING.Left = flxInvoiceDetails.CellLeft + flxInvoiceDetails.Left
        txtPYMTREMAINING.Width = flxInvoiceDetails.ColWidth(8) '8=>BILL AMOUNT
        lblPYMTREMAINING.Left = txtPYMTREMAINING.Left
        lblPYMTREMAINING.Width = txtPYMTREMAINING.Width
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    txtFlxCashAllocation.Visible = False
    If SSTab1.Tab = 1 Then
        Call EnableFields
        txtCashRcvd.SetFocus
    End If
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    flxAccnt.Visible = False
End Sub

Private Sub tmrAutoAllocate_Timer()
On Error GoTo er
    Dim DtlsRow As Integer
    Dim tmpRs As New ADODB.Recordset
    Dim ALLOCATEAMOUNT As String
    Dim AllocationID As Integer
    Dim ACCNTTYPE As String
    
    If tmrAutoAllocate.Enabled = False Then GoTo ExitGate
    
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    
    tmrAutoAllocate.Enabled = False
    
    ALLOCATEAMOUNT = Val(Replace(txtAllocateAmnt.Text, ",", ""))
    BillingNo = arrCurBill(cboBillingNo.ListIndex, 0)
    AllocationID = "-1"
    
    Call flxInvoiceDetailsColsChanger
    
    'cboAccntType is equal to the cboBillingNo
    ACCNTTYPE = arrCurBill(cboBillingNo.ListIndex, 4)
    
    'ACCNTTYPE = "SERVICE"
    'ACCNTTYPE = "SPARE PARTS"
    cboAccntType.ListIndex = modTools.getListIndex(cboAccntType, ACCNTTYPE)
    SSTab1.Tab = 0
    
    DoEvents
    cmdRemoveZeroAlloc.Enabled = False
    Call modAllocation.sp_allocation_auto(SysBranchID, CUSTOMERID, BillingNo, ALLOCATEAMOUNT, AllocationID, UserID, tmpRs)
    'ACCOUNTSTATUS, ACCOUNT_TYPE, REF_NO, REFAMOUNT, OSBAL, ALLOCAMOUNT, BALANCE, BILLAMOUNT, BILLINGNO, DUEDATE, TRANSDATE,TRANSDATE_NF, DUEDATE_NF, MONTHDIFF, JOBORDERID, Level, Message
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount <= 0 Then
            tmpRs.Close
            LoadflxAccount = False
            If tmpRs.State <> 0 Then tmpRs.Close
            GoTo ExitGate
        End If
    Else
        If tmpRs.State <> 0 Then tmpRs.Close
        GoTo ExitGate
    End If
    
    LoadflxAccount = True
    DtlsRow = 1
    
    tmpRs.MoveFirst
    
    
    If tmpRs!level = "Error" Then
        MsgBox tmpRs!message, vbCritical, "Error->tmrAutoAllocate_Timer()"
        If tmpRs.State <> 0 Then tmpRs.Close
        GoTo ExitGate
    End If
    
    With tmpRs
        Do While Not tmpRs.EOF
            '                    0         1            2        3            4          5              6          7              8              9           10           11          12       13           14       15
            '.FormatString = "      |^ACCOUNT TYPE |<STATUS |^REF. NO. |<REF. AMOUNT |<REF. BALANCE |<ALLOCATE |<NEW BALANCE |<BILL AMOUNT |^BILLING NO. |<REF. DATE |<DUE DATE |DUEDATE_NF |MONTHDIFF |JOBORDERID |TRANSDATE_NF "
            If DtlsRow = 1 Then
                flxInvoiceDetails.TextMatrix(DtlsRow, 0) = ""
                flxInvoiceDetails.TextMatrix(DtlsRow, 1) = !ACCOUNT_TYPE
                flxInvoiceDetails.TextMatrix(DtlsRow, 2) = !ACCOUNTSTATUS
                flxInvoiceDetails.TextMatrix(DtlsRow, 3) = !REF_NO
                flxInvoiceDetails.TextMatrix(DtlsRow, 4) = Format(!REFAMOUNT, "#,##0.00")
                flxInvoiceDetails.TextMatrix(DtlsRow, 5) = Format(!OSbal, "#,##0.00")
                flxInvoiceDetails.TextMatrix(DtlsRow, 6) = Format(!AllocAmount, "#,##0.00")
                flxInvoiceDetails.TextMatrix(DtlsRow, 7) = Format(!Balance, "#,##0.00")
                flxInvoiceDetails.TextMatrix(DtlsRow, 8) = Format(!BILLAMOUNT, "#,##0.00")
                flxInvoiceDetails.TextMatrix(DtlsRow, 9) = !BillingNo
                flxInvoiceDetails.TextMatrix(DtlsRow, 10) = !TRANSDATE
                flxInvoiceDetails.TextMatrix(DtlsRow, 11) = !DUEDATE
                flxInvoiceDetails.TextMatrix(DtlsRow, 12) = !DUEDATE_NF
                flxInvoiceDetails.TextMatrix(DtlsRow, 13) = !MONTHDIFF
                flxInvoiceDetails.TextMatrix(DtlsRow, 14) = !JOBORDERID
                flxInvoiceDetails.TextMatrix(DtlsRow, 15) = !TRANSDATE_NF
            Else
                flxInvoiceDetails.AddItem _
                    vbTab & !ACCOUNT_TYPE & _
                    vbTab & !ACCOUNTSTATUS & _
                    vbTab & !REF_NO & _
                    vbTab & Format(!REFAMOUNT, "#,##0.00") & _
                    vbTab & Format(!OSbal, "#,##0.00") & _
                    vbTab & Format(!AllocAmount, "#,##0.00") & _
                    vbTab & Format(!Balance, "#,##0.00") & _
                    vbTab & Format(!BILLAMOUNT, "#,##0.00") & _
                    vbTab & !BillingNo & _
                    vbTab & !TRANSDATE & _
                    vbTab & !DUEDATE & _
                    vbTab & !DUEDATE_NF & _
                    vbTab & !MONTHDIFF & _
                    vbTab & !JOBORDERID & _
                    vbTab & !TRANSDATE_NF
            End If
            DtlsRow = DtlsRow + 1
            DoEvents
            DoEvents
            tmpRs.MoveNext
        Loop
    End With
    If tmpRs.State <> 0 Then tmpRs.Close
    cmdRemoveZeroAlloc.Enabled = True
    fraAllocation.Visible = False
    txtAllocateAmnt.Text = ""
    txtAllocateAmnt.Locked = True
    txtAllocateAmnt.BackColor = txtHoldGrayBackColor.BackColor
    chkAutoAllocate.Value = 0
    Call EnableFields
    
    'Lock cboCustomer
    cboCustomer.Locked = True
    cboCustomer.BackColor = txtHoldGrayBackColor.BackColor
    cboAccntType.Locked = True
    cboAccntType.BackColor = txtHoldGrayBackColor.BackColor
    
    Call ComputeTTLAccountAllocAmount
    
    'Count Total Invoice Details
    txtTotalCnt.Text = flxInvoiceDetails.Rows - 1
    
    Call ComputeNetPayable
    
    Call ComputeNewCheckBalance
    
    cmdRemoveZeroAlloc.Enabled = True
    
ExitGate:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
Exit Sub
er:
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    
    MsgBox Err.Description, vbCritical, "Error->tmrAutoAllocate_Timer()"
End Sub

Private Sub tmrAutoAllocateLoadFrame_Timer()
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim BILLINGAMNT As Double
    Dim InvBal As Double
    Dim cbolistIndex As Integer
    
    tmrAutoAllocateLoadFrame.Enabled = False
    
    
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    
    DoEvents
    Call modAllocation.sp_allocation_list_curbilling(SysBranchID, CUSTOMERID, tmpRs)
    'BILLINGNO, TTLBILL, BILLINGCAPTION, INVBALANCE, ACCOUNT_TYPE
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount <= 0 Then
            MsgBox "No Record found!", vbExclamation, "Validation!"
            GoTo ExitGate
        End If
    End If
    DoEvents
    Call modTools.addToListObjectWithArray(cboBillingNo, tmpRs, arrCurBill, "BILLINGCAPTION", 5)
    If tmpRs.State <> 0 Then tmpRs.Close
    
    If cboBillingNo.ListCount > 0 Then cboBillingNo.ListIndex = 0
    
    DoEvents
    fraAllocation.Visible = True
    
    cmdTrans(0).Enabled = False
    cmdTrans(1).Enabled = False
    cmdTrans(2).Enabled = False
    cmdTrans(3).Enabled = False
    SSTab1.Enabled = False
    cmdShowLedger.Enabled = False
    
    If formAction = "NEW" Then
        cmdAutoAlloc.Enabled = True
    Else
        cmdAutoAlloc.Enabled = False
    End If
    
    cmdRemoveZeroAlloc.Enabled = False
    fraHoldBottom.Enabled = False
    
    
ExitGate:
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    
Exit Sub
er:
    If LoadOnProcess Then
        Call Unload(frm_OnProcess)
        Set frm_OnProcess = Nothing
    End If
    
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    tmrAutoAllocateLoadFrame.Enabled = False
    MsgBox Err.Description, vbCritical, "Error->tmrAutoAllocateFrame_Timer()"
End Sub

Private Sub tmrLoad_Timer()
On Error GoTo er
    tmrLoad.Enabled = False
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    
    If Trim(AllocationID) = "" Then GoTo ExitGate
    
    'Call the sub routine to find and display allocation record
    Call FindAndLoadRecord
    
    DoEvents
    Call EnableFields
    
    DoEvents
    Call RXResizer1_ResizeComplete
ExitGate:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
Exit Sub
er:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    tmrLoad.Enabled = False
    MsgBox Err.Description, vbCritical, "Error->tmrLoad_Timer()"
End Sub

Private Sub tmrLoadAccount_Timer()
On Error GoTo er
    Dim DtlsRow As Integer
    Dim tmpRs As New ADODB.Recordset
    Dim ACCNTTYPE As String
    
    If DidRecordEntry = False Then
        tmrLoadAccount.Enabled = False
        GoTo ExitGate
    End If
    If tmrLoadAccount.Enabled = False Then GoTo ExitGate
    tmrLoadAccount.Enabled = False
    
    
    'Clear List of On Customer Account
    Call flxAccntColsChanger
    
    ACCNTTYPE = cboAccntType.Text
    
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    
    If cboCustomer.ListIndex < 0 Then
        cboCustomer.Text = ""
        MsgBox "This customer name does not exist!", vbExclamation, "Validation!"
        GoTo ExitGate
    End If
    
    If Trim(cboAccntType.Text) = "" Then
        MsgBox "Account type is required!", vbExclamation, "Validation!"
        GoTo ExitGate
    End If
    
    CUSTOMERID = arrCustomer(cboCustomer.ListIndex, 0)
    
    DoEvents
    
    Call modAllocation.sp_onaccterms_list_allocation(SysBranchID, CUSTOMERID, Trim(cboAccntType.Text), tmpRs)
    'ACCOUNTSTATUS, ACCOUNT_TYPE, REFNO, DUEDATE, TRANSDATE, ACCTAMOUNT, ACCTBAL, BILLINGNO, BILLAMOUNT, DUEDATE_NF, MONTHDIFF, SPSORDERID, JOBORDERID,TRANSDATE_NF
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount <= 0 Then
            tmpRs.Close
            tmrLoadAccount.Enabled = False
            LoadflxAccount = False
            GoTo ExitGate
        End If
    Else
       GoTo ExitGate
    End If
    
    LoadflxAccount = True
    DtlsRow = 1
    
    tmpRs.MoveFirst
    With tmpRs
    Do While Not tmpRs.EOF
        '                          0     1             2       3         4           5            6           7         8             9            10          11             12          13           14
        'flxAccnt.FormatString = "   |^ACCOUNT TYPE |<STATUS |^REF. NO. |<REF. DATE |<DUE DATE |<REF. AMOUNT |<REF. BAL. |<BILL NO. |^BILLING AMOUNT. |<SPSORDERID |<DUEDATE NF |<MONTH DIFF. |<JOBORDERID |TRANSDATE_NF"
        
         If DtlsRow = 1 Then
             flxAccnt.TextMatrix(DtlsRow, 0) = ""
             flxAccnt.TextMatrix(DtlsRow, 1) = !ACCOUNT_TYPE
             flxAccnt.TextMatrix(DtlsRow, 2) = !ACCOUNTSTATUS
             flxAccnt.TextMatrix(DtlsRow, 3) = !RefNo
             flxAccnt.TextMatrix(DtlsRow, 4) = !TRANSDATE
             flxAccnt.TextMatrix(DtlsRow, 5) = !DUEDATE
             flxAccnt.TextMatrix(DtlsRow, 6) = Format(!AcctAmount, "#,##0.00")
             flxAccnt.TextMatrix(DtlsRow, 7) = Format(!ACCTBAL, "#,##0.00")
             flxAccnt.TextMatrix(DtlsRow, 8) = !BillingNo
             flxAccnt.TextMatrix(DtlsRow, 9) = Format(!BILLAMOUNT, "#,##0.00")
             flxAccnt.TextMatrix(DtlsRow, 10) = !SPSORDERID
             flxAccnt.TextMatrix(DtlsRow, 11) = !DUEDATE_NF
             flxAccnt.TextMatrix(DtlsRow, 12) = !MONTHDIFF
             flxAccnt.TextMatrix(DtlsRow, 13) = !JOBORDERID
             flxAccnt.TextMatrix(DtlsRow, 14) = !TRANSDATE_NF
             
         Else
             flxAccnt.AddItem _
                vbTab & !ACCOUNT_TYPE & _
                vbTab & !ACCOUNTSTATUS & _
                vbTab & !RefNo & _
                vbTab & !TRANSDATE & _
                vbTab & !DUEDATE & _
                vbTab & Format(!AcctAmount, "#,##0.00") & _
                vbTab & Format(!ACCTBAL, "#,##0.00") & _
                vbTab & !BillingNo & _
                vbTab & Format(!BILLAMOUNT, "#,##0.00") & _
                vbTab & !SPSORDERID & _
                vbTab & !DUEDATE_NF & _
                vbTab & !MONTHDIFF & _
                vbTab & !JOBORDERID & _
                vbTab & !TRANSDATE_NF
         End If
            DtlsRow = DtlsRow + 1
            tmpRs.MoveNext
         Loop
    End With
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    If formAction <> "UPDATE" Then
        cboPayType.ListIndex = -1
        cboPayType.Locked = False
        cboPayType.BackColor = txtHoldWhiteBackColor.BackColor
        Call flxCashAllocationColsChanger
    End If
ExitGate:
    If cboPayType.Text = "CHECK PAYMENT" Or cboPayType.Text = "CASH PAYMENT" Or cboPayType.Text = "CREDIT CARD PAYMENT" Or cboPayType.Text = "ONLINE PAYMENT" Then
        If formAction = "NEW" Then
            cmdAutoAlloc.Enabled = True
        Else
            cmdAutoAlloc.Enabled = False
        End If
    End If
 
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    
Exit Sub
er:
    tmrLoadAccount.Enabled = False
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    MsgBox Err.Description, vbCritical, "Error->tmrLoadAccount_Timer()"
End Sub

Private Sub tmrLoadLedger_Timer()
On Error GoTo er
    Dim tmpRs  As New ADODB.Recordset

    tmrLoadLedger.Enabled = False
    Call modAllocation.sp_onaccterms_list_ledger(SysBranchID, CUSTOMERID, cboAccntType.Text, txtRefNo.Text, tmpRs)
    
    DoEvents
    If tmpRs.RecordCount <= 0 Then
        Call flxAccntColsChanger
    Else
        flxLedger.FixedCols = 1
        Set flxLedger.DataSource = tmpRs
        DoEvents
        Call flxAccntColsChanger
    End If
    
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    
Exit Sub
er:
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    tmrLoadLedger.Enabled = False
    MsgBox Err.Description, vbCritical, "Error->tmrLoadLedger_Timer()"
End Sub

Private Sub txtAccntName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAccntName, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtNmOfBnk.SetFocus
End Sub

Private Sub txtAccntType_Click()
    flxAccnt.Visible = False
End Sub

Private Sub txtAllctedAmntGenTab_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAllctedAmntGenTab, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then txtTaxHeldPercentage.SetFocus
End Sub

Private Sub txtAllctedDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAllctedDate, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtAllctedAmntGenTab.SetFocus
End Sub

Private Sub txtAllocAmount_Click()
    txtAllocAmount.SelStart = 0
    txtAllocAmount.SelLength = Len(txtAllocAmount.Text)
    txtAllocAmount.SetFocus
    
    flxAccnt.Visible = False
End Sub

Private Sub txtAllocAmount_KeyPress(KeyAscii As Integer)
On Error GoTo er
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtCashRcvd, KeyAscii, NumbersWithDecimalPoint)

    If DidRecordEntry = False Then Exit Sub
    If txtAllocAmount.Text = "" Then Exit Sub
 
    Dim STATUS As String
    Dim ACCNTTYPE As String
    Dim RefNo As String
    Dim REFAMNT As Double
    Dim REFBAL As Double
    Dim BILLAMNT As Double
    Dim BillingNo As String
    Dim REFDATE As String
    Dim DUEDATE As String
    Dim Balance As Double
    Dim SPSORDERID As String
    Dim DUEDATENF As String
    Dim MONTHDIFF As String
    Dim JOBORID As String
    Dim TRANSDATE_NF As String
    Dim AllocAmnt As Double
    Dim NewActBal As Double
    
    AllocAmnt = Val(Replace(txtAllocAmount.Text, ",", ""))
    Balance = Val(Replace(txtBalance.Text, ",", " "))
    NewActBal = Balance - AllocAmnt
    If KeyAscii = 13 Then
    '                          0     1             2        3          4           5          6           7           8            9
    'flxAccnt.FormatString = "   |^ACCOUNT TYPE |<STATUS |^REF. NO. |<REF. DATE |<DUE DATE |<REF. AMOUNT |<REF. BAL. |<BILL NO. |^BILLING AMOUNT.
    '  10          11             12          13           14
    '|<SPSORDERID |<DUEDATE NF |<MONTH DIFF. |<JOBORDERID |TRANSDATE_NF"
        
    ACCNTTYPE = flxAccnt.TextMatrix(flxAccnt.Row, 1) ' 1=>ACCOUNT TYPE
    STATUS = flxAccnt.TextMatrix(flxAccnt.Row, 2) ' 2=>STATUS
    RefNo = flxAccnt.TextMatrix(flxAccnt.Row, 3) ' 3=>REF. NO.
    REFAMNT = flxAccnt.TextMatrix(flxAccnt.Row, 6) ' 6 =>REF. AMOUNT
    REFBAL = flxAccnt.TextMatrix(flxAccnt.Row, 7) ' 7 =>REF. BAL.
    BillingNo = flxAccnt.TextMatrix(flxAccnt.Row, 8) ' 8 =>BILL NO.
    BILLAMNT = flxAccnt.TextMatrix(flxAccnt.Row, 9) ' 9=>BILLING AMOUNT.
    REFDATE = flxAccnt.TextMatrix(flxAccnt.Row, 4) ' 4=>REF. DATE
    DUEDATE = flxAccnt.TextMatrix(flxAccnt.Row, 5) ' 5=>DUE DATE
    DUEDATENF = flxAccnt.TextMatrix(flxAccnt.Row, 11) '11=>DUEDATE NF
    MONTHDIFF = flxAccnt.TextMatrix(flxAccnt.Row, 12)  '12=>MONTH DIFF.
    JOBORID = flxAccnt.TextMatrix(flxAccnt.Row, 13) '13=>JOBORDERID
    TRANSDATE_NF = flxAccnt.TextMatrix(flxAccnt.Row, 14) '14=>TRANSDATE_NF
    
    If modTools.getFlxRowIndex(flxInvoiceDetails, 1, RefNo, 3) >= 1 Then
        MsgBox "Record already exist!", vbExclamation, "Validation!"
        Exit Sub
    End If

   If AllocAmnt > Balance Then
        MsgBox "Allocate amount is greater than balance!", vbExclamation, "Unable to allocate!"
        txtAllocAmount.SelStart = 0
        txtAllocAmount.SelLength = Len(txtFlxInvDtlsAllocateAmount.Text)
        txtAllocAmount.SetFocus
        Exit Sub
        
    ElseIf AllocAmnt = 0 Or txtAllocAmount.Text = "" Then
        MsgBox "Please input allocate amount", vbExclamation, "Validation!"
        txtAllocAmount.SelStart = 0
        txtAllocAmount.SelLength = Len(txtFlxInvDtlsAllocateAmount.Text)
        txtAllocAmount.SetFocus
        Exit Sub
        
    End If

    '     1             2          3        4               5           6           7               8           9           10          11          12          13          14          15
    '"|^ACCOUNT TYPE |<STATUS |^REF. NO. |<REF. AMOUNT |<REF. BALANCE |<ALLOCATE |<NEW BALANCE |<BILL AMOUNT |^BILLING NO. |<REF. DATE |<DUE DATE |DUEDATE_NF |MONTHDIFF |JOBORDERID |TRANSDATE_NF "
    DtlsRow = 1
    If Trim(flxInvoiceDetails.TextMatrix(DtlsRow, 1)) = "" Then
        flxInvoiceDetails.TextMatrix(DtlsRow, 0) = ""
        flxInvoiceDetails.TextMatrix(DtlsRow, 1) = ACCNTTYPE
        flxInvoiceDetails.TextMatrix(DtlsRow, 2) = STATUS
        flxInvoiceDetails.TextMatrix(DtlsRow, 3) = RefNo
        flxInvoiceDetails.TextMatrix(DtlsRow, 4) = Format(REFAMNT, "#,##0.00")
        flxInvoiceDetails.TextMatrix(DtlsRow, 5) = Format(REFBAL, "#,##0.00")
        flxInvoiceDetails.TextMatrix(DtlsRow, 6) = Format(txtAllocAmount.Text, "#,##0.00")
        flxInvoiceDetails.TextMatrix(DtlsRow, 7) = Format(NewActBal, "#,##0.00")
        flxInvoiceDetails.TextMatrix(DtlsRow, 8) = Format(BILLAMNT, "#,##0.00")
        flxInvoiceDetails.TextMatrix(DtlsRow, 9) = BillingNo
        flxInvoiceDetails.TextMatrix(DtlsRow, 10) = REFDATE
        flxInvoiceDetails.TextMatrix(DtlsRow, 11) = DUEDATE
        flxInvoiceDetails.TextMatrix(DtlsRow, 12) = DUEDATENF
        flxInvoiceDetails.TextMatrix(DtlsRow, 13) = MONTHDIFF
        flxInvoiceDetails.TextMatrix(DtlsRow, 14) = JOBORID
        flxInvoiceDetails.TextMatrix(DtlsRow, 15) = TRANSDATE_NF

    Else
        flxInvoiceDetails.AddItem "" & _
            vbTab & ACCNTTYPE & _
            vbTab & STATUS & _
            vbTab & RefNo & _
            vbTab & Format(REFAMNT, "#,##0.00") & _
            vbTab & Format(REFBAL, "#,##0.00") & _
            vbTab & Format(txtAllocAmount.Text, "#,##0.00") & _
            vbTab & Format(NewActBal, "#,##0.00") & _
            vbTab & Format(BILLAMNT, "#,##0.00") & _
            vbTab & BillingNo & _
            vbTab & REFDATE & _
            vbTab & DUEDATE & _
            vbTab & DUEDATENF & _
            vbTab & MONTHDIFF & _
            vbTab & JOBORID & _
            vbTab & TRANSDATE_NF
    End If
        
    Call ComputeTTLAccountAllocAmount
    
    DoEvents
    ActiveTaxWHeldObject = "txtTaxHeldPercentage"
    Call txtTaxHeldPercentage_Change
    ActiveTaxWHeldObject = ""
    
    'Count Total Invoice Details
    txtTotalCnt.Text = flxInvoiceDetails.Rows - 1
    
    txtRefNo.Text = ""
    txtAllocAmount.Text = ""
    txtBalance.Text = ""
    txtBillAmnt.Text = ""
    txtAccntType.Text = ""
    
    If Trim(cboPayType.Text) = "CASH PAYMENT" Or Trim(cboPayType.Text) = "ACKNOWLEDGEMENT" Or Trim(cboPayType.Text) = "CREDIT CARD PAYMENT" Or Trim(cboPayType.Text) = "ONLINE PAYMENT" Then
        
    Else
        Call ComputeNewCheckBalance
    End If
    
    flxAccnt.Visible = False
    cmdShowLedger.Enabled = False
    cboAccntType.Locked = True
    cboAccntType.BackColor = txtHoldGrayBackColor.BackColor
    cboCustomer.Locked = True
    cboCustomer.BackColor = txtHoldGrayBackColor.BackColor
    
    cmdRemoveInvoiceDetails.Enabled = True
    cmdRemoveZeroAlloc.Enabled = True
    
    DoEvents
    txtRefNo.SetFocus
End If
Exit Sub
er:
     MsgBox Err.Description, vbCritical, "Error->txtAllocAmount_KeyPress()"
End Sub

Sub ComputeNewCheckBalance()
On Error GoTo er

    Dim NetPay As Double
    Dim ChckBal As Double
    Dim NetPayable As Double
       
    If txtChckBal.Text = "" Then Exit Sub
    
    ChckBal = Replace(Format(txtChckBal.Text), ",", "")
    txtPYMTRCVD.Text = Format(ChckBal, "#,##0.00")
    
    NetPayable = Replace(Format(txtNetPayable.Text), ",", "")

    NetPay = ChckBal - NetPayable
    
    txtNewChckBal.Text = Format(NetPay, "#,##0.00")
    If NetPay >= 0 Then
        txtNewChckBal.ForeColor = txtHoldWhiteBackColor.ForeColor
    Else
        txtNewChckBal.ForeColor = txtHoldGrayBackColor.ForeColor
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->ComputeNewCheckBalance()"
End Sub

Private Sub txtAllocateAmnt_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAllocateAmnt, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then cmdPopUpAllocate.SetFocus
End Sub

Private Sub txtAllocationNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAllocationNo, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then txtAllctedDate.SetFocus
End Sub

Private Sub txtBalance_Click()
    flxAccnt.Visible = False
End Sub

Private Sub txtBillAmnt_Click()
    flxAccnt.Visible = False
End Sub

Private Sub txtCashRcvd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtCashRcvd, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then Call cmdAddCashReceive_Click
End Sub

Private Sub txtChckAmnt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtChckAmnt, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then txtNewChckBal.SetFocus
End Sub

Private Sub txtChckBal_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtChckBal, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then txtPRno.SetFocus
End Sub

Private Sub txtChckDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtChckDate, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtChckAmnt.SetFocus
End Sub

Private Sub txtChckNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtChckNo, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then txtChckDate.SetFocus
End Sub

Private Sub txtChckStat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtChckStat, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtPRDate.SetFocus
End Sub

Private Sub txtChckTyp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtChckTyp, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtChckStat.SetFocus
End Sub

Private Sub txtFind_Click()
    flxAccnt.Visible = False
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
On Error GoTo er
    If KeyAscii = 13 Then
        Call cmdFind_Click
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error ->txtFind_KeyPress()"
End Sub

Private Sub txtFindApplicationNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtFindApplicationNo.Enabled = False Then Exit Sub
        txtFindApplicationNo.Enabled = False
        
        Call cmdPopUpFind_Click
        txtFindApplicationNo.Enabled = True
    End If
End Sub

Private Sub txtFlxCashAllocation_KeyPress(KeyAscii As Integer)
On Error GoTo er
    Dim AllocatedAmount As Double
    Dim CASHRECEIVED As Double
    Dim Balance As Double
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxCashAllocation, KeyAscii, NumbersWithDecimalPoint)
    
    '                  0     1            2             3                4         5            6           7               8
    '.FormatString = "    |<CASH NO |<CASH RECEIVED |<TOTAL DISPENSE |^BALANCE |^ALLOCATE |^NEW BALANCE |^DATE RECEIVED |<ACTUAL CASH ID"
    If KeyAscii = 13 Then
        If DidRecordEntry = False Then Exit Sub
        If flxCashAllocation.Row = 0 Then Exit Sub
        If txtFlxCashAllocation.Text = "" Then Exit Sub

        If IsNumeric(Trim(txtFlxCashAllocation.Text)) = False Then
            MsgBox "Invalid allocated amount!" & vbCrLf & "Numeric is required.", vbExclamation, "Validation!"
            txtFlxCashAllocation.SetFocus
            Exit Sub
        End If
        
        DtlsCol = flxCashAllocation.Col
        If DtlsCol = 2 Then 'CASH RECEIVED
            If cboPayType.Text = "CASH PAYMENT" Or cboPayType.Text = "CREDIT CARD PAYMENT" Or cboPayType.Text = "ONLINE PAYMENT" Then
                CASHRECEIVED = Val(Replace(flxCashAllocation.TextMatrix(flxCashAllocation.Row, 5), ",", ""))  'ALLOCATE
                If Val(txtFlxCashAllocation.Text) < CASHRECEIVED Then
                    MsgBox "Cash receive is lesser than allocated amount!", vbExclamation, "Validation!"
                    txtFlxCashAllocation.Visible = True
                    txtFlxCashAllocation.SetFocus
                    Exit Sub
                End If
                
                flxCashAllocation.TextMatrix(flxCashAllocation.Row, 2) = txtFlxCashAllocation.Text 'CASH RECEIVED
                flxCashAllocation.TextMatrix(flxCashAllocation.Row, 4) = txtFlxCashAllocation.Text 'BALANCE
                flxCashAllocation.TextMatrix(flxCashAllocation.Row, 6) = Val(txtFlxCashAllocation.Text) - CASHRECEIVED 'NEW BALANCE
            End If
            
        ElseIf DtlsCol = 5 Then 'ALLOCATE
            If cboPayType.Text = "CASH PAYMENT" Or cboPayType.Text = "CREDIT CARD PAYMENT" Or cboPayType.Text = "ONLINE PAYMENT" Then
                AllocatedAmount = Val(txtFlxCashAllocation.Text)
                Balance = Replace(flxCashAllocation.TextMatrix(flxCashAllocation.Row, 2), ",", "")
                
                If AllocatedAmount > Balance Then
                    MsgBox "Cash received is lesser than allocated amount!", vbExclamation, "Validation!"
                    txtFlxCashAllocation.Visible = True
                    txtFlxCashAllocation.SetFocus
                    Exit Sub
                End If
                
                flxCashAllocation.TextMatrix(flxCashAllocation.Row, 5) = AllocatedAmount 'BALANCE
                flxCashAllocation.TextMatrix(flxCashAllocation.Row, 6) = Balance - AllocatedAmount 'NEW BALANCE
                
            ElseIf cboPayType.Text = "ACKNOWLEDGEMENT" Then
                AllocatedAmount = Val(txtFlxCashAllocation.Text)
                Balance = Replace(flxCashAllocation.TextMatrix(flxCashAllocation.Row, 4), ",", "")
                
                If AllocatedAmount > Balance Then
                    MsgBox "Cash balance is lesser than allocated amount!", vbExclamation, "Validation!"
                    txtFlxCashAllocation.Visible = True
                    txtFlxCashAllocation.SetFocus
                    Exit Sub
                End If
                
                flxCashAllocation.TextMatrix(flxCashAllocation.Row, 5) = AllocatedAmount 'BALANCE
                flxCashAllocation.TextMatrix(flxCashAllocation.Row, 6) = Balance - AllocatedAmount 'NEW BALANCE
            End If
        End If
        
        Call ComputeTTLCashAllocation
        txtFlxCashAllocation.Visible = False
    End If
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->txtFlxCashAllocation_KeyPress()"
End Sub

Private Sub txtFlxCashAllocation_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then Exit Sub
End Sub

Private Sub txtFlxInvDtlsAllocateAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then txtFlxInvDtlsAllocateAmount.Visible = False
End Sub

Private Sub txtFlxInvDtlsAllocateAmount_KeyPress(KeyAscii As Integer)
On Error GoTo er
    Dim Balance As Double
    If DidRecordEntry = False Then Exit Sub
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxInvDtlsAllocateAmount, KeyAscii, NumbersWithDecimalPoint)
    
    '                 0    1              2        3             4         5               6               7         8          9           10
    '.FormatString = " |^STATUS |<ACCOUNT TYPE |^REF. NO. |<REF. AMOUNT |<REF. BALANCE |<ALLOCATE |<NEW BALANCE |<BILL AMOUNT |^BILLING NO. |<REF. DATE |<DUE DATE "
    If KeyAscii = 13 Then
        If flxInvoiceDetails.Row < 1 Then Exit Sub
        If Trim(txtFlxInvDtlsAllocateAmount.Text) = "" Then Exit Sub
        
        If IsNumeric(Trim(txtFlxInvDtlsAllocateAmount.Text)) = False Then
            MsgBox "Invalid allocated amount!" & vbCrLf & "Numeric is required.", vbExclamation, "VALIDATION."
            txtFlxInvDtlsAllocateAmount.SetFocus
            Exit Sub
        End If
        
        Balance = Val(Replace(flxInvoiceDetails.TextMatrix(flxInvoiceDetails.Row, 5), ",", ""))
        
        If Val(txtFlxInvDtlsAllocateAmount.Text) > Balance Then
            MsgBox "Allocate amount is greater than the balance!", vbExclamation, "Unable to allocate!"
            txtFlxInvDtlsAllocateAmount.SelStart = 0
            txtFlxInvDtlsAllocateAmount.SelLength = Len(txtFlxInvDtlsAllocateAmount.Text)
            txtFlxInvDtlsAllocateAmount.SetFocus
            Exit Sub
        End If
        
        Call ComputeNewBalInvDtls(txtFlxInvDtlsAllocateAmount.Text)
        Call ComputeNewCheckBalance
        
        DoEvents
        ActiveTaxWHeldObject = "txtTaxHeldPercentage"
        Call txtTaxHeldPercentage_Change
        ActiveTaxWHeldObject = ""
    End If
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->txtFlxCashAllocationID_KeyPress()"
End Sub

Sub ComputeNewBalInvDtls(AllocAmnt As Double)
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub
    Dim flxCol5Value As Double
    Dim flxCol7Value As Double
    Dim AllocAmntVal As Double
    Dim BalanceVal As Double
        
    flxInvoiceDetails.SetFocus
    DtlsRow = flxInvoiceDetails.Rows - 1
    AllocAmntVal = Val(Replace(AllocAmnt, ",", " "))
    
    If txtFlxInvDtlsAllocateAmount.Visible = True Then
        flxCol5Value = Val(Replace(flxInvoiceDetails.TextMatrix(flxInvoiceDetails.Row, 5), ",", " "))
        flxCol7Value = flxCol5Value - AllocAmntVal         '   Balance - Allocate
        flxInvoiceDetails.TextMatrix(flxInvoiceDetails.Row, 6) = Format(Trim(AllocAmntVal), "#,##0.00")      'Col Allocate Value
        flxInvoiceDetails.TextMatrix(flxInvoiceDetails.Row, 7) = Format(Trim(flxCol7Value), "#,##0.00")       'Col New Balance Value
    Else
        BalanceVal = Val(Replace(txtBalance.Text, ",", " "))
        
        flxCol7Value = BalanceVal - AllocAmntVal         '   Balance - Allocate
        flxInvoiceDetails.TextMatrix(DtlsRow, 6) = Format(Trim(AllocAmntVal), "#,##0.00")      'Col Allocate Value
        flxInvoiceDetails.TextMatrix(DtlsRow, 7) = Format(Trim(flxCol7Value), "#,##0.00")        'Col New Balance Value
    End If
    
    Call ComputeTTLAccountAllocAmount
    txtFlxInvDtlsAllocateAmount.Visible = False
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->ComputeNewBalID()"
End Sub

Private Sub txtNetPayable_Change()
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub
    
    If Trim(cboPayType.Text) = "CASH PAYMENT" Then
        Dim Balance As Double
        Dim ALLOCATE As Double
        '                  0   1              2             3                4         5            6           7               8
        '.FormatString = "   |<CASH NO |<CASH RECEIVED |<TOTAL DISPENSE |^BALANCE |^ALLOCATE |^NEW BALANCE |^DATE RECEIVED |<ACTUAL CASH ID"
        If flxCashAllocation.Rows = 2 Then
            If Trim(flxCashAllocation.TextMatrix(1, 1)) <> "" Then
                ALLOCATE = Val(Replace(txtNetPayable.Text, ",", ""))
                Balance = Val(Replace(flxCashAllocation.TextMatrix(1, 2), ",", "")) - Val(Replace(flxCashAllocation.TextMatrix(1, 3), ",", ""))
                flxCashAllocation.TextMatrix(1, 5) = ALLOCATE
                flxCashAllocation.TextMatrix(1, 6) = Format(Balance - ALLOCATE, "#,##0.00")
                DoEvents
                Call ComputeTTLCashAllocation
            End If
        End If
    ElseIf Trim(cboPayType.Text) = "CHECK PAYMENT" Then
        Call ComputeNewCheckBalance
    End If
    
    DoEvents
    Call txtPYMTRCVD_Change
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->txtNetPayable_Change()"
End Sub

Private Sub txtNetPayable_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtNetPayable, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then cboCustomer.SetFocus
End Sub

Private Sub txtNewChckBal_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtNewChckBal, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then txtChckBal.SetFocus
End Sub

Private Sub txtNmOfBnk_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtNmOfBnk, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtChckTyp.SetFocus
End Sub

Private Sub txtPRDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtPRDate, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtChckNo.SetFocus
End Sub

Private Sub txtPRno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtPRno, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then txtAccntName.SetFocus
End Sub

Sub LoadLedger(CustID As Integer, ACCNTTYPE As String, RefNo As String)
On Error GoTo er
    Dim DtlsRow As Integer
    Dim tmpRs As New ADODB.Recordset
    
    Call flxLedgerColsChanger
    
    DoEvents
    Call modAllocation.sp_onaccterms_list_ledger(SysBranchID, CustID, ACCNTTYPE, RefNo, tmpRs)
    ' REFTYPE , RefNo, TRANSDATE, ACCTAMOUNT, Terms, MONTHLYDUE, Status, ldgrREFTYPE, ldgrREFNO, RefDate, DEBIT
    ', CREDIT, OSBAL, SEQNO, JOBORDERID, SPINVNO, JOBORDERNO, ACCTBAL
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount <= 0 Then
            tmpRs.Close
            tmrLoadLedger.Enabled = False
            LoadflxAccount = False
            MsgBox "No record found!", vbExclamation, "Validation!"
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    DtlsRow = 1
    
    tmpRs.MoveFirst
    With tmpRs
        ' REFTYPE , RefNo, TRANSDATE, ACCTAMOUNT, Terms, MONTHLYDUE, Status
        txtLdgrAcctType.Text = !RefType
        txtLdgrRefNo.Text = !RefNo
        txtLdgrTransDate.Text = !TRANSDATE
        txtLdgrAcctAmnt.Text = !AcctAmount
        txtLdgrTerms.Text = !Terms
        txtLdgrMnthlyDue.Text = !MONTHLYDUE
        txtldgrStatus.Text = !STATUS
        
        Do While Not tmpRs.EOF
            '                0      1              2       3           4        5         6
            '.FormatString =  "   |<REF. TYPE |<REF. NO. |<REF. DATE |<DEBIT |<CREDIT |<OS BAL
            If DtlsRow = 1 Then
                flxLedger.TextMatrix(DtlsRow, 0) = ""
                flxLedger.TextMatrix(DtlsRow, 1) = !ldgrREFTYPE
                flxLedger.TextMatrix(DtlsRow, 2) = !ldgrREFNO
                flxLedger.TextMatrix(DtlsRow, 3) = !REFDATE
                flxLedger.TextMatrix(DtlsRow, 4) = IIf(Val(!Debit) = 0, "", Format(!Debit, "#,##0.00"))
                flxLedger.TextMatrix(DtlsRow, 5) = IIf(Val(!CREDIT) = 0, "", Format(!CREDIT, "#,##0.00"))
                flxLedger.TextMatrix(DtlsRow, 6) = IIf(Val(!OSbal) = 0, "", Format(!OSbal, "#,##0.00"))
            Else
                flxLedger.AddItem _
                    vbTab & !ldgrREFTYPE & _
                    vbTab & !ldgrREFNO & _
                    vbTab & !REFDATE & _
                    vbTab & IIf(Val(!Debit) = 0, "", Format(!Debit, "#,##0.00")) & _
                    vbTab & IIf(Val(!CREDIT) = 0, "", Format(!CREDIT, "#,##0.00")) & _
                    vbTab & IIf(Val(!OSbal) = 0, "", Format(!OSbal, "#,##0.00"))
            End If
            DtlsRow = DtlsRow + 1
            tmpRs.MoveNext
        Loop
    End With
    If tmpRs.State <> 0 Then tmpRs.Close
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->LoadLedger()"
End Sub
  
Private Sub txtPYMTRCVD_Change()
On Error Resume Next
Dim tmpRemaining As Double
tmpRemaining = Val(Replace(txtPYMTRCVD.Text, ",", "")) - Val(Replace(txtNetPayable.Text, ",", ""))
txtPYMTREMAINING.Text = Format(tmpRemaining, "#,##0.00")
End Sub

Private Sub txtRefNo_Change()
On Error GoTo er
    txtAllocAmount.Text = ""
    txtBalance.Text = ""
    txtBillAmnt.Text = ""
    txtAccntType.Text = ""
    
    With flxAccnt
        .Top = fraHoldBottom.Top + txtRefNo.Height + txtRefNo.Top
        .Left = fraHoldBottom.Left + txtRefNo.Left
        .Visible = True
    End With
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->txtRefNo_Change()"
End Sub

Private Sub txtrefno_KeyDown(KeyCode As Integer, Shift As Integer)
    If DidRecordEntry = False Then Exit Sub
    If KeyCode = vbKeyDown Then
        If flxAccnt.Visible Then flxAccnt.SetFocus
    End If
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
On Error GoTo er
    Dim i As Integer
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRefNo, KeyAscii, ValidCodeChars)
    
    If DidRecordEntry = False Then
        flxAccnt.Visible = False
        Exit Sub
    End If
    
    SearchRefNo = False

    If KeyAscii = 13 Then
        flxAccnt.HighLight = flexHighlightNever
        
        If Trim(txtRefNo.Text) = "" Then
            txtRefNo.Text = ""
            With flxAccnt
                .Top = fraHoldBottom.Top + txtRefNo.Height + txtRefNo.Top
                .Left = fraHoldBottom.Left + txtRefNo.Left
                .Visible = True
                .HighLight = flexHighlightAlways
            End With
            Exit Sub
        End If
    
        If SearchRefNo = True Then
            SearchRefNo = False
            Exit Sub
        End If
        
        If Trim(cboCustomer.Text) = "" Then
            SearchcboCust = True
            MsgBox "CUSTOMER NAME IS MISSING!", vbExclamation, "VALIDATION!"
            CUSTOMERID = 0
            SearchcboCust = True
            cboCustomer.ListIndex = -1
            SearchcboCust = True
            cboCustomer.Text = ""
            cboCustomer.Locked = False
            cboCustomer.BackColor = txtHoldWhiteBackColor.BackColor
            cboCustomer.SetFocus
            Exit Sub
        End If
        
        
        flxAccntSelRow = 0
        With flxAccnt
            .Top = fraHoldBottom.Top + txtRefNo.Height + txtRefNo.Top
            .Left = fraHoldBottom.Left + txtRefNo.Left
            .Visible = True
            DidflxAccntInUsed = True
            
            For i = 1 To .Rows - 1
            'flxAccnt.FormatString = "  ACCOUNTSTATUS, ACCOUNT_TYPE, REFNO, DUEDATE, TRANSDATE, ACCTAMOUNT, ACCTBAL, BILLINGNO, BILLAMOUNT, SPSORDERID, JOBORDERID
            If InStr(.TextMatrix(i, 3), txtRefNo.Text) > 0 Then
                flxAccntSelRow = i
                .TopRow = flxAccntSelRow
                .Row = flxAccntSelRow
                .Col = 3
                .ColSel = flxAccntTTLCol
                .HighLight = flexHighlightAlways
                Exit Sub
            Else
                flxAccntSelRow = 1
                .TopRow = flxAccntSelRow
                .Row = flxAccntSelRow
                .Col = 3
                .ColSel = 13
                .HighLight = flexHighlightAlways
            End If
            Next i
        End With
        
        SearchRefNo = True
    End If

Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->txtRefNo_KeyPress()"
End Sub

Private Sub txtTaxHeldAmount_Change()
On Error GoTo er
    Dim tmpAmount As Double
    Dim tmpPerc As Double
    Dim AllocatedAmount As Double
    Dim NetSales As Double
    
    If DidRecordEntry = False Then Exit Sub
    If Trim(ActiveTaxWHeldObject) = "Skip" Then Exit Sub
    
    AllocatedAmount = Val(Replace(txtAllctedAmntGenTab.Text, ",", ""))
    If AllocatedAmount <= 0 Then
        txtNetPayable.Text = Format(AllocatedAmount, "#,##0.00")
        txtTaxHeldPercentage.Text = ""
        txtTaxHeldAmount.Text = ""
        Exit Sub
    End If
    
    DoEvents
    If Trim(ActiveTaxWHeldObject) <> "txtTaxHeldAmount" And Trim(ActiveTaxWHeldObject) <> "" Then Exit Sub
    
    DoEvents
    'AMOUNT Computation
    tmpAmount = Val(Replace(txtAllctedAmntGenTab.Text, ",", ""))
    tmpPerc = Format(Val(Format((Val(txtTaxHeldAmount.Text) / Val(Round(tmpAmount, 2))), "##0.0000")) * 100, "##0.0000")
    
    If Val(tmpPerc) > MaxTAXWithHELD Then
        MsgBox "EXCEED TO THE MAXIMUM TAX WITH HELD ALLOWED!", vbExclamation, "VALIDATION!"
        txtTaxHeldPercentage.Text = ""
        txtTaxHeldAmount.Text = ""
        Exit Sub
    End If
    
    txtTaxHeldPercentage.Text = IIf(tmpPerc = 0, "", tmpPerc)
    
    DoEvents
    Call ComputeNetPayable
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->txtTaxHeldAmount_Change()"
End Sub

Sub ComputeNetPayable()
On Error GoTo er
    Dim txtAllctedAmntValue As Double
    Dim txtTaxHeldAmntValue As Double
    Dim txtNetPayValue As Double
    
    txtAllctedAmntValue = Val(Replace(txtAllctedAmntGenTab.Text, ",", ""))
    txtTaxHeldAmntValue = Val(Replace(txtTaxHeldAmount.Text, ",", ""))
    
    txtNetPayValue = txtAllctedAmntValue - txtTaxHeldAmntValue
    txtNetPayable.Text = Format(txtNetPayValue, "#,##0.00")
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->ComputeNetPayable"
End Sub

Private Sub txtTaxHeldAmount_Click()
    TaxHeldAmntValue = 0
End Sub

Private Sub txtTaxHeldAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If DidRecordEntry = False Then Exit Sub
    ActiveTaxWHeldObject = "txtTaxHeldAmount"
End Sub

Private Sub txtTaxHeldPercentage_Change()
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub
    If Trim(ActiveTaxWHeldObject) = "Skip" Then Exit Sub
    
    If Val(Replace(txtAllctedAmntGenTab.Text, ",", "")) <= 0 Then
        txtNetPayable.Text = txtAllctedAmntGenTab.Text
        txtNetPayable.Text = ""
        txtTaxHeldAmount.Text = ""
        Exit Sub
    End If
    
    DoEvents
    If Trim(ActiveTaxWHeldObject) <> "txtTaxHeldPercentage" And Trim(ActiveTaxWHeldObject) <> "" Then Exit Sub
    
    DoEvents
    Call ComputeTaxWithHeld

Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->txtTaxHeldPercentage_Change()"
End Sub

Sub ComputeTaxWithHeld()
On Error GoTo er
    Dim tmpNet As Double
    Dim strTotalAllocatedAmount As String
    Dim strAmtReceived As String
    
    Dim tmpTaxWHeld As Double
    Dim strTax As String
   
    
    'Get the Tax Percentage
    txtTaxHeldPercentage.Text = Trim(txtTaxHeldPercentage.Text)
    strTax = Replace(txtTaxHeldPercentage.Text, ",", "")
    If Trim(strTax) = "" Then strTax = "0.0000"
    If Val(strTax) > MaxTAXWithHELD Then
        strTax = "0.00" ' Set TAX to zero
        MsgBox "EXCEED TO THE MAXIMUM TAX WITH HELD ALLOWED!", vbExclamation, "VALIDATION!"
        txtNetPayable.Text = txtAllctedAmntGenTab.Text
        txtTaxHeldPercentage.Text = ""
        txtTaxHeldAmount.Text = ""
        txtTaxHeldPercentage.SetFocus
        Exit Sub
    End If
    
    DoEvents
    'Get the Total Payable
    strTotalAllocatedAmount = Replace(txtAllctedAmntGenTab.Text, ",", "")
    If Trim(strTotalAllocatedAmount) = "" Then strTotalAllocatedAmount = "0.00"
    
    'Compute the tax w/ held
    tmpTaxWHeld = Val(strTotalAllocatedAmount) * Val((Format((Val(strTax) / 100), "##0.0000")))
    
    If ActiveTaxWHeldObject <> "txtTaxHeldAmount" Then
        txtTaxHeldAmount.Text = Format(tmpTaxWHeld, "##0.00")
    End If
    
    DoEvents
    Call ComputeNetPayable
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->ComputeTaxWHeld()"
End Sub

Private Sub txtTaxHeldAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtTaxHeldAmount, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then txtNetPayable.SetFocus
End Sub

Private Sub txtTaxHeldPercentage_Click()
    TaxHeldPerValue = 0
End Sub

Private Sub txtTaxHeldPercentage_KeyDown(KeyCode As Integer, Shift As Integer)
    If DidRecordEntry = False Then Exit Sub
    ActiveTaxWHeldObject = "txtTaxHeldPercentage"
End Sub

Private Sub txtTaxHeldPercentage_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtTaxHeldPercentage, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then txtTaxHeldAmount.SetFocus
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As Integer)
On Error GoTo er
    If DidRecordEntry = False Then Exit Sub

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboCustomer, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cboCustomer.Enabled = False
        Call cboCustomer_Click
        cboCustomer.Enabled = True
    End If
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->cboCust_KeyPress()"
End Sub

Private Sub txtTTLCash_Change()
    txtPYMTRCVD.Text = Format(Trim(Replace(txtTTLCash.Text, ",", "")), "#,##0.00")
End Sub
