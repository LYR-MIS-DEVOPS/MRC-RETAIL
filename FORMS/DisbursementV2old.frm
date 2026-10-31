VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form DisbursementV2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disbursement"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9990
   Icon            =   "DisbursementV2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPayee 
      Height          =   1575
      Left            =   10920
      TabIndex        =   81
      Top             =   7200
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2778
      _Version        =   393216
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cboBranch 
      Height          =   315
      Left            =   5640
      TabIndex        =   88
      Top             =   720
      Width           =   4215
   End
   Begin VB.TextBox txtORamt 
      Height          =   285
      Left            =   10560
      TabIndex        =   87
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab sstabDisplayAcctTitle 
      Height          =   5175
      Left            =   10440
      TabIndex        =   69
      Top             =   1200
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "ACCOUNT TITLE"
      TabPicture(0)   =   "DisbursementV2.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxAcct_Title"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdInsert"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin lvButton.lvButtons_H cmdInsert 
         Height          =   375
         Left            =   3840
         TabIndex        =   70
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Insert"
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
         Image           =   "DisbursementV2.frx":57FE
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   375
         Left            =   5280
         TabIndex        =   71
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&OK"
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
         Image           =   "DisbursementV2.frx":5D98
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   6720
         TabIndex        =   72
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Cancel"
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
         Image           =   "DisbursementV2.frx":6332
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxAcct_Title 
         Height          =   4335
         Left            =   120
         TabIndex        =   73
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   7646
         _Version        =   393216
         FixedCols       =   0
      End
   End
   Begin VB.TextBox txtFlx 
      Height          =   285
      Left            =   10440
      TabIndex        =   85
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   84
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   10440
      TabIndex        =   83
      Text            =   "HoldBackColor"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCVid 
      Height          =   285
      Left            =   11400
      TabIndex        =   82
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   11880
      TabIndex        =   80
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin lvButton.lvButtons_H cmdMainButton 
      Height          =   375
      Index           =   6
      Left            =   9000
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "Print Voucher"
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Print"
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
      cFHover         =   255
      LockHover       =   2
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "DisbursementV2.frx":68CC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMainButton 
      Height          =   375
      Index           =   5
      Left            =   7440
      TabIndex        =   77
      TabStop         =   0   'False
      ToolTipText     =   "View Summary"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "View Summary"
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
      cFHover         =   255
      LockHover       =   2
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "DisbursementV2.frx":6E66
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMainButton 
      Height          =   375
      Index           =   4
      Left            =   6480
      TabIndex        =   79
      TabStop         =   0   'False
      ToolTipText     =   "Edit Voucher"
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "&Edit"
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
      cFHover         =   255
      LockHover       =   2
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "DisbursementV2.frx":7400
      cBack           =   -2147483633
   End
   Begin VB.ComboBox cboTransType 
      Height          =   315
      Left            =   1920
      TabIndex        =   67
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtApprovedBy 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Top             =   7080
      Width           =   8175
   End
   Begin VB.TextBox txtPreparedBy 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   6720
      Width           =   8175
   End
   Begin VB.Frame frameTtlDMCM 
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   5325
      Width           =   9615
      Begin VB.TextBox txtSupplierAddress 
         Height          =   315
         Left            =   960
         TabIndex        =   95
         Top             =   960
         Width           =   8535
      End
      Begin VB.TextBox txtSupplier 
         Height          =   315
         Left            =   4440
         TabIndex        =   94
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox txtTIN 
         Height          =   315
         Left            =   480
         TabIndex        =   91
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtTtlDebit 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox txtTtlCredit 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   180
         Width           =   1215
      End
      Begin VB.CheckBox chkAutoCompute 
         Caption         =   "AUTO"
         Height          =   255
         Left            =   8640
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtDisbAmnt 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6120
         TabIndex        =   4
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "ADDRESS:"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "SUPPLIER:"
         Height          =   255
         Left            =   3600
         TabIndex        =   93
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "TIN:"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Total Debit:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Total Credit:"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Disburse Amount:"
         Height          =   255
         Left            =   4800
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdMainButton 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Branch Expense"
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
      cFHover         =   255
      LockHover       =   2
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "DisbursementV2.frx":799A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMainButton 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "&Administrative Expense"
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
      cFHover         =   255
      LockHover       =   2
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "DisbursementV2.frx":E1FC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMainButton 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   661
      Caption         =   "&Other Branch Expense"
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
      Image           =   "DisbursementV2.frx":E796
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxCV 
      Height          =   3375
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   $"DisbursementV2.frx":13F88
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   240
      TabIndex        =   14
      Top             =   7560
      Visible         =   0   'False
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Commissions"
      TabPicture(0)   =   "DisbursementV2.frx":1403D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flxComm"
      Tab(0).Control(1)=   "lblHeading(9)"
      Tab(0).Control(2)=   "lblHeading(8)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Cash Advances"
      TabPicture(1)   =   "DisbursementV2.frx":14059
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtCAttl"
      Tab(1).Control(1)=   "txtCADue"
      Tab(1).Control(2)=   "fraCATerm"
      Tab(1).Control(3)=   "chkTerms"
      Tab(1).Control(4)=   "cmdDeduct"
      Tab(1).Control(5)=   "flxAdv"
      Tab(1).Control(6)=   "lblHeading(11)"
      Tab(1).Control(7)=   "Label7"
      Tab(1).Control(8)=   "lblHeading(10)"
      Tab(1).Control(9)=   "lblDeduct"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Incentives"
      TabPicture(2)   =   "DisbursementV2.frx":14075
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblHeading(13)"
      Tab(2).Control(1)=   "lblHeading(12)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Refunds"
      TabPicture(3)   =   "DisbursementV2.frx":14091
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblHeading(15)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblHeading(14)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame1"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.TextBox txtCAttl 
         Height          =   315
         Left            =   -68640
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtCADue 
         Height          =   315
         Left            =   -68640
         TabIndex        =   39
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Frame fraCATerm 
         Height          =   1935
         Left            =   -68640
         TabIndex        =   31
         Top             =   2760
         Width           =   1935
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Payment Terms:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Amount Due"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Payment Mode"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkTerms 
         Caption         =   "Apply Terms"
         Height          =   195
         Left            =   -68640
         TabIndex        =   30
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   9375
         Begin VB.ComboBox cboRefundSrc 
            Height          =   315
            ItemData        =   "DisbursementV2.frx":140AD
            Left            =   1560
            List            =   "DisbursementV2.frx":140AF
            TabIndex        =   21
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtRefundRef 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1560
            TabIndex        =   20
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtRefDate 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtRefAmt 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1560
            TabIndex        =   18
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox txtRefCust 
            Height          =   315
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox txtReason 
            Height          =   795
            Left            =   4080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   1080
            Width           =   5055
         End
         Begin lvButton.lvButtons_H cmdGoRefund 
            Height          =   495
            Left            =   8520
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Send to Particulars"
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
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
            cFHover         =   255
            LockHover       =   2
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            ImgAlign        =   4
            Image           =   "DisbursementV2.frx":140B1
            cBack           =   -2147483633
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Select Source:"
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label lblRefNo 
            Alignment       =   1  'Right Justify
            Caption         =   "Reference No.:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   585
            Width           =   1335
         End
         Begin VB.Label lblRefDate 
            Alignment       =   1  'Right Justify
            Caption         =   "Ref. Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblRefAmt 
            Alignment       =   1  'Right Justify
            Caption         =   "Refund Amount:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Customer Name:"
            Height          =   255
            Left            =   4080
            TabIndex        =   25
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Reason for Refund:"
            Height          =   255
            Left            =   4080
            TabIndex        =   24
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lblAddRef 
            BackStyle       =   0  'Transparent
            Height          =   495
            Left            =   120
            TabIndex        =   23
            Top             =   1680
            Width           =   3615
         End
      End
      Begin lvButton.lvButtons_H cmdDeduct 
         Height          =   375
         Left            =   -68640
         TabIndex        =   38
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Deduct C.A."
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
         ImgAlign        =   3
         Image           =   "DisbursementV2.frx":1464B
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxComm 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   13
         RowHeightMin    =   315
         BackColorSel    =   12648447
         ForeColorSel    =   12582912
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         FormatString    =   $"DisbursementV2.frx":14BE5
      End
      Begin MSFlexGridLib.MSFlexGrid flxAdv 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   42
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   -2147483633
         FormatString    =   "^     |^     C.A. Amount |^    Amount Paid  |^ Balance           |^           Remarks                "
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   13
         Left            =   -67650
         TabIndex        =   55
         Top             =   30
         Width           =   2415
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   -67650
         TabIndex        =   54
         Top             =   30
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Total Cash Advance:"
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
         Left            =   -68640
         TabIndex        =   53
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   -65010
         TabIndex        =   52
         Top             =   30
         Width           =   2415
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   -65010
         TabIndex        =   51
         Top             =   30
         Width           =   2415
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   4
         Left            =   -65040
         TabIndex        =   50
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   6
         Left            =   -65040
         TabIndex        =   49
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   -67650
         TabIndex        =   48
         Top             =   30
         Width           =   2415
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   8
         Left            =   -67680
         TabIndex        =   47
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   10
         Left            =   -67680
         TabIndex        =   46
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   12
         Left            =   -67680
         TabIndex        =   45
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblDeduct 
         Caption         =   "Total Amount Due"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -68640
         TabIndex        =   43
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   14
         Left            =   8520
         TabIndex        =   44
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   15
         Left            =   8520
         TabIndex        =   56
         Top             =   30
         Width           =   2415
      End
   End
   Begin lvButton.lvButtons_H cmdMainButton 
      Height          =   375
      Index           =   8
      Left            =   10440
      TabIndex        =   75
      Top             =   735
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Cancel Voucher"
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
      cFHover         =   255
      LockHover       =   2
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "DisbursementV2.frx":14D02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMainButton 
      Height          =   375
      Index           =   7
      Left            =   10440
      TabIndex        =   76
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Save Voucher  "
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
      cFHover         =   255
      LockHover       =   2
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "DisbursementV2.frx":1529C
      cBack           =   -2147483633
   End
   Begin VB.Frame FraHeading 
      Height          =   975
      Left            =   240
      TabIndex        =   57
      Top             =   960
      Width           =   9615
      Begin VB.TextBox txtPayee 
         Height          =   315
         Left            =   1320
         TabIndex        =   90
         Top             =   240
         Width           =   5655
      End
      Begin VB.ComboBox cboTransaction 
         Height          =   315
         Left            =   1320
         TabIndex        =   60
         Top             =   600
         Width           =   5655
      End
      Begin VB.TextBox txtPCV 
         Height          =   315
         Left            =   7800
         TabIndex        =   59
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   7800
         TabIndex        =   58
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "PCV NO."
         Height          =   255
         Index           =   6
         Left            =   7080
         TabIndex        =   64
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "DATE"
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   63
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "TRANSACTION"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "PAYEE"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdAltA 
      Caption         =   "&A"
      Height          =   375
      Left            =   4920
      TabIndex        =   86
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "BRANCH:"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   89
      Top             =   750
      Width           =   735
   End
   Begin VB.Label lblChk 
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   74
      Top             =   8760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "TRANSACTION TYPE:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   68
      Top             =   750
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "APPROVED BY:"
      Height          =   255
      Left            =   360
      TabIndex        =   66
      Top             =   7095
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "PREPARED BY:"
      Height          =   255
      Left            =   360
      TabIndex        =   65
      Top             =   6735
      Width           =   1335
   End
   Begin VB.Menu mnuVoucher 
      Caption         =   "Voucher"
      Visible         =   0   'False
      Begin VB.Menu mnuAddRow 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Remove Row"
      End
   End
End
Attribute VB_Name = "DisbursementV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPayee As New ADODB.Recordset
Dim rsMaxPCVNo As New ADODB.Recordset
Dim rsAcctDesc As New ADODB.Recordset

Dim arrAcctDesc() As String
Dim arrBRANCH() As String
Dim myJVNO As String
Dim myTransID As Long, EMPID As Long, CUSTOMERID As Long
Dim i As Integer, j As Integer
Dim ActRow As Integer, ActCol As Integer
Dim DP As Currency, SP As Currency, SVC As Currency, OT As Currency, myCreditAmnt As Currency

Dim DisbursementType As String, formAction As String
Dim Qry As String

Dim AcctSrch As Boolean, LoadTab As Boolean, refundBtn As Boolean, CustSrch As Boolean, CustBln As Boolean

Dim StatSrch As Boolean
Dim strPartial As String, strTotal As String
Dim PrtBln As Boolean, PrtSrch As Boolean
Dim branchCode As String

Public PCVID As Long

Private Sub DisplayORDetails()
    Dim rsOR As New ADODB.Recordset
    Dim Qry As String, myRem As String, PayID As Long
    Dim Ans As Integer
    
    '***Access
    'Qry = "SELECT PAYMENTS.* From PAYMENTS WHERE (((PAYMENTS.ORNO)='" & txtRefundRef.Text & "'));"
    'mySQL
    If cboRefundSrc.Text = "O.R." Then
        Qry = "SELECT PAYMENTS.* From PAYMENTS WHERE PAYMENTS.ORNO='" & txtRefundRef.Text & "';"
    ElseIf cboRefundSrc.Text = "C.R." Then
        Qry = "SELECT PAYMENTS.* From PAYMENTS WHERE PAYMENTS.CRNO='" & txtRefundRef.Text & "';"
    End If
    myRem = ""
    With rsOR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtRefDate.Text = Format(!PayDate, "mm/dd/yyyy")
            txtRefCust.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
            DP = IIf(IsNull(!DownPayment), 0, !DownPayment)
            SP = IIf(IsNull(!SPAREPARTS), 0, !SPAREPARTS)
            SVC = IIf(IsNull(!Service), 0, !Service)
            OT = IIf(IsNull(!Others), 0, !Others)
            If txtRefundRef.Tag = "" Then
                txtRefAmt.Text = Format(DP + SP + SVC + OT, "#,##0.00")
            Else
                txtRefAmt.Text = flxCV.TextMatrix(1, 2)
            End If
            txtORamt.Text = DP + SP + SVC + OT
            myRem = IIf(IsNull(!Remarks), "", !Remarks)
            txtRefAmt.Locked = False
            PayID = !PaymentID
        Else
            myRem = "No Record"
        End If
        .Close
    End With
    If myRem = "No Record" Then
        MsgBox "No record found for O.R. #" & txtRefundRef.Text & "!", vbExclamation + vbOKOnly, "Refund"
        Exit Sub
    End If
    If DP > 0 Then
        '***Access
        'Qry = "SELECT DPLEDGER.* From DPLEDGER WHERE (((DPLEDGER.ORNO)='" & txtRefundRef.Text & "'));"
        'mySQL
        Qry = "SELECT DPLEDGER.* From DPLEDGER WHERE DPLEDGER.ORNO='" & txtRefundRef.Text & "';"
        With rsOR
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                lblAddRef.Caption = "O.R. Amt: " & Format(!ORAmount, "#,##0.00") & " DP Status: " & !STATUS & " Bal. " & Format(!Balance, "#,##0.00") & " " & IIf(IsNull(!Remarks), "", !Remarks)
                txtRefAmt.Text = Format(!Balance, "#,##0.00")
            End If
            .Close
        End With
    ElseIf OT > 0 Then
        lblAddRef.Caption = "O.R. Amt. " & txtRefAmt.Text & " (Others) " & myRem
    ElseIf SP > 0 And SVC > 0 Then
        lblAddRef.Caption = "O.R. Amt. " & txtRefAmt.Text & " Payment for Spare Parts & Service"
    ElseIf SP > 0 Then
        lblAddRef.Caption = "O.R. Amt. " & txtRefAmt.Text & " Payment for Spare Parts"
    ElseIf SVC > 0 Then
        lblAddRef.Caption = "O.R. Amt. " & txtRefAmt.Text & " Payment for Service"
    End If
    '***Access
    'Qry = "SELECT Count(ORDETAILS.CUSTOMERID) AS CountOfCUSTOMERID From ORDETAILS WHERE (((ORDETAILS.PAYMENTID)=" & PayID & "));"
    'mySQL
    Qry = "SELECT Count(ORDETAILS.CUSTOMERID) AS CountOfCUSTOMERID From ORDETAILS WHERE ORDETAILS.PAYMENTID=" & PayID & ";"
    Ans = Val(GetFieldValue(Qry))
    If Ans > 1 Then MsgBox "Cannot Refund O.R. due to several Invoices or J.O. allocated with it!", vbExclamation + vbOKOnly
End Sub

Private Sub DisplaySalesReturn()
    Dim rsSRet As New ADODB.Recordset
    Dim Qry As String, myRem As String, MyPro As String
    Dim Ans As Integer
    
    'Qry = "SELECT SRETURNSP.RETNO, SRETURNSP.RETDATE, SRETURNSP.TOTALAMT, SRETURNSP.PROCESS, SPINVOICE.SPINVNO, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, SPINVOICE.AMOUNT, SPINVOICE.INVTYPE, SRETURNSP.PROREF, SRETURNSP.REASON " & _
          "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO WHERE (((SRETURNSP.RETNO)='" & txtRefundRef.Text & "'));"
    'mySQl
    If cboRefundSrc.Text = "SALES RETURN - CREDIT" Then
        Qry = "SELECT SRETURNSP.RETNO, SRETURNSP.RETDATE, SRETURNSP.TOTALAMT, SRETURNSP.PROCESS, SPINVOICE.SPINVNO, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, SPINVOICE.AMOUNT, SPINVOICE.INVTYPE, SRETURNSP.PROREF, SRETURNSP.REASON " & _
              "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO WHERE SRETURNSP.RETNO='" & txtRefundRef.Text & "' and SPINVOICE.CUSTOMERID = " & Val(CUSTOMERID) & " ORDER BY SPINVOICE.INVDATE DESC;"
    Else
        Qry = "SELECT SRETURNSP.RETNO, SRETURNSP.RETDATE, SRETURNSP.TOTALAMT, SRETURNSP.PROCESS, SPINVOICE.SPINVNO, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, SPINVOICE.AMOUNT, SPINVOICE.INVTYPE, SRETURNSP.PROREF, SRETURNSP.REASON " & _
          "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO WHERE SRETURNSP.RETNO='" & txtRefundRef.Text & "' ORDER BY SPINVOICE.INVDATE DESC;"
    End If
    
    myRem = ""
    With rsSRet
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtRefDate.Text = Format(!RetDate, "mm/dd/yyyy")
            txtRefAmt.Text = Format(!TotalAmt, "#,##0.00")
            txtRefCust.Text = !ChargeTo
            txtReason.Text = IIf(IsNull(!REASON), "", !REASON)
            lblAddRef.Caption = "Inv.# " & !SPINVNO & " Inv. Date: " & Format(!InvDate, "mm/dd/yy") & " Amt: " & Format(!Amount, "#,##0.00") '& IIf(IsNull(!InvType), "" & "[" & !InvType & "]")
            If Not IsNull(!InvType) Then lblAddRef.Caption = lblAddRef.Caption & " [" & !InvType & "]"
            MyPro = IIf(IsNull(!ProRef), "", !ProRef)
            If MyPro <> "" Then myRem = "Sales Return Slip #" & txtRefundRef.Text & " was already " & !PROCESS & " with Inv.# " & MyPro
            txtRefAmt.Locked = True
        Else
            myRem = "No Record"
        End If
        .Close
    End With
    If myRem = "No Record" Then
        MsgBox "No record found for Sales Return Slip #" & txtRefundRef.Text & "!", vbExclamation + vbOKOnly, "Refund"
    ElseIf myRem <> "" And txtRefundRef.Tag = "" Then
        Ans = MsgBox(myRem & vbCrLf & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo, "Refund")
        If Ans = vbNo Then
            txtRefundRef.Text = ""
            txtRefDate.Text = ""
            txtRefAmt.Text = ""
            txtRefCust.Text = ""
            lblAddRef.Caption = ""
        End If
    End If
End Sub


Private Sub cboBranch_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboBranch
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
        
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBranch-Change()")
End Sub

Private Sub cboBranch_Click()
    branchCode = arrBRANCH(cboBranch.ListIndex, 0)
    txtPayee.SetFocus
End Sub

Private Sub cboBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cboBranch, cboBranch.Text)
        
        If cbolistIndex < 0 Then
            cboBranch.SetFocus
        Else
            cboBranch.ListIndex = cbolistIndex
            Call cboBranch_Click
        End If
    End If
End Sub

Private Sub cboRefundSrc_Click()
    If cboRefundSrc.Text = "O.R." Then
        lblRefNo.Caption = "O.R. No:"
        lblRefDate.Caption = "O.R. Date:"
        lblRefAmt.Caption = "O.R. Amount:"
    ElseIf cboRefundSrc.Text = "SALES RETURN - CREDIT" Then
        lblRefNo.Caption = "Sales Ret#:"
        lblRefDate.Caption = "Ret. Date:"
        lblRefAmt.Caption = "Ret. Amount:"
    ElseIf cboRefundSrc.Text = "SALES RETURN - CASH" Then
        lblRefNo.Caption = "Sales Ret#:"
        lblRefDate.Caption = "Ret. Date:"
        lblRefAmt.Caption = "Ret. Amount:"
    End If
End Sub

Private Sub cboTransaction_Click()
Call cboTransaction_Change
End Sub

Private Sub cmdGoRefund_Click()
    Dim vMsg As String, Qry As String, PayID As Long
    
    If txtRefundRef.Text = "" Or txtRefDate.Text = "" Then
        MsgBox "No valid refund reference available!", vbExclamation + vbOKOnly, "Refund"
        Exit Sub
    End If
    vMsg = ""
    If txtReason.Text = "" Then vMsg = "Please indicate reason of refund!"
    If Val(Replace(txtRefAmt, ",", "")) = 0 Then vMsg = vMsg & vbCrLf & "No Refund amount specified!"
    If cboRefundSrc.Text = "O.R." And Val(Replace(txtRefAmt, ",", "")) > Val(txtORamt.Text) Then vMsg = vMsg & vbCrLf & "Refund amount must not be greater than O.R. amount!"
    If cboRefundSrc.Text = "O.R." Then
        'Qry = "SELECT PAYMENTS.PAYMENTID From PAYMENTS WHERE (((PAYMENTS.ORNO)='" & txtRefundRef.Text & "'));"
        'mySQl
        Qry = "SELECT PAYMENTS.PAYMENTID From PAYMENTS WHERE PAYMENTS.ORNO='" & txtRefundRef.Text & "';"
        PayID = Val(GetFieldValue(Qry))
        'Qry = "SELECT Count(ORDETAILS.CUSTOMERID) AS CountOfCUSTOMERID From ORDETAILS WHERE (((ORDETAILS.PAYMENTID)=" & PayID & "));"
        'mySQL
        Qry = "SELECT Count(ORDETAILS.CUSTOMERID) AS CountOfCUSTOMERID From ORDETAILS WHERE ORDETAILS.PAYMENTID=" & PayID & ";"
        
        If Val(GetFieldValue(Qry)) > 1 Then vMsg = vMsg & vbCrLf & "Cannot Refund O.R. due to several Invoices or J.O. allocated with it!"
    End If
    If vMsg <> "" Then
        MsgBox vMsg, vbCritical + vbOKOnly, "Validation"
        Exit Sub
    End If
    
    'Display Refund details in flxCV
    'flxCV.Rows = 1
    Dim myVatOutput As Currency, myAcctTitle As Currency
    If Val(chkAutoCompute.Value) = 1 Then
        For i = 1 To flxCV.Rows - 1
            myAcctTitle = Format((Val(Replace(txtRefAmt.Text, ",", "")) / 1.12), "#,###,###.00")
            myVatOutput = Format(Val(Replace(txtRefAmt.Text, ",", "")) - (Val(Replace(txtRefAmt.Text, ",", "")) / 1.12), "#,###,###.00")
            If flxCV.TextMatrix(i, 4) = "CREDIT" Then
                flxCV.TextMatrix(i, 3) = Format(txtRefAmt.Text, "#,###,###.00")
            End If
            If flxCV.TextMatrix(i, 6) = "20004" And flxCV.TextMatrix(i, 4) = "DEBIT" Then
                flxCV.TextMatrix(i, 2) = Format(myVatOutput, "#,###,###.00")
            ElseIf flxCV.TextMatrix(i, 6) = "20011" And flxCV.TextMatrix(i, 4) = "CREDIT" Then
                flxCV.TextMatrix(i, 3) = Format(myVatOutput, "#,###,###.00")
            ElseIf flxCV.TextMatrix(i, 6) <> "20004" And flxCV.TextMatrix(i, 4) = "DEBIT" Then
                flxCV.TextMatrix(i, 2) = Format(myAcctTitle, "#,###,###.00")
            End If
        Next i
    Else
        For i = 1 To flxCV.Rows - 1
            myAcctTitle = Format(Val(Replace(txtRefAmt.Text, ",", "")), "#,###,###.00")
            If flxCV.TextMatrix(i, 4) = "CREDIT" Then
                flxCV.TextMatrix(i, 3) = Format(txtRefAmt.Text, "#,###,###.00")
            End If
            If flxCV.TextMatrix(i, 6) = "20004" And flxCV.TextMatrix(i, 4) = "DEBIT" Then
                flxCV.TextMatrix(i, 2) = ""
            ElseIf flxCV.TextMatrix(i, 6) <> "20004" And flxCV.TextMatrix(i, 4) = "DEBIT" Then
                flxCV.TextMatrix(i, 2) = Format(myAcctTitle, "#,###,###.00")
            End If
        Next i
    End If
    
    If cboRefundSrc.Text = "O.R." Then
        flxCV.AddItem vbTab & "O.R.# " & txtRefundRef.Text & " Date: " & txtRefDate.Text
    ElseIf cboRefundSrc.Text = "SALES RETURN - CREDIT" Or cboRefundSrc.Text = "SALES RETURN - CASH" Then
        flxCV.AddItem vbTab & "Sales Return Slip# " & txtRefundRef.Text & " Date: " & txtRefDate.Text
    End If
    If lblAddRef.Caption <> "" Then flxCV.AddItem vbTab & lblAddRef.Caption
    flxCV.AddItem vbTab & "Reason: " & txtReason.Text
    Frame1.Enabled = False
    
    Call computeDMCM

End Sub

Private Sub flxPayee_Click()
    txtPayee.Text = flxPayee.TextMatrix(flxPayee.Row, 1)
    If Val(flxPayee.TextMatrix(flxPayee.Row, 2)) > 0 Then
        EMPID = flxPayee.TextMatrix(flxPayee.Row, 2)
        CUSTOMERID = 0
    ElseIf Val(flxPayee.TextMatrix(flxPayee.Row, 3)) > 0 Then
        EMPID = 0
        CUSTOMERID = flxPayee.TextMatrix(flxPayee.Row, 3)
    End If
    flxPayee.Visible = False
    Call txtPayee_LostFocus
End Sub

Private Sub flxPayee_EnterCell()
    CustSrch = True
    txtPayee.Text = flxPayee.TextMatrix(flxPayee.Row, 1)
    If Val(flxPayee.TextMatrix(flxPayee.Row, 2)) > 0 Then
        EMPID = flxPayee.TextMatrix(flxPayee.Row, 2)
        CUSTOMERID = 0
    ElseIf Val(flxPayee.TextMatrix(flxPayee.Row, 3)) > 0 Then
        EMPID = 0
        CUSTOMERID = flxPayee.TextMatrix(flxPayee.Row, 3)
    End If
End Sub

Private Sub flxPayee_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxPayee.Visible = False
        txtPayee.SetFocus
    End If
End Sub

Private Sub flxPayee_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPayee_Change
        flxPayee.Visible = False
        Call txtPayee_LostFocus
        cboTransaction.SetFocus
    End If
End Sub

Private Sub flxPayee_LostFocus()
    If Val(flxPayee.TextMatrix(flxPayee.Row, 2)) > 0 Then
        EMPID = flxPayee.TextMatrix(flxPayee.Row, 2)
        CUSTOMERID = 0
    ElseIf Val(flxPayee.TextMatrix(flxPayee.Row, 3)) > 0 Then
        EMPID = 0
        CUSTOMERID = flxPayee.TextMatrix(flxPayee.Row, 3)
    End If
    flxPayee.Visible = False
    Call txtPayee_LostFocus
End Sub

Private Sub mnuAddRow_Click()
Dim Ans As VbMsgBoxResult
    Ans = MsgBox("Do you want to add Account Title?, If No I will add Blank Row", vbQuestion + vbYesNo, "Validation")
    
    If Ans = vbNo Then
        flxCV.AddItem ""
    Else
        Call LoadAcctTitleTab
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete Row")
    If Ans = vbYes Then
        If (InStr(cboTransaction.Text, "COLLECTION REFUND") > 0 Or InStr(cboTransaction.Text, "SALES RETURN REFUND") > 0 Or InStr(cboTransaction.Text, "SERVICE RETURN REFUND") > 0) And txtRefundRef.Text <> "" And ActRow = 1 Then
            flxCV.Rows = 1
            flxCV.Rows = 2
            txtRefundRef.Text = ""
            txtRefDate.Text = ""
            txtRefAmt.Text = ""
            txtORamt.Text = ""
            txtRefCust.Text = ""
            txtReason.Text = ""
            lblAddRef.Caption = ""
            Frame1.Enabled = True
        End If
        If flxCV.Rows = 2 Then
            flxCV.Rows = 1
            flxCV.Rows = 2
        Else
            flxCV.RemoveItem ActRow
        End If
    End If
End Sub

Private Sub txtPayee_LostFocus()
Dim Ans As Integer
    If formAction = "" Then Exit Sub
    If txtPayee.Text = "" Then
        MsgBox "No Payee Found!.", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
    If flxPayee.Visible = False Then
        If Val(EMPID) > 0 Then
            Ans = MsgBox("This Payee is an Employee!. Do you want to Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Validation")
            If Ans = vbNo Then
                txtPayee.SetFocus
                Exit Sub
            End If
            txtPCV.SetFocus
        ElseIf Val(CUSTOMERID) > 0 Then
            Ans = MsgBox("This Payee is a Customer!. Do you want to Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Validation")
            If Ans = vbNo Then
                txtPayee.SetFocus
                Exit Sub
            End If
            txtPCV.SetFocus
        End If
    End If
    If flxPayee.Visible = True Then
        flxPayee.Visible = False
    End If
End Sub

Private Sub txtRefundRef_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    'strValid = "1234567890"
    'If KeyAscii > 26 Then
    '    If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    'End If
    If KeyAscii = 13 Then
        If cboRefundSrc.Text = "SALES RETURN - CREDIT" Or cboRefundSrc.Text = "SALES RETURN - CASH" Then
            txtReason.SetFocus
        Else
            txtRefAmt.SetFocus
        End If
    End If
End Sub

Private Sub txtRefundRef_LostFocus()
    If cboRefundSrc.Text = "SALES RETURN - CREDIT" And txtRefundRef.Text <> "" Then Call DisplaySalesReturn
    If cboRefundSrc.Text = "SALES RETURN - CASH" And txtRefundRef.Text <> "" Then Call DisplaySalesReturn
    
    If (cboRefundSrc.Text = "O.R." Or cboRefundSrc.Text = "C.R.") And txtRefundRef.Text <> "" Then Call DisplayORDetails
    If cboRefundSrc.Text = "" Then
        txtRefDate.Text = ""
        txtRefCust.Text = ""
        DP = ""
        SP = ""
        SVC = ""
        OT = ""
        txtRefAmt.Text = ""
        txtORamt.Text = ""
    End If
End Sub

Public Sub DisplayVoucher()
    Dim rsDisb As New ADODB.Recordset
    Dim Qry As String, Chrg As String, myBCODE As String, myBranch As String
    Dim i As Integer, Details As Variant
    
    myBCODE = ""
    myBranch = ""
    
    If PCVID = 0 Then Exit Sub
    '***Access
    'Qry = "SELECT DISBURSEMENT.* From DISBURSEMENT " & _
          "WHERE (((DISBURSEMENT.VOUCHERID)=" & txtCVid.Text & "));"
    'mySQL
    Qry = "SELECT DISBURSEMENT.* From DISBURSEMENT " & _
          "WHERE DISBURSEMENT.VOUCHERID=" & Val(PCVID) & ";"
    rsDisb.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsDisb.RecordCount = 0 Then Exit Sub
    rsDisb.MoveFirst
    cboTransType.Text = rsDisb!DEPARTMENT
    'Load data on Combobox
    DisbursementType = rsDisb!DISBURSETYPE
    Call CboDisplay
    
    
    myBCODE = IIf(IsNull(rsDisb!branchCode), "", rsDisb!branchCode)
    If DisbursementType = "ADMINISTRATIVE EXPENSE" Then
        myBranch = GetFieldValue("Select DESCRIPTION From settings Where `VALUE` = '" & myBCODE & "' ")
    Else
        myBranch = GetFieldValue("Select BRANCHNAME From branch Where BRANCHCODE = '" & myBCODE & "' ")
    End If
    cboBranch.Text = myBranch
    
    txtPayee.Text = rsDisb!payee
    cboTransaction.Text = rsDisb!AcctDesc
    txtPCV.Text = IIf(IsNull(rsDisb!CVNUMBER), "", rsDisb!CVNUMBER)
    txtDate.Text = rsDisb!CVDate
    txtApprovedBy.Text = IIf(IsNull(rsDisb!APPROVEDBY), "", rsDisb!APPROVEDBY)
    txtPreparedBy.Text = IIf(IsNull(rsDisb!PREPAREDBY), "", rsDisb!PREPAREDBY)
    txtTIN.Text = IIf(IsNull(rsDisb!TIN), "", rsDisb!TIN)
    txtSupplier.Text = IIf(IsNull(rsDisb!Supplier), "", rsDisb!Supplier)
    txtSupplierAddress.Text = IIf(IsNull(rsDisb!SupplierAddress), "", rsDisb!SupplierAddress)
    myJVNO = GetFieldValue("SELECT JVNO FROM journal WHERE REFNO ='" & txtPCV.Text & "' LIMIT 1")
    If myJVNO = "" Then
        With flxCV
            .Rows = 1
            .AddItem vbTab & rsDisb!Particulars
            If Not IsNull(rsDisb!Details) Then
                Details = Split(rsDisb!Details, "|")
                For i = LBound(Details) To UBound(Details)
                    .AddItem vbTab & Details(i)
                Next i
            End If
            .TextMatrix(1, 2) = FormatNumber(rsDisb!Amount, 2)
            txtFlx.Visible = False
        End With
    Else
        Qry = "SELECT ACCT_CODE,DEBIT, CREDIT,DRCR FROM journaldtls WHERE JVNO = '" & myJVNO & "' ORDER BY SQNO"
        With rsAnyUse
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxCV.Rows = 1
                flxCV.AddItem vbTab & rsDisb!Particulars
                If Not IsNull(rsDisb!Details) Then
                    Details = Split(rsDisb!Details, "|")
                    For i = LBound(Details) To UBound(Details)
                        flxCV.AddItem vbTab & Details(i)
                    Next i
                End If
                
                For i = 1 To .RecordCount
                    flxCV.TextMatrix(i, 6) = IIf(IsNull(!acct_code), "", !acct_code)
                    flxCV.TextMatrix(i, 2) = IIf((!Debit = "0"), "", Format(!Debit, "#,###,###.00"))
                    flxCV.TextMatrix(i, 3) = IIf((!Credit = "0"), "", Format(!Credit, "#,###,###.00"))
                    flxCV.TextMatrix(i, 4) = IIf(IsNull(!DRCR), "", !DRCR)
                    .MoveNext
                Next i
            End If
            .Close
        End With
        txtFlx.Visible = False
    End If
    
    For i = 1 To flxCV.Rows - 1
        If flxCV.TextMatrix(i, 6) = "10401" Or flxCV.TextMatrix(i, 6) = "20004" Or flxCV.TextMatrix(i, 6) = "20011" Then
            chkAutoCompute.Visible = True
            chkAutoCompute.Value = 1
        End If
    Next i
    
    txtPrnCopy.Text = IIf(IsNull(rsDisb!PRINTCOPY), "0", rsDisb!PRINTCOPY)
    If InStr(cboTransaction.Text, "COLLECTION REFUND") > 0 Or InStr(cboTransaction.Text, "SALES RETURN REFUND") > 0 Or InStr(cboTransaction.Text, "SERVICE RETURN REFUND") > 0 Then Chrg = IIf(IsNull(rsDisb!ChargeTo), "", rsDisb!ChargeTo)
    rsDisb.Close
    If (InStr(cboTransaction.Text, "COLLECTION REFUND") > 0 Or InStr(cboTransaction.Text, "SALES RETURN REFUND") > 0 Or InStr(cboTransaction.Text, "SERVICE RETURN REFUND") > 0) And Chrg <> "" Then
        SSTab1.Tab = 3
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.Visible = True
        If InStr(Chrg, "SR#") > 0 Then
            cboRefundSrc.Text = "SALES RETURN - CREDIT"
            cboRefundSrc.Text = "SALES RETURN - CASH"
            txtRefundRef.Text = Right$(Chrg, Len(Chrg) - 3)
        ElseIf InStr(Chrg, "OR#") > 0 Then
            cboRefundSrc.Text = "O.R."
            txtRefundRef.Text = Right$(Chrg, Len(Chrg) - 3)
        End If
        txtRefundRef.Tag = Chrg
        Call txtRefundRef_LostFocus
        For i = flxCV.Rows - 1 To 1 Step -1
            If InStr(flxCV.TextMatrix(i, 1), "Reason:") > 0 Then
                txtReason.Text = Replace(flxCV.TextMatrix(i, 1), "Reason:", "")
            End If
        Next i
        Frame1.Enabled = False
    End If
    Call LockData(False)
    frameTtlDMCM.Visible = True
    flxPayee.Visible = False
    Call computeDMCM
    
End Sub

Function ValidatePCV() As String
    Dim vMsg As String, Qry As String, myBCODE As String
    Dim i As Integer
    Dim TtlDebit As Currency, TtlCredit As Currency
    
    vMsg = ""
    TtlDebit = 0
    TtlCredit = 0
    If txtPayee.Text = "" Then vMsg = "Incomplete entry on field Payee."
    If cboTransType.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Transaction Type"
    If cboTransaction.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Account Title"
    'old code If PayeeType < 3 And txtPayeeID.Text = "" Then VMsg = VMsg & vbCrLf & "Please select a Payee from list."
    If txtPCV.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field PCV Number"
    Else
        If formAction = "Add" Then
            '***Access
            'Qry = "SELECT DISBURSEMENT.CVNUMBER From DISBURSEMENT WHERE (((DISBURSEMENT.CVNUMBER)=" & txtPCV.Text & "));"
            'mySQL
            Qry = "SELECT DISBURSEMENT.CVNUMBER From DISBURSEMENT WHERE DISBURSEMENT.CVNUMBER='" & txtPCV.Text & "';"
            If CtrlNumExist(Qry) Then vMsg = vMsg & vbCrLf & "PCV number already exist"
        End If
    End If

    'remove last empty row
    For i = flxCV.Rows - 1 To 2 Step -1
        If flxCV.TextMatrix(i, 1) = "" Then flxCV.RemoveItem i
    Next i
    If flxCV.Rows >= 2 Then
        If flxCV.TextMatrix(1, 1) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Particulars"
        If flxCV.TextMatrix(1, 2) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Amount"
    End If
    If txtApprovedBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Approved By"
    If txtPreparedBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Prepared By"
    
    For i = 1 To flxCV.Rows - 1
        If flxCV.TextMatrix(i, 4) = "DEBIT" Then
            TtlDebit = TtlDebit + Val(Replace(flxCV.TextMatrix(i, 2), ",", ""))
        Else
            TtlCredit = TtlCredit + Val(Replace(flxCV.TextMatrix(i, 3), ",", ""))
        End If
        
        If flxCV.TextMatrix(i, 6) = "10401" Then
            If txtTIN.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field TIN"
            If txtSupplier.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field SUPPLIER"
            If txtSupplierAddress.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field SUPPLIER ADDRESS"
        End If
    Next i
    If TtlCredit <> TtlDebit Then
        If txtDisbAmnt.Text = "" Or Val(txtDisbAmnt.Text) = 0 Then
            vMsg = vMsg & vbCrLf & "No Disbursement Amount Found!."
        End If
        'vMsg = vMsg & vbCrLf & "Total Credit Amount is not Equal to Total Debit Amount"
    End If
    
    If DisbursementType = "ADMINISTRATIVE EXPENSE" Then
        myBCODE = GetFieldValue("Select `VALUE` From settings Where DESCRIPTION = '" & cboBranch.Text & "' ")
    Else
        myBCODE = GetFieldValue("Select BRANCHCODE From branch Where BRANCHNAME = '" & cboBranch.Text & "' ")
    End If
    If myBCODE = "" Then
        vMsg = vMsg & vbCrLf & "No Branch Found!."
    End If
    
    
    ValidatePCV = vMsg
End Function


Public Sub ComputeAcctTitle()
    Call CreateTempTableAcctTitle
    Qry = "call accounting_corp.sp_auto_compute_acct_entries(" & Val(myTransID) & "," & IIf(Val(Replace(txtFlx.Text, ",", "")) = 0, Val(Replace(txtTtlCredit.Text, ",", "")), Val(Replace(txtFlx.Text, ",", ""))) & "," & Val(UserID) & ")"
    
    With rsAnyUse
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            For i = 1 To flxCV.Rows - 1
                If flxCV.TextMatrix(i, 6) <> "" Then
                    If flxCV.TextMatrix(i, 4) = "DEBIT" Then
                        flxCV.TextMatrix(i, 2) = Format(!Debit, "#,###,###.00")
                    Else
                        flxCV.TextMatrix(i, 3) = Format(!Credit, "#,###,###.00")
                    End If
                    .MoveNext
                End If
            Next i
        End If
        .Close
    End With
End Sub


Sub CreateTempTableAcctTitle()
On Error GoTo er

        LYRcon.Execute "drop temporary table if exists accounting_corp.tmp_acct_entries;"

         DoEvents
        LYRcon.Execute " create temporary table accounting_corp.tmp_acct_entries( " & _
                       "    Acct_Code Varchar(30)    " & _
                       "   ,DEBIT     Decimal(24,2)  " & _
                       "   ,CREDIT    Decimal(24,2)  " & _
        ");"
        
        Qry = "Insert into accounting_corp.tmp_acct_entries(Acct_Code,DEBIT,CREDIT) VALUES "
        For i = 1 To flxCV.Rows - 1
            If i > 1 Then Qry = Qry & ","
            Qry = Qry & "('" & flxCV.TextMatrix(i, 6) & "'," & Val(Replace(flxCV.TextMatrix(i, 2), ",", "")) & "," & Val(Replace(flxCV.TextMatrix(i, 3), ",", "")) & ") "
        Next i
        
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTableAcctTitle()")
End Sub

Sub CreateTempTable()
Dim SEQNO As Integer
On Error GoTo er

        LYRcon.Execute "Drop Temporary Table IF  Exists accounting_corp.TMP_CVDETAILS;"

         DoEvents
        LYRcon.Execute " Create Temporary Table IF Not Exists accounting_corp.TMP_CVDETAILS( " & _
        "    Acct_Code      Varchar(30)                 " & _
        "   ,DRCR           Varchar(30)                 " & _
        "   ,DEBIT          Decimal(24,2)               " & _
        "   ,CREDIT         Decimal(24,2)               " & _
        "   ,SEQNO          INT(11)                     " & _
        ");"
        
        Qry = "Insert into accounting_corp.TMP_CVDETAILS(Acct_Code,DRCR,DEBIT,CREDIT,SEQNO) VALUES "
        SEQNO = 0
        For i = 1 To flxCV.Rows - 1
            If i > 1 Then Qry = Qry & ","
            SEQNO = SEQNO + 1
            Qry = Qry & "('" & flxCV.TextMatrix(i, 6) & "','" & flxCV.TextMatrix(i, 4) & "'," & Val(Replace(flxCV.TextMatrix(i, 2), ",", "")) & "," & Val(Replace(flxCV.TextMatrix(i, 3), ",", "")) & "," & SEQNO & ") "
        Next i
        
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Function CVDetails() As String
    Dim Tmp As String
    Dim i As Integer
    
    Tmp = ""
'    If flxCV.Rows <= 2 Then
'        If DisburseType = 1 Then    'Cash Advance
'            CVDetails = Right$(cboTransaction.Text, Len(cboTransaction.Text) - (InStr(cboTransaction.Text, "[") - 1))
'        Else
'            CVDetails = ""
'        End If
'    Else
        For i = 2 To flxCV.Rows - 1
            Tmp = IIf(Tmp = "", flxCV.TextMatrix(i, 1), Tmp & "|" & flxCV.TextMatrix(i, 1))
        Next i
'        If DisburseType = 1 Then    'Cash Advance
'            Tmp = Right$(cboTransaction.Text, Len(cboTransaction.Text) - (InStr(cboTransaction.Text, "[") - 1)) & "|" & Tmp
'        End If
        CVDetails = Tmp
'    End If
End Function

Private Sub SaveVoucher()
    Dim rsDisburse As New ADODB.Recordset
    Dim rsCashAdv As New ADODB.Recordset
    Dim rsAgentOrder As New ADODB.Recordset
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String, nBal As Single, X As Integer, MyDP As Single, MyOR As String
    Dim MyDisburseType As String, myBCODE As String
    Dim myAmount As Currency
    
    For i = 1 To flxCV.Rows - 1
        If flxCV.TextMatrix(i, 4) = "CREDIT" Then
            myAmount = Val(Replace(flxCV.TextMatrix(i, 3), ",", ""))
        End If
    Next i
    
    If DisbursementType = "ADMINISTRATIVE EXPENSE" Then
        myBCODE = GetFieldValue("Select `VALUE` From settings Where DESCRIPTION = '" & cboBranch.Text & "' ")
    Else
        myBCODE = GetFieldValue("Select BRANCHCODE From branch Where BRANCHNAME = '" & cboBranch.Text & "' ")
    End If
    If formAction = "Add" Then
        Call CreateTempTable
        Qry = "call accounting_corp.sp_journal_create('" & sysBranchCode & "','CAR','PCV'," & Val(myTransID) & ",'" & txtPCV.Text & "'," & Val(UserID) & ",'" & Format(txtDate.Text, "YYYY-MM-DD") & "')"
        With rsAnyUse
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        End With
        
        LYRcon.Execute "INSERT INTO disbursement(CVDATE, `STATUS`, CVNUMBER, DISBURSETYPE, ACCTDESC, " & _
            "DEPARTMENT, PAYEE, PARTICULARS, AMOUNT, DETAILS, PREPAREDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE, DISBAMNT, BRANCHCODE,TIN,SUPPLIER,SUPPLIERADDRESS,CUSTOMERID,EMPID)" & _
            "VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "', 'Temp', '" & txtPCV.Text & "', " & _
            "'" & DisbursementType & "', '" & cboTransaction.Text & "', '" & cboTransType.Text & "', '" & txtPayee.Text & "', " & _
            "'" & flxCV.TextMatrix(1, 1) & "', '" & myAmount & "', '" & CVDetails & "', " & _
            "'" & txtPreparedBy.Text & "', '" & txtApprovedBy.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', '" & IIf((txtDisbAmnt.Text = "" Or Val(txtDisbAmnt.Text) = 0), myAmount, Val(Replace(txtDisbAmnt.Text, ",", ""))) & "','" & myBCODE & "','" & txtTIN.Text & "','" & txtSupplier.Text & "','" & txtSupplierAddress.Text & "', " & _
            "" & Val(CUSTOMERID) & "," & Val(EMPID) & ");"
            
        PCVID = GetFieldValue("SELECT MAX(VOUCHERID) FROM disbursement")
        myJVNO = GetFieldValue("SELECT JVNO FROM journal WHERE REFNO = '" & txtPCV.Text & "' LIMIT 1 ")
        LYRcon.Execute "UPDATE DISBURSEMENT SET JVNO = '" & myJVNO & "' WHERE VOUCHERID = " & Val(PCVID) & ";"
    Else
        Call CreateTempTable
        Qry = "call accounting_corp.sp_journal_update('" & sysBranchCode & "','CAR','PCV'," & Val(myTransID) & ",'" & txtPCV.Text & "'," & Val(UserID) & ",'" & Format(txtDate.Text, "YYYY-MM-DD") & "','" & myJVNO & "')"
        With rsAnyUse
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        End With
    
        LYRcon.Execute "UPDATE Disbursement SET CVDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "', `STATUS` = 'Temp', " & _
            "CVNUMBER = '" & txtPCV.Text & "', DISBURSETYPE = '" & DisbursementType & "', " & _
            "ACCTDESC = '" & cboTransaction.Text & "', DEPARTMENT = '" & cboTransType.Text & "', PAYEE = '" & txtPayee.Text & "', " & _
            "PARTICULARS = '" & flxCV.TextMatrix(1, 1) & "', AMOUNT = '" & myAmount & "', DISBAMNT = '" & myAmount & "', " & _
            "DETAILS = '" & CVDetails & "', PREPAREDBY = '" & txtPreparedBy.Text & "', APPROVEDBY = '" & txtApprovedBy.Text & "', " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', BRANCHCODE = '" & myBCODE & "', TIN = '" & txtTIN.Text & "', SUPPLIER = '" & txtSupplier.Text & "', SUPPLIERADDRESS = '" & txtSupplierAddress.Text & "', " & _
            "CUSTOMERID = " & Val(CUSTOMERID) & ", EMPID = " & Val(EMPID) & " " & _
            "Where Disbursement.VOUCHERID = " & Val(PCVID) & ";"
    End If
    
    Qry = "SELECT Disbursement.VOUCHERID From Disbursement Where DISBURSEMENT.CVNUMBER = '" & txtPCV.Text & "';"
    If formAction = "Add" Then PCVID = GetFieldValue(Qry)

    If InStr(cboTransaction.Text, "COLLECTION REFUND") > 0 Or InStr(cboTransaction.Text, "SALES RETURN REFUND") > 0 Or InStr(cboTransaction.Text, "SERVICE RETURN REFUND") > 0 Then
        If cboRefundSrc.Text = "SALES RETURN - CREDIT" Or cboRefundSrc.Text = "SALES RETURN - CASH" Then
            LYRcon.Execute "Update Disbursement SET CHARGETO = 'SR#" & txtRefundRef.Text & "' " & _
                "Where Disbursement.VOUCHERID = " & Val(PCVID) & ";"
        ElseIf cboRefundSrc.Text = "O.R." Then
            LYRcon.Execute "Update Disbursement SET CHARGETO = 'OR#" & txtRefundRef.Text & "' " & _
                "Where Disbursement.VOUCHERID = " & Val(PCVID) & ";"
        End If
    End If
End Sub


Private Sub UpdateFinalVoucher()
    Dim rsInv As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim Qry As String, InvType As String, MyInv As String, RetRem As String, ProRef As String, MyStat As String
    Dim OldBal As Currency, myTtl As Currency, CustID As Integer, MyRetID As Long, mySEQNO As Long
    
    If InStr(cboTransaction.Text, "COLLECTION REFUND") > 0 Or InStr(cboTransaction.Text, "SALES RETURN REFUND") > 0 Or InStr(cboTransaction.Text, "SERVICE RETURN REFUND") > 0 Then
        If InStr(cboRefundSrc.Text, "SALES RETURN") > 0 And txtRefundRef.Text <> "" Then
            '***Access
            'Qry = "SELECT SRETURNSP.RETNO, SRETURNSP.RETURNID, SRETURNSP.PROREF, SRETURNSP.REMARKS, SPINVOICE.SPINVNO, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, SPINVOICE.AMOUNT, SPINVOICE.INVTYPE, SPINVOICE.CUSTOMERID " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO WHERE (((SRETURNSP.RETNO)='" & txtRefundRef.Text & "'));"
            'mySQL
            If cboRefundSrc.Text = "SALES RETURN - CREDIT" Then
                Qry = "SELECT SRETURNSP.RETNO, SRETURNSP.RETURNID, SRETURNSP.PROREF, SRETURNSP.REMARKS, SPINVOICE.SPINVNO, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, SPINVOICE.AMOUNT, SPINVOICE.INVTYPE, SPINVOICE.CUSTOMERID " & _
                      "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO WHERE SRETURNSP.RETNO='" & txtRefundRef.Text & "' and SPINVOICE.CUSTOMERID = " & Val(CUSTOMERID) & ";"
            Else
                Qry = "SELECT SRETURNSP.RETNO, SRETURNSP.RETURNID, SRETURNSP.PROREF, SRETURNSP.REMARKS, SPINVOICE.SPINVNO, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, SPINVOICE.AMOUNT, SPINVOICE.INVTYPE, SPINVOICE.CUSTOMERID " & _
                      "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO WHERE SRETURNSP.RETNO='" & txtRefundRef.Text & "' ;"
            End If
            
            
            With rsInv
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    InvType = IIf(IsNull(!InvType), "", !InvType)
                    MyInv = !SPINVNO
                    CustID = IIf(IsNull(!CUSTOMERID), 0, !CUSTOMERID)
                    MyRetID = !RETURNID
                    RetRem = IIf(IsNull(!Remarks), "", !Remarks)
                    ProRef = IIf(IsNull(!ProRef), "", !ProRef)
                End If
                .Close
            End With
            If InvType = "CREDIT" Then  'Update Ledger
                 myTtl = Val(Replace(txtDisbAmnt.Text, ",", ""))
                 Qry = "SELECT ONACCTERMS.BALANCE AS BALANCE From ONACCTERMS " & _
                      "WHERE ONACCTERMS.CUSTOMERID=" & CustID & " AND SPINVNO = '" & MyInv & "';"
                 OldBal = Val(GetFieldValue(Qry))
                
                 Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From LEDGER WHERE (((LEDGER.SPINVNO)='" & MyInv & "'));"
                 mySEQNO = Val(GetFieldValue(Qry)) + 1
                 LYRcon.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, DEBIT, OSBAL, CUSTOMERID, SEQNO) " & _
                                "VALUES('" & MyInv & "', 'Refund', '" & txtPCV.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
                                "" & myTtl & ", " & OldBal + myTtl & ", " & CustID & "," & mySEQNO & " );"
                
                'Update Customer and Invoice Balance
                MyStat = IIf(OldBal - myTtl >= 1, "OPEN", "CLOSED")
                
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & OldBal + myTtl & ", ONACCTERMS.`STATUS` = '" & MyStat & "' " & _
                               "WHERE ONACCTERMS.SPINVNO='" & MyInv & "';"
                
                'Update Customer Balance
                Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                      "WHERE ONACCTERMS.CUSTOMERID=" & CustID & " AND ONACCTERMS.`STATUS`='OPEN';"
                OldBal = Val(GetFieldValue(Qry))
                
                LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = " & OldBal & " WHERE CUSTOMER.CUSTOMERID=" & CustID & ";"
            End If
            RetRem = IIf(RetRem <> "", RetRem & ", ", "") & "Refund from PCV#" & txtPCV.Text & " Amt: " & flxCV.TextMatrix(1, 2)
            If ProRef <> "" Then
                LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = '" & Left(RetRem, 250) & "' WHERE SRETURNSP.RETURNID=" & MyRetID & ";"
            Else
                LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = '" & Left(RetRem, 250) & "', SRETURNSP.PROREF = '" & txtPCV.Text & "' WHERE SRETURNSP.RETURNID=" & MyRetID & ";"
            End If
        ElseIf (cboRefundSrc.Text = "O.R." Or cboRefundSrc.Text = "C.R.") And txtRefundRef.Text <> "" Then
            
            'Update O.R. if present in DP Ledger
            Qry = "SELECT DPLEDGER.* From DPLEDGER WHERE DPLEDGER.ORNO='" & txtRefundRef.Text & "';"
            With rsInv
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    OldBal = !Balance
                    !Balance = OldBal - Val(Replace(flxCV.TextMatrix(1, 2), ",", ""))
                    !STATUS = IIf(OldBal - Val(Replace(flxCV.TextMatrix(1, 2), ",", "")) > 0, "OPEN", "CLOSED")
                    RetRem = "Refund from PCV#" & txtPCV.Text & " Amt:" & flxCV.TextMatrix(1, 2) & IIf(IsNull(!Remarks), "", ", " & !Remarks)
                    !Remarks = Left$(RetRem, 200)
                    .Update
                End If
                .Close
            End With
            'Update Ledger if OR exist
            Qry = "SELECT LEDGER.* From Ledger WHERE LEDGER.REFTYPE='Payment' AND LEDGER.REFNO='" & txtRefundRef.Text & "';"
            With rsInv
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    MyInv = !SPINVNO
                    CustID = !CUSTOMERID
                    MyRetID = !SEQNO + 1
                Else
                    MyInv = ""
                End If
                .Close
            End With
            If MyInv <> "" Then 'Add Ledger
                Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE ONACCTERMS.SPINVNO='" & MyInv & "';"
                OldBal = Val(GetFieldValue(Qry))
                myTtl = Val(Replace(flxCV.TextMatrix(1, 2), ",", ""))
                
                LYRcon.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, DEBIT, OSBAL, CUSTOMERID, SEQNO) " & _
                                "VALUES('" & MyInv & "', 'Refund', '" & txtPCV.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
                                "" & myTtl & ", " & OldBal + myTtl & ", " & CustID & ", " & MyRetID & ");"
            
            End If
        End If
    End If
            
End Sub

Private Sub PrintPCV()
    Dim PrnVal(0 To 11) As String

    PrnVal(0) = "" 'txtPCV.Text 'PCVNO
    PrnVal(1) = txtPayee.Text   'PAYEE
    PrnVal(2) = txtDate.Text    'DATE
    PrnVal(3) = cboTransaction.Text   'TITLE
    PrnVal(4) = ""  'Amount
    PrnVal(5) = ""  'TOTAL
    PrnVal(6) = ""  'txtPreparedBy.Text
    PrnVal(7) = ""  'txtApprovedBy.Text
    PrnVal(8) = ""  'txtPayee.Text
    PrnVal(10) = flxCV.TextMatrix(0, 2)
    PrnVal(11) = flxCV.TextMatrix(0, 3)
    PrnVal(9) = cboTransType.Text
    Call PrintMyForm(PrnVal, "Petty Cash Voucher", flxCV)
End Sub

Function SPinFlx(myPN As String) As Boolean
    Dim i As Integer
    
    With flxCV
        For i = 0 To .Rows - 1
            'If myPN = .TextMatrix(i, 1) Then
            If myPN = .TextMatrix(i, 6) Then
                SPinFlx = True
                Exit Function
            End If
        Next i
    End With
    SPinFlx = False
    Exit Function
End Function

Public Sub LoadAcctTitleTab()
    If formAction = "" Then Exit Sub
    With flxAcct_Title
        .Cols = 5
        .FormatString = " |^ACCOUNT CODE |^ACCOUNT TITLE"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.2       'ACCOUNT CODE
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.7       'ACCOUNT TITLE
        .ColWidth(3) = 0 '(.Width - .ColWidth(0)) * 0.23   'DRCR
        .ColWidth(4) = 0 '(.Width - .ColWidth(0)) * 0.45   'Formula
    End With
    LoadTab = False
    myTransID = 0
    If cboTransaction.Text = "" Then Exit Sub
    myTransID = arrAcctDesc(cboTransaction.ListIndex, 0)
    chkAutoCompute.Visible = False

    Qry = "call accounting_corp.sp_display_acctTitle('PCV','LYRMparts','" & Val(myTransID) & "')"
    If rsAnyUse.State = adStateOpen Then rsAnyUse.Close
    rsAnyUse.Open Qry, LYRcon, adOpenKeyset, adOpenKeyset
    If rsAnyUse.RecordCount > 0 Then
        rsAnyUse.MoveFirst
        flxCV.Rows = rsAnyUse.RecordCount + 1
        'flxAcct_Title.Rows = rsAnyUse.RecordCount + 1
        'sstabDisplayAcctTitle.Visible = True
        'sstabDisplayAcctTitle.Left = (Me.Width - sstabDisplayAcctTitle.Width) / 2
        
        For i = 1 To flxCV.Rows - 1
'            flxAcct_Title.TextMatrix(i, 1) = rsAnyUse!acct_code
'            flxAcct_Title.TextMatrix(i, 2) = rsAnyUse!Acct_Tittle
'            flxAcct_Title.TextMatrix(i, 3) = rsAnyUse!DRCR
'            flxAcct_Title.TextMatrix(i, 4) = rsAnyUse!Formula
'            If rsAnyUse!acct_code = "10401" Or rsAnyUse!acct_code = "20004" Then
'                chkAutoCompute.Visible = True
'            End If
'.AddItem vbTab & flxAcct_Title.TextMatrix(i, 2) & vbTab & vbTab & vbTab & flxAcct_Title.TextMatrix(i, 3) & vbTab & flxAcct_Title.TextMatrix(i, 4) & vbTab & flxAcct_Title.TextMatrix(i, 1)
            flxCV.TextMatrix(i, 1) = rsAnyUse!Acct_Tittle
            flxCV.TextMatrix(i, 2) = "" 'debit
            flxCV.TextMatrix(i, 3) = "" 'credit
            flxCV.TextMatrix(i, 4) = rsAnyUse!DRCR
            flxCV.TextMatrix(i, 5) = rsAnyUse!Formula 'credit
            flxCV.TextMatrix(i, 6) = rsAnyUse!acct_code
            If rsAnyUse!acct_code = "10401" Or rsAnyUse!acct_code = "20004" Or rsAnyUse!acct_code = "20011" Then
                chkAutoCompute.Visible = True
            End If
            rsAnyUse.MoveNext
        Next i
        rsAnyUse.Close
    End If
    flxCV.SetFocus
End Sub


Public Sub CboDisplay()
Dim myTransType As String
On Error GoTo er

    If DisbursementType = "ADMINISTRATIVE EXPENSE" Then
        If cboTransType.Text <> "OTHER DISBURSEMENT" Then
            myTransType = "ADMINISTRATIVE EXPENSE"
        Else
            myTransType = "OTHER DISBURSEMENT"
        End If
    ElseIf DisbursementType = "OTHER BRANCH EXPENSE" Then
        If cboTransType.Text <> "OTHER DISBURSEMENT" Then
            myTransType = "OPERATING EXPENSE"
        Else
            myTransType = "OTHER DISBURSEMENT"
        End If
    Else
        myTransType = cboTransType.Text
    End If
    'Get Data from database
    DoEvents
    Qry = "call accounting_corp.sp_trans_legend('PCV','LYRMParts','" & myTransType & "','DBM')"
    If rsAcctDesc.State = adStateOpen Then rsAcctDesc.Close
    rsAcctDesc.Open Qry, LYRcon, adOpenStatic, adLockOptimistic
    If rsAcctDesc.State = 1 Then
       Call modObjTools.addToListObjectWithArrayGivenField(cboTransaction, rsAcctDesc, arrAcctDesc, "Trans_Legend_ID", "Transaction")
    End If
    If rsAcctDesc.State <> 0 Then rsAcctDesc.Close

    Call DisplayPayee
    'flxPayee.Visible = False
    
    If DisbursementType = "ADMINISTRATIVE EXPENSE" Then
        Label2(3).Caption = "DEPT.:"
        If cboTransType.Text = "LYRMP ADMIN EXPENSE" Then
            'Get Data from database sp_get_Department
            DoEvents
            Qry = " Select DESCRIPTION as DEPARTMENT, `VALUE` as DEPCODE From settings Where REMARKS = 'LYRMP ADMIN'"
            If rsAcctDesc.State = adStateOpen Then rsAcctDesc.Close
            rsAcctDesc.Open Qry, LYRcon, adOpenStatic, adLockOptimistic
            If rsAcctDesc.State = 1 Then
                Call modObjTools.addToListObjectWithArrayGivenField(cboBranch, rsAcctDesc, arrBRANCH, "DEPCODE", "DEPARTMENT")
            End If
            If rsAcctDesc.State <> 0 Then rsAcctDesc.Close
        ElseIf cboTransType.Text = "HEAD OFFICE EXPENSE" Then
            'Get Data from database sp_get_Department
            DoEvents
            Qry = " Select DESCRIPTION as DEPARTMENT, `VALUE` as DEPCODE From settings Where REMARKS = 'SUPPORT GROUP'"
            If rsAcctDesc.State = adStateOpen Then rsAcctDesc.Close
            rsAcctDesc.Open Qry, LYRcon, adOpenStatic, adLockOptimistic
            If rsAcctDesc.State = 1 Then
                Call modObjTools.addToListObjectWithArrayGivenField(cboBranch, rsAcctDesc, arrBRANCH, "DEPCODE", "DEPARTMENT")
            End If
            If rsAcctDesc.State <> 0 Then rsAcctDesc.Close
        Else
            'Get Data from database sp_get_Department
            DoEvents
            Qry = " Select DESCRIPTION as DEPARTMENT, `VALUE` as DEPCODE From settings Where REMARKS IN ('SUPPORT GROUP','LYRMP ADMIN')"
            If rsAcctDesc.State = adStateOpen Then rsAcctDesc.Close
            rsAcctDesc.Open Qry, LYRcon, adOpenStatic, adLockOptimistic
            If rsAcctDesc.State = 1 Then
                Call modObjTools.addToListObjectWithArrayGivenField(cboBranch, rsAcctDesc, arrBRANCH, "DEPCODE", "DEPARTMENT")
            End If
            If rsAcctDesc.State <> 0 Then rsAcctDesc.Close
        End If
    ElseIf DisbursementType = "OTHER BRANCH EXPENSE" Then
        DoEvents
        Call modMotoR.sp_get_Branch(rsAcctDesc)
        If rsAcctDesc.State = 1 Then
            Call modObjTools.addToListObjectWithArrayGivenField(cboBranch, rsAcctDesc, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
        End If
        If rsAcctDesc.State <> 0 Then rsAcctDesc.Close
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub computeDMCM()
Dim TtlDebit As Currency, TtlCredit As Currency, TtlDisbAmnt As Currency

    For i = 1 To flxCV.Rows - 1
        TtlDebit = TtlDebit + Val(Replace(flxCV.TextMatrix(i, 2), ",", ""))
        TtlCredit = TtlCredit + Val(Replace(flxCV.TextMatrix(i, 3), ",", ""))
    Next i
    
    For i = 1 To flxCV.Rows - 1
        If flxCV.TextMatrix(i, 6) = "10001" Then
            TtlDisbAmnt = TtlDisbAmnt + Val(Replace(flxCV.TextMatrix(i, 3), ",", ""))
        End If
    Next i
    txtTtlDebit.Text = Format(TtlDebit, "#,###,###.00")
    txtTtlCredit.Text = Format(TtlCredit, "#,###,###.00")
    txtDisbAmnt.Text = Format(TtlDisbAmnt, "#,###,###.00")
End Sub

Public Sub DisplayPayee()
    Qry = "SELECT  concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS PAYEE,EMPLOYEES.EMPID, 0 CUSTOMERID " & _
          "From EMPLOYEES WHERE EMPLOYEES.DEPARTMENT<>'Default' AND EMPLOYEES.INACTIVE=0 " & _
          "Union All " & _
          "SELECT CUSTOMER.CUSTOMERNAME as Payee, 0, CUSTOMER.CUSTOMERID " & _
          "From CUSTOMER " & _
          "GROUP BY PAYEE " & _
          "ORDER BY PAYEE;"
    
    If rsPayee.State = adStateOpen Then rsPayee.Close
    rsPayee.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxPayee.DataSource = rsPayee
    flxPayee.Refresh
    flxPayee.ColWidth(0) = 0
    flxPayee.ColWidth(1) = txtPayee.Width
End Sub
Private Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry
    cboTransType.Locked = DataEntry
    cboBranch.Locked = DataEntry
    
    txtPayee.Locked = DataEntry
    cboTransaction.Locked = DataEntry
    txtPCV.Locked = DataEntry
    txtDate.Locked = DataEntry
    
    txtTtlCredit.Locked = DataEntry
    txtTtlDebit.Locked = DataEntry
    txtDisbAmnt.Locked = DataEntry
    
    txtPreparedBy.Locked = DataEntry
    txtApprovedBy.Locked = DataEntry
    
    txtTIN.Locked = DataEntry
    txtSupplier.Locked = DataEntry
    txtSupplierAddress.Locked = DataEntry
    
    If cboTransType.Locked Then
        cboTransType.BackColor = txtHoldGrayBackColor.BackColor
        cboBranch.BackColor = txtHoldGrayBackColor.BackColor
        txtPayee.BackColor = txtHoldGrayBackColor.BackColor
        cboTransaction.BackColor = txtHoldGrayBackColor.BackColor
        txtPCV.BackColor = txtHoldGrayBackColor.BackColor
        txtDate.BackColor = txtHoldGrayBackColor.BackColor
        txtTtlCredit.BackColor = txtHoldGrayBackColor.BackColor
        txtTtlDebit.BackColor = txtHoldGrayBackColor.BackColor
        txtDisbAmnt.BackColor = txtHoldGrayBackColor.BackColor
        txtPreparedBy.BackColor = txtHoldGrayBackColor.BackColor
        txtApprovedBy.BackColor = txtHoldGrayBackColor.BackColor
        txtTIN.BackColor = txtHoldGrayBackColor.BackColor
        txtSupplier.BackColor = txtHoldGrayBackColor.BackColor
        txtSupplierAddress.BackColor = txtHoldGrayBackColor.BackColor
    Else
        cboTransType.BackColor = txtHoldWhiteBackColor.BackColor
        cboBranch.BackColor = txtHoldWhiteBackColor.BackColor
        txtPayee.BackColor = txtHoldWhiteBackColor.BackColor
        cboTransaction.BackColor = txtHoldWhiteBackColor.BackColor
        txtPCV.BackColor = txtHoldWhiteBackColor.BackColor
        txtDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtTtlCredit.Locked = True
        txtTtlDebit.Locked = True
        txtDisbAmnt.Locked = True
        txtPreparedBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtApprovedBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtTIN.BackColor = txtHoldWhiteBackColor.BackColor
        txtSupplier.BackColor = txtHoldWhiteBackColor.BackColor
        txtSupplierAddress.BackColor = txtHoldWhiteBackColor.BackColor
    End If
End Sub


Private Sub Clearform()
    cboTransType.Text = ""
    
    'Frame 1
    txtPayee.Text = ""
    cboTransaction.Text = ""
    cboBranch.Text = ""
    txtPCV.Text = ""
    txtDate.Text = ""
    
    'clear flx
    flxCV.Rows = 1
    flxCV.Rows = 2
    
    'Frame 2
    txtTtlCredit.Text = ""
    txtTtlDebit.Text = ""
    txtDisbAmnt.Text = ""
    txtTIN.Text = ""
    txtSupplier.Text = ""
    txtSupplierAddress.Text = ""
    chkAutoCompute.Value = 0
    
    'signatory
    txtPreparedBy.Text = ""
    txtApprovedBy.Text = ""
    
    flxPayee.Visible = False
    txtPrnCopy.Text = ""
   
    'refund
    txtRefundRef.Text = ""
    txtRefundRef.Tag = ""
    txtRefDate.Text = ""
    txtRefCust.Text = ""
    txtRefAmt.Text = ""
   
    txtReason.Text = ""
    lblAddRef.Caption = ""
    cboRefundSrc.Text = ""
    Frame1.Enabled = True


    PCVID = 0
    myJVNO = ""
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->PRINT
'5 ->SAVE
'6 ->CANCEL

    'hide all button
    For i = 0 To 1
        cmdMainButton(i).Visible = False
    Next i

    For i = 3 To 8
        cmdMainButton(i).Visible = False
    Next i
    
    If DataEntry Then
        cmdMainButton(7).Visible = True
        cmdMainButton(8).Visible = True
    Else
        For i = 0 To 1
            cmdMainButton(i).Visible = True
        Next i
        For i = 3 To 8
            cmdMainButton(i).Visible = True
            cmdMainButton(7).Visible = False
            cmdMainButton(8).Visible = False
        Next i
    End If
End Sub

Private Sub cboTransaction_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If AcctSrch Then
        AcctSrch = False
        Exit Sub
    End If
    With cboTransaction
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                AcctSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
End Sub

Private Sub cboTransaction_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then AcctSrch = True
End Sub

Private Sub cboTransaction_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then

        If cboTransaction.Text = "" Then
            Exit Sub
        End If
        
        cbolistIndex = modObjTools.getListIndex(cboTransaction, cboTransaction.Text)
        
        If cbolistIndex < 0 Then
            cboTransaction.SetFocus
        Else
            cboTransaction.ListIndex = cbolistIndex
        End If
        
        txtPCV.SetFocus
    End If
End Sub

Private Sub cboTransaction_LostFocus()
    flxCV.Rows = 1
    flxCV.Rows = 2
    'flxAcct_Title.Rows = 1
    'flxAcct_Title.Rows = 2
    Call LoadAcctTitleTab
    txtTtlCredit.Text = ""
    txtTtlDebit.Text = ""
    txtDisbAmnt.Text = ""
    
    If cboTransaction.Text <> "" Then
        If InStr(cboTransaction.Text, "COLLECTION REFUND") > 0 Or InStr(cboTransaction.Text, "SALES RETURN REFUND") > 0 Or InStr(cboTransaction.Text, "SERVICE RETURN REFUND") > 0 Then
            If refundBtn = False Then
                Me.Height = Me.Height + SSTab1.Height
            End If
            SSTab1.Tab = 3
            SSTab1.TabEnabled(3) = True
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(0) = False
            SSTab1.Visible = True
            
            cboRefundSrc.Clear
            If cboTransaction.Text = "COLLECTION REFUND" Then
                cboRefundSrc.AddItem "C.R."
            ElseIf cboTransaction.Text = "SALES RETURN REFUND" Then
                cboRefundSrc.AddItem "SALES RETURN - CREDIT"
                cboRefundSrc.AddItem "SALES RETURN - CASH"
            ElseIf cboTransaction.Text = "SERVICE RETURN REFUND" Then
                cboRefundSrc.AddItem "O.R."
            End If
            refundBtn = True
        Else
            If refundBtn = True Then
                Me.Height = Me.Height - SSTab1.Height
                refundBtn = False
            End If
            SSTab1.Tab = 3
            SSTab1.TabEnabled(3) = False
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(0) = False
            SSTab1.Visible = False
        End If
    End If
End Sub

Private Sub cboTransType_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If AcctSrch Then
        AcctSrch = False
        Exit Sub
    End If
    With cboTransType
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                AcctSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With

End Sub

Private Sub cboTransType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then AcctSrch = True
End Sub

Private Sub cboTransType_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then

        If cboTransType.Text = "" Then
            Exit Sub
        End If
        
        cbolistIndex = modObjTools.getListIndex(cboTransType, cboTransType.Text)
        
        If cbolistIndex < 0 Then
            cboTransType.SetFocus
        Else
            cboTransType.ListIndex = cbolistIndex
        End If
        
        txtPayee.SetFocus
    End If
End Sub

Private Sub cboTransType_LostFocus()
Call CboDisplay

End Sub

Private Sub chkAutoCompute_Click()
Dim myCreditAmnt As Currency

If chkAutoCompute.Value = 1 Then
    myCreditAmnt = 0
    For i = 1 To flxCV.Rows - 1
        If flxCV.TextMatrix(i, 3) <> "" Then
            myCreditAmnt = Val(Replace(flxCV.TextMatrix(i, 3), ",", ""))
        ElseIf myTransID = "364" And flxCV.TextMatrix(i, 4) = "DEBIT" Then
            myCreditAmnt = Val(Replace(flxCV.TextMatrix(i, 2), ",", ""))
        End If
    Next i
    
    If Val(myCreditAmnt) > 0 Then
        txtFlx.Text = myCreditAmnt
        Call ComputeAcctTitle
    ElseIf flxCV.TextMatrix(flxCV.Rows - 1, 4) = "CREDIT" And flxCV.TextMatrix(flxCV.Rows - 1, 3) = "" Then
        MsgBox "No Credit Amount Found!", vbInformation + vbOKOnly, "Validation"
        chkAutoCompute.Value = 0
        Exit Sub
    End If
Else
    If formAction = "Add" Then
        For i = 1 To flxCV.Rows - 1
            If flxCV.TextMatrix(i, 4) = "DEBIT" Then
                flxCV.TextMatrix(i, 2) = ""
            End If
        Next i
    End If
End If
txtTtlCredit.Text = ""
txtTtlDebit.Text = ""
txtDisbAmnt.Text = ""
Call computeDMCM
End Sub

Private Sub cmdCancel_Click()
    flxAcct_Title.Clear
    sstabDisplayAcctTitle.Visible = False
    LoadTab = False
    flxCV.SetFocus
    flxCV.Row = flxCV.Rows - 1
    flxCV.Col = 3
End Sub

Private Sub cmdInsert_Click()
With flxCV
    If .TextMatrix(.Rows - 1, 1) = "" Then 'Remove last row if empty
        If .Rows = 2 Then
            .Rows = 1
        Else
            .RemoveItem .Rows - 1
        End If
    End If
    For i = 1 To flxAcct_Title.Rows - 1
        If flxAcct_Title.TextMatrix(i, 0) <> "" Then
            If Not SPinFlx(flxAcct_Title.TextMatrix(i, 1)) Then
                .AddItem vbTab & flxAcct_Title.TextMatrix(i, 2) & vbTab & vbTab & vbTab & flxAcct_Title.TextMatrix(i, 3) & vbTab & flxAcct_Title.TextMatrix(i, 4) & vbTab & flxAcct_Title.TextMatrix(i, 1)
            End If
        End If
    Next i
    If .Rows = 1 Then .AddItem ""
End With
End Sub

Private Sub cmdMainButton_Click(Index As Integer)
    Dim Qry As String, MyAcc As String, vMsg As String
    Dim Ans As Integer
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=57;"

    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    If Index = 0 Then 'ope expense
        Call Clearform
        Call LockData(True)
        Call SetButtonVisibility(True)
        cboTransType.Clear
        cboTransType.Text = ""
        cboTransType.AddItem "OPERATING EXPENSE"
        cboTransType.AddItem "OTHER DISBURSEMENT"
        formAction = "Add"
        DisbursementType = "OPERATING EXPENSE"
        txtPreparedBy.Text = UserName
        txtApprovedBy.Text = GetFieldValue("Select MANAGER From Branch Where BRANCHCODE = '" & sysBranchCode & "'")
        txtDate.Text = Format(Now(), "MM/DD/YYYY")
        cboTransType.SetFocus
        refundBtn = False
        If refundBtn = False Then
            If SSTab1.Visible = True Then
                SSTab1.Visible = False
                Me.Height = Me.Height - SSTab1.Height
            End If
        End If
        Label2(3).Caption = "BRANCH:"
        cboBranch.BackColor = txtHoldGrayBackColor.BackColor
        Call CboDisplay
        cboBranch.Text = SysBranch
        cboBranch.Locked = True
        
    ElseIf Index = 1 Then 'admin expense
        Call Clearform
        Call LockData(True)
        Call SetButtonVisibility(True)
        cboTransType.Clear
        cboTransType.Text = ""
        cboTransType.AddItem ""
        cboTransType.AddItem "LYRMP ADMIN EXPENSE"
        cboTransType.AddItem "HEAD OFFICE EXPENSE"
        cboTransType.AddItem "HEAD LEASE FEE EXPENSE"
        cboTransType.AddItem "OTHER DISBURSEMENT"
        formAction = "Add"
        DisbursementType = "ADMINISTRATIVE EXPENSE"
        txtPreparedBy.Text = UserName
        txtApprovedBy.Text = GetFieldValue("Select MANAGER From Branch Where BRANCHCODE = '" & sysBranchCode & "'")
        txtDate.Text = Format(Now(), "MM/DD/YYYY")
        refundBtn = False
        cboTransType.SetFocus
        If refundBtn = False Then
            If SSTab1.Visible = True Then
                SSTab1.Visible = False
                Me.Height = Me.Height - SSTab1.Height
            End If
        End If
        Label2(3).Caption = "DEPT:"
'    ElseIf Index = 2 Then 'refund
'        Call Clearform
'        Call LockData(True)
'        Call SetButtonVisibility(True)
'        cboTransType.Clear
'        cboTransType.Text = ""
'        cboTransType.Text = "REFUNDS"
'        cboTransType.Locked = True
'        lblTransButton(0).Caption = "DISBURSEMENT"
'        lblTransButton(1).Caption = "DISBURSEMENT"
'        lblTransButton(0).Left = 7100
'        lblTransButton(1).Left = 7100
'        formAction = "Add"
'        DisbursementType = "REFUNDS"
'        txtPreparedBy.Text = UserName
'        txtApprovedBy.Text = GetFieldValue("Select MANAGER From Branch Where BRANCHCODE = '" & sysBranchCode & "'")
'        cboRefundSrc.Clear
'        cboRefundSrc.AddItem "SALES RETURN"
'        cboRefundSrc.AddItem "O.R."
'        cboRefundSrc.AddItem "C.R."
'        cboTransType.BackColor = txtHoldGrayBackColor.BackColor
'        txtDate.Text = Format(Now(), "MM/DD/YYYY")
'        Me.Height = Me.Height + SSTab1.Height + 180
'        refundBtn = True
'        SSTab1.Tab = 3
'        SSTab1.TabEnabled(3) = True
'        SSTab1.TabEnabled(2) = False
'        SSTab1.TabEnabled(1) = False
'        SSTab1.TabEnabled(0) = False
'        SSTab1.Visible = True
'        txtPayee.SetFocus
'        Call CboDisplay
    ElseIf Index = 3 Then 'other branch expense
        Call Clearform
        Call LockData(True)
        Call SetButtonVisibility(True)
        cboTransType.Clear
        cboTransType.Text = ""
        cboTransType.AddItem "OTHER BRANCH EXPENSE"
        cboTransType.AddItem "OTHER DISBURSEMENT"
        formAction = "Add"
        DisbursementType = "OTHER BRANCH EXPENSE"
        txtPreparedBy.Text = UserName
        txtApprovedBy.Text = GetFieldValue("Select MANAGER From Branch Where BRANCHCODE = '" & sysBranchCode & "'")
        txtDate.Text = Format(Now(), "MM/DD/YYYY")
        refundBtn = False
        If refundBtn = False Then
            If SSTab1.Visible = True Then
                SSTab1.Visible = False
                Me.Height = Me.Height - SSTab1.Height
            End If
        End If
        cboTransType.SetFocus
        Label2(3).Caption = "BRANCH"
        Call CboDisplay
    
    ElseIf Index = 4 Then 'edit
        If Val(PCVID) = 0 And Val(txtPrnCopy.Text) = 0 Then
            MsgBox "No PCV to be Edit", vbInformation + vbOKOnly, "Validation"
            Exit Sub
        ElseIf Val(PCVID) > 0 And Val(txtPrnCopy.Text) > 0 Then
            MsgBox "This PCV is already Final!", vbInformation + vbOKOnly, "Validation"
            Exit Sub
        End If
        formAction = "Edit"
        cboTransType.SetFocus
        Call LockData(True)
        Call SetButtonVisibility(True)
    ElseIf Index = 5 Then 'view summary
        Unload Me
        MotoRep.toolBarButton.Visible = False
        DisburseSumm.Show
    ElseIf Index = 6 Then 'print
        Ans = MsgBox("Do you want to Print Voucher?", vbQuestion + vbYesNo, "Print")
        If Ans = vbNo Then Exit Sub
        Call PrintPCV
        If Val(txtPrnCopy.Text) = 0 Then
            LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.PRINTCOPY = DISBURSEMENT.PRINTCOPY + 1, DISBURSEMENT.STATUS = 'Final' " & _
                           "WHERE DISBURSEMENT.VOUCHERID=" & Val(PCVID) & ";"
            txtPrnCopy.Text = "1"
            'Final Saving of involve tables
            'If DisburseType = 1 Then  'Cash Advance
            '    Call SaveCashAdvance
            '    txtCAnotice.Visible = False
            'End If
            Call UpdateFinalVoucher
        Else
            LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.PRINTCOPY = DISBURSEMENT.PRINTCOPY + 1 " & _
                           "WHERE DISBURSEMENT.VOUCHERID=" & Val(PCVID) & ";"
            txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
        End If
    ElseIf Index = 7 Then 'save
        vMsg = ValidatePCV
        If vMsg <> "" Then
            MsgBox vMsg, vbExclamation + vbOKOnly, "Validation Error"
            Exit Sub
        End If
        Ans = MsgBox("Do you want to Save Voucher?", vbQuestion + vbYesNo, "Save")
        If Ans = vbNo Then Exit Sub
        Call SaveVoucher
        Ans = MsgBox("Do you want to Print Voucher?", vbQuestion + vbYesNo, "Print")
        If Ans = vbYes Then
            Call PrintPCV
            LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.PRINTCOPY = 1, DISBURSEMENT.`STATUS` = 'Final' " & _
                           "WHERE DISBURSEMENT.VOUCHERID=" & Val(PCVID) & ";"
            txtPrnCopy.Text = "1"
            Call UpdateFinalVoucher
        Else
            txtPrnCopy.Text = "0"
        End If
        Call LockData(False)
        Call SetButtonVisibility(False)
        formAction = ""
    ElseIf Index = 8 Then 'cancel
        Ans = MsgBox("Do you want to Cancel " & IIf(formAction = "Add", "Voucher?", "Editing?"), vbQuestion + vbYesNo, "Cancel")
        If Ans = vbNo Then Exit Sub
        Call SetButtonVisibility(False)
        Call LockData(False)
        formAction = ""
    End If
    
End Sub

Private Sub cmdOK_Click()
    Call cmdInsert_Click
    
    For i = 1 To flxAcct_Title.Rows - 1
        If flxAcct_Title.TextMatrix(i, 0) = "" And flxAcct_Title.TextMatrix(i, 1) = "10001" Then
            MsgBox "Please Check the Cash On Hand Account Title"
            Exit Sub
        End If
    Next i
    sstabDisplayAcctTitle.Visible = False
    flxCV.SetFocus
    flxCV.Row = flxCV.Rows - 1
    flxCV.Col = 3
    LoadTab = False
End Sub

Private Sub flxAcct_Title_DblClick()
    If flxAcct_Title.TextMatrix(flxAcct_Title.Row, 0) = "" Then
        With flxAcct_Title
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
        End With
    Else
        flxAcct_Title.TextMatrix(flxAcct_Title.Row, 0) = ""
    End If
End Sub

Private Sub flxAcct_Title_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If flxAcct_Title.TextMatrix(flxAcct_Title.Row, 0) = "" Then
            With flxAcct_Title
                .Col = 0
                .CellFontName = "WINGDINGS"
                .CellFontSize = 12
                .Text = lblChk.Caption
                .CellAlignment = 4
            End With
        Else
            flxAcct_Title.TextMatrix(flxAcct_Title.Row, 0) = ""
        End If
    End If
End Sub

Private Sub flxAcct_Title_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxAcct_Title_DblClick
End Sub

Private Sub flxCV_Click()
    ActRow = flxCV.Row
    ActCol = flxCV.Col
    If LoadTab = True Then
        Exit Sub
    End If
    If ActRow = 0 Or ActCol = 0 Then Exit Sub
    If formAction = "" Then Exit Sub
    If flxCV.TextMatrix(ActRow, 4) = "CREDIT" Then
        If ActCol = 3 Then
            With txtFlx
                .Top = flxCV.CellTop + flxCV.Top
                .Left = flxCV.CellLeft + flxCV.Left
                .Width = flxCV.CellWidth
                .Text = flxCV.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        End If
    ElseIf flxCV.TextMatrix(ActRow, 4) = "DEBIT" Then
        If ActCol = 2 Then
            With txtFlx
                .Top = flxCV.CellTop + flxCV.Top
                .Left = flxCV.CellLeft + flxCV.Left
                .Width = flxCV.CellWidth
                .Text = flxCV.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        End If
    Else
        If ActCol = 1 And flxCV.TextMatrix(ActRow, 0) = "" Then
            With txtFlx
                .Top = flxCV.CellTop + flxCV.Top
                .Left = flxCV.CellLeft + flxCV.Left
                .Width = flxCV.CellWidth
                .Text = flxCV.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub flxCV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxCV_Click
End Sub

Private Sub flxCV_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuVoucher)
    End If
End Sub

Private Sub Form_Load()
    
    'Get Data from database sp_get_Department
'    DoEvents
'    Call modMotoR.sp_get_Branch(rsAcctDesc)
'    If rsAcctDesc.State = 1 Then
'        Call modObjTools.addToListObjectWithArrayGivenField(cboBranch, rsAcctDesc, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
'    End If
'    If rsAcctDesc.State <> 0 Then rsAcctDesc.Close
        
    flxCV.ColWidth(4) = 0
    flxCV.ColWidth(5) = 0
    flxCV.ColWidth(6) = 0
    Call LockData(False)
    '===================BUTTON POSITION===================
    If cmdMainButton(0).Visible = False Then
        'Align Save Button
        cmdMainButton(7).Left = cmdMainButton(0).Left
        cmdMainButton(7).Top = cmdMainButton(0).Top
        'Align Cancel Button
        cmdMainButton(8).Left = cmdMainButton(1).Left
        cmdMainButton(8).Top = cmdMainButton(1).Top
    End If
    MotoRep.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If formAction = "Add" Or formAction = "Edit" Then Exit Sub
MotoRep.Enabled = True
Unload Me
End Sub

Private Sub txtApprovedBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtApprovedBy, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtPayee.SetFocus
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxCV.SetFocus
        flxCV.Col = 1
    End If
End Sub

Private Sub txtDisbAmnt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtDisbAmnt, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then txtPreparedBy.SetFocus
End Sub

Private Sub txtFlx_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    strValid = "1234567890."
    
    If KeyAscii > 26 And ActCol = 2 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Call txtFlx_LostFocus
        flxCV.SetFocus
        If ActCol = 1 Then
            flxCV.Col = 2
        Else
            If ActRow = flxCV.Rows - 1 Then
                flxCV.AddItem ""
                flxCV.Col = 1: flxCV.Row = flxCV.Rows - 1
            ElseIf ActRow < flxCV.Rows - 1 Then
                flxCV.Col = 1: flxCV.Row = ActRow + 1
            End If
        End If
    End If
End Sub

Private Sub txtFlx_LostFocus()
    flxCV.TextMatrix(ActRow, ActCol) = Format(txtFlx.Text, "#,###,###.00")
    txtFlx.Visible = False
    Call computeDMCM
End Sub

Private Sub txtPayee_Change()
    If formAction = "" Then Exit Sub
    With flxPayee
        If CustSrch Then
            CustSrch = False
            Exit Sub
        End If
        If txtPayee.Text = "" Then .Visible = False: Exit Sub
        rsPayee.Find "[PAYEE] LIKE '" + txtPayee.Text + "%'", , adSearchForward, 1
        If Not rsPayee.EOF Then
            .Top = FraHeading.Top + txtPayee.Top + txtPayee.Height + 30
            .Left = FraHeading.Left + txtPayee.Left
            .Width = txtPayee.Width
            .Visible = True
            If Not CustBln Then
                txtPayee.Tag = Len(txtPayee.Text)
                CustBln = True
            End If
            .TopRow = rsPayee.AbsolutePosition - 1
            .Row = rsPayee.AbsolutePosition - 1
            .Col = 0: .ColSel = .Cols - 1
            
            txtPayee.Text = rsPayee!payee
            If Val(rsPayee!EMPID) > 0 Then
                EMPID = rsPayee!EMPID
                CUSTOMERID = 0
            ElseIf Val(rsPayee!CUSTOMERID) > 0 Then
                EMPID = 0
                CUSTOMERID = rsPayee!CUSTOMERID
            End If
            
            txtPayee.SelStart = Val(txtPayee.Tag)
            txtPayee.SelLength = Len(txtPayee.Text)
            CustBln = False
        Else
            EMPID = 0
            CUSTOMERID = 0
            txtPayee.Text = ""
            .Visible = False
        End If
    End With
End Sub

Private Sub txtPayee_Click()
Call txtPayee_Change
End Sub

Private Sub txtPayee_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If flxPayee.Visible Then
                flxPayee.SetFocus
            Else
                cboTransaction.SetFocus
            End If
        Case vbKeyEscape
            txtPayee.Text = ""
            flxPayee.Visible = False
        Case vbKeyDelete
            CustSrch = True
        Case vbKeyBack
            CustSrch = True
    End Select
End Sub

Private Sub txtPayee_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxPayee.Visible = False
        txtPCV.SetFocus
    End If
End Sub

Private Sub txtPCV_GotFocus()
    Dim rsMaxPCVNo As New ADODB.Recordset
    Dim Qry As String
    
    If formAction <> "Add" Then Exit Sub
    Qry = "SELECT Max(round(DISBURSEMENT.CVNUMBER)) AS MaxOfCVNUMBER FROM DISBURSEMENT;"

    With rsMaxPCVNo
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtPCV.Text = IIf(IsNull(.Fields(0)), "0001", Format(.Fields(0) + 1, "0000"))
            txtPCV.SelStart = 1
            txtPCV.SelLength = Len(txtPCV.Text)
        Else
            txtPCV.Text = ""
        End If
        .Close
    End With
End Sub

Private Sub txtPCV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboTransaction.SetFocus
End Sub

Private Sub txtPCV_LostFocus()
    Dim Qry As String
    
    If formAction <> "Add" Then Exit Sub
    If txtPCV.Text <> "" Then
        '***Access
        'Qry = "SELECT DISBURSEMENT.CVNUMBER From DISBURSEMENT WHERE (((DISBURSEMENT.CVNUMBER)=" & txtPCV.Text & "));"
        'mySQL
        Qry = "SELECT DISBURSEMENT.CVNUMBER From DISBURSEMENT WHERE DISBURSEMENT.CVNUMBER='" & txtPCV.Text & "';"
        If CtrlNumExist(Qry) Then
            MsgBox "PCV number already exist!", vbExclamation + vbOKOnly, "Validation Error"
        End If
    End If
End Sub

Private Sub txtPreparedBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtPreparedBy, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtApprovedBy.SetFocus
End Sub

