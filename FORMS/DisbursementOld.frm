VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form DisbursementOld 
   Caption         =   "Disbursement"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   11580
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtORamt 
      Height          =   285
      Left            =   8880
      TabIndex        =   97
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSalesRetID 
      Height          =   285
      Left            =   8280
      TabIndex        =   81
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAltA 
      Caption         =   "&A"
      Height          =   375
      Left            =   10200
      TabIndex        =   80
      Top             =   7440
      Width           =   135
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   3120
      TabIndex        =   79
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin lvButton.lvButtons_H cmdOR 
      Height          =   375
      Left            =   9240
      TabIndex        =   78
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Get O.R. Number"
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
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtCAlessCOM 
      Height          =   285
      Left            =   6720
      TabIndex        =   69
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCommOrigBal 
      Height          =   285
      Left            =   5520
      TabIndex        =   68
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   5040
      TabIndex        =   63
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboLYRg 
      Height          =   315
      Left            =   3240
      TabIndex        =   62
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame fraProfile 
      Height          =   4095
      Left            =   3240
      TabIndex        =   57
      Top             =   8400
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Image imgPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   120
         TabIndex        =   61
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label lblAddress 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   120
         TabIndex        =   60
         Top             =   2820
         Width           =   2535
      End
      Begin VB.Label lblPosition 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   120
         TabIndex        =   59
         Top             =   3315
         Width           =   2535
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   120
         TabIndex        =   58
         Top             =   3615
         Width           =   2535
      End
   End
   Begin VB.TextBox txtPreparedBy 
      Height          =   285
      Left            =   4320
      TabIndex        =   36
      Top             =   3840
      Width           =   3495
   End
   Begin VB.TextBox txtApprovedBy 
      Height          =   285
      Left            =   4320
      TabIndex        =   35
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox txtFlx 
      Height          =   285
      Left            =   7920
      TabIndex        =   34
      Top             =   4320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtCAnotice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   855
      Left            =   9240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   33
      Text            =   "DisbursementOld.frx":0000
      Top             =   2240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtCVid 
      Height          =   285
      Left            =   3120
      TabIndex        =   32
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPayeeID 
      Height          =   285
      Left            =   3600
      TabIndex        =   31
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCommOrderID 
      Height          =   285
      Left            =   4080
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboServiceUnit 
      Height          =   315
      Left            =   7200
      TabIndex        =   29
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtMotorID 
      Height          =   285
      Left            =   4560
      TabIndex        =   28
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPayee 
      Height          =   1575
      Left            =   4080
      TabIndex        =   27
      Top             =   7920
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
   Begin VB.Frame FraHeading 
      Height          =   1095
      Left            =   3000
      TabIndex        =   17
      Top             =   360
      Width           =   8175
      Begin VB.ComboBox cboAcctTitle 
         Height          =   315
         Left            =   1200
         TabIndex        =   25
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtPayee 
         Height          =   315
         Left            =   1200
         TabIndex        =   24
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   6360
         TabIndex        =   19
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtPCV 
         Height          =   315
         Left            =   6360
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "PAYEE"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "ACCT. TITLE"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "DATE"
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   21
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "PCV NO."
         Height          =   255
         Index           =   6
         Left            =   5640
         TabIndex        =   20
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Timer HideBtn2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9480
      Top             =   7800
   End
   Begin VB.Timer MoveUp2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9480
      Top             =   7440
   End
   Begin VB.Timer MoveUp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   7440
   End
   Begin VB.Timer HideBtn 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   7800
   End
   Begin VB.Timer ReArrange 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   8160
   End
   Begin lvButton.lvButtons_H cmdPayee 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   8520
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&Employees "
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
      Image           =   "DisbursementOld.frx":0051
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPayee 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   9015
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "Sales &Agent"
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
      Image           =   "DisbursementOld.frx":0D2B
      ImgSize         =   24
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDisburse 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   615
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&2] Service Center         "
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
      Image           =   "DisbursementOld.frx":117B
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDisburse 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1110
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&3] Rebuilt Factory         "
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
      Image           =   "DisbursementOld.frx":1715
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDisburse 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&1] Spare Parts              "
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
      Image           =   "DisbursementOld.frx":7F77
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPayee 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   9510
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&Customer   "
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
      Image           =   "DisbursementOld.frx":8511
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPayee 
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   10005
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&Others     "
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
      Image           =   "DisbursementOld.frx":8C0B
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDisburse 
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   3090
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&7] Cash Voucher        "
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
      Image           =   "DisbursementOld.frx":91A5
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDisburse 
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2595
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&6] Non-Operating Exp."
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
      Image           =   "DisbursementOld.frx":973F
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDisburse 
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1605
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&4] On-Acct. Charges     "
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
      Image           =   "DisbursementOld.frx":9CD9
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDisburse 
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2100
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&5] Personal Drawings  "
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
      Image           =   "DisbursementOld.frx":A273
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMain 
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   7755
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
      Image           =   "DisbursementOld.frx":A80D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMain 
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   7260
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
      Image           =   "DisbursementOld.frx":ADA7
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMain 
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6765
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "E&xit Disbursement"
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
      Image           =   "DisbursementOld.frx":B341
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMain 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5655
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&View Summary     "
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
      Image           =   "DisbursementOld.frx":B8DB
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMain 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6150
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&Print Voucher      "
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
      Image           =   "DisbursementOld.frx":BE75
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdMain 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "&Edit Voucher       "
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
      Image           =   "DisbursementOld.frx":C40F
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxCV 
      Height          =   2295
      Left            =   3000
      TabIndex        =   37
      Top             =   1560
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   $"DisbursementOld.frx":C9A9
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   3000
      TabIndex        =   38
      Top             =   4680
      Visible         =   0   'False
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Commissions"
      TabPicture(0)   =   "DisbursementOld.frx":CA40
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblHeading(9)"
      Tab(0).Control(1)=   "lblHeading(8)"
      Tab(0).Control(2)=   "flxComm"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Cash Advances"
      TabPicture(1)   =   "DisbursementOld.frx":CA5C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblHeading(11)"
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(2)=   "lblHeading(10)"
      Tab(1).Control(3)=   "lblDeduct"
      Tab(1).Control(4)=   "flxAdv"
      Tab(1).Control(5)=   "txtCAttl"
      Tab(1).Control(6)=   "txtCADue"
      Tab(1).Control(7)=   "cmdDeduct"
      Tab(1).Control(8)=   "fraCATerm"
      Tab(1).Control(9)=   "chkTerms"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Incentives"
      TabPicture(2)   =   "DisbursementOld.frx":CA78
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblHeading(13)"
      Tab(2).Control(1)=   "lblHeading(12)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Refunds"
      TabPicture(3)   =   "DisbursementOld.frx":CA94
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblHeading(15)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblHeading(14)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame1"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   120
         TabIndex        =   82
         Top             =   360
         Width           =   8175
         Begin VB.TextBox txtReason 
            Height          =   795
            Left            =   4080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   94
            Top             =   1080
            Width           =   3855
         End
         Begin VB.TextBox txtRefCust 
            Height          =   315
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   91
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox txtRefAmt 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1560
            TabIndex        =   90
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox txtRefDate 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   89
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtRefundRef 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1560
            TabIndex        =   86
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox cboRefundSrc 
            Height          =   315
            Left            =   1560
            TabIndex        =   84
            Top             =   240
            Width           =   2055
         End
         Begin lvButton.lvButtons_H cmdGoRefund 
            Height          =   495
            Left            =   7320
            TabIndex        =   95
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
            Image           =   "DisbursementOld.frx":CAB0
            cBack           =   -2147483633
         End
         Begin VB.Label lblAddRef 
            BackStyle       =   0  'Transparent
            Height          =   495
            Left            =   120
            TabIndex        =   96
            Top             =   1680
            Width           =   3615
         End
         Begin VB.Label Label8 
            Caption         =   "Reason for Refund:"
            Height          =   255
            Left            =   4080
            TabIndex        =   93
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Customer Name:"
            Height          =   255
            Left            =   4080
            TabIndex        =   92
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblRefAmt 
            Alignment       =   1  'Right Justify
            Caption         =   "Refund Amount:"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label lblRefDate 
            Alignment       =   1  'Right Justify
            Caption         =   "Ref. Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblRefNo 
            Alignment       =   1  'Right Justify
            Caption         =   "Reference No.:"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   585
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Select Source:"
            Height          =   255
            Left            =   360
            TabIndex        =   83
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkTerms 
         Caption         =   "Apply Terms"
         Height          =   195
         Left            =   -68640
         TabIndex        =   77
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame fraCATerm 
         Height          =   1935
         Left            =   -68640
         TabIndex        =   70
         Top             =   2760
         Width           =   1935
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            TabIndex        =   76
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   74
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Payment Mode"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Amount Due"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Payment Terms:"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   120
            Width           =   1215
         End
      End
      Begin lvButton.lvButtons_H cmdDeduct 
         Height          =   375
         Left            =   -68640
         TabIndex        =   67
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
         Image           =   "DisbursementOld.frx":D04A
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtCADue 
         Height          =   315
         Left            =   -68640
         TabIndex        =   65
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtCAttl 
         Height          =   315
         Left            =   -68640
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   840
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid flxComm 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   40
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
         FormatString    =   $"DisbursementOld.frx":D5E4
      End
      Begin MSFlexGridLib.MSFlexGrid flxAdv 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   41
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
      Begin VB.Label lblDeduct 
         Caption         =   "Total Amount Due"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -68640
         TabIndex        =   64
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
         Left            =   7320
         TabIndex        =   42
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
         Index           =   8
         Left            =   -67680
         TabIndex        =   53
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
         TabIndex        =   54
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
         Index           =   6
         Left            =   -65040
         TabIndex        =   52
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
         Index           =   4
         Left            =   -65040
         TabIndex        =   51
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
         Index           =   5
         Left            =   -65010
         TabIndex        =   50
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
         Index           =   7
         Left            =   -65010
         TabIndex        =   49
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
         TabIndex        =   48
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
         Index           =   11
         Left            =   -67650
         TabIndex        =   47
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
         Index           =   13
         Left            =   -67650
         TabIndex        =   45
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
         Index           =   15
         Left            =   7350
         TabIndex        =   43
         Top             =   30
         Width           =   2415
      End
   End
   Begin VB.TextBox txtCashAdvID 
      Height          =   285
      Left            =   6120
      TabIndex        =   66
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DISBURSEMENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   7680
      TabIndex        =   98
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "PREPARED BY:"
      Height          =   255
      Left            =   3000
      TabIndex        =   56
      Top             =   3855
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "APPROVED BY:"
      Height          =   255
      Left            =   3000
      TabIndex        =   55
      Top             =   4215
      Width           =   1335
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DISBURSEMENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   7710
      TabIndex        =   26
      Top             =   150
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1275
      Left            =   120
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   3735
      Left            =   120
      Top             =   240
      Width           =   2775
   End
   Begin VB.Menu mnuVoucher 
      Caption         =   "Voucher"
      Visible         =   0   'False
      Begin VB.Menu mnuAddRow 
         Caption         =   "Add Row                  Alt+A"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsMC 
         Caption         =   "Insert Service Unit"
      End
   End
   Begin VB.Menu mnuComm 
      Caption         =   "Commissions"
      Visible         =   0   'False
      Begin VB.Menu mnuMove 
         Caption         =   "Move to Particulars"
      End
   End
End
Attribute VB_Name = "DisbursementOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAcctDesc As New ADODB.Recordset
Dim rsPayee As New ADODB.Recordset
Dim rsService As New ADODB.Recordset
Dim rsOnAcct As New ADODB.Recordset

Dim DisburseType As Integer, PayeeType As Integer
Dim CustSrch As Boolean, CustBln As Boolean, AcctSrch As Boolean, GrpSrch As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim NewPCV As Boolean, EditBtnClicked As Boolean
Dim MyDisburseType, MyPayeeType
Dim RefundType As String, MyDept As String
Dim ActiveBtn As Integer
Dim CurrentBtn As Integer
Dim CmdTop(0 To 6) As Integer
Dim ActiveBtn2 As Integer
Dim CurrentBtn2 As Integer
Dim CmdTop2(0 To 6) As Integer

Public Sub DisplayVoucher()
    Dim rsDisb As New ADODB.Recordset
    Dim Qry As String, Chrg As String
    Dim i As Integer, Details As Variant
    
    If txtCVid.Text = "" Then Exit Sub
    '***Access
    'Qry = "SELECT DISBURSEMENT.* From DISBURSEMENT " & _
          "WHERE (((DISBURSEMENT.VOUCHERID)=" & txtCVid.Text & "));"
    'mySQL
    Qry = "SELECT DISBURSEMENT.* From DISBURSEMENT " & _
          "WHERE DISBURSEMENT.VOUCHERID=" & txtCVid.Text & ";"
    rsDisb.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsDisb.RecordCount = 0 Then Exit Sub
    rsDisb.MoveFirst
    MyDept = rsDisb!DEPARTMENT
    If rsDisb!DisburseType = "Operating Expenses" Then
        If rsDisb!DEPARTMENT = "SPARE PARTS" Then
            DisburseType = 0
        ElseIf rsDisb!DEPARTMENT = "SERVICE CENTER" Then
            DisburseType = 1
        Else
            DisburseType = 4    'PERSONAL DRAWINGS
        End If
    ElseIf rsDisb!DisburseType = "Non-Operating Expense" Then
        If rsDisb!DEPARTMENT = "REBUILT FACTORY" Then
            DisburseType = 2
        ElseIf rsDisb!DEPARTMENT = "SPARE PARTS" Then
            DisburseType = 5
        Else
            DisburseType = 3
        End If
    Else
        If rsDisb!DEPARTMENT = "SPARE PARTS" Then
            DisburseType = 0
        ElseIf rsDisb!DEPARTMENT = "SERVICE CENTER" Then
            DisburseType = 1
        ElseIf rsDisb!DEPARTMENT = "REBUILT FACTORY" Then
            DisburseType = 2
        Else
            DisburseType = 3
        End If
    End If
    If rsDisb!EmpId <> 0 Then
        PayeeType = 0
        txtPayeeID.Text = rsDisb!EmpId
    ElseIf rsDisb!CUSTOMERID <> 0 Then
        PayeeType = 2
        txtPayeeID.Text = rsDisb!CUSTOMERID
        txtCustID.Text = rsDisb!CUSTOMERID
    ElseIf rsDisb!AGENTID <> 0 Then
        PayeeType = 1
        txtPayeeID.Text = rsDisb!AGENTID
    Else
        PayeeType = 3
    End If
    CustSrch = True
    txtPayee.Text = rsDisb!payee
    cboAcctTitle.Text = rsDisb!AcctDesc
    txtPCV.Text = IIf(IsNull(rsDisb!CVNUMBER), "", Format(rsDisb!CVNUMBER, "0000"))
    txtDate.Text = rsDisb!CVDate
    txtApprovedBy.Text = IIf(IsNull(rsDisb!ApprovedBy), "", rsDisb!ApprovedBy)
    txtPreparedBy.Text = IIf(IsNull(rsDisb!PREPAREDBY), "", rsDisb!PREPAREDBY)
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
    txtPrnCopy.Text = IIf(IsNull(rsDisb!PRINTCOPY), "0", rsDisb!PRINTCOPY)
    If DisburseType = 0 And InStr(cboAcctTitle.Text, "REFUND") > 0 Then Chrg = IIf(IsNull(rsDisb!ChargeTo), "", rsDisb!ChargeTo)
    rsDisb.Close
    If DisburseType = 0 And InStr(cboAcctTitle.Text, "REFUND") > 0 And Chrg <> "" Then
        SSTab1.Tab = 3
        SSTab1.Visible = True
        If InStr(Chrg, "SR#") > 0 Then
            cboRefundSrc.Text = "SALES RETURN"
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
    If txtPrnCopy.Text = "0" Then
        Call DisplayMainButton(5)
    Else
        Call DisplayMainButton(4)
    End If
    Call LockData(True)
End Sub
Private Sub SaveVoucher()
    Dim rsDisburse As New ADODB.Recordset
    Dim rsCashAdv As New ADODB.Recordset
    Dim rsAgentOrder As New ADODB.Recordset
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String, nBal As Single, X As Integer, MyDP As Single, MyOR As String
    Dim MyDisburseType As String
    
    Select Case DisburseType
        Case 0  'Spare Parts
            MyDisburseType = "Operating Expenses"
            MyDept = "SPARE PARTS"
        Case 1  'Service Center
            MyDisburseType = "Operating Expenses"
            MyDept = "SERVICE CENTER"
        Case 2  'Rebuilt
            MyDisburseType = "Non-Operating Expense"
            MyDept = "REBUILT FACTORY"
            txtCustID.Text = 67 'Only in Moto-R Tagum
        Case 3  'On Acct
            MyDisburseType = "Non-Operating Expense"
            MyDept = cboLYRg.Text
        Case 4  'Personal Drawings
            MyDisburseType = "Operating Expenses"
            MyDept = "SPARE PARTS"
        Case 5  'Non operating
            MyDisburseType = "Non-Operating Expense"
            MyDept = "SPARE PARTS"
    End Select
    
    '***Access
    'If NewPCV Then
    '    rsDisburse.Open "Disbursement", LYRcon, adOpenKeyset, adLockOptimistic
    '    rsDisburse.AddNew
    'Else
    '    Qry = "SELECT DISBURSEMENT.* From DISBURSEMENT " & _
    '          "WHERE (((DISBURSEMENT.VOUCHERID)=" & txtCVid.Text & "));"
    '    rsDisburse.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    If rsDisburse.RecordCount > 0 Then rsDisburse.MoveFirst
    'End If
    'rsDisburse!CVDate = txtDate.Text
    'rsDisburse!Status = "Temp"
    'rsDisburse!CVNUMBER = IIf(IsNumeric(txtPCV.Text), Val(txtPCV.Text), 0)
    'rsDisburse!DisburseType = MyDisburseType '& IIf(DisburseType <> 5, "", " " & cboLYRg.Text)
    'rsDisburse!ACCTDESC = cboAcctTitle.Text
    'rsDisburse!DEPARTMENT = MyDept
    'rsDisburse!PAYEE = txtPayee.Text
    'rsDisburse!PARTICULARS = flxCV.TextMatrix(1, 1)
    'rsDisburse!AMOUNT = PureNum(flxCV.TextMatrix(1, 2))
    'rsDisburse!Details = CVDetails
    'rsDisburse!PREPAREDBY = txtPreparedBy.Text
    'rsDisburse!APPROVEDBY = txtApprovedBy.Text
    'Select Case PayeeType
    '    Case 0  'EMPLOYEE
    '        rsDisburse!EmpId = Val(txtPayeeID.Text)
    '        If DisburseType = 2 And txtMotorID.Text <> "" Then rsDisburse!CUSTOMERID = txtMotorID.Text
    '    Case 1  'SALES AGENT
    '        rsDisburse!AgentId = Val(txtPayeeID.Text)
    'End Select
    'If DisburseType = 3 Then    'On Acct
    '    rsDisburse!ChargeTo = cboLYRg.Text
    '    rsDisburse!CUSTOMERID = Val(txtCustID.Text)
    'ElseIf DisburseType = 2 Then    'Rebuilt Factory
    '    rsDisburse!ChargeTo = MyDept
    '    rsDisburse!CUSTOMERID = Val(txtCustID.Text)
    'End If
    'If DisburseType = 0 And InStr(cboAcctTitle.Text, "REFUND") > 0 Then
    '    If cboRefundSrc.Text = "SALES RETURN" Then
    '        rsDisburse!ChargeTo = "SR#" & txtRefundRef.Text
    '    ElseIf cboRefundSrc.Text = "O.R." Then
    '        rsDisburse!ChargeTo = "OR#" & txtRefundRef.Text
    '    End If
    'End If
    'rsDisburse!UPDATEDBY = UserID
    'rsDisburse!LASTUPDATE = Date
    'rsDisburse.Update
    'If NewPCV Then txtCVid.Text = rsDisburse!VOUCHERID
    'rsDisburse.Close
    
    'mySQL
    If NewPCV Then
'        LYRcon.Execute "INSERT INTO disbursement(CVDATE, `STATUS`, CVNUMBER, DISBURSETYPE, ACCTDESC, " & _
'            "DEPARTMENT, PAYEE, PARTICULARS, AMOUNT, DETAILS, PREPAREDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE)" & _
'            "VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "', 'Temp', " & IIf(IsNumeric(txtPCV.Text), Val(txtPCV.Text), 0) & ", " & _
'            "'" & MyDisburseType & "', '" & cboAcctTitle.Text & "', '" & MyDept & "', '" & txtPayee.Text & "', " & _
'            "'" & flxCV.TextMatrix(1, 1) & "', '" & Val(Replace(flxCV.TextMatrix(1, 2), ",", "")) & "', '" & CVDetails & "', " & _
'            "'" & txtPreparedBy.Text & "', '" & txtApprovedBy.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
'
'     LYRcon.Execute "INSERT INTO disbursement(CVDATE, `STATUS`, CVNUMBER, DISBURSETYPE, ACCTDESC, " & _
'            "DEPARTMENT, PAYEE, PARTICULARS, AMOUNT, DETAILS, PREPAREDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE)" & _
'            "VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "', 'Temp', " & IIf(IsNumeric(txtPCV.Text), Val(txtPCV.Text), 0) & ", " & _
'            "'" & MyDisburseType & "', '" & cboAcctTitle.Text & "', '" & MyDept & "', '" & txtPayee.Text & "', " & _
'            "'" & flxCV.TextMatrix(1, 1) & "', '" & Val(Replace(flxCV.TextMatrix(1, 2), ",", "")) & "', '" & CVDetails & "', " & _
'            "'" & txtPreparedBy.Text & "', '" & txtApprovedBy.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
'
    LYRcon.Execute "INSERT INTO disbursement(CVDATE, `STATUS`, CVNUMBER, DISBURSETYPE, ACCTDESC, " & _
            "DEPARTMENT, PAYEE, PARTICULARS, AMOUNT, DETAILS, PREPAREDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE)" & _
            "VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "', 'Temp', '" & txtPCV.Text & "', " & _
            "'" & MyDisburseType & "', '" & cboAcctTitle.Text & "', '" & MyDept & "', '" & txtPayee.Text & "', " & _
            "'" & flxCV.TextMatrix(1, 1) & "', '" & Val(Replace(flxCV.TextMatrix(1, 2), ",", "")) & "', '" & CVDetails & "', " & _
            "'" & txtPreparedBy.Text & "', '" & txtApprovedBy.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
    Else
'        LYRcon.Execute "UPDATE Disbursement SET CVDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "', `STATUS` = 'Temp', " & _
'            "CVNUMBER = " & IIf(IsNumeric(txtPCV.Text), Val(txtPCV.Text), 0) & ", DISBURSETYPE = '" & MyDisburseType & "', " & _
'            "ACCTDESC = '" & cboAcctTitle.Text & "', DEPARTMENT = '" & MyDept & "', PAYEE = '" & txtPayee.Text & "', " & _
'            "PARTICULARS = '" & flxCV.TextMatrix(1, 1) & "', AMOUNT = '" & Val(Replace(flxCV.TextMatrix(1, 2), ",", "")) & "', " & _
'            "DETAILS = '" & CVDetails & "', PREPAREDBY = '" & txtPreparedBy.Text & "', APPROVEDBY = '" & txtApprovedBy.Text & "', " & _
'            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
'            "Where Disbursement.VOUCHERID = " & txtCVid.Text & ";"
         LYRcon.Execute "UPDATE Disbursement SET CVDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "', `STATUS` = 'Temp', " & _
            "CVNUMBER = '" & txtPCV.Text & "', DISBURSETYPE = '" & MyDisburseType & "', " & _
            "ACCTDESC = '" & cboAcctTitle.Text & "', DEPARTMENT = '" & MyDept & "', PAYEE = '" & txtPayee.Text & "', " & _
            "PARTICULARS = '" & flxCV.TextMatrix(1, 1) & "', AMOUNT = '" & Val(Replace(flxCV.TextMatrix(1, 2), ",", "")) & "', " & _
            "DETAILS = '" & CVDetails & "', PREPAREDBY = '" & txtPreparedBy.Text & "', APPROVEDBY = '" & txtApprovedBy.Text & "', " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
            "Where Disbursement.VOUCHERID = " & txtCVid.Text & ";"
    
    End If
    
    Qry = "SELECT Disbursement.VOUCHERID From Disbursement Where DISBURSEMENT.CVNUMBER = " & IIf(IsNumeric(txtPCV.Text), Val(txtPCV.Text), 0) & ";"
    If NewPCV Then txtCVid.Text = GetFieldValue(Qry)

    Select Case PayeeType
        Case 0  'EMPLOYEE
            LYRcon.Execute "Update Disbursement SET EmpId = " & Val(txtPayeeID.Text) & " " & _
                " Where Disbursement.VOUCHERID = " & txtCVid.Text & ";"
            If DisburseType = 2 And txtMotorID.Text <> "" Then
                LYRcon.Execute "UPDATE Disbursement SET CUSTOMERID = " & txtMotorID.Text & " " & _
                    "Where Disbursement.VOUCHERID = " & txtCVid.Text & ";"
            End If
        Case 1  'SALES AGENT
            LYRcon.Execute "UPDATE Disbursement SET AGENTID = " & Val(txtPayeeID.Text) & " " & _
                "Where Disbursement.VOUCHERID = " & txtCVid.Text & ";"
    End Select
    
    If DisburseType = 3 Then    'On Acct
        LYRcon.Execute "UPDATE Disbursement SET CHARGETO = '" & cboLYRg.Text & "', CUSTOMERID = " & Val(txtCustID.Text) & " " & _
            "Where Disbursement.VOUCHERID = " & txtCVid.Text & ";"
    ElseIf DisburseType = 2 Then    'Rebuilt Factory
        LYRcon.Execute "UPDATE Disbursement SET CHARGETO = '" & MyDept & "', CUSTOMERID = " & Val(txtCustID.Text) & " " & _
            "Where Disbursement.VOUCHERID = " & txtCVid.Text & ";"
    End If
    If DisburseType = 0 And InStr(cboAcctTitle.Text, "REFUND") > 0 Then
        If cboRefundSrc.Text = "SALES RETURN" Then
            LYRcon.Execute "Update Disbursement SET CHARGETO = 'SR#" & txtRefundRef.Text & "' " & _
                "Where Disbursement.VOUCHERID = " & txtCVid.Text & ";"
        ElseIf cboRefundSrc.Text = "O.R." Then
            'rsDisburse!ChargeTo = "OR#" & txtRefundRef.Text
               LYRcon.Execute "Update Disbursement SET CHARGETO = 'OR#" & txtRefundRef.Text & "' " & _
                "Where Disbursement.VOUCHERID = " & txtCVid.Text & ";"
        End If
    End If

'old code Save
'    If DisburseType = 0 And InStr(cboAcctTitle.Text, "REFUND") > 0 And InStr(cboAcctTitle.Text, "SALES RETURN") > 0 Then
'        If Val(txtSalesRetID.Text) > 0 Then
'            LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.PROREF = '" & txtPCV.Text & "' WHERE (((SRETURNSP.RETURNID)=" & txtSalesRetID.Text & "));"
'        End If
'    End If
    
End Sub

Private Sub SaveCashAdvance()
    Dim rsCashAdv As New ADODB.Recordset
    Dim rsCALedger As New ADODB.Recordset
    Dim Qry As String, CAdvID As Long
    Dim myCA As Single, MyBal As Single
    
    If txtPayeeID.Text = "" Then
        MsgBox "No record of Payee!", vbCritical + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    'If Cash Advance exist then update only else add new
    
    If PayeeType = 0 Then   'Employee           1 record and the same remarks (charge to)
        Qry = "SELECT CASHADVANCE.* From CASHADVANCE WHERE (((CASHADVANCE.EMPID)=" & Val(txtPayeeID.Text) & ") AND ((CASHADVANCE.REMARKS)='" & flxCV.TextMatrix(1, 1) & "'));"
    ElseIf PayeeType = 1 Then 'Agent
        Qry = "SELECT CASHADVANCE.* From CASHADVANCE WHERE (((CASHADVANCE.AGENTID)=" & Val(txtPayeeID.Text) & "));"
    End If
    
    
    myCA = PureNum(flxCV.TextMatrix(1, 2))
    If Not CtrlNumExist(Qry) Then
        rsCashAdv.Open "CashAdvance", LYRcon, adOpenKeyset, adLockOptimistic
        rsCashAdv.AddNew
        If PayeeType = 1 Then rsCashAdv!AGENTID = txtPayeeID.Text
        If PayeeType = 0 Then rsCashAdv!EmpId = txtPayeeID.Text
        rsCashAdv!CAamount = myCA
        rsCashAdv!Balance = myCA
        MyBal = myCA
    Else
        'Qry = "SELECT CASHADVANCE.* From CASHADVANCE WHERE (((CASHADVANCE.AGENTID)=" & txtPayeeID.Text & "));"
        rsCashAdv.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        rsCashAdv.MoveFirst
        If rsCashAdv!Balance = 0 Then
            rsCashAdv!CAamount = myCA
            rsCashAdv!AmountPaid = 0
            rsCashAdv!Balance = myCA
            MyBal = myCA
        Else
            MyBal = rsCashAdv!Balance + myCA
            rsCashAdv!Balance = MyBal
            rsCashAdv!CAamount = rsCashAdv!CAamount + myCA
        End If
    End If
    rsCashAdv!Remarks = flxCV.TextMatrix(1, 1)
    rsCashAdv!UPDATEDBY = UserID
    rsCashAdv!LASTUPDATE = Date
    rsCashAdv.Update
    CAdvID = rsCashAdv!TransNo
    rsCashAdv.Close
    
    'mySQL
    'If Not CtrlNumExist(Qry) Then
    '    If PayeeType = 1 Then
            'Insert payee type Agent
    '        LYRcon.Execute "INSERT INTO CashAdvance(AGENTID)VALUES(" & txtPayeeID.Text & ");"
    '    End If
    '    If PayeeType = 0 Then
            'Insert payee type employee
    '        LYRcon.Execute "INSERT INTO CashAdvance(EMPID)VALUES(" & txtPayeeID.Text & ");"
    '    End If
       'Insert CA of Agent and Emplyee
    '   LYRcon.Execute ""
    'Else
        'Update CA here
    '    If PayeeType = 1 Then
            'Insert payee type Agent
            
    '    End If
    '    If PayeeType = 0 Then
            'Insert payee type employee
        
    '    End If
    'End If
    'Add new Cash Advance Ledger
    
    With rsCALedger
        .Open "CALedger", LYRcon, adOpenKeyset, adLockOptimistic
        .AddNew
        !CASHADVID = CAdvID
        !TRANSDATE = txtDate.Text
        !RefNo = txtPCV.Text
        !CAamount = myCA
        !CurrentBal = MyBal 'Balance
        !UPDATEDBY = UserID
        !LASTUPDATE = Now
        !STATUS = "ACTIVE"
        .Update
        .Close
    End With
    
End Sub
                                  

Function CVDetails() As String
    Dim Tmp As String
    Dim i As Integer
    
    Tmp = ""
'    If flxCV.Rows <= 2 Then
'        If DisburseType = 1 Then    'Cash Advance
'            CVDetails = Right$(cboAcctTitle.Text, Len(cboAcctTitle.Text) - (InStr(cboAcctTitle.Text, "[") - 1))
'        Else
'            CVDetails = ""
'        End If
'    Else
        For i = 2 To flxCV.Rows - 1
            Tmp = IIf(Tmp = "", flxCV.TextMatrix(i, 1), Tmp & "|" & flxCV.TextMatrix(i, 1))
        Next i
'        If DisburseType = 1 Then    'Cash Advance
'            Tmp = Right$(cboAcctTitle.Text, Len(cboAcctTitle.Text) - (InStr(cboAcctTitle.Text, "[") - 1)) & "|" & Tmp
'        End If
        CVDetails = Tmp
'    End If
End Function
Function ValidatePCV() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer
    
    vMsg = ""
    If txtPayee.Text = "" Then vMsg = "Incomplete entry on field Payee."
    If cboAcctTitle.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Account Title"
    'old code If PayeeType < 3 And txtPayeeID.Text = "" Then VMsg = VMsg & vbCrLf & "Please select a Payee from list."
    If txtPCV.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field PCV Number"
    Else
        If NewPCV Then
            '***Access
            'Qry = "SELECT DISBURSEMENT.CVNUMBER From DISBURSEMENT WHERE (((DISBURSEMENT.CVNUMBER)=" & txtPCV.Text & "));"
            'mySQL
            Qry = "SELECT DISBURSEMENT.CVNUMBER From DISBURSEMENT WHERE DISBURSEMENT.CVNUMBER='" & txtPCV.Text & "';"
            If CtrlNumExist(Qry) Then vMsg = vMsg & vbCrLf & "PCV number already exist"
        End If
    End If
    If PCVexist Then vMsg = vMsg & vbCrLf & "PCV Number already Exist"
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
    If DisburseType = 0 And PayeeType = 1 And txtPayeeID.Text = "" Then vMsg = vMsg & vbCrLf & "No record of Sales Agent" 'Operating Exp / Sales Agent
    If DisburseType = 3 And Val(txtCustID.Text) = 0 Then vMsg = vMsg & vbCrLf & "Invalid Customer Name! Please select from list"
    ValidatePCV = vMsg
End Function


Private Sub ClearEntries()
    txtPayee.Text = ""
    txtPayeeID.Text = ""
    cboAcctTitle.Text = ""
    txtPCV.Text = ""
    txtDate.Text = ""
    txtPreparedBy.Text = ""
    txtApprovedBy.Text = ""
    flxCV.Rows = 1
    flxCV.Rows = 2
    flxPayee.Visible = False
    txtMotorID.Text = ""
    txtPrnCopy.Text = ""
    imgPic.Picture = LoadPicture("")
    lblName.Caption = ""
    lblAddress.Caption = ""
    lblPosition.Caption = ""
    lblStatus.Caption = ""
    'Clear Flx
    flxComm.Rows = 1
    flxAdv.Rows = 1
    'flxRefund.Rows = 1
    flxComm.Rows = 2
    flxAdv.Rows = 2
    'flxRefund.Rows = 2
    txtCAttl.Text = ""
    txtCADue.Text = ""
    txtCashAdvID.Text = ""
    txtCAlessCOM.Text = ""
    txtCommOrigBal.Text = ""
    txtSalesRetID.Text = ""
    
    txtRefundRef.Text = ""
    txtRefundRef.Tag = ""
    txtRefDate.Text = ""
    txtRefCust.Text = ""
    txtRefAmt.Text = ""
    txtORamt.Text = ""
    txtReason.Text = ""
    lblAddRef.Caption = ""
    cboRefundSrc.Text = ""
    Frame1.Enabled = True
    
End Sub
Private Sub DisplayCashAdvance(MyID As Integer)
    Dim rsAdvance As New ADODB.Recordset
    Dim rsCAcomm As New ADODB.Recordset
    Dim Qry As String
    Dim Total As Single, i As Integer
    
    If PayeeType = 1 Then 'Sales agent
        Qry = "SELECT CASHADVANCE.CAAMOUNT, CASHADVANCE.AMOUNTPAID, CASHADVANCE.BALANCE, CASHADVANCE.REMARKS, CASHADVANCE.TRANSNO " & _
              "From CASHADVANCE WHERE (((CASHADVANCE.BALANCE)>0) AND ((CASHADVANCE.AGENTID)=" & MyID & "));"
        'mySQL Qry = "SELECT CASHADVANCE.CAAMOUNT, CASHADVANCE.AMOUNTPAID, CASHADVANCE.BALANCE, CASHADVANCE.REMARKS, CASHADVANCE.TRANSNO " & _
              "From CASHADVANCE WHERE CASHADVANCE.BALANCE>0 AND CASHADVANCE.AGENTID=" & MyID & ";"
        
    ElseIf PayeeType = 0 Then 'Employee
        Qry = "SELECT CASHADVANCE.CAAMOUNT, CASHADVANCE.AMOUNTPAID, CASHADVANCE.BALANCE, CASHADVANCE.REMARKS, CASHADVANCE.TRANSNO " & _
              "From CASHADVANCE WHERE (((CASHADVANCE.BALANCE)>0) AND ((CASHADVANCE.EMPID)=" & MyID & "));"
        'mySQL Qry = "SELECT CASHADVANCE.CAAMOUNT, CASHADVANCE.AMOUNTPAID, CASHADVANCE.BALANCE, CASHADVANCE.REMARKS, CASHADVANCE.TRANSNO " & _
              "From CASHADVANCE WHERE CASHADVANCE.BALANCE>0 AND CASHADVANCE.EMPID=" & MyID & ";"
 
'        Qry = "SELECT CASHADVANCE.CADATE, CASHADVANCE.CVNUMBER, CASHADVANCE.CAAMOUNT, CASHADVANCE.BALANCE, CASHADVANCE.TERMS, CASHADVANCE.PAYMENTMODE, CASHADVANCE.AMOUNTDUE, CASHADVANCE.REMARKS, CASHADVANCE.CASHADVID " & _
              "From CASHADVANCE Where (((CASHADVANCE.BALANCE) > 0) And ((CASHADVANCE.EMPID) = " & myID & ")) ORDER BY CASHADVANCE.CADATE;"
    End If
    With rsAdvance
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Total = 0
            .MoveFirst
            flxAdv.Rows = 1
            Do While Not .EOF
                flxAdv.AddItem vbTab & Format$(.Fields(0), "###,##0.00") & vbTab & Format$(.Fields(1), "###,##0.00") & vbTab & Format$(.Fields(2), "###,##0.00") & vbTab & IIf(IsNull(.Fields(3)), "", .Fields(3)) & vbTab & vbTab & .Fields(4)
                Total = Total + .Fields(2)
                txtCashAdvID.Text = !TransNo
                If CtrlNumExist("SELECT CALEDGER.* From CALEDGER WHERE (((CALEDGER.CASHADVID)=" & rsAdvance!TransNo & "));") Then flxAdv.TextMatrix(flxAdv.Rows - 1, 0) = "+"
                'mysql If CtrlNumExist("SELECT CALEDGER.* From CALEDGER WHERE CALEDGER.CASHADVID=" & rsAdvance!Transno & ";" Then flxAdv.TextMatrix(flxAdv.Rows - 1, 0) = "+"
                .MoveNext
            Loop
            flxAdv.ColWidth(6) = 0
            If PayeeType = 2 Then SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.Tab = 1
            SSTab1.Visible = True
        Else
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.Tab = 1
            flxAdv.Rows = 1
            SSTab1.Visible = False
        End If
        .Close
    End With
    If EditBtnClicked Then
        txtCAttl.Text = ""
        txtCADue.Text = ""
    Else
        txtCAttl.Text = Format$(Total, "P###,##0.00")
        txtCADue.Text = Format$(Total, "P###,##0.00")
        'If flxAdv.Rows > 1 Then txtCADue.Text = flxAdv.TextMatrix(1, 7)
    End If
    
    If DisburseType = 0 And PayeeType = 1 And txtCVid.Text <> "" Then
        Qry = "SELECT CACOMMTEMP.* From CACOMMTEMP WHERE (((CACOMMTEMP.VOUCHERID)=" & txtCVid.Text & "));"
        'mySQL Qry = "SELECT CACOMMTEMP.* From CACOMMTEMP WHERE CACOMMTEMP.VOUCHERID=" & txtCVid.Text & ";"
        With rsCAcomm
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                txtCommOrderID.Text = !CommOrderID
                txtCommOrigBal.Text = !CommOrigBal
                txtCAlessCOM.Text = !CAlessComm
            End If
            .Close
        End With
        For i = flxCV.Rows - 1 To 1 Step -1
            If InStr(flxCV.TextMatrix(i, 1), "LESS: Cash Advance") > 0 Then
                flxCV.TextMatrix(i, 3) = "CA DEDUCTION"
                Exit For
            End If
        Next i
    End If
    
End Sub

Private Sub DisplayAgentComm(MyAgent As Integer)
    Dim rsComm As New ADODB.Recordset
    Dim Qry As String

    If DisburseType = 1 Then 'Cash Advance show approved commissions only
        Qry = "SELECT AGENTORDERS.SALESDATE, [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS CUSTOMER, UNITORDERED.MODEL, AGENTORDERS.COD, AGENTORDERS.LCP, AGENTORDERS.COMMISSION, AGENTORDERS.BALANCE, AGENTORDERS.REMARKS, AGENTORDERS.STATUS, UNITORDERED.ACCOUNTNO, AGENTORDERS.PCTCOMM, AGENTORDERS.ORDERID " & _
              "FROM (AGENTORDERS INNER JOIN UNITORDERED ON AGENTORDERS.ORDERID = UNITORDERED.ORDERID) INNER JOIN CUSTOMER ON UNITORDERED.CUSTOMERID = CUSTOMER.CUSTOMERID " & _
              "WHERE (((AGENTORDERS.BALANCE)>0) AND ((AGENTORDERS.STATUS)='Approved') AND ((AGENTORDERS.AGENTID)=" & MyAgent & ")) " & _
              "ORDER BY AGENTORDERS.SALESDATE DESC;"
      
    ElseIf DisburseType = 0 Then
        Qry = "SELECT AGENTORDERS.SALESDATE, [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS CUSTOMER, UNITORDERED.MODEL, AGENTORDERS.COD, AGENTORDERS.LCP, AGENTORDERS.COMMISSION, AGENTORDERS.BALANCE, AGENTORDERS.REMARKS, AGENTORDERS.STATUS, UNITORDERED.ACCOUNTNO, AGENTORDERS.PCTCOMM, AGENTORDERS.ORDERID " & _
              "FROM (AGENTORDERS INNER JOIN UNITORDERED ON AGENTORDERS.ORDERID = UNITORDERED.ORDERID) INNER JOIN CUSTOMER ON UNITORDERED.CUSTOMERID = CUSTOMER.CUSTOMERID " & _
              "WHERE (((AGENTORDERS.BALANCE)>0) AND ((AGENTORDERS.STATUS)='Approved' Or (AGENTORDERS.STATUS)='Pending') AND ((AGENTORDERS.AGENTID)=" & MyAgent & ")) " & _
              "ORDER BY AGENTORDERS.SALESDATE DESC;"
    End If
    If Qry = "" Then Exit Sub
    With rsComm
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxComm.Rows = 1
            Do While Not .EOF
                If Not IsNull(!ACCOUNTNO) Then
                    If InStr(flxCV.TextMatrix(1, 1), !ACCOUNTNO) <> 0 Then txtCommOrderID.Text = !OrderID
                End If
                flxComm.AddItem vbTab & .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2) & vbTab & IIf(IsNull(.Fields(3)), "", Format$(.Fields(3), "P###,##0.00")) & vbTab & IIf(IsNull(.Fields(4)), "", Format$(.Fields(4), "P###,##0.00")) & vbTab & Format$(Int(.Fields(5)), "P###,##0.00") & vbTab & Format$(Int(.Fields(6)), "P###,##0.00") & vbTab & IIf(IsNull(.Fields(7)), "", .Fields(7)) & _
                vbTab & .Fields(8) & vbTab & IIf(IsNull(.Fields(9)), "", .Fields(9)) & vbTab & Format$(.Fields(10) / 100, "#.0%") & vbTab & .Fields(11)
                .MoveNext
            Loop
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.Tab = 0
            SSTab1.Visible = True
        Else
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.Tab = 0
            flxComm.Rows = 1
            SSTab1.Visible = True
        End If
        .Close
    End With
    flxComm.ColWidth(12) = 10
    
End Sub


Public Sub ShowPic()
    On Error Resume Next
    
    Open "pictemp" For Binary Access Write As lngDataFile
        lngLengh = rsPayee.Fields(4).ActualSize
        intChunks = lngLengh \ ChunkSize
        intFragment = lngLengh Mod ChunkSize
        ReDim Chunk(intFragment)
        Chunk() = rsPayee.Fields(4).GetChunk(intFragment)
        Put lngDataFile, , Chunk()
        For i = 1 To intChunks
           ReDim Buffer(ChunkSize)
            Chunk() = rsPayee.Fields(4).GetChunk(ChunkSize)
            Put lngDataFile, , Chunk()
        Next i
    Close lngDataFile
    
    filename = "pictemp"
    imgPic.Picture = LoadPicture(filename)
End Sub

Private Sub LockData(bln As Boolean)
    FraHeading.Enabled = Not bln
    flxCV.Enabled = Not bln
    txtPreparedBy.Enabled = Not bln
    txtApprovedBy.Enabled = Not bln
    SSTab1.Enabled = Not bln
End Sub

Private Sub cboAcctTitle_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If AcctSrch Then
        AcctSrch = False
        Exit Sub
    End If
    With cboAcctTitle
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsAcctDesc.MoveFirst
            rsAcctDesc.Move i
            'If .Text = "COMMISSION" And DisburseType = 0 And PayeeType = 1 Then
            '    SSTab1.Tab = 0
            'End If
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

Private Sub cboAcctTitle_Click()
    Call cboAcctTitle_Change
End Sub

Private Sub cboAcctTitle_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cboAcctTitle.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboAcctTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then AcctSrch = True
End Sub

Private Sub cboAcctTitle_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPCV.SetFocus
End Sub

Private Sub cboAcctTitle_LostFocus()
    If DisburseType = 0 And InStr(cboAcctTitle.Text, "REFUND") > 0 Then
        SSTab1.Tab = 3
        SSTab1.TabEnabled(3) = True
        SSTab1.Visible = True
    End If
End Sub


Private Sub cboLYRg_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If GrpSrch Then
        GrpSrch = False
        Exit Sub
    End If
    With cboLYRg
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsOnAcct.MoveFirst
            rsOnAcct.Move i
            txtCustID.Text = rsOnAcct!CUSTOMERID
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                GrpSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustID.Text = ""
        End If
    End With
End Sub

Private Sub cboLYRg_Click()
    Call cboLYRg_Change
End Sub

Private Sub cboLYRg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then GrpSrch = True
End Sub

Private Sub cboLYRg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cboLYRg_LostFocus
End Sub

Private Sub cboLYRg_LostFocus()
    If Val(txtCustID.Text) > 0 Then
        Shape1.Height = 1575
        Call DisplayMainButton(3)
        txtPayee.SetFocus
    End If
End Sub

Private Sub cboRefundSrc_Click()
    If cboRefundSrc.Text = "O.R." Then
        lblRefNo.Caption = "O.R. No:"
        lblRefDate.Caption = "O.R. Date:"
        lblRefAmt.Caption = "O.R. Amount:"
    ElseIf cboRefundSrc.Text = "SALES RETURN" Then
        lblRefNo.Caption = "Sales Ret#:"
        lblRefDate.Caption = "Ret. Date:"
        lblRefAmt.Caption = "Ret. Amount:"
    End If
End Sub

Private Sub cboServiceUnit_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If SrvcSrch Then
        SrvcSrch = False
        Exit Sub
    End If
    With cboServiceUnit
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsService.MoveFirst
            rsService.Move i
            txtMotorID.Text = rsService!MotorID
            
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SrvcSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtMotorID.Text = ""
        End If
    End With
End Sub

Private Sub cboServiceUnit_Click()
    Call cboServiceUnit_Change
End Sub

Private Sub cboServiceUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SrvcSrch = True
End Sub

Private Sub cboServiceUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cboServiceUnit_LostFocus
        flxCV.SetFocus
        flxCV.Col = 2
    End If
End Sub

Private Sub cboServiceUnit_LostFocus()
    flxCV.TextMatrix(ActRow, ActCol) = cboServiceUnit.Text
    cboServiceUnit.Visible = False
    If txtMotorID.Text <> "" Then
        flxCV.TextMatrix(ActRow, 3) = "MC"
    Else
        flxCV.TextMatrix(ActRow, 3) = ""
    End If
End Sub

Private Sub cmdAltA_Click()
    If flxCV.TextMatrix(flxCV.Rows - 1, 1) <> "" Then
        flxCV.AddItem ""
        flxCV.SetFocus
        flxCV.Row = flxCV.Rows - 1
        flxCV.Col = 1
    Else
        MsgBox "Incomplete entry on last row!", vbExclamation + vbOKOnly, "Validation Error"
    End If
End Sub

Private Sub cmdDeduct_Click()
    Dim vMsg As String
    Dim myTtlCom As Single, myCAdue As Single, myTtlCA As Single
    Dim i As Integer
    
    myCAdue = PureNum(txtCADue.Text)
    myTtlCom = PureNum(txtCommOrigBal.Text)
    myTtlCA = PureNum(txtCAttl.Text)
    vMsg = "" 'Validate
    If myCAdue > myTtlCA Then vMsg = "Total Amount Due must not be greater than Total Cash Advance."
    If myCAdue = 0 Then vMsg = vMsg & vbCrLf & "Please specify C.A. amount to be deducted in commission."
    If PayeeType = 1 And InStr(UCase$(cboAcctTitle.Text), "COMMISSION") > 0 Then
        If myCAdue > myTtlCom Then vMsg = vMsg & vbCrLf & "Total Cash Advance Due must not be greater than Commission."
    End If
    If vMsg <> "" Then
        MsgBox vMsg, vbExclamation + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    For i = flxCV.Rows - 1 To 1 Step -1
        If flxCV.TextMatrix(i, 3) = "CA DEDUCTION" Then flxCV.RemoveItem i
    Next i
    If PayeeType = 1 And InStr(UCase$(cboAcctTitle.Text), "COMMISSION") > 0 Then
        flxCV.AddItem vbTab & "LESS: Cash Advance = " & Format$(myCAdue, "###,##0.00") & vbTab & vbTab & "CA DEDUCTION"
        flxCV.TextMatrix(1, 2) = Format$(myTtlCom - myCAdue, "###,##0.00")
        txtCAlessCOM.Text = myCAdue
    End If
'    ElseIf PayeeType = 1 And InStr(UCase$(cboAcctTitle.Text), "INCENTIVES") > 0 And txtCashAdvID.Text <> "" Then
    
End Sub

Private Sub DisplayBtnDisburse(DispCode As Integer)
    Dim i As Integer
    
    If DispCode = 10 Then
        For i = 0 To 6
            cmdDisburse(i).Top = 360 + IIf(i > 0, cmdDisburse(i).Height * i, 0)
            cmdDisburse(i).Visible = True
            cmdDisburse(i).Enabled = True
        Next i
        Shape1.Height = 3735
    Else
        For i = 0 To 6
            cmdDisburse(i).Visible = False
        Next i
        cmdDisburse(DispCode).Top = 360
        cmdDisburse(DispCode).Visible = True
        cmdDisburse(DispCode).Enabled = False
    End If
End Sub
Private Sub cmdDisburse_Click(Index As Integer)
    Dim Qry As String, MyAcc As String
        
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=57));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=57;"

    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
        
    DisburseType = Index
    MyDept = ""
    If EditBtnClicked Then
        NewPCV = False
    Else
        NewPCV = True
        Call ClearEntries
    End If
    '    ActiveBtn = Index
    '    CurrentBtn = 6
    '    If Index = 6 Then
    '        CurrentBtn = 5
    '        cmdDisburse(6).ZOrder
    '        MoveUp.Enabled = True
    '    Else
    '        HideBtn.Enabled = True
    '    End If
    'ElseIf Index <> 6 And DispBtn Then
    If Index <> 6 And Not DispBtn Then
        Call DisplayBtnDisburse(Index)
        Call DisplayButtonPayee
    End If
    Select Case Index
        Case Is <= 3, 5 'Operating Expense, LYRGCOM OP.EXP., NON-OP. EXP.
            Qry = "SELECT DISBURSEACCTDESC.ACCOUNTDESCRIPTION From DISBURSEACCTDESC " & _
                  "ORDER BY DISBURSEACCTDESC.ACCOUNTDESCRIPTION;"
            If rsAcctDesc.State = adStateOpen Then rsAcctDesc.Close
            With rsAcctDesc
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    If Not EditBtnClicked Then cboAcctTitle.Clear
                    Do While Not .EOF
                        cboAcctTitle.AddItem .Fields(0)
                        .MoveNext
                    Loop
                End If
            End With
            lblHeading(2).Caption = Right$(cmdDisburse(Index).Caption, Len(cmdDisburse(Index).Caption) - 4)
            lblHeading(3).Caption = Right$(cmdDisburse(Index).Caption, Len(cmdDisburse(Index).Caption) - 4)

        Case 4  'Personal Drawings
            'myACCESS
            'Qry = "SELECT UCase$([VALUE]) AS PAYEE, SETTINGS.DESCRIPTION, SETTINGS.REMARKS From SETTINGS " & _
                  "WHERE (((SETTINGS.DESCRIPTION)='DISBURSEMENT') AND ((SETTINGS.REMARKS)='Drawings'));"
            'mySQL
            Qry = "SELECT UCase(SETTINGS.VALUE) AS PAYEE, SETTINGS.DESCRIPTION, SETTINGS.REMARKS " & _
                  "From SETTINGS Where SETTINGS.DESCRIPTION AND 'DISBURSEMENT' AND SETTINGS.REMARKS AND 'Drawings';"
            
            If rsPayee.State = adStateOpen Then rsPayee.Close
            rsPayee.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            Set flxPayee.DataSource = rsPayee
            flxPayee.Refresh
            flxPayee.ColWidth(0) = 0
            flxPayee.ColWidth(1) = flxPayee.Width
            PayeeType = 4
        Case 6  'LYR Cash out & Transfer of Funds
            CashVoucher.CVtype = "EXPENSE"
            CashVoucher.Show vbModal
        Case Else
        
    End Select
    If EditBtnClicked Then Exit Sub
    txtPreparedBy.Text = GetSettings("DISBURSER")
    txtApprovedBy.Text = GetSettings("MANAGER")
    txtDate.Text = Date
End Sub

Private Sub DisplayMainButton(DispCode As Integer)
    Dim i As Integer
    
    For i = 0 To 5
        cmdMain(i).Visible = False
    Next i
    Shape2.Left = Shape1.Left
    Shape2.Top = Shape1.Top + Shape1.Height + 105
    Select Case DispCode
        Case 1  'View & Exit
            Shape2.Height = cmdMain(0).Height * 2 + 210
            cmdMain(1).Top = Shape2.Top + 105
            cmdMain(1).Left = Shape2.Left + 120
            cmdMain(1).Visible = True
            cmdMain(3).Top = cmdMain(1).Top + cmdMain(1).Height
            cmdMain(3).Left = Shape2.Left + 120
            cmdMain(3).Visible = True
        Case 2  'Cancel
            Shape2.Height = cmdMain(0).Height + 210
            cmdMain(5).Top = Shape2.Top + 105
            cmdMain(5).Left = Shape2.Left + 120
            cmdMain(5).Visible = True
        Case 3  'Save Cancel
            If DisburseType = 4 Then    'Personal Drawings
                Shape2.Height = 100
                Shape2.Visible = False
            Else
                Shape2.Height = cmdMain(0).Height * 2 + 210
                Shape2.Visible = True
            End If
            cmdMain(4).Top = Shape2.Top + 105
            cmdMain(4).Left = Shape2.Left + 120
            cmdMain(4).Visible = True
            cmdMain(5).Top = cmdMain(4).Top + cmdMain(4).Height
            cmdMain(5).Left = Shape2.Left + 120
            cmdMain(5).Visible = True
        Case 4  'View, Print, Exit
            Shape2.Height = cmdMain(0).Height * 3 + 210
            cmdMain(1).Top = Shape2.Top + 105
            cmdMain(1).Left = Shape2.Left + 120
            cmdMain(1).Visible = True
            For i = 2 To 3
                cmdMain(i).Top = cmdMain(i - 1).Top + cmdMain(i - 1).Height
                cmdMain(i).Left = Shape2.Left + 120
                cmdMain(i).Visible = True
            Next i
        Case 5  'Edit, View, Print, Exit
            Shape2.Height = cmdMain(0).Height * 4 + 210
            cmdMain(0).Top = Shape2.Top + 105
            cmdMain(0).Left = Shape2.Left + 120
            cmdMain(0).Visible = True
            For i = 1 To 3
                cmdMain(i).Top = cmdMain(i - 1).Top + cmdMain(i - 1).Height
                cmdMain(i).Left = Shape2.Left + 120
                cmdMain(i).Visible = True
            Next i
    End Select
    If (PayeeType = 0 Or PayeeType = 1) And (DispCode <> 4 And DispCode <> 5) Then 'Employee payee display profile
        fraProfile.Top = Shape2.Top + Shape2.Height + 120
        fraProfile.Left = Shape2.Left
        fraProfile.Visible = True
    End If
    If DisburseType = 4 Then
        Call LockData(False)
        txtPayee.SetFocus
    End If
End Sub

Private Sub DisplayButtonPayee()
    Dim i As Integer
                
    If DisburseType = 3 Then 'Operating Exp LYRGCOM
        cboLYRg.Top = cmdDisburse(0).Top + 495 + 105
        cboLYRg.Left = cmdDisburse(0).Left
        cboLYRg.Visible = True
        Call LockData(False)
        cboLYRg.SetFocus
        Shape1.Height = 2795
    ElseIf DisburseType = 4 Then 'Personal Drawings
        Shape1.Height = 735
        cmdDisburse(ActiveBtn).Enabled = False
        Call DisplayMainButton(3)
        Exit Sub
    Else
        cmdPayee(0).Top = cmdDisburse(0).Top + 495 + 105
        For i = 1 To 3
            cmdPayee(i).Top = cmdPayee(i - 1).Top + 495
        Next i
        For i = 0 To 3
            cmdPayee(i).Left = cmdDisburse(0).Left
            cmdPayee(i).Visible = True
        Next i
        Shape1.Height = 2795
    End If
    cmdDisburse(ActiveBtn).Enabled = False
    Call DisplayMainButton(2)
    If EditBtnClicked Then Call cmdPayee_Click(PayeeType)
        
End Sub

Private Sub PrintPCV()
    Dim PrnVal(0 To 9) As String

    PrnVal(0) = "" 'txtPCV.Text
    PrnVal(1) = txtPayee.Text
    PrnVal(2) = txtDate.Text
    PrnVal(3) = cboAcctTitle.Text
    PrnVal(4) = flxCV.TextMatrix(1, 2)
    PrnVal(5) = "" 'flxCV.TextMatrix(1, 2)
    PrnVal(6) = ""  'txtPreparedBy.Text
    PrnVal(7) = ""  'txtApprovedBy.Text
    PrnVal(8) = ""  'txtPayee.Text
    If DisburseType = 2 Then
        PrnVal(9) = "* Charge to REBUILT FACTORY"
    ElseIf DisburseType = 3 Then
        PrnVal(9) = "* Charge to " & MyDept
    Else
        PrnVal(9) = ""
    End If
    Call PrintMyForm(PrnVal, "Petty Cash Voucher", flxCV)
End Sub

Private Sub cmdDisburse_GotFocus(Index As Integer)
    cmdDisburse(Index).ForeColor = vbRed
End Sub

Private Sub cmdDisburse_LostFocus(Index As Integer)
    cmdDisburse(Index).ForeColor = vbBlack
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
    flxCV.Rows = 1
    If cboRefundSrc.Text = "O.R." Then
        flxCV.AddItem vbTab & "O.R.# " & txtRefundRef.Text & " Date: " & txtRefDate.Text & vbTab & txtRefAmt.Text
    ElseIf cboRefundSrc.Text = "SALES RETURN" Then
        flxCV.AddItem vbTab & "Sales Return Slip# " & txtRefundRef.Text & " Date: " & txtRefDate.Text & vbTab & txtRefAmt.Text
    End If
    If lblAddRef.Caption <> "" Then flxCV.AddItem vbTab & lblAddRef.Caption
    flxCV.AddItem vbTab & "Reason: " & txtReason.Text
    Frame1.Enabled = False
        
End Sub

Private Sub cmdMain_Click(Index As Integer)
    Dim i As Integer, Ans As Integer
    Dim vMsg As String
    
    Select Case Index
        Case 0  'Edit
            EditBtnClicked = True
            NewPCV = False
            Call cmdDisburse_Click(DisburseType)
            Call LockData(False)
        Case 1  'View
            Unload Me
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
            DisburseSumm.Show
        Case 2  'Print
            Ans = MsgBox("Do you want to Print Voucher?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            Call PrintPCV
            If Val(txtPrnCopy.Text) = 0 Then
                'LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.PRINTCOPY = [DISBURSEMENT].[PRINTCOPY]+ 1, DISBURSEMENT.STATUS = 'Final' " & _
                               "WHERE (((DISBURSEMENT.VOUCHERID)=" & txtCVid.Text & "));"
                'mySQL
                LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.PRINTCOPY = DISBURSEMENT.PRINTCOPY + 1, DISBURSEMENT.STATUS = 'Final' " & _
                               "WHERE DISBURSEMENT.VOUCHERID=" & txtCVid.Text & ";"
                txtPrnCopy.Text = "1"
                'Final Saving of involve tables
                'If DisburseType = 1 Then  'Cash Advance
                '    Call SaveCashAdvance
                '    txtCAnotice.Visible = False
                'End If
                Call UpdateFinalVoucher
            Else
                'LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.PRINTCOPY = [DISBURSEMENT].[PRINTCOPY]+ 1 " & _
                               "WHERE (((DISBURSEMENT.VOUCHERID)=" & txtCVid.Text & "));"
                'mySQL
                LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.PRINTCOPY = DISBURSEMENT.PRINTCOPY + 1 " & _
                               "WHERE DISBURSEMENT.VOUCHERID=" & txtCVid.Text & ";"
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
            End If
            DisburseType = -1
            PayeeType = -1
            Call DisplayMainButton(4)
        Case 3  'Exit
            Unload Me     'Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
        Case 4  'Save
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
                'Access
                'LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.PRINTCOPY = 1, DISBURSEMENT.STATUS = 'Final' " & _
                               "WHERE (((DISBURSEMENT.VOUCHERID)=" & txtCVid.Text & "));"
                'mySQL
                LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.PRINTCOPY = 1, DISBURSEMENT.`STATUS` = 'Final' " & _
                               "WHERE DISBURSEMENT.VOUCHERID=" & txtCVid.Text & ";"
                txtPrnCopy.Text = "1"
                Call UpdateFinalVoucher
            Else
                txtPrnCopy.Text = "0"
                If DisburseType = 0 And PayeeType = 1 And txtCAlessCOM.Text <> "" Then
                    '***Access
                    'LYRcon.Execute "INSERT INTO CACOMMTEMP ( COMMORDERID, AGENTID, VOUCHERID, COMMORIGBAL, CALESSCOMM) " & _
                                   "VALUES (" & txtCommOrderID.Text & ", " & txtPayeeID.Text & ", " & txtCVid.Text & ", " & txtCommOrigBal.Text & ", " & txtCAlessCOM.Text & ");"
                    'mySQl
                    LYRcon.Execute "INSERT INTO CACOMMTEMP ( COMMORDERID, AGENTID, VOUCHERID, COMMORIGBAL, CALESSCOMM) " & _
                                   "VALUES (" & txtCommOrderID.Text & ", " & txtPayeeID.Text & ", " & txtCVid.Text & ", " & txtCommOrigBal.Text & ", " & txtCAlessCOM.Text & ");"

                End If
            End If
            If DisburseType = 1 And txtPrnCopy.Text = "1" Then 'Cash Advance
                'Call SaveCashAdvance
                'txtCAnotice.Visible = False
            End If
            If NewPCV Then  'Prompt to process same disbursement type & payee
                Ans = MsgBox("Do you want to Process another voucher" & vbCrLf & "using disbursement type " & MyDisburseType(DisburseType) & "?", vbQuestion + vbYesNo, "Another One?")
                If Ans = vbYes Then
                    Call ClearEntries
                    txtPreparedBy.Text = GetSettings("DISBURSER")
                    txtApprovedBy.Text = GetSettings("MANAGER")
                    txtDate.Text = Date
                    txtPayee.SetFocus
                    Exit Sub
                End If
            End If
            For i = 0 To 3  'Hide Payee Buttons
                cmdPayee(i).Visible = False
                cmdPayee(i).Enabled = True
            Next i
            Shape1.Height = 3735    'Display all Disbursement Button
            cmdDisburse(ActiveBtn).Enabled = True
            'If DispBtn Then
                Call DisplayBtnDisburse(10)
            'Else
            '    If ActiveBtn = 0 Then
            '        cmdDisburse(0).ZOrder
            '        CurrentBtn = 1
            '    Else
            '        cmdDisburse(ActiveBtn).ZOrder
            '        CurrentBtn = 0
            '    End If
            '    cmdDisburse(CurrentBtn).Top = cmdDisburse(ActiveBtn).Top
            '    cmdDisburse(CurrentBtn).Visible = True
            '    ReArrange.Enabled = True
            'End If
            NewPCV = False
            EditBtnClicked = False
            If txtPrnCopy.Text = "0" Then
                Call DisplayMainButton(5)
            Else
                Call DisplayMainButton(4)
            End If
            cboLYRg.Visible = False
            cboLYRg.Text = ""
            SSTab1.Visible = False
            lblHeading(2).Caption = "DISBURSEMENT"
            lblHeading(3).Caption = "DISBURSEMENT"
            Call LockData(True)
            fraProfile.Visible = False
            cmdMain(1).SetFocus
            
        Case 5  'Cancel
            Ans = MsgBox("Do you want to Cancel " & IIf(NewPCV, "Voucher?", "Editing?"), vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            For i = 0 To 3
                cmdPayee(i).Visible = False
                cmdPayee(i).Enabled = True
            Next i
            Shape1.Height = 3735
            cmdDisburse(ActiveBtn).Enabled = True
            'If DispBtn Then
                Call DisplayBtnDisburse(10)
            'Else
            '    If ActiveBtn = 0 Then
            '        cmdDisburse(0).ZOrder
            '        CurrentBtn = 1
            '    Else
            '        cmdDisburse(ActiveBtn).ZOrder
            '        CurrentBtn = 0
            '    End If
            '    cmdDisburse(CurrentBtn).Top = cmdDisburse(ActiveBtn).Top
            '    cmdDisburse(CurrentBtn).Visible = True
            '    ReArrange.Enabled = True
            'End If
            cboLYRg.Visible = False
            cboLYRg.Text = ""
            SSTab1.Visible = False
            lblHeading(2).Caption = "DISBURSEMENT"
            lblHeading(3).Caption = "DISBURSEMENT"
            If EditBtnClicked Then
                Call DisplayMainButton(5)
            Else
                Call DisplayMainButton(1)
                Call ClearEntries
            End If
            NewPCV = False
            EditBtnClicked = False
            Call LockData(True)
            fraProfile.Visible = False
            'tabCashAdv.Visible = False
            txtCAnotice.Visible = False
            cmdMain(1).SetFocus
    End Select
End Sub

Private Sub cmdMain_GotFocus(Index As Integer)
    cmdMain(Index).ForeColor = vbRed
End Sub
  
Private Sub cmdMain_LostFocus(Index As Integer)
    cmdMain(Index).ForeColor = vbBlack
End Sub


Private Sub cmdOR_Click()
    Dim rsPayment As New ADODB.Recordset
    Dim MyOR As String, Qry As String
    Dim Ans As Integer, MyAmt As Single, GetOut As Boolean
    
    If flxCV.TextMatrix(1, 1) <> "" And txtPayee.Text <> "" Then
        Ans = MsgBox("Do you want Add another refund to existing Payee?", vbQuestion + vbYesNo, "Refund")
        If Ans = vbNo Then
            flxCV.Rows = 1
            flxCV.Rows = 2
            txtPayee.Text = ""
        End If
    End If
    MyOR = InputBox("Please type O.R. number to refund:", "Refund")
    If MyOR = "" Then Exit Sub
    '***Access
    'Qry = "SELECT PAYMENTS.* From PAYMENTS WHERE (((PAYMENTS.ORNO)='" & MyOR & "'));"
    'mysql
    Qry = "SELECT PAYMENTS.* From PAYMENTS WHERE PAYMENTS.ORNO='" & MyOR & "';"
    With rsPayment
        If .State = adStateOpen Then Exit Sub
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If flxCV.TextMatrix(1, 1) <> "" And txtPayee.Text <> "" And txtPayee.Text <> !CustomerName Then
                MsgBox "Payee Name and Customer Name in OR does not match!", vbExclamation + vbOKOnly, "Refund"
                .Close
                Exit Sub
            End If
            If !Cash > 0 Then MyAmt = !Cash: RefundType = "CASH"
            If !DownPayment > 0 Then MyAmt = !DownPayment: RefundType = "DOWNPAYMENT"
            If !Installment > 0 Then MyAmt = !Installment: RefundType = "INSTALLMENT"
            If !SPAREPARTS > 0 Then MyAmt = !SPAREPARTS: RefundType = "SPAREPARTS"
            If !Service > 0 Then MyAmt = !Service: RefundType = "SERVICE"
            If !Others > 0 Then MyAmt = !Others: RefundType = "OTHERS"
            Ans = MsgBox("Do you want to refund the following payment:?" & vbCrLf & _
                "Customer Name: " & !CustomerName & vbCrLf & "O.R. Number: " & !ORNO & vbCrLf & "O.R. Date: " & !PayDate & vbCrLf & _
                "O.R. Amount: " & Format$(MyAmt, "#,##0.00"), vbQuestion + vbYesNo, "Refund")
            If Ans = vbYes Then
                If flxCV.TextMatrix(1, 1) = "" Then flxCV.Rows = 1
                flxCV.AddItem vbTab & "OR#:" & !ORNO & " OR Date:" & Format$(!PayDate, "mm/dd/yy") & " OR Amt. " & Format$(MyAmt, "#,##0.00") & vbTab & vbTab & !ORNO
                If Not IsNull(!Remarks) Then flxCV.AddItem vbTab & IIf(RefundType = "SPAREPARTS" Or RefundType = "SERVICE", "Payment for ", "") & !Remarks
                flxCV.TextMatrix(1, 2) = Format(PureNum(flxCV.TextMatrix(1, 2)) + MyAmt, "#,##0.00")
                CustSrch = True
                txtPayee.Text = !CustomerName
                txtPayeeID.Text = !CUSTOMERID
                cmdOR.Tag = Val(cmdOR.Tag) + 1
            End If
        Else
            MsgBox "No record of OR Number " & MyOR & "!", vbExclamation + vbOKOnly, "Refund"
        End If
        .Close
    End With
End Sub

Private Sub cmdPayee_Click(Index As Integer)
    Dim Qry As String, QryD As String
    Dim i As Integer
    
    PayeeType = Index
    'Move selected Button
'    If DispBtn Then
        For i = 0 To 3
            cmdPayee(i).Visible = False
        Next i
        cmdPayee(Index).Top = cmdDisburse(0).Top + 495 + 105
        cmdPayee(Index).Visible = True
        cmdPayee(Index).Enabled = False
        Shape1.Height = 1375
        Call DisplayMainButton(3)
'    Else
'        ActiveBtn2 = Index
'        CurrentBtn2 = 3
'        If Index = 3 Then
'            CurrentBtn2 = 2
'            cmdPayee(3).ZOrder
'            MoveUp2.Enabled = True
'        Else
'            HideBtn2.Enabled = True
'        End If
 '   End If
    Select Case Index
        Case 0  'Employees
            '***Access
            'Qry = "SELECT UCase$([LASTNAME] & ', ' & [FIRSTNAME] & ' ' & IIf(IsNull([MIDDLENAME]),'',Left$([MIDDLENAME],1) & '.')) AS PAYEE, EMPLOYEES.EMPID, EMPLOYEES.ADDRESS, EMPLOYEES.POSITION, EMPLOYEES.EMPPICTURE, EMPLOYEES.EMPSTATUS " & _
                  "From EMPLOYEES WHERE (((EMPLOYEES.DEPARTMENT)<>'Default') AND ((EMPLOYEES.INACTIVE)=False)) ORDER BY UCase$([LASTNAME] & ', ' & [FIRSTNAME] & ' ' & IIf(IsNull([MIDDLENAME]),'',Left$([MIDDLENAME],1) & '.'));"
            'mySQL
            Qry = "SELECT  concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS PAYEE, EMPLOYEES.EMPID, EMPLOYEES.ADDRESS, EMPLOYEES.POSITION, EMPLOYEES.EMPPICTURE, EMPLOYEES.EMPSTATUS " & _
                  "From EMPLOYEES WHERE EMPLOYEES.DEPARTMENT<>'Default' AND EMPLOYEES.INACTIVE=0 ORDER BY concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.')));"
            
            'Old Code
            'If DisburseType = 1 Then 'Cash Advance
            '    If rsAcctDesc.State = adStateOpen Then rsAcctDesc.Close
            '    With rsAcctDesc
            '        QryD = "SELECT DISBURSEACCTDESC.ACCOUNTDESCRIPTION From DISBURSEACCTDESC " & _
            '               "WHERE (((DISBURSEACCTDESC.ACCOUNTTYPE)='C/A EMPLOYEE')) ORDER BY DISBURSEACCTDESC.ACCOUNTDESCRIPTION;"
            '        .Open QryD, LYRcon, adOpenKeyset, adLockOptimistic
            '        If .RecordCount > 0 Then
            '            .MoveFirst
            ''            If Not EditBtnClicked Then cboAcctTitle.Clear
            '            Do While Not .EOF
            '                cboAcctTitle.AddItem .Fields(0)
            '                .MoveNext
            '            Loop
            '        End If
            '    End With
            'End If
        'Case 1  'Sales Agent
        '    Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & IIf(IsNull([MIDDLENAME]),'',Left$([MIDDLENAME],1) & '.') AS PAYEE, AGENTS.AGENTID, AGENTS.ADDRESS, AGENTS.RANK, AGENTS.PICTURE " & _
        '           "FROM AGENTS ORDER BY [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & IIf(IsNull([MIDDLENAME]),'',Left$([MIDDLENAME],1) & '.');"
        '    If DisburseType = 1 Then 'Cash Advance
        '        If rsAcctDesc.State = adStateOpen Then rsAcctDesc.Close
        '        With rsAcctDesc
        '            QryD = "SELECT DISBURSEACCTDESC.ACCOUNTDESCRIPTION From DISBURSEACCTDESC " & _
        '                   "WHERE (((DISBURSEACCTDESC.ACCOUNTTYPE)='C/A AGENT')) ORDER BY DISBURSEACCTDESC.ACCOUNTDESCRIPTION;"
        '            .Open QryD, LYRcon, adOpenKeyset, adLockOptimistic
        '            If .RecordCount > 0 Then
        '                .MoveFirst
        '                If Not EditBtnClicked Then cboAcctTitle.Clear
        '                Do While Not .EOF
        '                    cboAcctTitle.AddItem .Fields(0)
        '                    .MoveNext
        '                Loop
        '            End If
        '        End With
         '   End If
        Case 2  'Customer
            Qry = "SELECT CUSTOMER.CUSTOMERNAME as Payee, CUSTOMER.CUSTOMERID, CUSTOMER.ADDRESS, CUSTOMER.OSBAL From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"

            'If DisburseType = 2 Then cmdOR.Visible = True
                
        Case 3  'Others
            'If DisburseType = 2 Then cmdOR.Visible = True
    End Select
    If Index <> 3 Then  'Load data to recordset except for payee others
        If rsPayee.State = adStateOpen Then rsPayee.Close
        rsPayee.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        Set flxPayee.DataSource = rsPayee
        flxPayee.Refresh
        flxPayee.ColWidth(0) = 0
        flxPayee.ColWidth(1) = flxPayee.Width
    End If
    cmdPayee(Index).Enabled = False
    Call LockData(False)
    txtPayee.SetFocus
    'If EditBtnClicked And DisburseType = 1 And PayeeType <= 1 Then Call DisplayCashAdvance(txtPayeeID.Text)
        
End Sub

Private Sub cmdPayee_GotFocus(Index As Integer)
    cmdPayee(Index).ForeColor = vbRed
End Sub

Private Sub cmdPayee_LostFocus(Index As Integer)
    cmdPayee(Index).ForeColor = vbBlack
End Sub

Private Sub flxAdv_Click()
    Dim rsLgr As New ADODB.Recordset
    Dim Qry As String, MyID As Long, MyRow As Integer
    Dim i As Integer
    
    With flxAdv
        If .Col <> 0 Or .Row = 0 Or .TextMatrix(.Row, 0) = "" Then Exit Sub
        If .TextMatrix(.Row, 0) = "+" Then
            .TextMatrix(.Row, 0) = "-"
            MyID = .TextMatrix(.Row, 6)
            '***Access
            'Qry = "SELECT CALEDGER.* From CALEDGER WHERE (((CALEDGER.CASHADVID)=" & MyId & ") AND ((CALEDGER.STATUS)='ACTIVE'));"
            'mySQL
            Qry = "SELECT CALEDGER.* From CALEDGER WHERE CALEDGER.CASHADVID=" & MyID & " AND CALEDGER.STATUS='ACTIVE';"
            rsLgr.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsLgr.RecordCount > 0 Then
                MyRow = .Row + 1
                i = .Row + 1
                .AddItem vbTab & "Date" & vbTab & "PCV No." & vbTab & "C.A. Amt." & vbTab & "Amt. Paid" & vbTab & "O.S. Bal." & vbTab & .Row, i
                rsLgr.MoveFirst
                Do While Not rsLgr.EOF
                    i = i + 1
                    .AddItem vbTab & rsLgr!TRANSDATE & vbTab & rsLgr!RefNo & vbTab & Format$(rsLgr!CAamount, "###,##0.00") & vbTab & Format$(rsLgr!AmountPaid, "###,##0.00") & vbTab & Format$(rsLgr!CurrentBal, "###,##0.00") & vbTab & .Row, i
                    rsLgr.MoveNext
                Loop
            End If
            rsLgr.Close
            'format cell sub title
            .Row = MyRow: .Col = 0: .RowSel = MyRow: .ColSel = 6
            .FillStyle = flexFillRepeat
            .CellBackColor = .BackColorBkg
            .Row = MyRow: .Col = 0: .RowSel = i: .ColSel = 0
            .FillStyle = flexFillRepeat
            .CellBackColor = .BackColorBkg
        ElseIf .TextMatrix(.Row, 0) = "-" Then
            .TextMatrix(.Row, 0) = "+"
            MyRow = .Row
            For i = .Rows - 1 To MyRow + 1 Step -1
                If .TextMatrix(i, 6) = MyRow Then .RemoveItem i
            Next i
        End If
    End With
End Sub

Private Sub flxComm_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim ShowMenu As Boolean
    
    ShowMenu = False
    If DisburseType = 0 And PayeeType = 1 And cboAcctTitle.Text = "COMMISSION" Then ShowMenu = True
    If DisburseType = 1 And PayeeType = 1 And InStr(UCase$(cboAcctTitle.Text), "COMMISSION") > 0 Then ShowMenu = True
    If flxComm.Row = 0 Or Not ShowMenu Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuComm)
    End If
End Sub

Private Sub flxCV_Click()
    ActRow = flxCV.Row
    ActCol = flxCV.Col
    
    If ActRow = 0 Or ActCol = 0 Then Exit Sub
    If ActRow >= 2 And ActCol = 2 Then Exit Sub
    If flxCV.TextMatrix(ActRow, 3) = "MC" And ActCol = 1 Then
        With cboServiceUnit
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
    Else
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
End Sub

Private Sub flxCV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxCV_Click
End Sub

Private Sub flxCV_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If DisburseType = 0 And PayeeType = 0 And UCase$(cboAcctTitle.Text) = "REPAIR & MAINTENANCE" Then
            mnuSep.Visible = True
            mnuInsMC.Visible = True
        Else
            mnuSep.Visible = False
            mnuInsMC.Visible = False
        End If
        Call PopupMenu(mnuVoucher)
    End If
End Sub


Private Sub flxPayee_Click()
    txtPayee.Text = flxPayee.TextMatrix(flxPayee.Row, 1)
End Sub

Private Sub flxPayee_EnterCell()
    CustSrch = True
    txtPayee.Text = flxPayee.TextMatrix(flxPayee.Row, 1)
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
        cboAcctTitle.SetFocus
    End If
End Sub



Private Sub Form_Load()
    Dim Qry As String
    Dim i As Integer
    
    '***Access
    'Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID, CUSTOMER.APPLYSERVICE " & _
          "From CUSTOMER Where (((CUSTOMER.APPLYSERVICE) = True)) ORDER BY CUSTOMER.CUSTOMERNAME;"
    'mySQL
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID, CUSTOMER.APPLYSERVICE " & _
          "From CUSTOMER Where CUSTOMER.APPLYSERVICE = 1 ORDER BY CUSTOMER.CUSTOMERNAME;"
 
    With rsOnAcct
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboLYRg.AddItem .Fields(0)
                .MoveNext
            Loop
        End If
    End With
    
    MyDisburseType = Array("Spare Parts", "Service Center", "Rebuilt Factory", "On-Acct. Charges", "Personal Drawings", "Non-Operating Expenses", "Cash Voucher")
    MyPayeeType = Array("Employee", "Customer", "Sales Agent", "Others")
    For i = 0 To 6
        CmdTop(i) = cmdDisburse(i).Top
    Next i
    Shape1.BackColor = RGB(105, 125, 255)
    Shape2.BackColor = RGB(105, 125, 255)
    PayeeType = 5   'payee type out of bounce pra dli ma display
    Call DisplayMainButton(1)
    Call LockData(True)
    cboRefundSrc.AddItem "SALES RETURN"
    cboRefundSrc.AddItem "O.R."
    cboRefundSrc.AddItem "C.R."
    cmdAltA.Top = 20000
    'flxRefund.ColWidth(6) = 0
    'cboPayMode.AddItem "Weekly"
    'cboPayMode.AddItem "Bi-Monthly"
    'cboPayMode.AddItem "Monthly"
    'cboPayMode.Text = cboPayMode.List(0)
    'tabCashAdv.Caption = ""
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsAcctDesc.State = adStateOpen Then rsAcctDesc.Close
    If rsPayee.State = adStateOpen Then rsPayee.Close
    If rsService.State = adStateOpen Then rsService.Close
    If rsOnAcct.State = adStateOpen Then rsOnAcct.Close
    'LYRcon.Close
End Sub


Private Sub HideBtn_Timer()
    cmdDisburse(CurrentBtn - 1).ZOrder
    cmdDisburse(CurrentBtn).Top = cmdDisburse(CurrentBtn).Top - 100
    If cmdDisburse(CurrentBtn).Top <= cmdDisburse(CurrentBtn - 1).Top Then
        If CurrentBtn - 1 = ActiveBtn Then
            cmdDisburse(CurrentBtn).Visible = False
            CurrentBtn = ActiveBtn - 1
            HideBtn.Enabled = False
            cmdDisburse(ActiveBtn).ZOrder
            If ActiveBtn <> 0 Then MoveUp.Enabled = True
            If ActiveBtn = 0 Then Call DisplayButtonPayee
        Else
            cmdDisburse(CurrentBtn).Visible = False
            CurrentBtn = CurrentBtn - 1
        End If
    End If
End Sub

Private Sub HideBtn2_Timer()
    cmdPayee(CurrentBtn2 - 1).ZOrder
    cmdPayee(CurrentBtn2).Top = cmdPayee(CurrentBtn2).Top - 50
    If cmdPayee(CurrentBtn2).Top <= cmdPayee(CurrentBtn2 - 1).Top Then
        If CurrentBtn2 - 1 = ActiveBtn2 Then
            cmdPayee(CurrentBtn2).Visible = False
            CurrentBtn2 = ActiveBtn2 - 1
            HideBtn2.Enabled = False
            cmdPayee(ActiveBtn2).ZOrder
            If ActiveBtn2 <> 0 Then MoveUp2.Enabled = True
            If ActiveBtn2 = 0 Then
                Shape1.Height = 1375
                Call DisplayMainButton(3)
            End If
        Else
            cmdPayee(CurrentBtn2).Visible = False
            CurrentBtn2 = CurrentBtn2 - 1
        End If
    End If
End Sub

Private Sub mnuAddRow_Click()
    flxCV.AddItem ""
End Sub



Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete Row")
    If Ans = vbYes Then
        If InStr(cboAcctTitle.Text, "REFUND") > 0 And txtRefundRef.Text <> "" And ActRow = 1 Then
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

Private Sub mnuInsert_Click()
    flxCV.AddItem "", ActRow
End Sub
'For Marketing
Private Sub mnuInsMC_Click()
    Dim Qry As String
    
    'Qry = "SELECT [MODEL] & ' ' & [COLOR] & IIf(IsNull([PLATENO]),'',' Plate No. ' & [PLATENO]) AS SERVICE, MOTORCYCLE.MOTORID " & _
          "From MOTORCYCLE WHERE (((MOTORCYCLE.STATUS)='SERVICE UNIT'));"
    'mySQL
    Qry = "SELECT MOTORCYCLE.MODEL & ' ' & MOTORCYCLE.COLOR & ' Plate No. ' & MOTORCYCLE.PLATENO AS SERVICE, MOTORCYCLE.MOTORID " & _
          "From MOTORCYCLE WHERE MOTORCYCLE.STATUS='SERVICE UNIT';"
 
    With rsService
        If .State = adStateClosed Then
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                cboServiceUnit.Clear
                Do While Not .EOF
                    cboServiceUnit.AddItem !Service
                    .MoveNext
                Loop
            End If
        End If
    End With
    If flxCV.TextMatrix(ActRow, 1) <> "" Then
        flxCV.AddItem ""
        flxCV.Row = flxCV.Rows - 1
        flxCV.Col = 1
        ActRow = flxCV.Row
        ActCol = 1
    Else
        flxCV.Col = 1
        ActCol = 1
    End If
    With cboServiceUnit
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
End Sub

Private Sub mnuMove_Click()
    Dim myCAdue As Single, MyBal As Single, myDeduct As Single
    Dim Ans As Integer
    
    If flxComm.Rows = 1 Then Exit Sub
    With flxComm
        If UCase$(.TextMatrix(.Row, 9)) <> "APPROVED" Then
            MsgBox "Commission status not yet Approved!", vbExclamation + vbOKOnly, "Claim failed"
            Exit Sub
        End If
        If DisburseType = 0 Then flxCV.Rows = 1
        flxCV.AddItem vbTab & "Account No: " & .TextMatrix(.Row, 10) & vbTab & IIf(DisburseType = 0, .TextMatrix(.Row, 7), "")
        flxCV.AddItem vbTab & "Customer: " & .TextMatrix(.Row, 2)
        flxCV.AddItem vbTab & "Unit/Model: " & .TextMatrix(.Row, 3)
        If .TextMatrix(.Row, 5) = "P0.00" Then
            flxCV.AddItem vbTab & "COD: " & .TextMatrix(.Row, 4) & " * " & .TextMatrix(.Row, 11) & IIf(DisburseType = 0, " = " & .TextMatrix(.Row, 7), "")
        Else
            flxCV.AddItem vbTab & "LCP: " & .TextMatrix(.Row, 5) & " * " & .TextMatrix(.Row, 11) & IIf(DisburseType = 0, " = " & .TextMatrix(.Row, 7), "")
        End If
        If PureNum(.TextMatrix(.Row, 7)) < PureNum(.TextMatrix(.Row, 6)) And DisburseType = 0 Then
            flxCV.AddItem vbTab & "Balance: " & .TextMatrix(.Row, 7)
        End If
        txtCommOrderID.Text = .TextMatrix(.Row, 12)
        txtCommOrigBal.Text = .TextMatrix(.Row, 7)
        'Validate Cash Advance
        myCAdue = PureNum(txtCADue.Text)
        MyBal = PureNum(.TextMatrix(.Row, 7))
        myDeduct = IIf(myCAdue > MyBal, MyBal, myCAdue)
        If myCAdue > 0 And DisburseType = 0 Then
            Ans = MsgBox("Do you want to deduct Cash Advance [" & Format$(myDeduct, "###,##0.00") & "] in commission?", vbQuestion + vbYesNo, "Deduct C/A")
            If Ans = vbYes Then
                flxCV.AddItem vbTab & "LESS: Cash Advance = " & Format$(myDeduct, "###,##0.00") & vbTab & vbTab & "CA DEDUCTION"
                flxCV.TextMatrix(1, 2) = Format$(MyBal - myDeduct, "###,##0.00")
                txtCAlessCOM.Text = myDeduct
            Else
                MsgBox "You are required to deduct C.A. in commission", vbExclamation + vbOKOnly, "Cash Advance"
            End If
            cmdDeduct.Visible = True
        End If
    End With
End Sub

Private Sub MoveUp_Timer()
    cmdDisburse(ActiveBtn).Top = cmdDisburse(ActiveBtn).Top - 100
    If cmdDisburse(ActiveBtn).Top <= cmdDisburse(CurrentBtn).Top Then
        cmdDisburse(CurrentBtn).Visible = False
        CurrentBtn = CurrentBtn - 1
    End If
    If cmdDisburse(ActiveBtn).Top <= cmdDisburse(0).Top Then
        MoveUp.Enabled = False
        cmdDisburse(ActiveBtn).Top = cmdDisburse(0).Top
        Call DisplayButtonPayee
    End If
End Sub

Private Sub MoveUp2_Timer()
    cmdPayee(ActiveBtn2).Top = cmdPayee(ActiveBtn2).Top - 50
    If cmdPayee(ActiveBtn2).Top <= cmdPayee(CurrentBtn2).Top Then
        cmdPayee(CurrentBtn2).Visible = False
        CurrentBtn2 = CurrentBtn2 - 1
    End If
    If cmdPayee(ActiveBtn2).Top <= cmdPayee(0).Top Then
        MoveUp2.Enabled = False
        cmdPayee(ActiveBtn2).Top = cmdPayee(0).Top
        Shape1.Height = 1375
        Call DisplayMainButton(3)
    End If
End Sub

Private Sub ReArrange_Timer()
    If ActiveBtn > CurrentBtn Then  'pag move dapat active button ang makita
        cmdDisburse(ActiveBtn).Top = cmdDisburse(ActiveBtn).Top + 50
        If cmdDisburse(ActiveBtn).Top >= CmdTop(CurrentBtn + 1) Then
            cmdDisburse(ActiveBtn).Top = CmdTop(CurrentBtn + 1)
            If ActiveBtn = CurrentBtn + 1 And ActiveBtn <> 6 Then
                CurrentBtn = CurrentBtn + 2
                cmdDisburse(CurrentBtn).Top = CmdTop(CurrentBtn - 1)
            Else
                CurrentBtn = CurrentBtn + 1
                cmdDisburse(CurrentBtn).Top = CmdTop(CurrentBtn)
            End If
            cmdDisburse(CurrentBtn).Visible = True
        End If
    Else
        cmdDisburse(CurrentBtn).Top = cmdDisburse(CurrentBtn).Top + 50
        If cmdDisburse(CurrentBtn).Top >= CmdTop(CurrentBtn) Then
            cmdDisburse(CurrentBtn).Top = CmdTop(CurrentBtn)
            cmdDisburse(CurrentBtn).ZOrder
            If CurrentBtn = 6 Then
                ReArrange.Enabled = False
                Exit Sub
            End If
            cmdDisburse(CurrentBtn + 1).Top = cmdDisburse(CurrentBtn).Top
            cmdDisburse(CurrentBtn + 1).Visible = True
            CurrentBtn = CurrentBtn + 1
        End If
    End If
End Sub


Private Sub txtCADue_Change()
    Dim strValid As String
    strValid = "1234567890."
    
    If KeyAscii > 26 And ActCol = 2 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxCV.SetFocus
        flxCV.Col = 1
    End If
End Sub

Private Sub txtFlx_GotFocus()
    If DisburseType = 1 And PayeeType = 1 Then  'CA - Sales Agent
        If lblPosition.Caption = "Unit Sales Supervisor" Then
            txtCAnotice.Text = lblPosition.Caption & vbCrLf & "Allowable Cash Advance" & vbCrLf & "must NOT exceed" & _
                                vbCrLf & "P300 per week!"
            txtCAnotice.Visible = True
        ElseIf lblPosition.Caption = "Unit Sales Manager" Then
            txtCAnotice.Text = lblPosition.Caption & vbCrLf & "Allowable Cash Advance" & vbCrLf & "must NOT exceed" & _
                                vbCrLf & "P500 per week!"
            txtCAnotice.Visible = True
        End If
    End If
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
    If ActCol = 2 Then
        flxCV.TextMatrix(ActRow, ActCol) = Format$(txtFlx.Text, "###,##0.00")
    Else
        flxCV.TextMatrix(ActRow, ActCol) = txtFlx.Text
    End If
    txtFlx.Visible = False
End Sub

Private Sub txtPayee_Change()
    With flxPayee
        If CustSrch Or DisburseType = 3 Or rsPayee.State = adStateClosed Then
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
            imgPic.Picture = LoadPicture("")
            If PayeeType <= 1 Then
                If Not IsNull(rsPayee.Fields(4)) Then Call ShowPic
            End If
            txtPayee.Text = rsPayee!payee
            If PayeeType < 3 Then
                lblName.Caption = rsPayee!payee
                txtPayeeID.Text = rsPayee.Fields(1)
                lblAddress.Caption = IIf(IsNull(rsPayee.Fields(2)), "", rsPayee.Fields(2))
                lblPosition.Caption = IIf(IsNull(rsPayee.Fields(3)), "", rsPayee.Fields(3))
            End If
            flxCV.Rows = 1: flxCV.Rows = 2: txtCommOrderID.Text = ""
            Select Case PayeeType
                Case 0  'Employee
                    lblStatus.Caption = IIf(IsNull(rsPayee.Fields(5)), "", rsPayee.Fields(5))
                    'If DisburseType = 1 Then    'Cash Advance
                    '    Call DisplayCashAdvance(rsPayee.Fields(1))
                    'End If
                    'SSTab1.Visible = True
                Case 1  'Agent
                    lblStatus.Caption = ""
                    'Display commission and advances for type Cash Advance - Sales Agent
                    If DisburseType = 0 Or DisburseType = 1 Then    'Operating Expenses & C.A.
                        Call DisplayCashAdvance(rsPayee.Fields(1))
                        Call DisplayAgentComm(rsPayee.Fields(1))
                    End If
                    SSTab1.Visible = True
                
                Case 2  'Customer
                    lblPosition.Caption = IIf(lblPosition.Caption = "", "", lblPosition.Caption & "[Spouse]")
                    lblStatus.Caption = ""
                    'If DisburseType = 2 Then    'Cash Refund
                    '    Call DisplayCustRefund
                    'End If
            End Select
            
            txtPayee.SelStart = Val(txtPayee.Tag)
            txtPayee.SelLength = Len(txtPayee.Text)
            CustBln = False
        Else
            rsPayee.Requery
            imgPic.Picture = LoadPicture("")
            lblName.Caption = ""
            lblAddress.Caption = ""
            lblPosition.Caption = ""
            lblStatus.Caption = ""
            txtPayeeID.Text = ""
            SSTab1.Visible = False
            .Visible = False
        End If
    End With
End Sub


Private Sub txtPayee_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If flxPayee.Visible Then
                flxPayee.SetFocus
            Else
                cboAcctTitle.SetFocus
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
        cboAcctTitle.SetFocus
    End If
End Sub


Private Sub txtPCV_GotFocus()
    Dim rsMaxPCVNo As New ADODB.Recordset
    Dim Qry As String
    
    If Not NewPCV Then Exit Sub
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
    If KeyAscii = 13 Then txtDate.SetFocus
End Sub

Private Sub txtPCV_LostFocus()
    Dim Qry As String
    
    If Not NewPCV Then Exit Sub
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

Private Sub txtRefundRef_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    'strValid = "1234567890"
    'If KeyAscii > 26 Then
    '    If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    'End If
    If KeyAscii = 13 Then
        If cboRefundSrc.Text = "SALES RETURN" Then
            txtReason.SetFocus
        Else
            txtRefAmt.SetFocus
        End If
    End If
End Sub

Private Sub txtRefundRef_LostFocus()
    If cboRefundSrc.Text = "SALES RETURN" And txtRefundRef.Text <> "" Then Call DisplaySalesReturn
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

Private Sub DisplayORDetails()
    Dim rsOR As New ADODB.Recordset
    Dim Qry As String, myRem As String, PayID As Long
    Dim Ans As Integer, DP As Currency, SP As Currency, SVC As Currency, OT As Currency
    
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
    Qry = "SELECT SRETURNSP.RETNO, SRETURNSP.RETDATE, SRETURNSP.TOTALAMT, SRETURNSP.PROCESS, SPINVOICE.SPINVNO, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, SPINVOICE.AMOUNT, SPINVOICE.INVTYPE, SRETURNSP.PROREF, SRETURNSP.REASON " & _
          "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO WHERE SRETURNSP.RETNO='" & txtRefundRef.Text & "' ORDER BY SPINVOICE.INVDATE DESC;"
    
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

Private Sub UpdateFinalVoucher()
    Dim rsInv As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim Qry As String, InvType As String, MyInv As String, RetRem As String, ProRef As String
    Dim OldBal As Currency, myTtl As Currency, CustID As Integer, MyRetID As Long
    
    If DisburseType = 0 And InStr(cboAcctTitle.Text, "REFUND") > 0 Then
        If cboRefundSrc.Text = "SALES RETURN" And txtRefundRef.Text <> "" Then
            '***Access
            'Qry = "SELECT SRETURNSP.RETNO, SRETURNSP.RETURNID, SRETURNSP.PROREF, SRETURNSP.REMARKS, SPINVOICE.SPINVNO, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, SPINVOICE.AMOUNT, SPINVOICE.INVTYPE, SPINVOICE.CUSTOMERID " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO WHERE (((SRETURNSP.RETNO)='" & txtRefundRef.Text & "'));"
            'mySQL
            Qry = "SELECT SRETURNSP.RETNO, SRETURNSP.RETURNID, SRETURNSP.PROREF, SRETURNSP.REMARKS, SPINVOICE.SPINVNO, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, SPINVOICE.AMOUNT, SPINVOICE.INVTYPE, SPINVOICE.CUSTOMERID " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO WHERE SRETURNSP.RETNO='" & txtRefundRef.Text & "';"
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
                '***Access
                'Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO)='" & MyInv & "'));"
                'OldBal = Val(GetFieldValue(Qry))
                'myTtl = Val(Replace(flxCV.TextMatrix(1, 2), ",", ""))
                'With rsLedger
                '    .Open "Ledger", LYRcon, adOpenKeyset, adLockOptimistic
                '    .AddNew
                '    !SPINVNO = MyInv
                '    !RefType = "Refund"
                '    !RefNo = txtPCV.Text
                '    !RefDate = txtDate.Text
                '    !Debit = myTtl
                '    !OSbal = OldBal + myTtl
                '    !CUSTOMERID = CustID
                '     ***Access
                '     Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From LEDGER WHERE (((LEDGER.SPINVNO)='" & MyInv & "'));"
                '    !SEQNO = Val(GetFieldValue(Qry)) + 1
                '    .Update
                '    .Close
                'End With
                
                'mySQL
                 myTtl = Val(Replace(flxCV.TextMatrix(1, 2), ",", ""))
                 Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From LEDGER WHERE (((LEDGER.SPINVNO)='" & MyInv & "'));"
                 
                 LYRcon.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, DEBIT, OSBAL, CUSTOMERID, SEQNO) " & _
                                "VALUES('" & MyInv & "', 'Refund', '" & txtPCV.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
                                "" & myTtl & ", " & OldBal + myTtl & ", " & CustID & ", Val(GetFieldValue(Qry)) + 1);"
                
                'Update Customer and Invoice Balance
                MyStat = IIf(OldBal - myTtl >= 1, "OPEN", "CLOSED")
                '***Access
                'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & OldBal + myTtl & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE (((ONACCTERMS.SPINVNO)='" & MyInv & "'));"
                'mySQL
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & OldBal + myTtl & ", ONACCTERMS.`STATUS` = '" & MyStat & "' " & _
                               "WHERE ONACCTERMS.SPINVNO='" & MyInv & "';"
                
                'Update Customer Balance
                '***Access
                'Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                      "WHERE (((ONACCTERMS.CUSTOMERID)=" & CustID & ") AND ((ONACCTERMS.STATUS)='OPEN'));"
                'mySQL
                Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                      "WHERE ONACCTERMS.CUSTOMERID=" & CustID & " AND ONACCTERMS.`STATUS`='OPEN';"
                OldBal = Val(GetFieldValue(Qry))
                '***Access
                'LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = " & OldBal & " WHERE (((CUSTOMER.CUSTOMERID)=" & CustID & "));"
                'mySQL
                LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = " & OldBal & " WHERE CUSTOMER.CUSTOMERID=" & CustID & ";"
            End If
            RetRem = IIf(RetRem <> "", RetRem & ", ", "") & "Refund from PCV#" & txtPCV.Text & " Amt: " & flxCV.TextMatrix(1, 2)
            If ProRef <> "" Then
                '***Access
                'LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = '" & Left$(RetRem, 250) & "' WHERE (((SRETURNSP.RETURNID)=" & MyRetID & "));"
                'mySQL
                LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = '" & Left(RetRem, 250) & "' WHERE SRETURNSP.RETURNID=" & MyRetID & ";"
            Else
                '***Access
                'LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = '" & Left$(RetRem, 250) & "', SRETURNSP.PROREF = '" & txtPCV.Text & "' WHERE (((SRETURNSP.RETURNID)=" & MyRetID & "));"
                'mySQL
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
            '***Access
            'Qry = "SELECT LEDGER.* From Ledger WHERE (((LEDGER.REFTYPE)='Payment') AND ((LEDGER.REFNO)='" & txtRefundRef.Text & "'));"
            
            'mysql
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
                '***Access
                'Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE ONACCTERMS.SPINVNO='" & MyInv & "';"
                'OldBal = Val(GetFieldValue(Qry))
                'myTtl = Val(Replace(flxCV.TextMatrix(1, 2), ",", ""))
                'With rsLedger
                '    .Open "Ledger", LYRcon, adOpenKeyset, adLockOptimistic
                '    .AddNew
                '    !SPINVNO = MyInv
                '    !RefType = "Refund"
                '    !RefNo = txtPCV.Text
                '    !REFDATE = txtDate.Text
                '    !DEBIT = myTtl
                '    !OSbal = OldBal + myTtl
                '    !CUSTOMERID = CustID
                '    !SEQNO = MyRetID
                '    .Update
                '    .Close
                'End With
                
                'mySQL
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
