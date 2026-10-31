VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form SPSorderOLD 
   Caption         =   "Spare Parts Sales Order"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   ControlBox      =   0   'False
   Icon            =   "SPSorderOld.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   11535
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCPOid 
      Height          =   285
      Left            =   10440
      TabIndex        =   76
      Top             =   9240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSQtyBck 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   8880
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CheckBox chkDP 
      Caption         =   "DP"
      Height          =   255
      Left            =   360
      TabIndex        =   73
      Top             =   9600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtJOid 
      Height          =   285
      Left            =   1800
      TabIndex        =   72
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEmpID 
      Height          =   285
      Left            =   10920
      TabIndex        =   70
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame fraAdmin 
      Height          =   735
      Left            =   5520
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cboOrderNo 
         Height          =   315
         Left            =   4200
         TabIndex        =   69
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cboPartsMan 
         Height          =   315
         Left            =   960
         TabIndex        =   67
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Order No."
         Height          =   255
         Left            =   3360
         TabIndex        =   68
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Partsman:"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.TextBox txtOSBal 
      Height          =   285
      Left            =   6480
      TabIndex        =   64
      Top             =   10440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   5040
      TabIndex        =   63
      Top             =   10440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPONO 
      Height          =   285
      Left            =   7200
      TabIndex        =   62
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtBillAmt 
      Height          =   285
      Left            =   4200
      TabIndex        =   61
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPayStart 
      Height          =   285
      Left            =   3360
      TabIndex        =   60
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtNPays 
      Height          =   285
      Left            =   2520
      TabIndex        =   59
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPayScheme 
      Height          =   285
      Left            =   1560
      TabIndex        =   58
      Top             =   10440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTerms 
      Height          =   285
      Left            =   480
      TabIndex        =   57
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCustName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   55
      Tag             =   "XYF"
      Top             =   6720
      Width           =   4335
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   360
      TabIndex        =   52
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAppModel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   51
      Top             =   8520
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.TextBox txtNoDiscount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "No Discount"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtGrandTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   47
      Tag             =   "XYF"
      Top             =   6855
      Width           =   1620
   End
   Begin VB.TextBox txtLessDiscAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   46
      Tag             =   "XYF"
      Top             =   6570
      Width           =   1620
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   45
      Tag             =   "XYF"
      Top             =   6285
      Width           =   1620
   End
   Begin VB.TextBox txtPctDisc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   42
      Tag             =   "XYF"
      Text            =   "0%"
      Top             =   6607
      Width           =   495
   End
   Begin VB.TextBox txtTtlDisc 
      Height          =   285
      Left            =   9840
      TabIndex        =   40
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8760
      TabIndex        =   39
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   8040
      TabIndex        =   38
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtRefNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   36
      Tag             =   "XYF"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox txtInvNo 
      Height          =   285
      Left            =   5760
      TabIndex        =   32
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid flxPND 
      Height          =   320
      Left            =   5400
      TabIndex        =   31
      Top             =   150
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   556
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   16777215
      BackColorBkg    =   12648447
      GridColor       =   16777215
      GridColorFixed  =   16777215
      ScrollBars      =   1
      BorderStyle     =   0
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   315
      Left            =   8160
      TabIndex        =   27
      Tag             =   "XYF"
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "Find"
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
      Image           =   "SPSorderOld.frx":08CA
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   9120
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtSPSorderID 
      Height          =   285
      Left            =   480
      TabIndex        =   25
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   720
      Left            =   240
      TabIndex        =   23
      Tag             =   "XYF"
      Top             =   7200
      Width           =   11055
      Begin VB.TextBox txtFrnt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   48
         Tag             =   "XYF"
         Top             =   240
         Width           =   4335
      End
      Begin lvButton.lvButtons_H cmdAdvFind 
         Height          =   495
         Left            =   10440
         TabIndex        =   28
         Tag             =   "XYF"
         ToolTipText     =   "[F12] Advance Search"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderOld.frx":0E64
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCashier 
         Height          =   495
         Left            =   7680
         TabIndex        =   29
         Tag             =   "XYF"
         ToolTipText     =   "Cashier"
         Top             =   165
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderOld.frx":13FE
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdWareHouse 
         Height          =   495
         Left            =   8784
         TabIndex        =   30
         Tag             =   "XYF"
         ToolTipText     =   "Warehouse"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderOld.frx":1AD4
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   8232
         TabIndex        =   33
         Tag             =   "XYF"
         ToolTipText     =   "Cancel Order"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderOld.frx":206E
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdReq 
         Height          =   495
         Left            =   9888
         TabIndex        =   35
         Tag             =   "XYF"
         ToolTipText     =   "Insert Requisition"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderOld.frx":2F48
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPO 
         Height          =   495
         Left            =   9336
         TabIndex        =   37
         Tag             =   "XYF"
         ToolTipText     =   "[F11] Insert On Account"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderOld.frx":34E2
         cBack           =   -2147483633
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FRONTLINER:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   49
         Tag             =   "XYF"
         Top             =   292
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPrintCopy 
      Height          =   285
      Left            =   4920
      TabIndex        =   22
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   8280
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   8760
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   8280
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   8760
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   1560
      TabIndex        =   19
      Top             =   8040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCust 
      Height          =   1815
      Left            =   2640
      TabIndex        =   18
      Top             =   8400
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3201
      _Version        =   393216
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1815
      Left            =   6000
      TabIndex        =   17
      Top             =   8400
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtReqID 
      Height          =   285
      Left            =   6240
      TabIndex        =   16
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAltA 
      Caption         =   "&A"
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Tag             =   "XY"
      Top             =   10920
      Width           =   135
   End
   Begin VB.Timer tmrQtyTip 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   8400
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   600
      TabIndex        =   14
      Text            =   "Qty. On Hand: 0"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtFlx 
      Height          =   315
      Left            =   3720
      TabIndex        =   13
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAcctNo 
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Top             =   9240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   1800
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   0
      FormatString    =   $"SPSorderOld.frx":3A7C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   240
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   720
      Width           =   11055
      Begin VB.TextBox txtOrderno 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   180
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   315
         Left            =   9240
         TabIndex        =   71
         Tag             =   "XYF"
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER NO."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   210
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorderOld.frx":3B0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorderOld.frx":49E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorderOld.frx":58C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorderOld.frx":679B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorderOld.frx":7675
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&New Order"
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
      Image           =   "SPSorderOld.frx":854F
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Order"
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
      Image           =   "SPSorderOld.frx":8AE9
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Save Order"
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
      Image           =   "SPSorderOld.frx":9083
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Cancel Order"
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
      Image           =   "SPSorderOld.frx":961D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   24
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Edit Order"
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
      Image           =   "SPSorderOld.frx":9BB7
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxBO 
      Height          =   1095
      Left            =   5400
      TabIndex        =   75
      Top             =   9360
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Date   |<PO Number    |<Orders"
   End
   Begin VB.Label lblRemarks 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3840
      TabIndex        =   56
      Tag             =   "XYF"
      Top             =   6360
      Width           =   2325
   End
   Begin VB.Label lblCustomer 
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   54
      Tag             =   "XYF"
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label lblRef 
      Caption         =   "Reference No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   53
      Tag             =   "XYF"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label lblGrandTtl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "  GRAND TOTAL:  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8190
      TabIndex        =   44
      Tag             =   "XYF"
      Top             =   6885
      Width           =   1335
   End
   Begin VB.Label lblDiscount 
      AutoSize        =   -1  'True
      Caption         =   "DISCOUNT:  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8520
      TabIndex        =   43
      Tag             =   "XYF"
      Top             =   6607
      Width           =   900
   End
   Begin VB.Label lblLess 
      AutoSize        =   -1  'True
      Caption         =   "Less:  "
      Height          =   195
      Left            =   7440
      TabIndex        =   41
      Tag             =   "XYF"
      Top             =   6615
      Width           =   465
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
      Left            =   0
      TabIndex        =   34
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   10800
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11655
      m_MDI_Height    =   8445
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "  TOTAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8760
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   6330
      Width           =   600
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDERED SPARE PARTS"
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
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   1545
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
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
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   450
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
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
      Index           =   1
      Left            =   270
      TabIndex        =   8
      Top             =   480
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDERED SPARE PARTS"
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
      Height          =   240
      Index           =   5
      Left            =   270
      TabIndex        =   10
      Top             =   1575
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   60
      Width           =   615
   End
   Begin VB.Menu mnuRows 
      Caption         =   "Spare Parts Invoice"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row       Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Just Scroll"
      Visible         =   0   'False
      Begin VB.Menu Mark1 
         Caption         =   "Mark Item"
      End
      Begin VB.Menu Mark0 
         Caption         =   "Unmark Item"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu MarkAll 
         Caption         =   "Mark All Item"
      End
      Begin VB.Menu UnmarkAll 
         Caption         =   "Unmark All Item"
      End
   End
End
Attribute VB_Name = "SPSorderOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim rsPartsMan As New ADODB.Recordset

Dim PrtSrch As Boolean, ActSrch As Boolean, PrtBln As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim tmrCount As Integer, BtnIndx As Integer
Dim AddNew As Boolean, PmanSrch As Boolean
Public JustScroll As Boolean
Dim fRow As Integer, fCol As Integer
Public EmpId As Integer

Private Sub ClearEntries()
    txtOrderno.Text = "<Auto>"
    'txtDate.Text = ""
    lblRemarks.Caption = ""
    lblRef.Caption = "Reference No."
    lblCustomer.Caption = "Customer"
    txtTotal.Text = ""
    txtSPSorderID.Text = ""
    txtInvNo.Text = ""
    txtReqID.Text = ""
    txtPONO.Text = ""
    txtRefNo.Text = ""
    txtCustID.Text = ""
    txtTtlDisc.Text = ""
    txtTerms.Text = ""
    txtPayScheme.Text = ""
    txtNPays.Text = ""
    txtPayStart.Text = ""
    txtBillAmt.Text = ""
    txtOSBal.Text = ""
    txtAddress.Text = ""
    txtCustName.Text = ""
    txtGrandTotal.Text = ""
    txtLessDiscAmt.Text = ""
    flxOrd.Rows = 1
    flxOrd.Rows = 2
    txtPrintCopy = ""
End Sub
                                                                
Private Sub LockData(Bln As Boolean)
    If UserType = "Administrator" Then
        fraAdmin.Visible = Bln
    End If
    JustScroll = Bln
End Sub
   
Private Sub cboOrderNo_Change()
    Call flxPND_Click
End Sub

Private Sub cboOrderNo_Click()
    Call flxPND_Click
End Sub

Private Sub cboOrderNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPfind.MySource = "SPSORDER"
        SPfind.Show vbModal
    End If
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
End Sub

Private Sub cboPartsMan_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboPartsMan
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsPartsMan.MoveFirst
            rsPartsMan.Move i
            txtEmpID.Text = rsPartsMan!EmpId
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PmanSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtEmpID.Text = ""
        End If
    End With
    cboOrderNo.Clear
    cboOrderNo.Text = ""
    Call ClearEntries
End Sub

Private Sub cboPartsMan_Click()
    Call cboPartsMan_Change
End Sub

Private Sub cboPartsMan_GotFocus()
    Dim Tmp
    
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cboPartsMan.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboPartsMan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PmanSrch = True
End Sub

Private Sub cboPartsMan_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPfind.MySource = "SPSORDER"
        SPfind.Show vbModal
    End If
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
End Sub

Private Sub cboPartsMan_LostFocus()
    Dim rsPmanID As New ADODB.Recordset
    Dim Qry As String
    
    If txtEmpID.Text <> "" Then
        Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & txtEmpID.Text & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#)) ORDER BY ORDERNO;"
        With rsPmanID
            cboOrderNo.Clear
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    cboOrderNo.AddItem !ORDERNO
                    .MoveNext
                Loop
            End If
            .Close
        End With
    End If
End Sub

Private Sub cmdAdvFind_Click()
    SPfind.MySource = "SPSORDER"
    SPfind.Show vbModal
    'MotoRep.sBar.Panels(6).Text = "Spare Parts Sales Order"
End Sub

Private Sub cmdAltA_Click()
    Call mnuAdd_Click
End Sub

Private Sub cmdCancel_Click()
    Dim Ans As Integer
    
    If txtOrderno.Text = "" Or Not IsNumeric(txtOrderno.Text) Then Exit Sub
    Ans = MsgBox("Do you want to Cancel Order No. " & txtOrderno.Text & "?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = 'CANCELLED' " & _
                   "WHERE (((SPSALESORDER.SPSORDERID)=" & txtSPSorderID.Text & "));"
    MsgBox "Order successfully Cancelled!", vbInformation + vbOKOnly, "Cancel"
    Call cboPartsMan_Change
    Call ClearEntries
    If UserType <> "Administrator" Then Call DisplayPendingOrders
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSI
        For i = fRow To .Rows - 1
            For j = fCol To 9
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 9:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 9, i + 1, i)
                    fCol = IIf(j = 9, 1, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1
End Sub

Private Sub cmdPO_Click()
    InsertPO.Show vbModal
End Sub

Private Sub cmdReq_Click()
    Dim rsReq As New ADODB.Recordset
    Dim rsReqDtls As New ADODB.Recordset
    Dim MyReq As String, Qry As String
    Dim Ans As Integer, ReqID As Long
    
    If flxOrd.Rows >= 2 And flxOrd.TextMatrix(flxOrd.Rows - 1, 10) <> "" Then
        Ans = MsgBox("All Spareparts  listed in the Order Slip will be removed!" & vbCrLf & "Do you want to Continue?", vbQuestion + vbYesNo)
        If Ans = vbNo Then Exit Sub
    End If
    MyReq = InputBox("Please enter the Requisition Number:", "Req. No.")
    If MyReq = "" Then Exit Sub
    Qry = "SELECT REQUISITION.REQNO, REQUISITION.JOBNO, REQUISITION.CUSTOMER, REQUISITION.REQID, REQUISITION.JOBORDERID, REQUISITION.CUSTOMERID " & _
          "From Requisition WHERE (((REQUISITION.REQNO)='" & MyReq & "'));"
    With rsReq
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ReqID = !ReqID
            lblRef.Caption = "Requisition No."
            txtRefNo.Text = !REQNO
            lblRemarks.Caption = "Job Order No. " & !JobNo
            txtCustName.Text = !CUSTOMER
            lblCustomer.Caption = "Service Customer"
            txtJOid.Text = !JoborderID
            txtCustID.Text = !CUSTOMERID
        Else
            ReqID = 0
        End If
        .Close
    End With
    If ReqID = 0 Then
        MsgBox "Requisition No. " & MyReq & " does not exist!", vbExclamation + vbOKOnly, "Requisition"
        Exit Sub
    End If
    flxOrd.Rows = 1
    Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, [PRICE]*[QTY] AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE (((SPSORDERDETAIL.REQID)=" & ReqID & "));"
    With rsReqDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = 1
            Do While Not .EOF
                flxOrd.AddItem vbTab & !QtyOnHand & vbTab & !Qty & vbTab & IIf(IsNull(!Unit), "", !Unit) & vbTab & !PartNumber & vbTab & !Description & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlPrice, "###,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & !SparepartID
                .MoveNext
            Loop
        End If
        .Close
    End With
    txtReqID.Text = ReqID
    Call ComputeTotal
    If Val(txtCustID.Text) > 0 Then
        Call GetCustomerDefaultTerm
        lblCustomer.Caption = "On-Acct. Customer"
    End If
    Call LockData(False)
End Sub

Private Sub GetCustomerDefaultTerm()
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String, Gtotal As Currency
    
    'Saving On Account Job Order to On Account Terms
    'Get Terms
    Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.STATUS)='Default') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
    With rsTerm
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtTerms.Text = !Terms
            txtPayScheme.Text = !PayScheme
            txtNPays.Text = !Npays
        Else
            txtTerms.Text = 30
            txtPayScheme.Text = "MONTHLY"
            txtNPays.Text = 1
        End If
        .Close
    End With
    txtPayStart.Text = StartPay(txtPayScheme.Text)
    Gtotal = Val(Replace(txtGrandTotal.Text, ",", ""))
    If Val(txtNPays.Text) > 0 Then
        txtBillAmt.Text = Gtotal / Val(txtNPays.Text)
    Else
        txtBillAmt.Text = Gtotal
    End If
End Sub

Function StartPay(S As String) As Date
    Dim iPS As Integer, sPay As Date, MyDate As Date
    
    MyDate = txtDate.Text
    If S = "WEEKLY" Then iPS = 7
    If S = "BI-MONTHLY" Then iPS = 15
    If S = "MONTHLY" Then iPS = 30
    If iPS = 7 Then
        If Day(MyDate) < 5 Then
            sPay = DateSerial(Year(MyDate), Month(MyDate), 7)
        ElseIf Day(MyDate) < 12 Then
            sPay = DateSerial(Year(MyDate), Month(MyDate), 14)
        ElseIf Day(MyDate) < 19 Then
            sPay = DateSerial(Year(MyDate), Month(MyDate), 21)
        ElseIf Day(MyDate) < 28 Then
            sPay = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, MyDate)), Month(DateAdd("m", 1, MyDate)), 1))
        Else
            sPay = DateAdd("m", 1, DateSerial(Year(MyDate), Month(MyDate), 7))
        End If
        StartPay = sPay
    ElseIf iPS = 15 Then
        StartPay = DateAdd("d", iPS, MyDate)
    ElseIf iPS = 30 Then
        StartPay = DateAdd("m", 1, MyDate)
    End If
End Function


Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 12) As String
    Dim MyAcc As String, Qry As String
    
    On Error GoTo LocalErr
    Select Case Index
        Case 0  'New
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=25));"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            Call ClearEntries
            Call LockData(False)
            'txtDate.Text = Date
            txtOrderno.Text = "<Auto>"
            txtFrnt.Text = getFrontLiner
            AddNew = True
            Call DisplayMainButton(1)
            cmdCancel.Visible = False
            'cmdCashier.Visible = False
            cmdWareHouse.Visible = False
            flxPND.Cols = 0
            flxOrd.SetFocus
            flxOrd.Col = 4
        
        Case 1  'Edit
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=25));"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If txtSPSorderID.Text = "" Then
                MsgBox "No available Sales Order to edit", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If txtInvNo.Text <> "" Then
                MsgBox "Unable to edit Order# " & txtOrderno.Text & " because an Invoice was already issued.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If InStr(lblRemarks.Caption, "CANCELLED") > 0 Then
                MsgBox "Unable to edit Sales Order " & txtOrderno.Text & " because it has already been Cancelled.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            flxPND.Cols = 0
            Call DisplayMainButton(1)
            AddNew = False
        Case 2  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
        Case 3  'Save
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            myMsg = ValidateSalesOrder
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save order?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'Save Spare Parts Invoice
            
            Call SaveSalesOrder
            Ans = MsgBox("Sales Order Successfully Saved!" & vbCrLf & vbCrLf & "Process Another Order?", vbQuestion + vbYesNo, "Order")
            If Ans = vbNo Then
                Call DisplayMainButton(2)
                cmdCancel.Visible = True
                'cmdCashier.Visible = True
                cmdWareHouse.Visible = True
                If UserType <> "Administrator" Then
                    Call DisplayPendingOrders
                End If
            Else
                Call cmdTrans_Click(0)
            End If
        Case 4  'Cancel
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            Ans = MsgBox("Do you want to cancel Order?", vbQuestion + vbYesNo)
            If Ans = vbNo Then Exit Sub
            Call DisplayMainButton(2)
            Call ClearEntries
            Ans = MsgBox("Process Another Order?", vbQuestion + vbYesNo, "Order")
            If Ans = vbNo Then
                MotoRep.tbMain.Visible = True
                Unload Me
            Else
                cmdCancel.Visible = True
                'cmdCashier.Visible = True
                cmdWareHouse.Visible = True
                If UserType <> "Administrator" Then
                    Call DisplayPendingOrders
                End If
                Call LockData(True)
            End If
    End Select
    Exit Sub
LocalErr:
    MsgBox "Error Code: " & CStr(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly, "Erro"
    
End Sub

Private Sub cmdTrans_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPfind.MySource = "SPSORDER"
        SPfind.Show vbModal
    End If
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
End Sub

Private Sub cmdWareHouse_Click()
    Dim i As Integer, vSkip As Boolean
    
    'Validate items if all are found
    With flxOrd
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "" Then
                vSkip = True
                Exit For
            Else
                vSkip = False
            End If
        Next i
    End With
    If Not vSkip Then
        MsgBox "No available items for Picking!", vbExclamation + vbOKOnly, "Warning"
        Exit Sub
    End If
    LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.DESTINATION = 'WAREHOUSE' " & _
                   "WHERE (((SPSALESORDER.SPSORDERID)=" & txtSPSorderID.Text & "));"
    MsgBox "Sales Order can now be viewed in Warehouse!", vbInformation + vbOKOnly, "Updated in Warehouse"
    
End Sub

Function GetBackOrder(MyID As Integer) As Integer
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String
    Dim xRow As Integer, xBal As Integer, i As Integer
    
    'Clear FlxStock
    Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & MyID & ")) ORDER BY SPITR.TRANSDATE;"
    flxBO.Rows = 1
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxBO.AddItem Format(!TransDate, "mm/dd/yyyy") & vbTab & !Reference & vbTab & !Order
                .MoveNext
            Loop
        End If
        .Close
    End With
    xBal = 0
    With flxBO
        For i = 1 To .Rows - 1
            'Display PO in Stock Card and obtain Balance
            xBal = xBal + Val(.TextMatrix(i, 2))
            'flxStock.AddItem .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & vbTab & .TextMatrix(i, 2) & vbTab & vbTab & xBal   '.TextMatrix(i, 2)
            'Acquire Packing list for P.O.
            MyPO = .TextMatrix(i, 1)
            Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.PACKINGLIST)='" & MyPO & "')) ORDER BY SPITR.TRANSDATE;"
            rsRec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsRec.RecordCount > 0 Then
                rsRec.MoveFirst
                Do While Not rsRec.EOF
                    xBal = xBal - rsRec!Receive
                    'flxStock.AddItem Format(rsRec!TransDate, "mm/dd/yyyy") & vbTab & MyPO & vbTab & rsRec!Reference & vbTab & vbTab & rsRec!Receive & vbTab & xBal
                    rsRec.MoveNext
                Loop
            End If
            rsRec.Close
        Next i
    End With
    GetBackOrder = xBal
End Function

Function GetSalesQty(MyID As Long) As Integer
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String, StartDate As Date, EndDate As Date
    Dim MySales As Integer, MyRet As Integer
    
    StartDate = DateSerial(Year(DateAdd("m", -3, Date)), Month(DateAdd("m", -3, Date)), 1)
    EndDate = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), Day(DateAdd("d", -1, DateSerial(Year(Date), Month(Date), 1))))
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & MyID & "));"
    With rsQty
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#));"
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function


Private Sub flxOrd_Click()
    Dim MyDesc As String, PrvW As Integer, CrnW As Integer, MaxW As Integer, MyBO As Integer, MySQ As Integer
    
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
    
    If ActCol = 0 Or ActRow = 0 Or JustScroll Then Exit Sub
    Select Case ActCol
        Case 1      'Available
            If Val(flxOrd.TextMatrix(ActRow, 10)) > 0 Then
                MyBO = GetBackOrder(flxOrd.TextMatrix(ActRow, 10))
                MySQ = GetSalesQty(flxOrd.TextMatrix(ActRow, 10))
                With txtSQtyBck
                    .Top = flxOrd.CellTop + flxOrd.Top + 300
                    .Left = flxOrd.CellLeft + flxOrd.Left
                    .Width = 1800   'flxOrd.CellWidth * 2
                    .Height = 315
                    .Text = "BO=" & MyBO & ", SQ=" & MySQ & " [3 mos.]"
                    .Visible = True
                    .ZOrder
                End With
            End If
        Case 2    'qty
            With txtFlx
                If flxOrd.TextMatrix(ActRow, 2) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 2)
                .Top = flxOrd.CellTop + flxOrd.Top
                .Left = flxOrd.CellLeft + flxOrd.Left
                .Width = flxOrd.CellWidth
                .Text = flxOrd.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 4 'PartNumber
            With txtPartNo
                If flxOrd.TextMatrix(ActRow, 4) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 1) & "|" & flxOrd.TextMatrix(ActRow, 2) & "|" & flxOrd.TextMatrix(ActRow, 3) & "|" & flxOrd.TextMatrix(ActRow, 4) & "|" & flxOrd.TextMatrix(ActRow, 5) & "|" & flxOrd.TextMatrix(ActRow, 6) & "|" & flxOrd.TextMatrix(ActRow, 7) & "|" & flxOrd.TextMatrix(ActRow, 8) & "|" & flxOrd.TextMatrix(ActRow, 9) & "|" & flxOrd.TextMatrix(ActRow, 10) & "|" & flxOrd.TextMatrix(ActRow, 11)
                .Top = flxOrd.CellTop + flxOrd.Top
                .Left = flxOrd.CellLeft + flxOrd.Left
                .Width = flxOrd.CellWidth
                .Text = flxOrd.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 5  'Description show Applic. Model tool tip
            MyDesc = ""
            If Len(flxOrd.TextMatrix(ActRow, 5)) > 23 Then
                MyDesc = flxOrd.TextMatrix(ActRow, 5)
            End If
            PrvW = Len(MyDesc)
            If flxOrd.TextMatrix(ActRow, 11) <> "" Then
                MyDesc = IIf(MyDesc = "", "App. Model:" & flxOrd.TextMatrix(ActRow, 11), MyDesc & vbCrLf & "App. Model:" & flxOrd.TextMatrix(ActRow, 11))
            End If
            CrnW = Len(flxOrd.TextMatrix(ActRow, 11)) + 12
            If Len(MyDesc) > 0 Then
                With txtAppModel
                    .Top = flxOrd.CellTop + flxOrd.Top + 300
                    .Left = flxOrd.CellLeft + flxOrd.Left
                    .Width = IIf(PrvW < CrnW, CrnW * 100, PrvW * 100)
                    .Height = IIf(PrvW > 0 And CrnW - 12 > 0, 560, 315)
                    .Text = MyDesc
                    .Visible = True
                    .ZOrder
                End With
            End If
        Case 6  'Discount
            With txtDiscount
                If flxOrd.TextMatrix(ActRow, 4) <> "" And Val(flxOrd.TextMatrix(ActRow, 12)) > 0 Then
                    If flxOrd.TextMatrix(ActRow, 6) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 7)
                    'Display discount
                    txtNoDiscount.Top = flxOrd.CellTop + flxOrd.Top + flxOrd.CellHeight
                    txtNoDiscount.Left = flxOrd.CellLeft + flxOrd.Left
                    txtNoDiscount.Text = "Max. " & flxOrd.TextMatrix(ActRow, 12) & IIf(Right$(flxOrd.TextMatrix(ActRow, 12), 1) <> "%", "%", "")
                    txtNoDiscount.Visible = True
                    .ZOrder
                    'Display Textbox
                    .Top = flxOrd.CellTop + flxOrd.Top
                    .Left = flxOrd.CellLeft + flxOrd.Left
                    .Width = flxOrd.CellWidth
                    If flxOrd.Text <> "" Then
                        .Text = IIf(Right$(flxOrd.Text, 1) <> "%", flxOrd.Text, Left$(flxOrd.Text, Len(flxOrd.Text) - 1))
                    Else
                        .Text = flxOrd.Text
                    End If
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                Else
                    txtNoDiscount.Top = flxOrd.CellTop + flxOrd.Top + 250
                    txtNoDiscount.Left = flxOrd.CellLeft + flxOrd.Left
                    txtNoDiscount.Visible = True
                    .ZOrder
                End If
            End With
        Case 7         'Allow Administrators to Change the Price ... just temporary
            With txtPrice
                If flxOrd.TextMatrix(ActRow, 4) <> "" Then 'And UserType = "Administrator" Then
                    If flxOrd.TextMatrix(ActRow, 7) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 7)
                    .Top = flxOrd.CellTop + flxOrd.Top
                    .Left = flxOrd.CellLeft + flxOrd.Left
                    .Width = flxOrd.CellWidth
                    .Text = flxOrd.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End If
            End With
    End Select
End Sub
                             
Private Sub flxOrd_DblClick()
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
                       
    If ActCol = 0 And txtSPSorderID.Text <> "" Then
        If flxOrd.TextMatrix(ActRow, 0) = "" Then
            flxOrd.Row = ActRow: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
            LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'FOUND' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(ActRow, 10) & "));"
        Else
            flxOrd.TextMatrix(ActRow, 0) = ""
            LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'BLANK' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(ActRow, 10) & "));"
        End If
    End If
End Sub


Private Sub flxOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxOrd_Click
End Sub
                                                  
Private Sub flxOrd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPfind.MySource = "SPSORDER"
        SPfind.Show vbModal
    End If
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
End Sub

Private Sub flxOrd_LeaveCell()
    txtNoDiscount.Visible = False
    txtAppModel.Visible = False
    txtSQtyBck.Visible = False
End Sub

Private Sub flxOrd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If JustScroll Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuStatus)
        End If
    Else
        If Button = vbRightButton Then
            Call PopupMenu(mnuRows)
        End If
    End If
End Sub

Private Sub flxParts_Click()
    Dim PctDisc As Single
    Dim DiscPrice As Currency, TtlPrice As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxOrd.TextMatrix(ActRow, 1) = flxParts.TextMatrix(flxParts.Row, 5) 'QtyOnHand
        flxOrd.TextMatrix(ActRow, 5) = flxParts.TextMatrix(flxParts.Row, 3) 'Description
        flxOrd.TextMatrix(ActRow, 6) = ""   'Discount
        flxOrd.TextMatrix(ActRow, 9) = flxParts.TextMatrix(flxParts.Row, 6) 'Location
        flxOrd.TextMatrix(ActRow, 11) = flxParts.TextMatrix(flxParts.Row, 7) 'Applic Model
        flxOrd.TextMatrix(ActRow, 10) = flxParts.TextMatrix(flxParts.Row, 1) 'SPid
        flxOrd.TextMatrix(ActRow, 12) = flxParts.TextMatrix(flxParts.Row, 8) 'Max Discount
        flxOrd.TextMatrix(ActRow, 14) = flxParts.TextMatrix(flxParts.Row, 9) 'COST
        If Val(flxParts.TextMatrix(flxParts.Row, 9)) = 0 Then
            flxOrd.TextMatrix(ActRow, 7) = "" 'Price
            flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
        Else
            flxOrd.TextMatrix(ActRow, 7) = Format$(Val(flxParts.TextMatrix(flxParts.Row, 4)), "###,##0.00") 'Price
            flxOrd.TextMatrix(ActRow, 13) = Val(flxParts.TextMatrix(flxParts.Row, 4)) 'SRP
        End If
        'Qty and SRP >0
        If Val(flxOrd.TextMatrix(ActRow, 2)) > 0 Then
            TtlPrice = Val(flxOrd.TextMatrix(ActRow, 2)) * Val(.TextMatrix(.Row, 4))
            flxOrd.TextMatrix(ActRow, 8) = Format(TtlPrice, "#,###,##0.00")
        Else
            flxOrd.TextMatrix(ActRow, 8) = "0.00"
        End If
        Call ComputeTotal
        If InStr(flxOrd.TextMatrix(ActRow, 5), "ENGINE OIL") > 0 Or InStr(flxOrd.TextMatrix(ActRow, 5), "OIL 4T") > 0 Then
            flxOrd.TextMatrix(ActRow, 3) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Qrts.", "Qrt.")
        Else
            flxOrd.TextMatrix(ActRow, 3) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Pcs.", "Pc.")
        End If
    End With
End Sub

Private Sub flxParts_EnterCell()
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxOrd.TextMatrix(ActRow, 1) = flxParts.TextMatrix(flxParts.Row, 5) 'QtyOnHand
        flxOrd.TextMatrix(ActRow, 5) = flxParts.TextMatrix(flxParts.Row, 3) 'Description
        flxOrd.TextMatrix(ActRow, 6) = ""   'Discount
        flxOrd.TextMatrix(ActRow, 9) = flxParts.TextMatrix(flxParts.Row, 6) 'Location
        flxOrd.TextMatrix(ActRow, 11) = flxParts.TextMatrix(flxParts.Row, 7) 'Applic Model
        flxOrd.TextMatrix(ActRow, 10) = flxParts.TextMatrix(flxParts.Row, 1) 'SPid
        flxOrd.TextMatrix(ActRow, 12) = flxParts.TextMatrix(flxParts.Row, 8) 'Max Discount
        flxOrd.TextMatrix(ActRow, 14) = flxParts.TextMatrix(flxParts.Row, 9) 'COST
        If Val(flxParts.TextMatrix(flxParts.Row, 9)) = 0 Then
            flxOrd.TextMatrix(ActRow, 7) = "" 'Price
            flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
        Else
            flxOrd.TextMatrix(ActRow, 7) = Format$(Val(flxParts.TextMatrix(flxParts.Row, 4)), "###,##0.00") 'Price
            flxOrd.TextMatrix(ActRow, 13) = Val(flxParts.TextMatrix(flxParts.Row, 4)) 'SRP
        End If
        'Qty and SRP >0
        If Val(flxOrd.TextMatrix(ActRow, 2)) > 0 Then
            TtlPrice = Val(flxOrd.TextMatrix(ActRow, 2)) * Val(.TextMatrix(.Row, 4))
            flxOrd.TextMatrix(ActRow, 8) = Format(TtlPrice, "#,###,##0.00")
        Else
            flxOrd.TextMatrix(ActRow, 8) = "0.00"
        End If
        Call ComputeTotal
        If InStr(flxOrd.TextMatrix(ActRow, 5), "ENGINE OIL") > 0 Or InStr(flxOrd.TextMatrix(ActRow, 5), "OIL 4T") > 0 Then
            flxOrd.TextMatrix(ActRow, 3) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Qrts.", "Qrt.")
        Else
            flxOrd.TextMatrix(ActRow, 3) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Pcs.", "Pc.")
        End If
    End With
End Sub

Private Sub flxParts_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        Tmp = Split(txtPartNo.Tag, "|")
        For i = LBound(Tmp) To UBound(Tmp)
            flxOrd.TextMatrix(ActRow, i + 1) = Tmp(i)
        Next i
        txtPartNo.Visible = False
        flxParts.Visible = False
        flxOrd.SetFocus
        flxOrd.Col = 4
    End If

End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        'If all fields are filled then addrow focus on next entry
        If Val(flxOrd.TextMatrix(ActRow, 2)) > Val(flxOrd.TextMatrix(ActRow, 1)) And flxOrd.TextMatrix(ActRow, 2) <> "" Then
            MsgBox "Qty Ordered must not be greater than Qty Available!", vbExclamation + vbOKOnly, "Validation"
            flxOrd.Row = ActRow
            flxOrd.Col = 2
            Call flxOrd_Click
            Exit Sub
        End If
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" Then
            If ActRow = flxOrd.Rows - 1 Then flxOrd.AddItem ""
            flxOrd.SetFocus
            flxOrd.Row = flxOrd.Rows - 1
            flxOrd.Col = 4
        ElseIf flxOrd.TextMatrix(ActRow, 5) = "" Then   'No Description focus on Partnumber of current row
            flxOrd.SetFocus
            flxOrd.Col = 4
        ElseIf flxOrd.TextMatrix(ActRow, 2) = "" Then   'No Qty Focus on same column
            flxOrd.SetFocus
            flxOrd.Col = 2
        End If
        flxParts.Visible = False
    End If
End Sub

Function getFrontLiner() As String
    Dim rsFrontline As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT EMPLOYEES.FIRSTNAME, EMPLOYEES.EMPID From Employees " & _
          "WHERE (((EMPLOYEES.EMPID)=" & EmpId & "));"
    With rsFrontline
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            getFrontLiner = !FirstName
        Else
            getFrontLiner = ""
        End If
        .Close
    End With
End Function

Private Sub DisplayPendingOrders()
    Dim rsPending As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
          "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & EmpId & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#)) ORDER BY SPSALESORDER.ORDERNO;;"
    With rsPending
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxPND.Cols = 0
            Do While Not .EOF
                flxPND.Cols = flxPND.Cols + 1
                flxPND.ColAlignment(flxPND.Cols - 1) = 4
                flxPND.TextMatrix(0, flxPND.Cols - 1) = !ORDERNO
                .MoveNext
            Loop
        Else
            flxPND.Cols = 0
        End If
        .Close
    End With
End Sub

Private Sub flxPND_Click()
    Dim rsOrder As New ADODB.Recordset
    Dim rsOrdDtls As New ADODB.Recordset
    Dim Qry As String
    
    If UserType <> "Administrator" Then
        If flxPND.Cols = 0 Then Exit Sub
        If flxPND.TextMatrix(0, flxPND.Col) = "" Then Exit Sub
        
        Qry = "SELECT SPSALESORDER.* From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & flxPND.TextMatrix(0, flxPND.Col) & ") AND ((SPSALESORDER.ORDERDATE)=#" & Date & "#) AND ((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & EmpId & "));"
    Else
        Qry = "SELECT SPSALESORDER.* From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & Val(cboOrderNo.Text) & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#) AND ((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & txtEmpID.Text & "));"
    End If
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtOrderno.Text = !ORDERNO
            txtDate.Text = Format(!ORDERDATE, "mm/dd/yyyy")
            txtFrnt.Text = !Frontliner
            txtSPSorderID.Text = !SPSORDERID
            lblRemarks.Caption = IIf(IsNull(!Remarks), "", !Remarks)
            txtCustID.Text = IIf(IsNull(!CUSTOMERID), "", !CUSTOMERID)
            txtRefNo.Text = IIf(IsNull(!Reference), "", !Reference)
            txtInvNo.Text = IIf(IsNull(!INVOICENO), "", !INVOICENO)
            txtReqID.Text = IIf(IsNull(!ReqID), "", !ReqID)
            EmpId = rsOrder!EmpId
            txtSPSorderID.Text = !SPSORDERID
            If Val(txtCustID.Text) > 0 Then
                myRem = IIf(IsNull(!Remarks), "", !Remarks)
                If myRem = "CUSTOMER W/ DP" Then
                    lblRef.Caption = "O.R. Number"
                    lblCustomer.Caption = "Customer w/ DP"
                    chkDP.Value = 1
                    Call GetCustomerTerm
                Else
                    Call GetCustomerTerm
                    lblRef.Caption = "P.O. Number"
                    lblCustomer.Caption = "On-Acct. Customer"
                End If
            ElseIf Val(txtReqID.Text) > 0 Then
                lblRef.Caption = "Requisition No."
                txtRefNo.Text = GetReqInfo(txtReqID.Text, 2)
                txtCustName.Text = GetReqInfo(txtReqID.Text, 1)
            Else
                txtCustName.Text = ""
            End If
        Else
            txtOrderno.Text = ""
            'txtDate.Text = Format(Date, "mm/dd/yyyy")
            txtFrnt.Text = ""
            txtSPSorderID.Text = ""
            'lblRemarks.Caption = ""
            txtCustID.Text = ""
            lblRef.Caption = ""
            txtInvNo.Text = ""
            txtReqID.Text = ""
            txtPONO.Text = ""
            lblRef.Caption = ""
            'EmpId = 0
            Call DisplayPendingOrders
        End If
        .Close
    End With
    If txtSPSorderID.Text = "" Then Exit Sub
    Qry = "SELECT SPSORDERDETAIL.STATUS, SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.DISCOUNT, SPSORDERDETAIL.PRICE, [QTY]*[PRICE] AS TTLPRICE, SPAREPARTS.LOCATION, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.STOREPRICE, SPAREPARTS.COST " & _
          "FROM SPSORDERDETAIL INNER JOIN SPAREPARTS ON SPSORDERDETAIL.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    With rsOrdDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = 1
            Do While Not .EOF
                flxOrd.AddItem IIf(IsNull(!Status), "", !Status) & vbTab & !QtyOnHand & vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(!Discount > 0, !Discount & "%", "") & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlPrice, "###,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SparepartID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MaxDiscount), "", !MaxDiscount) & vbTab & !StorePrice & vbTab & !Cost
                If flxOrd.TextMatrix(flxOrd.Rows - 1, 0) = "FOUND" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 0) = "DROP" Then
                    flxOrd.Row = flxOrd.Rows - 1: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
                Else
                    flxOrd.TextMatrix(flxOrd.Rows - 1, 0) = ""
                End If
                .MoveNext
            Loop
            Call ComputeTotal
        Else
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        End If
        .Close
    End With
    
    cmdCancel.Visible = True
    'cmdCashier.Visible = True
    cmdWareHouse.Visible = True
    cmdAdvFind.Visible = False
    cmdReq.Visible = False
    cmdPO.Visible = False
End Sub

Function GetReqInfo(MyID As Long, myCode As Integer) As String
    Dim rsReqCust As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT REQUISITION.CUSTOMER, REQUISITION.REQNO From REQUISITION WHERE (((REQUISITION.REQID)=" & MyID & "));"
    With rsReqCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If myCode = 1 Then
                GetReqInfo = IIf(IsNull(!CUSTOMER), "", !CUSTOMER)
            Else
                GetReqInfo = IIf(IsNull(!REQNO), "", !REQNO)
            End If
        Else
            GetReqInfo = ""
        End If
        .Close
    End With
End Function


Private Sub GetCustomerTerm()
    Dim rsCust As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.OSBAL From CUSTOMER " & _
          "WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"

    With rsCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtCustName.Text = !CustomerName
            txtAddress.Text = !Address
            txtOSBal.Text = !OSbal
        End If
        .Close
    End With
    If chkDP.Value = 1 Then Exit Sub
    If Val(txtReqID.Text) > 0 And Val(txtCustID.Text) > 0 Then
        Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
                "WHERE (((ONACCTERMS.REQUISITIONID)=" & txtReqID.Text & ") AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
    Else
        Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
                "WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
    End If
    With rsTerm
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtTerms.Text = !Terms
            txtPayScheme.Text = !PayScheme
            txtNPays.Text = !Npays
            txtPayStart.Text = !PayStart
            txtBillAmt.Text = !BillingAmt
        End If
        .Close
    End With
    If txtTerms.Text = "" Then Call GetCustomerDefaultTerm

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPfind.MySource = "SPSORDER"
        SPfind.Show vbModal
    End If
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
End Sub


Private Sub Form_Load()
    Dim Qry As String
    
    ActRow = 0
    ActCol = 0
    If rsParts.State = adStateClosed Then Call LoadSpareParts
    txtDate.Text = Format(Date, "mm/dd/yyyy")
    If UserType = "Administrator" Then
        Call DisplayMainButton(2)
        flxPND.Visible = False
        Shape1.Visible = False
        Shape2.Visible = False
        fraAdmin.Visible = True
        Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER Where (((PCUSER.PartsMan) = True)) ORDER BY PCUSER.USERNAME;"
        With rsPartsMan
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                cboPartsMan.Clear
                Do While Not .EOF
                    cboPartsMan.AddItem !UserName
                    .MoveNext
                Loop
            End If
        End With
    Else
        Call DisplayPendingOrders
        If flxPND.Cols = 0 Then
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=25));"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                Call DisplayMainButton(2)
                Exit Sub
            End If
            Call DisplayMainButton(1)
            txtOrderno.Text = "<Auto>"
            'flxOrd.SetFocus
            flxOrd.Col = 4
            AddNew = True
        Else
            Call DisplayMainButton(2)
        End If
    End If
    EmpId = UserID
    txtDate.Text = Format(Date, "mm/dd/yyyy")
    MotoRep.sBar.Panels(6).Text = "Spare Parts Sales Order"
End Sub

Private Sub DisplayMainButton(myCode As Integer)
    If myCode = 1 Then
        cmdTrans(3).Visible = True
        cmdTrans(4).Visible = True
        cmdTrans(0).Visible = False
        cmdTrans(1).Visible = False
        cmdTrans(2).Visible = False
        Call LockData(False)
        cmdPO.Visible = True
        cmdReq.Visible = True
        cmdAdvFind.Visible = True
        cmdCancel.Visible = False
        'cmdCashier.Visible = False
        cmdWareHouse.Visible = False
    Else
        cmdTrans(3).Visible = False
        cmdTrans(4).Visible = False
        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = True
        cmdTrans(2).Visible = True
        Call LockData(True)
        cmdPO.Visible = False
        cmdReq.Visible = False
        cmdAdvFind.Visible = False
        cmdCancel.Visible = False
        'cmdCashier.Visible = False
        cmdWareHouse.Visible = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsParts.State = adStateOpen Then rsParts.Close
    If rsPartsMan.State = adStateOpen Then rsPartsMan.Close
    MotoRep.sBar.Panels(6).Text = ""
    MotoRep.tbMain.Visible = True
End Sub


Private Sub mnuAdd_Click()
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
'    ElseIf flxOrd.Row = 10 Then
'        MsgBox "Sorry, there are only 10 rows allowed per invoice!", vbCritical + vbOKOnly, "Add Row failed"
'        Exit Sub
    Else
        flxOrd.AddItem ""
        flxOrd.SetFocus
        flxOrd.Row = flxOrd.Rows - 1
        flxOrd.Col = 4
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        'If Not AddNew Then lstDelete.AddItem flxOrd.TextMatrix(ActRow, 9)
        If flxOrd.Rows = 2 Then
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        Else
            flxOrd.RemoveItem ActRow
        End If
        Call ComputeTotal
    End If
End Sub


Public Sub DisplaySalesOrder(MyTrans As String)
    Dim rsOrder As New ADODB.Recordset
    Dim rsODetail As New ADODB.Recordset
    Dim Qry As String, mySrc As String, MyUnit As String, myRem As String
    Dim TtlPrice As Single, OwnCode As Integer
    
    Call ClearEntries
        
    Qry = "SELECT SPSALESORDER.* From SPSALESORDER WHERE (((SPSALESORDER.SPSORDERID)=" & MyTrans & "));"
    
    rsOrder.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsOrder.RecordCount > 0 Then
        rsOrder.MoveFirst
        'txtOrderBy.Text = rsOrder!Customer
        'txtAddress.Text = rsOrder!Address
        txtOrderno.Text = rsOrder!ORDERNO
        txtDate.Text = Format(rsOrder!ORDERDATE, "mm/dd/yyyy")
        'lblRemarks.Caption = IIf(IsNull(rsOrder!REMARKS), "", rsOrder!REMARKS)
        txtTotal.Text = Format$(rsOrder!TotalPrice, "###,##0.00")
        txtFrnt.Text = rsOrder!Frontliner
        txtCustID.Text = IIf(IsNull(rsOrder!CUSTOMERID), "", rsOrder!CUSTOMERID)
        txtRefNo.Text = IIf(IsNull(rsOrder!Reference), "", rsOrder!Reference)
        txtInvNo.Text = IIf(IsNull(rsOrder!INVOICENO), "", rsOrder!INVOICENO)
        txtReqID.Text = IIf(IsNull(rsOrder!ReqID), "", rsOrder!ReqID)
        EmpId = rsOrder!EmpId
        txtSPSorderID.Text = rsOrder!SPSORDERID
        If txtCustID.Text <> "" Then
            myRem = IIf(IsNull(rsOrder!Remarks), "", rsOrder!Remarks)
            If myRem = "CUSTOMER W/ DP" Then
                lblRef.Caption = "O.R. Number"
                lblCustomer.Caption = "Customer w/ DP"
                chkDP.Value = 1
                Call GetCustomerTerm
            Else
                Call GetCustomerTerm
                lblRef.Caption = "P.O. Number"
                lblCustomer.Caption = "On-Acct. Customer"
            End If
        End If
    End If
    rsOrder.Close
    With rsODetail
        Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.DISCOUNT, SPSORDERDETAIL.PRICE, [PRICE]*[QTY] AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID " & _
              "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE (((SPSORDERDETAIL.SPSORDERID)=" & MyTrans & "));"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = 1
            Do While Not .EOF
                flxOrd.AddItem vbTab & !QtyOnHand & vbTab & !Qty & vbTab & IIf(IsNull(!Unit), "", !Unit) & vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(!Discount <> 0, Format$(!Discount, "##0.00") & "%", "") & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlPrice, "###,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & !SparepartID
                .MoveNext
            Loop
        End If
        .Close
    End With
    'If txtInvNo.Text = "" Then
    '    Call DisplayMainButton(2)
    'Else
    '    Call DisplayMainButton(2)
    'End If
End Sub


Private Sub SaveSalesOrder()
    Dim rsOrder As New ADODB.Recordset
    Dim rsODetails As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim i As Integer, myTtl As Single, Ans As Integer, NewOnAcct As Boolean, MyDisc As Currency
    Dim eMsg As String, Qry As String
                                  
    'Saving Table Spare Parts Invoice
    On Error GoTo LocalErr
    myTtl = Val(Replace(txtGrandTotal.Text, ",", ""))
    eMsg = "Saving SP Sales Order"
    If AddNew Then
        rsOrder.Open "SPSalesOrder", LYRcon, adOpenKeyset, adLockOptimistic
        rsOrder.AddNew
        txtOrderno.Text = GetOrderNo + 1
    Else
        Qry = "SELECT SPSALESORDER.* From SPSALESORDER WHERE (((SPSALESORDER.SPSORDERID)=" & txtSPSorderID.Text & "));"
        rsOrder.Open Qry, LYRcon, adOpenDynamic, adLockPessimistic
        rsOrder.MoveFirst
    End If
    rsOrder!ORDERNO = Val(txtOrderno.Text)
    rsOrder!ORDERDATE = txtDate.Text
    rsOrder!TotalPrice = myTtl
    rsOrder!LessDiscPrice = Val(Replace(txtLessDiscAmt.Text, ",", ""))
    rsOrder!GrossPrice = Val(Replace(txtTotal.Text, ",", ""))
    rsOrder!InDiscAmt = Val(txtTtlDisc.Text)
    If InStr(txtPctDisc.Text, "%") > 0 Then
        rsOrder!PctDisc = Val(Left$(txtPctDisc.Text, Len(txtPctDisc.Text) - 1))
    Else
        rsOrder!PctDisc = Val(txtPctDisc.Text)
    End If
    rsOrder!Frontliner = txtFrnt.Text
    rsOrder!EmpId = EmpId
    If Val(txtReqID.Text) > 0 Then
        rsOrder!ReqID = txtReqID.Text
        rsOrder!Remarks = lblRemarks.Caption
    End If
    If Val(txtCustID.Text) > 0 Then
        rsOrder!CUSTOMERID = txtCustID.Text
        rsOrder!Reference = txtRefNo.Text
        If chkDP.Value = 1 Then rsOrder!Remarks = "CUSTOMER W/ DP"
    ElseIf Val(txtCustID.Text) = 0 And Not AddNew Then
        rsOrder!CUSTOMERID = 0
        rsOrder!Reference = ""
        rsOrder!Remarks = ""
    End If
    If AddNew Then
        rsOrder.Update
        txtSPSorderID.Text = rsOrder!SPSORDERID
    Else
        rsOrder.MoveNext
    End If
    rsOrder.Close
    'Saving Table Spare Parts Order Details
    
    If Not AddNew Then
        LYRcon.Execute "DELETE SPSORDERDETAIL.*, SPSORDERDETAIL.SPSORDERID From SPSORDERDETAIL " & _
                       "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    End If
    With flxOrd
        For i = 1 To flxOrd.Rows - 1
            If Val(txtReqID.Text) <> 0 Then
                Qry = "SELECT SPSORDERDETAIL.SPSORDERID, SPSORDERDETAIL.SPAREPARTID From SPSORDERDETAIL " & _
                      "WHERE (((SPSORDERDETAIL.REQID)=" & Val(txtReqID.Text) & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & .TextMatrix(i, 10) & "));"
                If CtrlNumExist(Qry) Then
                    eMsg = "Line 1110: Updating Sales Order Details"
                    LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.SPSORDERID = " & txtSPSorderID.Text & ", SPSORDERDETAIL.QTY = " & .TextMatrix(i, 2) & ", SPSORDERDETAIL.PRICE = " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", SPSORDERDETAIL.STATUS = '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "' " & _
                                   "WHERE (((SPSORDERDETAIL.SPAREPARTID)=" & .TextMatrix(i, 10) & ") AND ((SPSORDERDETAIL.REQID)=" & txtReqID.Text & "));"
                Else
                    If InStr(.TextMatrix(i, 6), "%") > 0 Then
                        MyDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                    Else
                        MyDisc = Val(.TextMatrix(i, 6))
                    End If
                    eMsg = "Line 1119: Adding record to Sales Order Details"
                    LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, REQID, QTY, UNIT, PRICE, STATUS, DISCOUNT) " & _
                                   "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & Val(txtReqID.Text) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & MyDisc & ");"
                End If
            Else
                If InStr(.TextMatrix(i, 6), "%") > 0 Then
                    MyDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                Else
                    MyDisc = Val(.TextMatrix(i, 6))
                End If
                eMsg = "Line 1136: Adding New Sales Order Details @ row " & i
                LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, QTY, UNIT, PRICE, STATUS, DISCOUNT) " & _
                               "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & MyDisc & ");"
            End If
        Next i
    End With
    
    'Update Terms if Present
    If Val(txtCustID.Text) > 0 And chkDP.Value = 0 Then
        'Check if on-acct record exist
        If Val(txtReqID.Text) > 0 Then
            Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.REQUISITIONID)=" & txtReqID.Text & ") AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
        Else
            Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
        End If
        If CtrlNumExist(Qry) Then
            'Update terms
            NewOnAcct = False
            If Val(txtNPays.Text) > 0 Then
                txtBillAmt.Text = myTtl / Val(txtNPays.Text)
            End If
        Else
            NewOnAcct = True
        End If
        With rsTerm
            eMsg = "Line 1164: Saving Terms to On Account table"
            If NewOnAcct Then
                .Open "ONACCTERMS", LYRcon, adOpenKeyset, adLockOptimistic
                .AddNew
            Else
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                .MoveFirst
            End If
            !CUSTOMERID = txtCustID.Text
            !PayScheme = txtPayScheme.Text
            !Terms = Val(txtTerms.Text)
            !Npays = Val(txtNPays.Text)
            If IsDate(txtPayStart.Text) Then !PayStart = txtPayStart.Text
            'txtBillAmt.Text = Format$(PureNum(txtTotal.Text) / Val(txtNPays.Text), "####0.00")
            !BillingAmt = Val(Replace(txtBillAmt.Text, ",", ""))
            !SPSORDERID = txtSPSorderID.Text
            !Balance = myTtl
            !AcctAmount = myTtl
            !Status = "Temp"
            .Update
            .Close
        End With
    End If
    If txtCustID.Text = "" And Not AddNew Then 'remove order in table onaccount
        'Make sure On-acct is remove
        'LYRcon.Execute "DELETE ONACCTERMS.* From ONACCTERMS WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & "));"
        eMsg = "Line 1190: Updating On Account table - Removing Customer ID"
        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.CUSTOMERID = 0 WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & "));"
    End If
    Exit Sub
LocalErr:
    If eMsg = "Saving SP Sales Order" Or Err.Number = 3219 Then
        Call SaveFromError3219
    ElseIf InStr(eMsg, "Line 1136:") > 0 Then
        Call SaveDetailsOnly
    Else
        MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
    
End Sub

Private Sub SaveDetailsOnly()
    Dim i As Integer, MyDisc As Currency
    
    On Error GoTo LocalErr
    LYRcon.Execute "DELETE SPSORDERDETAIL.*, SPSORDERDETAIL.SPSORDERID From SPSORDERDETAIL " & _
                   "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    With flxOrd
        For i = 1 To .Rows - 1
            If InStr(.TextMatrix(i, 6), "%") > 0 Then
                MyDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
            Else
                MyDisc = Val(.TextMatrix(i, 6))
            End If
            LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, QTY, UNIT, PRICE, STATUS, DISCOUNT) " & _
                           "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & MyDisc & ");"
        Next i
    End With
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub SaveFromError3219()
    Dim rsOrder As New ADODB.Recordset
    Dim rsODetails As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String, myTtl As Single, Ans As Integer, NewOnAcct As Boolean
    Dim eMsg As String, eOrderID As Long, MyDisc As Currency, i As Integer
                                  
    'Saving Table Spare Parts Invoice
    On Error GoTo LocalErr
    myTtl = Val(Replace(txtGrandTotal.Text, ",", ""))
    eMsg = "Saving SP Sales Order"
    If Not AddNew Then
        eOrderID = Val(txtSPSorderID.Text)
        LYRcon.Execute "DELETE SPSALESORDER.* From SPSALESORDER WHERE (((SPSALESORDER.SPSORDERID)=" & eOrderID & "));"
        AddNew = True
    End If
        
    rsOrder.Open "SPSalesOrder", LYRcon, adOpenKeyset, adLockOptimistic
    rsOrder.AddNew
    txtOrderno.Text = GetOrderNo + 1
    rsOrder!ORDERNO = Val(txtOrderno.Text)
    rsOrder!ORDERDATE = txtDate.Text
    rsOrder!TotalPrice = myTtl
    rsOrder!LessDiscPrice = Val(Replace(txtLessDiscAmt.Text, ",", ""))
    rsOrder!GrossPrice = Val(Replace(txtTotal.Text, ",", ""))
    rsOrder!InDiscAmt = Val(txtTtlDisc.Text)
    If InStr(txtPctDisc.Text, "%") > 0 Then
        rsOrder!PctDisc = Val(Left$(txtPctDisc.Text, Len(txtPctDisc.Text) - 1))
    Else
        rsOrder!PctDisc = Val(txtPctDisc.Text)
    End If
    rsOrder!Frontliner = txtFrnt.Text
    rsOrder!EmpId = EmpId
    If Val(txtReqID.Text) > 0 Then
        rsOrder!ReqID = txtReqID.Text
        rsOrder!Remarks = lblRemarks.Caption
    End If
    If Val(txtCustID.Text) > 0 Then
        rsOrder!CUSTOMERID = txtCustID.Text
        rsOrder!Reference = txtRefNo.Text
        If chkDP.Value = 1 Then rsOrder!Remarks = "CUSTOMER W/ DP"
    ElseIf Val(txtCustID.Text) = 0 And Not AddNew Then
        rsOrder!CUSTOMERID = 0
        rsOrder!Reference = ""
        rsOrder!Remarks = ""
    End If
    rsOrder.Update
    txtSPSorderID.Text = rsOrder!SPSORDERID
    rsOrder.Close
    'Saving Table Spare Parts Order Details
    
    If Not AddNew Then
        LYRcon.Execute "DELETE SPSORDERDETAIL.*, SPSORDERDETAIL.SPSORDERID From SPSORDERDETAIL " & _
                       "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & eOrderID & "));"
    End If
    With flxOrd
        For i = 1 To flxOrd.Rows - 1
            If Val(txtReqID.Text) <> 0 Then
                Qry = "SELECT SPSORDERDETAIL.SPSORDERID, SPSORDERDETAIL.SPAREPARTID From SPSORDERDETAIL " & _
                      "WHERE (((SPSORDERDETAIL.REQID)=" & Val(txtReqID.Text) & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & .TextMatrix(i, 10) & "));"
                If CtrlNumExist(Qry) Then
                    eMsg = "Line 1110: Updating Sales Order Details"
                    LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.SPSORDERID = " & txtSPSorderID.Text & ", SPSORDERDETAIL.QTY = " & .TextMatrix(i, 2) & ", SPSORDERDETAIL.PRICE = " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", SPSORDERDETAIL.STATUS = '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "' " & _
                                   "WHERE (((SPSORDERDETAIL.SPAREPARTID)=" & .TextMatrix(i, 10) & ") AND ((SPSORDERDETAIL.REQID)=" & txtReqID.Text & "));"
                Else
                    If InStr(.TextMatrix(i, 6), "%") > 0 Then
                        MyDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                    Else
                        MyDisc = Val(.TextMatrix(i, 6))
                    End If
                    eMsg = "Line 1119: Adding record to Sales Order Details"
                    LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, REQID, QTY, UNIT, PRICE, STATUS, DISCOUNT) " & _
                                   "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & Val(txtReqID.Text) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & MyDisc & ");"
                End If
            Else
                If InStr(.TextMatrix(i, 6), "%") > 0 Then
                    MyDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                Else
                    MyDisc = Val(.TextMatrix(i, 6))
                End If
                eMsg = "Line 1136: Adding New Sales Order Details"
                LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, QTY, UNIT, PRICE, STATUS, DISCOUNT) " & _
                               "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & MyDisc & ");"
            End If
        Next i
    End With
    
    'Update Terms if Present
    If Val(txtCustID.Text) > 0 And chkDP.Value = 0 Then
        'Check if on-acct record exist
        If Val(txtReqID.Text) > 0 Then
            Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.REQUISITIONID)=" & txtReqID.Text & ") AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
        Else
            Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
        End If
        If CtrlNumExist(Qry) Then
            'Update terms
            NewOnAcct = False
            If Val(txtNPays.Text) > 0 Then
                txtBillAmt.Text = myTtl / Val(txtNPays.Text)
            End If
        Else
            NewOnAcct = True
        End If
        With rsTerm
            eMsg = "Line 1164: Saving Terms to On Account table"
            If NewOnAcct Then
                .Open "ONACCTERMS", LYRcon, adOpenKeyset, adLockOptimistic
                .AddNew
            Else
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                .MoveFirst
            End If
            !CUSTOMERID = txtCustID.Text
            !PayScheme = txtPayScheme.Text
            !Terms = Val(txtTerms.Text)
            !Npays = Val(txtNPays.Text)
            If IsDate(txtPayStart.Text) Then !PayStart = txtPayStart.Text
            'txtBillAmt.Text = Format$(PureNum(txtTotal.Text) / Val(txtNPays.Text), "####0.00")
            !BillingAmt = Val(Replace(txtBillAmt.Text, ",", ""))
            !SPSORDERID = txtSPSorderID.Text
            !Balance = myTtl
            !AcctAmount = myTtl
            !Status = "Temp"
            .Update
            .Close
        End With
    End If
    If txtCustID.Text = "" And Not AddNew Then 'remove order in table onaccount
        'Make sure On-acct is remove
        'LYRcon.Execute "DELETE ONACCTERMS.* From ONACCTERMS WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & "));"
        eMsg = "Line 1190: Updating On Account table - Removing Customer ID"
        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.CUSTOMERID = 0 WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & "));"
    End If
    Qry = "SELECT SPSALESORDER.SPSORDERID From SPSALESORDER WHERE (((SPSALESORDER.SPSORDERID)=" & eOrderID & "));"
    If CtrlNumExist(Qry) Then
        LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = 'CANCELLED' " & _
                       "WHERE (((SPSALESORDER.SPSORDERID)=" & eOrderID & "));"
    End If
    Exit Sub
    MsgBox "This is Order Number " & txtOrderno.Text, vbInformation + vbOKOnly, "New Order was made"
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Function GetOrderNo() As Integer
    Dim rsMyOrder As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT Max(SPSALESORDER.ORDERNO) AS MaxOfORDERNO From SPSALESORDER " & _
          "WHERE (((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#));"
    With rsMyOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetOrderNo = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            GetOrderNo = 0
        End If
        .Close
    End With
End Function

Function ValidateSalesOrder() As String
    Dim MyErr As String, Qry As String
    Dim i As Integer
    
    MyErr = ""
    If Not IsDate(txtDate.Text) Then MyErr = MyErr & vbCrLf & "Invalid entry on field Date"
    If txtFrnt.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field Frontliner"
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" And flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then
        If flxOrd.Rows = 2 Then
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        Else
            flxOrd.RemoveItem flxOrd.Rows - 1
        End If
    End If
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row of Ordered Spare Parts"
    For i = 1 To flxOrd.Rows - 1
        If Val(flxOrd.TextMatrix(i, 2)) = 0 Then MyErr = MyErr & vbCrLf & "Ordered Spare Parts @ Row " & i & " no Qty.!"
        If Val(flxOrd.TextMatrix(i, 7)) = 0 Then MyErr = MyErr & vbCrLf & "No Price specified @ Row " & i & "!"
    Next i
    ValidateSalesOrder = MyErr
End Function

Private Sub LoadSpareParts()
    Dim Qry As String
        

    Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    flxParts.Refresh
    With flxParts
        .Width = flxOrd.ColWidth(4)
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = flxOrd.ColWidth(4)
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
    End With
End Sub


Private Sub RXResizer1_ResizeComplete()
    
    On Error GoTo LocalErr
    With flxOrd
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.05
        .ColWidth(2) = (.Width - 300) * 0.08
        .ColWidth(3) = (.Width - 300) * 0.08
        .ColWidth(4) = (.Width - 300) * 0.16
        .ColWidth(5) = (.Width - 300) * 0.2
        .ColWidth(6) = (.Width - 300) * 0.1
        .ColWidth(7) = (.Width - 300) * 0.1
        .ColWidth(8) = (.Width - 300) * 0.1
        .ColWidth(9) = (.Width - 300) * 0.12
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .Col = 6: .Row = 1
        txtTotal.Left = flxOrd.Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7)
        txtTotal.Width = .ColWidth(8)
        flxParts.ColWidth(2) = .ColWidth(4)
        Label3.Left = txtTotal.Left - (Label3.Width + 50)
        Label3.Top = txtTotal.Top
        Label3.FontSize = txtTotal.FontSize
    End With
    txtLessDiscAmt.Left = txtTotal.Left
    txtLessDiscAmt.Width = txtTotal.Width
    txtLessDiscAmt.Top = txtTotal.Top + txtTotal.Height
    txtGrandTotal.Left = txtTotal.Left
    txtGrandTotal.Width = txtTotal.Width
    txtGrandTotal.Top = txtLessDiscAmt.Top + txtLessDiscAmt.Height
    lblDiscount.Left = txtTotal.Left - (lblDiscount.Width + 50)
    lblDiscount.Top = txtLessDiscAmt.Top
    txtPctDisc.Left = lblDiscount.Left - (txtPctDisc.Width + 50)
    txtPctDisc.Top = lblDiscount.Top
    lblLess.Left = txtPctDisc.Left - (lblLess.Width + 50)
    lblLess.Top = txtPctDisc.Top
    lblGrandTtl.Left = txtGrandTotal.Left - (lblGrandTtl.Width + 50)
    lblGrandTtl.Top = txtGrandTotal.Top
    
    If ActRow > 0 And ActCol > 0 Then
        flxOrd.Row = ActRow
        flxOrd.Col = ActCol
        txtPartNo.Top = flxOrd.CellTop + flxOrd.Top
        txtPartNo.Left = flxOrd.CellLeft + flxOrd.Left
        txtPartNo.Width = flxOrd.CellWidth
    End If
    With flxParts
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
    End With
'    lblHeading(2).Left = Frame1.Left + Frame1.Width - lblHeading(2).Width
'    lblHeading(3).Left = Frame1.Left + Frame1.Width - lblHeading(3).Width - 30
    lblHeading(4).Top = flxOrd.Top - lblHeading(4).Height + 30
    lblHeading(5).Top = flxOrd.Top - lblHeading(5).Height
    lblHeading(4).Left = flxOrd.Left
    lblHeading(5).Left = flxOrd.Left + 30
    flxPND.Left = cmdTrans(2).Left + cmdTrans(2).Width + 1000 'Me.ScaleWidth - flxPND.Width
    flxPND.Width = Me.ScaleWidth - flxPND.Left
    Shape2.Left = flxPND.Left
    Shape2.Width = flxPND.Width
    Shape1.Left = Shape2.Left - 240
    fraAdmin.Left = Frame1.Left + (Frame1.Width - fraAdmin.Width)
    Exit Sub
LocalErr:
    'Unload Me
End Sub

Private Sub tmrQtyTip_Timer()
    If tmrCount = 0 Then
        With txtQtyTip
            .Top = flxOrd.CellTop + flxOrd.Top + 290
            .Left = flxOrd.Left + flxOrd.ColWidth(0) + flxOrd.ColWidth(1) + flxOrd.ColWidth(2) + flxOrd.ColWidth(3) + flxOrd.ColWidth(4) + flxOrd.ColWidth(5) + flxOrd.ColWidth(6)
            .Text = "Cost: " & Format(GetCost(flxOrd.TextMatrix(ActRow, 10)), "#,##0.00")
            .Visible = True
            .ZOrder
        End With
        tmrCount = 1
        tmrQtyTip.Interval = 3000
    Else
        txtQtyTip.Visible = False
        tmrQtyTip.Enabled = False
        tmrCount = 0
        tmrQtyTip.Interval = 300
    End If
End Sub

Function GetCost(SPid As Long) As Single
    Dim rsCost As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT SPAREPARTS.COST From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & SPid & "));"
    With rsCost
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetCost = !Cost * 1.12
        Else
            GetCost = 0
        End If
        .Close
    End With

End Function

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPfind.MySource = "SPSORDER"
        SPfind.Show vbModal
    End If
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
End Sub

Private Sub txtDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtDiscount.Text = Empty
        flxOrd.TextMatrix(ActRow, 6) = IIf(txtDiscount.Tag = "", "", txtDiscount.Tag)
        txtDiscount.Tag = "skiplostfocus"
        flxOrd.SetFocus
        flxOrd.Col = 2
        txtDiscount.Visible = False
    End If
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    Dim StrValid As String
    
    StrValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(StrValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Call txtDiscount_LostFocus
        If txtDiscount.Tag = "Xcd100" Then txtDiscount.Tag = "": Exit Sub
        'If all fields are filled then addrow focus on next entry
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" Then
            If ActRow = flxOrd.Rows - 1 Then flxOrd.AddItem ""
            flxOrd.SetFocus
            flxOrd.Row = flxOrd.Rows - 1
            flxOrd.Col = 4
        ElseIf flxOrd.TextMatrix(ActRow, 5) = "" Then   'No Description focus on Partnumber of current row
            flxOrd.SetFocus
            flxOrd.Col = 4
        ElseIf flxOrd.TextMatrix(ActRow, 2) = "" Then   'No Qty Focus on same column
            flxOrd.SetFocus
            flxOrd.Col = 2
        End If
    End If
End Sub

Private Sub txtDiscount_LostFocus()
    Dim DiscAmt As Currency, DiscPrice As Currency
    
    If txtDiscount.Tag = "skiplostfocus" Then
        txtDiscount.Tag = ""
        Exit Sub
    End If
    With flxOrd
        If Val(txtDiscount.Text) > Val(flxOrd.TextMatrix(ActRow, 12)) Then
            MsgBox "Discount must not exceed " & flxOrd.TextMatrix(ActRow, 12) & "%", vbCritical + vbOKOnly, "Validation"
            txtDiscount.Tag = "Xcd100"
            txtDiscount.Text = flxOrd.TextMatrix(ActRow, 12)
            Exit Sub
        End If
        If Val(txtDiscount.Text) > 0 Then
            DiscAmt = Val(.TextMatrix(ActRow, 13)) * (Val(txtDiscount.Text) / 100)
            DiscPrice = Val(.TextMatrix(ActRow, 13)) - DiscAmt 'Discounted
            .TextMatrix(ActRow, 7) = Format$(DiscPrice, "#,##0.00")
        End If
        .TextMatrix(ActRow, 3) = Replace(.TextMatrix(ActRow, 3), "*", "")
    End With
    flxOrd.TextMatrix(ActRow, 6) = IIf(Val(txtDiscount.Text) > 0, txtDiscount.Text & "%", "")
    Call ComputeTotal
    txtDiscount.Visible = False
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPfind.MySource = "SPSORDER"
        SPfind.Show vbModal
    End If
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
End Sub

Private Sub txtFlx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtFlx.Text = Empty
        flxOrd.TextMatrix(ActRow, 2) = IIf(txtFlx.Tag = "XcdQty", "", txtFlx.Tag)
        txtFlx.Tag = "skiplostfocus"
        flxOrd.SetFocus
        flxOrd.Col = 2
        txtFlx.Visible = False
    End If
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtOrderno.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        flxOrd.SetFocus
        flxOrd.Col = 4  'partnumber
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxOrd.SetFocus
        flxOrd.Col = 4  'partnumber
    End If
End Sub

Private Sub txtFlx_KeyPress(KeyAscii As Integer)
    Dim StrValid As String
    
    StrValid = "0123456789"
    If ActCol = 1 Then
        If KeyAscii > 26 Then
            If InStr(StrValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then
        Call txtFlx_LostFocus
        If txtFlx.Tag = "XcdQty" Then txtFlx.Tag = "": Exit Sub
        'If all fields are filled then addrow focus on next entry
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" Then
            If ActRow = flxOrd.Rows - 1 Then flxOrd.AddItem ""
            flxOrd.SetFocus
            flxOrd.Row = flxOrd.Rows - 1
            flxOrd.Col = 4
        ElseIf flxOrd.TextMatrix(ActRow, 5) = "" Then   'No Description focus on Partnumber of current row
            flxOrd.SetFocus
            flxOrd.Col = 4
        ElseIf flxOrd.TextMatrix(ActRow, 2) = "" Then   'No Qty Focus on same column
            flxOrd.SetFocus
            flxOrd.Col = 2
        End If
    End If
End Sub

Private Sub txtFlx_LostFocus()
    If txtFlx.Tag = "skiplostfocus" Then
        txtFlx.Tag = ""
        Exit Sub
    End If
    With flxOrd
        If Val(txtFlx.Text) > Val(.TextMatrix(ActRow, 1)) And .TextMatrix(ActRow, 10) <> "" Then
            MsgBox "Order Qty. must not exceed Qty. On-Hand (" & .TextMatrix(ActRow, 1) & ")!", vbCritical + vbOKOnly, "Validation"
            txtFlx.Tag = "XcdQty"
            txtFlx.SelStart = 0
            txtFlx.SelLength = Len(txtFlx.Text)
            txtFlx.SetFocus
            Exit Sub
        End If
    End With
    If InStr(flxOrd.TextMatrix(ActRow, 5), "ENGINE OIL") > 0 Or InStr(flxOrd.TextMatrix(ActRow, 5), "OIL 4T") > 0 Then
        flxOrd.TextMatrix(ActRow, 3) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Qrts.", "Qrt.")
    Else
        flxOrd.TextMatrix(ActRow, 3) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Pcs.", "Pc.")
    End If
    flxOrd.TextMatrix(ActRow, ActCol) = txtFlx.Text
    Call ComputeTotal
    txtFlx.Visible = False
End Sub

Public Sub ComputeTotal()
    Dim i As Integer
    Dim TtlPrice As Currency, GrndTtl As Currency, DiscAmt As Currency, TtlDiscAmt As Currency
    Dim PctD As Single, myPrice As Currency
    
    With flxOrd
        GrndTtl = 0
        TtlDiscAmt = 0
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 2)) > 0 And .TextMatrix(i, 7) <> "" Then
                TtlPrice = Val(.TextMatrix(i, 2)) * Val(Replace(.TextMatrix(i, 7), ",", ""))
                GrndTtl = GrndTtl + TtlPrice
                .TextMatrix(i, 8) = Format$(TtlPrice, "#,##0.00")
            End If
            If .TextMatrix(i, 6) <> "" Then
                If Right$(.TextMatrix(i, 6), 1) = "%" Then
                    PctD = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                Else
                    PctD = Val(.TextMatrix(i, 6))
                End If
                myPrice = Val(.TextMatrix(i, 13))
                DiscAmt = myPrice * (PctD / 100)
                TtlDiscAmt = TtlDiscAmt + DiscAmt
            End If
        Next i
        txtTtlDisc.Text = TtlDiscAmt
        txtTotal.Text = Format$(GrndTtl, "#,##0.00")
        txtGrandTotal.Text = Format$(GrndTtl - (Val(Replace(txtLessDiscAmt.Text, ",", ""))), "#,##0.00")
    End With
            
End Sub

Private Sub txtOrderno_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPfind.MySource = "SPSORDER"
        SPfind.Show vbModal
    End If
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
End Sub

Private Sub txtPartNo_Change()
    Dim TtlPrice As Currency, DiscPrice As Currency
    Dim PctDisc As Single
    With flxParts
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
        .Visible = True
        .ZOrder
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsParts.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsParts.EOF Then
            If Not PrtBln Then
                txtPartNo.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            .TopRow = rsParts.AbsolutePosition - 1
            .Row = rsParts.AbsolutePosition - 1
            .Col = 0: .ColSel = 8
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            flxOrd.TextMatrix(ActRow, 1) = flxParts.TextMatrix(flxParts.Row, 5) 'QtyOnHand
            flxOrd.TextMatrix(ActRow, 5) = flxParts.TextMatrix(flxParts.Row, 3) 'Description
            flxOrd.TextMatrix(ActRow, 6) = ""   'Discount
            flxOrd.TextMatrix(ActRow, 9) = flxParts.TextMatrix(flxParts.Row, 6) 'Location
            flxOrd.TextMatrix(ActRow, 11) = flxParts.TextMatrix(flxParts.Row, 7) 'Applic Model
            flxOrd.TextMatrix(ActRow, 10) = flxParts.TextMatrix(flxParts.Row, 1) 'SPid
            flxOrd.TextMatrix(ActRow, 12) = flxParts.TextMatrix(flxParts.Row, 8) 'Max Discount
            flxOrd.TextMatrix(ActRow, 14) = flxParts.TextMatrix(flxParts.Row, 9) 'COST
            If Val(flxParts.TextMatrix(flxParts.Row, 9)) = 0 Then
                flxOrd.TextMatrix(ActRow, 7) = "" 'Price
                flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
            Else
                flxOrd.TextMatrix(ActRow, 7) = Format$(Val(flxParts.TextMatrix(flxParts.Row, 4)), "###,##0.00") 'Price
                flxOrd.TextMatrix(ActRow, 13) = Val(flxParts.TextMatrix(flxParts.Row, 4)) 'SRP
            End If
            'Qty and SRP >0
            If Val(flxOrd.TextMatrix(ActRow, 2)) > 0 Then
                TtlPrice = Val(flxOrd.TextMatrix(ActRow, 2)) * Val(.TextMatrix(.Row, 4))
                flxOrd.TextMatrix(ActRow, 8) = Format(TtlPrice, "#,###,##0.00")
            Else
                flxOrd.TextMatrix(ActRow, 8) = "0.00"
            End If
            Call ComputeTotal
            If InStr(flxOrd.TextMatrix(ActRow, 5), "ENGINE OIL") > 0 Or InStr(flxOrd.TextMatrix(ActRow, 5), "OIL 4T") > 0 Then
                flxOrd.TextMatrix(ActRow, 3) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Qrts.", "Qrt.")
            Else
                flxOrd.TextMatrix(ActRow, 3) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Pcs.", "Pc.")
            End If
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
'            rsParts.Requery
            flxOrd.TextMatrix(ActRow, 1) = ""
            flxOrd.TextMatrix(ActRow, 5) = ""
            flxOrd.TextMatrix(ActRow, 6) = ""
            flxOrd.TextMatrix(ActRow, 7) = ""
            flxOrd.TextMatrix(ActRow, 8) = ""
            flxOrd.TextMatrix(ActRow, 9) = ""
            flxOrd.TextMatrix(ActRow, 10) = ""
            flxOrd.TextMatrix(ActRow, 11) = ""
            flxOrd.TextMatrix(ActRow, 12) = ""
            flxOrd.TextMatrix(ActRow, 13) = ""
            flxOrd.TextMatrix(ActRow, 14) = ""
            .Visible = False
            Call ComputeTotal
        End If
    End With
End Sub

Private Sub txtPartNo_GotFocus()
    With flxParts
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
    End With
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            txtPartNo.Tag = "skiplostfocus"
            flxParts.Visible = True
            flxParts.SetFocus
        Case vbKeyEscape
            txtPartNo.Text = Empty
            Tmp = Split(txtPartNo.Tag, "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxOrd.TextMatrix(ActRow, i + 1) = Tmp(i)
            Next i
            txtPartNo.Tag = "skiplostfocus"
            flxOrd.SetFocus
            flxOrd.Col = 4
            txtPartNo.Visible = False
            flxParts.Visible = False
        Case vbKeyDelete
            PrtSrch = True
            rsParts.MoveFirst
        Case vbKeyBack
            PrtSrch = True
            rsParts.MoveFirst
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxOrd.SetFocus
        'Call txtPartNo_LostFocus
        'If all fields are filled then addrow focus on next entry
        If Val(flxOrd.TextMatrix(ActRow, 2)) > Val(flxOrd.TextMatrix(ActRow, 1)) And flxOrd.TextMatrix(ActRow, 2) <> "" Then
            MsgBox "Qty Ordered must not be greater than Qty Available!", vbExclamation + vbOKOnly, "Validation"
            flxOrd.Row = ActRow
            flxOrd.Col = 2
            Call flxOrd_Click
            Exit Sub
        End If
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" Then
            If ActRow = flxOrd.Rows - 1 Then flxOrd.AddItem ""
            flxOrd.SetFocus
            flxOrd.Row = flxOrd.Rows - 1
            flxOrd.Col = 4
        ElseIf flxOrd.TextMatrix(ActRow, 5) = "" Then   'No Description focus on Partnumber of current row
            flxOrd.SetFocus
            flxOrd.Col = 4
        ElseIf flxOrd.TextMatrix(ActRow, 2) = "" Then   'No Qty Focus on same column
            flxOrd.SetFocus
            flxOrd.Col = 2
        End If
        flxParts.Visible = False
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    Dim i As Integer
    
    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    If flxOrd.TextMatrix(ActRow, 10) = "" Or txtPartNo.Text = "" Then
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
        txtPartNo.Visible = False
        flxParts.Visible = False
        For i = 1 To 14
            flxOrd.TextMatrix(ActRow, i) = ""
        Next i
        Exit Sub
    End If
    flxOrd.TextMatrix(ActRow, 4) = txtPartNo.Text
    txtPartNo.Visible = False
    flxParts.Visible = False
End Sub

Private Sub txtOrderNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtAddress.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        txtDate.SetFocus
    End If
End Sub

Private Sub txtOrderNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDate.SetFocus
End Sub

Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtPrice.Text = Empty
        flxOrd.TextMatrix(ActRow, 7) = IIf(txtPrice.Tag = "XcdQty", "", txtPrice.Tag)
        txtPrice.Tag = "skiplostfocus"
        flxOrd.SetFocus
        flxOrd.Col = 7
        txtPrice.Visible = False
    End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    Dim StrValid As String
    
    StrValid = "0123456789.X"
    If KeyAscii > 26 Then
        If InStr(StrValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If Chr(KeyAscii) = "X" Then
        tmrQtyTip.Enabled = True
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxOrd.SetFocus
        'Call txtPrice_LostFocus
        If txtPrice.Tag = "XcdQty" Then txtPrice.Tag = "": Exit Sub
        'If all fields are filled then addrow focus on next entry
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" Then
            If ActRow = flxOrd.Rows - 1 Then flxOrd.AddItem ""
            flxOrd.SetFocus
            flxOrd.Row = flxOrd.Rows - 1
            flxOrd.Col = 4
        ElseIf flxOrd.TextMatrix(ActRow, 5) = "" Then   'No Description focus on Partnumber of current row
            flxOrd.SetFocus
            flxOrd.Col = 4
        ElseIf flxOrd.TextMatrix(ActRow, 2) = "" Then   'No Qty Focus on same column
            flxOrd.SetFocus
            flxOrd.Col = 2
        End If
    End If
End Sub


Private Sub txtPrice_LostFocus()
    Dim AmtDsc As Currency, LessAmt As Currency, PctD As Double
    
    If txtPrice.Tag = "skiplostfocus" Then
        txtPrice.Tag = ""
        Exit Sub
    End If
    With flxOrd
        'Validate if cost=0 no price will be displayed
        
        If Val(.TextMatrix(ActRow, 14)) = 0 Then
            MsgBox "No available COST for this item." & vbCrLf & "Please inform your Manager.", vbOKOnly + vbExclamation, "Zero Cost"
            flxOrd.TextMatrix(ActRow, 7) = ""
            txtPrice.Visible = False
            Call ComputeTotal
            Exit Sub
        Else
            AmtDsc = Val(.TextMatrix(ActRow, 14)) * 1.12
            If Val(Replace(txtPrice.Text, ",", "")) < Format(AmtDsc, "###0.00") Then
                MsgBox "Invalid Price entered!" & vbCrLf & "Please follow the correct Price", vbOKOnly + vbExclamation, "Invalid Price"
                flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 13), "#,##0.00")
                flxOrd.TextMatrix(ActRow, 6) = ""
                txtPrice.Visible = False
                Call ComputeTotal
                Exit Sub
            End If
        End If
        LessAmt = Val(.TextMatrix(ActRow, 13)) * (Val(.TextMatrix(ActRow, 12)) / 100)
        AmtDsc = Val(.TextMatrix(ActRow, 13)) - LessAmt
        If UserType <> "Administrator" Then
            If Val(Replace(txtPrice.Text, ",", "")) < AmtDsc Then
                MsgBox "Price must not be less than maximum discounted Price: " & Format(AmtDsc, "#,##0.00") & "!", vbCritical + vbOKOnly, "Validation"
                txtPrice.Tag = "XcdQty"
                txtPrice.SelStart = 0
                txtPrice.SelLength = Len(txtPrice.Text)
                txtPrice.SetFocus
                Exit Sub
            End If
        End If
        If Val(Replace(txtPrice.Text, ",", "")) < Val(.TextMatrix(ActRow, 13)) Then
            PctD = (Val(.TextMatrix(ActRow, 13)) - Val(Replace(txtPrice.Text, ",", ""))) / Val(.TextMatrix(ActRow, 13))
            If Val(Replace(txtPrice.Text, ",", "")) > 0 Then .TextMatrix(ActRow, 6) = Format(PctD, "#0.000%")
        End If
        If Val(Replace(txtPrice.Text, ",", "")) < AmtDsc Then 'Price is less than allowable discount
            .TextMatrix(ActRow, 3) = IIf(InStr(.TextMatrix(ActRow, 3), "*") > 0, "", .TextMatrix(ActRow, 3) & "*")
        Else
            .TextMatrix(ActRow, 3) = Replace(.TextMatrix(ActRow, 3), "*", "")
        End If
    End With
    flxOrd.TextMatrix(ActRow, 7) = Format(txtPrice.Text, "#,##0.00")
    Call ComputeTotal
    txtPrice.Visible = False

End Sub

