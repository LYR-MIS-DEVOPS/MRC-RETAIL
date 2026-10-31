VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form SPSorder 
   Caption         =   "Spare Parts Sales Order"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   ControlBox      =   0   'False
   Icon            =   "SPSorder.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   11505
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   1920
      TabIndex        =   97
      Tag             =   "XYF"
      Top             =   6525
      Width           =   4335
   End
   Begin VB.TextBox txtTotPen 
      Height          =   285
      Left            =   2400
      TabIndex        =   96
      Top             =   9780
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   240
      TabIndex        =   93
      Tag             =   "XYF"
      Top             =   7170
      Width           =   11055
      Begin MSFlexGridLib.MSFlexGrid flxOrder 
         Height          =   555
         Left            =   60
         TabIndex        =   94
         Tag             =   "XYF"
         Top             =   150
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   979
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         ForeColor       =   255
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         ScrollBars      =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin lvButton.lvButtons_H cmdRemoveWS 
      Height          =   375
      Left            =   6240
      TabIndex        =   91
      Tag             =   "XYF"
      ToolTipText     =   "Remove Wholesale Customer"
      Top             =   6180
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
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
      Image           =   "SPSorder.frx":08CA
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtVat 
      Height          =   285
      Left            =   2760
      TabIndex        =   90
      Top             =   9390
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid flxSPec 
      Height          =   615
      Left            =   2040
      TabIndex        =   89
      Top             =   10350
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.CheckBox chkWholesaleCash 
      Caption         =   "Check1"
      Height          =   195
      Left            =   1320
      TabIndex        =   88
      Top             =   10110
      Visible         =   0   'False
      Width           =   1455
   End
   Begin TabDlg.SSTab tabWS 
      Height          =   4455
      Left            =   5850
      TabIndex        =   81
      Top             =   8700
      Visible         =   0   'False
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPSorder.frx":0E64
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdWSCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdWSOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboWS"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OptCash"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OptOnAcct"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.OptionButton OptOnAcct 
         Caption         =   "ON-ACCOUNT"
         Height          =   255
         Left            =   360
         TabIndex        =   87
         Top             =   3510
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptCash 
         Caption         =   "CASH"
         Height          =   255
         Left            =   2040
         TabIndex        =   86
         Top             =   3480
         Width           =   975
      End
      Begin VB.ComboBox cboWS 
         Height          =   2910
         Left            =   120
         Style           =   1  'Simple Combo
         TabIndex        =   82
         Top             =   480
         Width           =   3615
      End
      Begin lvButton.lvButtons_H cmdWSOK 
         Height          =   375
         Left            =   240
         TabIndex        =   83
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
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
         Image           =   "SPSorder.frx":0E80
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdWSCancel 
         Height          =   375
         Left            =   2040
         TabIndex        =   84
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
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
         Image           =   "SPSorder.frx":141A
         cBack           =   -2147483633
      End
      Begin VB.Label Label6 
         Caption         =   "Select Wholesale:"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxWDisc 
      Height          =   735
      Left            =   8760
      TabIndex        =   73
      Top             =   10110
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox txtCusType 
      Height          =   285
      Left            =   360
      TabIndex        =   72
      Top             =   10350
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4560
      Top             =   1440
   End
   Begin VB.TextBox txtCPOid 
      Height          =   285
      Left            =   5280
      TabIndex        =   71
      Top             =   9750
      Visible         =   0   'False
      Width           =   735
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
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   69
      Top             =   8760
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CheckBox chkDP 
      Caption         =   "DP"
      Height          =   255
      Left            =   360
      TabIndex        =   68
      Top             =   9990
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstDelete 
      Height          =   645
      ItemData        =   "SPSorder.frx":19B4
      Left            =   8760
      List            =   "SPSorder.frx":19B6
      TabIndex        =   67
      Top             =   9390
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtJOid 
      Height          =   285
      Left            =   1800
      TabIndex        =   66
      Top             =   9390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEmpID 
      Height          =   285
      Left            =   10920
      TabIndex        =   64
      Top             =   8550
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame fraAdmin 
      Height          =   735
      Left            =   5520
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cboOrderNo 
         Height          =   315
         Left            =   4200
         TabIndex        =   63
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cboPartsMan 
         Height          =   315
         Left            =   960
         TabIndex        =   61
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Order No."
         Height          =   255
         Left            =   3360
         TabIndex        =   62
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Partsman:"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.TextBox txtOSBal 
      Height          =   285
      Left            =   6480
      TabIndex        =   58
      Top             =   10950
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   5040
      TabIndex        =   57
      Top             =   10950
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPONO 
      Height          =   285
      Left            =   7200
      TabIndex        =   56
      Top             =   8550
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtBillAmt 
      Height          =   285
      Left            =   4200
      TabIndex        =   55
      Top             =   10950
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPayStart 
      Height          =   285
      Left            =   3360
      TabIndex        =   54
      Top             =   10950
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtNPays 
      Height          =   285
      Left            =   2520
      TabIndex        =   53
      Top             =   10950
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPayScheme 
      Height          =   285
      Left            =   1560
      TabIndex        =   52
      Top             =   10950
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTerms 
      Height          =   285
      Left            =   480
      TabIndex        =   51
      Top             =   10950
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
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   49
      Tag             =   "XYF"
      Top             =   6210
      Width           =   4335
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   360
      TabIndex        =   46
      Top             =   9390
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
      TabIndex        =   45
      Top             =   9030
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
      TabIndex        =   44
      Text            =   "No Discount"
      Top             =   8910
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
      TabIndex        =   41
      Tag             =   "XYF"
      Top             =   6465
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
      TabIndex        =   40
      Tag             =   "XYF"
      Top             =   6180
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
      TabIndex        =   39
      Tag             =   "XYF"
      Top             =   5895
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
      TabIndex        =   36
      Tag             =   "XYF"
      Text            =   "0%"
      Top             =   6210
      Width           =   495
   End
   Begin VB.TextBox txtTtlDisc 
      Height          =   285
      Left            =   9840
      TabIndex        =   34
      Top             =   8550
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8760
      TabIndex        =   33
      Top             =   8550
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   8040
      TabIndex        =   32
      Top             =   8550
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
      Height          =   315
      Left            =   1920
      TabIndex        =   31
      Tag             =   "XYF"
      Top             =   5895
      Width           =   1695
   End
   Begin VB.TextBox txtInvNo 
      Height          =   285
      Left            =   5760
      TabIndex        =   29
      Top             =   8550
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid flxPND 
      Height          =   600
      Left            =   5400
      TabIndex        =   28
      Top             =   150
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1058
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
      Left            =   8175
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
      Image           =   "SPSorder.frx":19B8
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   9240
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtSPSorderID 
      Height          =   285
      Left            =   480
      TabIndex        =   25
      Top             =   8550
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   720
      Left            =   240
      TabIndex        =   23
      Tag             =   "XYF"
      Top             =   7800
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
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   240
         Width           =   4335
      End
      Begin lvButton.lvButtons_H cmdAdvFind 
         Height          =   495
         Left            =   10440
         TabIndex        =   74
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
         Image           =   "SPSorder.frx":1F52
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCashier 
         Height          =   495
         Left            =   7080
         TabIndex        =   75
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
         Image           =   "SPSorder.frx":24EC
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdWareHouse 
         Height          =   495
         Left            =   8200
         TabIndex        =   76
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
         Image           =   "SPSorder.frx":2BC2
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   7640
         TabIndex        =   77
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
         Image           =   "SPSorder.frx":315C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdReq 
         Height          =   495
         Left            =   9880
         TabIndex        =   78
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
         Image           =   "SPSorder.frx":4036
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPO 
         Height          =   495
         Left            =   9315
         TabIndex        =   79
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
         Image           =   "SPSorder.frx":45D0
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdWholesale 
         Height          =   495
         Left            =   8760
         TabIndex        =   80
         Tag             =   "XYF"
         ToolTipText     =   "[F10] Insert Wholesale"
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
         Image           =   "SPSorder.frx":4B6A
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
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   292
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPrintCopy 
      Height          =   285
      Left            =   4920
      TabIndex        =   22
      Top             =   8550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   8790
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   9270
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   8790
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   9270
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   1560
      TabIndex        =   19
      Top             =   8550
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCust 
      Height          =   1815
      Left            =   4440
      TabIndex        =   18
      Top             =   8790
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
      Top             =   8910
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
      Top             =   8550
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
      Left            =   2760
      Top             =   8910
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   600
      TabIndex        =   14
      Text            =   "Qty. On Hand: 0"
      Top             =   9030
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtFlx 
      Height          =   315
      Left            =   3720
      TabIndex        =   13
      Top             =   8550
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAcctNo 
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Top             =   9750
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   3945
      Left            =   240
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   1800
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6959
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   0
      FormatString    =   $"SPSorder.frx":4F04
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
         TabIndex        =   65
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
            Picture         =   "SPSorder.frx":4F95
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorder.frx":5E6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorder.frx":6D49
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorder.frx":7C23
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorder.frx":8AFD
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
      Image           =   "SPSorder.frx":99D7
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
      Image           =   "SPSorder.frx":9F71
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
      Image           =   "SPSorder.frx":A50B
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
      Image           =   "SPSorder.frx":AAA5
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
      Image           =   "SPSorder.frx":B03F
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxBO 
      Height          =   1095
      Left            =   7560
      TabIndex        =   70
      Top             =   10830
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
   Begin lvButton.lvButtons_H cmdExcel 
      Height          =   315
      Left            =   7200
      TabIndex        =   99
      Tag             =   "XYF"
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "Excel"
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
      Image           =   "SPSorder.frx":B5D9
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog dlgDEMO 
      Left            =   7200
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label8 
      Caption         =   "Customer Address"
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
      TabIndex        =   98
      Tag             =   "XYF"
      Top             =   6585
      Width           =   1455
   End
   Begin VB.Label lblSPNumber 
      Height          =   315
      Left            =   240
      TabIndex        =   95
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label lblTerms 
      Height          =   255
      Left            =   1920
      TabIndex        =   92
      Tag             =   "XYF"
      Top             =   6840
      Width           =   4815
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
      TabIndex        =   50
      Tag             =   "XYF"
      Top             =   5970
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
      TabIndex        =   48
      Tag             =   "XYF"
      Top             =   6240
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
      TabIndex        =   47
      Tag             =   "XYF"
      Top             =   5925
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
      TabIndex        =   38
      Tag             =   "XYF"
      Top             =   6495
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
      TabIndex        =   37
      Tag             =   "XYF"
      Top             =   6210
      Width           =   900
   End
   Begin VB.Label lblLess 
      AutoSize        =   -1  'True
      Caption         =   "Less:  "
      Height          =   195
      Left            =   7440
      TabIndex        =   35
      Tag             =   "XYF"
      Top             =   6225
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
      TabIndex        =   30
      Top             =   6090
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   10800
      Top             =   8910
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11745
      m_MDI_Height    =   9180
      OriginalBorder  =   10
      OriginalCaption =   31
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
      Top             =   5940
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
Attribute VB_Name = "SPSorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset

Dim PrtSrch As Boolean, ActSrch As Boolean, PrtBln As Boolean, ProperExit As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim tmrCount As Integer, BtnIndx As Integer
Public AddNew As Boolean
Dim PmanSrch As Boolean
Public JustScroll As Boolean, StatSrch As Boolean
Dim fRow As Integer, fCol As Integer
Public EmpId As Integer

Private Sub ClearEntries()
    txtOrderNo.Text = "<Auto>"
    'txtDate.Text = Format(Date, "##/##/####")
    lblRemarks.Caption = ""
    lblRef.Caption = "Reference No."
    lblCustomer.Caption = "Customer"
    txtTotal.Text = ""
    txtSPSorderID.Text = ""
    txtInvNo.Text = ""
    txtReqID.Text = ""
    txtPOno.Text = ""
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
    cboAddress.Text = ""
    txtGrandTotal.Text = ""
    txtLessDiscAmt.Text = ""
    flxOrd.Rows = 1
    flxOrd.Rows = 2
    txtPrintCopy = ""
    txtCusType.Text = ""
    lblTerms.Caption = ""
    cmdRemoveWS.Visible = False
    chkDP.Value = 0
End Sub
                                                                
Private Sub LockData(Bln As Boolean)
    If UserType = "Administrator" Then
        fraAdmin.Visible = Bln
    End If
    JustScroll = Bln
    cboAddress.Locked = Bln
    txtCustName.Locked = Bln
End Sub
   
Private Sub cboAddress_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If
    With cboAddress
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            'txtSrcID.Tag = ""
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            'txtSrcID.Tag = "invalid"
        End If
    End With

End Sub

Private Sub cboAddress_Click()
    Call cboAddress_Change
End Sub

Private Sub cboAddress_GotFocus()
   Dim Tmp

   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboAddress.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub cboOrderNo_Change()
    Call flxPND_Click
End Sub

Private Sub cboOrderNo_Click()
    Call flxPND_Click
End Sub

Private Sub cboOrderNo_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
End Sub

Private Sub cboPartsMan_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboPartsMan
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            txtEmpID.Text = cboPartsMan.ItemData(i)
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
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
End Sub

Private Sub cboPartsMan_LostFocus()
    Dim rsPmanID As New ADODB.Recordset
    Dim Qry As String
    
    If txtEmpID.Text <> "" Then
        'Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & txtEmpID.Text & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#)) ORDER BY ORDERNO;"
        'mySQL
         Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & txtEmpID.Text & ") AND ((SPSALESORDER.ORDERDATE)='" & Format(txtDate.Text, "yyyy-mm-dd") & "')) ORDER BY ORDERNO;"
       
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

Private Sub cboWS_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboWS
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            'txtCustID.Text = .ItemData(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PmanSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustID.Text = ""
        End If
    End With
End Sub

Private Sub cboWS_Click()
    Call cboWS_Change
End Sub

Private Sub cboWS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PmanSrch = True
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
    Dim Ans As Integer, eMsg As String
    
    On Error GoTo LocalErr
    If txtOrderNo.Text = "" Or Not IsNumeric(txtOrderNo.Text) Then Exit Sub
    Ans = MsgBox("Do you want to Cancel Order No. " & txtOrderNo.Text & "?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    eMsg = "Updating Sales Order for Cancellation"
    'LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = 'CANCELLED' " & _
                   "WHERE (((SPSALESORDER.SPSORDERID)=" & txtSPSorderID.Text & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = 'CANCELLED' " & _
                   "WHERE SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ";"
    
    eMsg = "Updating Sales Order Details for Cancellation"
    'LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.SPSORDERID = 0 WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.SPSORDERID = 0 WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & ";"
    
    MsgBox "Order successfully Cancelled!", vbInformation + vbOKOnly, "Cancel"
    Call cboPartsMan_Change
    Call ClearEntries
    If UserType <> "Administrator" Then Call DisplayPendingOrders
    Exit Sub
LocalErr:
    MsgBox "Error Updating table SP Sales Order during Cancel" & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub


Private Sub cmdExcel_Click()
    Dim XLapp As Excel.Application
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet
    Dim sSaveAsFileName
    Dim xRow As Integer
    
    If Val(txtOrderNo.Text) = 0 Then Exit Sub
    With dlgDEMO
        .CancelError = True
        .InitDir = mstrLastDir
        .Flags = cdlOFNHideReadOnly
        .filename = Day(Date) & UCase$(Left$(MonthName(Month(Date)), 3)) & Year(Date) & "ORDER" & txtOrderNo.Text
        .Filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
        .ShowOpen
        sSaveAsFileName = .filename
    End With
    'cdlCancel &H&H7FF3 & Cancel was selected

    Set XLapp = CreateObject("Excel.Application")
    XLapp.Visible = True
    XLapp.Workbooks.Add
    Set wbkMVR = XLapp.ActiveWorkbook
    Set wksLineA1 = wbkMVR.Sheets(1)

    'Show Customer
    '1:PART NUMBER, 2:DESCRIPTION, 3:CLASSIFICATION 4:DATE SOLD, 5:SALES INV#., 6:QTY., 7:COST W/O VAT, 8:UNIT PRICE, 9:SUPPLIER, 10:TRANS. TYPE,
    '11:CUSTOMER/DEALER, 12:ADDRESS, 13:CITY/TOWN, 14:PROVINCE, 15:REGION, 16:BRANCHES, 17:CUSTOMER TYPE, 18:FRONTLINE 18:CUST.ID

    wksLineA1.Name = "Sales Order"
    wksLineA1.Cells(1, 1).Value = "SALES ORDER NO. " & txtOrderNo.Text
    wksLineA1.Cells(2, 1).Value = "SALES DATE: " & txtDate.Text
    wksLineA1.Cells(3, 1).Value = "QTY."
    wksLineA1.Cells(3, 2).Value = "UNIT"
    wksLineA1.Cells(3, 3).Value = "PARTNUMBER"
    wksLineA1.Cells(3, 4).Value = "DESCRIPTION"
    wksLineA1.Cells(3, 5).Value = "DISCOUNT"
    wksLineA1.Cells(3, 6).Value = "PRICE"
    wksLineA1.Cells(3, 7).Value = "TOTAL"
    With flxOrd
        'pBar.Value = 0
        'pBar.Max = .Rows - 1
        xRow = 4
        For i = 1 To .Rows - 1
            For j = 2 To 8
                wksLineA1.Cells(i + 3, j - 1).Value = .TextMatrix(i, j)
            Next j
            'pBar.Value = pBar.Value + 1
            DoEvents
            xRow = xRow + 1
        Next i
        wksLineA1.Cells(xRow, 6).Value = "TOTAL:"
        wksLineA1.Cells(xRow, 7).Value = txtTotal.Text
        xRow = xRow + 1
        wksLineA1.Cells(xRow, 1).Value = "CUSTOMER: " & txtCustName.Text
        xRow = xRow + 1
        wksLineA1.Cells(xRow, 1).Value = "ADDRESS: " & cboAddress.Text
    End With
    XLapp.DisplayAlerts = False
    'MyQry = GetFieldValue("SELECT pcuser.Password From pcuser Where pcuser.EmpId = 4;")
    wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False

    wbkMVR.Close SaveChanges:=False
    XLapp.Quit
 
    Set XLapp = Nothing
    Set wbkMVR = Nothing
    Set wksLineA1 = Nothing
    MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, "Sales DB"

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
    If rsParts.State = adStateClosed Then Timer1.Enabled = True
    InsertPO.Show vbModal
End Sub

Private Sub cmdRemoveWS_Click()
    Dim Ans As Integer, i As Integer
    
    If flxOrd.Rows = 2 And flxOrd.TextMatrix(1, 10) = "" Then
        'No Amount to Change Remove without prompting user
        txtCustName.Text = ""
        txtCustID.Text = ""
        txtCusType.Text = ""
        cboAddress.Text = ""
        cmdRemoveWS.Visible = False
        Exit Sub
    Else
        Ans = MsgBox("Do you want to remove Wholesale customer " & txtCustName.Text & "?" & vbCrLf & _
            "The prices of all items ordered will change to original SRP!", vbQuestion + vbYesNo, "Remove Wholesale")
        If Ans = vbNo Then Exit Sub
    End If
    'txtRefNo.Text = ""
    'txtCustName.Text = ""
    'lblRemarks.Caption = ""
    txtCustID.Text = ""
    txtCusType.Text = ""
    cmdRemoveWS.Visible = False
    txtPayScheme.Text = ""
    txtTerms.Text = ""
    txtNPays.Text = ""
    txtBillAmt.Text = ""
    txtPayStart.Text = ""
    
    With flxOrd
        For i = 1 To .Rows - 1
            .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 13)), "#,##0.00")
            .TextMatrix(i, 6) = ""
        Next i
    End With
    Call ComputeTotal
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
    'Qry = "SELECT REQUISITION.REQNO, REQUISITION.JOBNO, REQUISITION.CUSTOMER, REQUISITION.REQID, REQUISITION.JOBORDERID, REQUISITION.CUSTOMERID " & _
          "From Requisition WHERE (((REQUISITION.REQNO)='" & MyReq & "'));"
    'mySQL
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
            txtCustName.Text = !customer
            lblCustomer.Caption = "Service Customer"
            txtJOid.Text = !JOBORDERID
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
    If Val(txtCustID.Text) > 0 Then
        cboAddress.Text = GetFieldValue("SELECT CONCAT(customer.CITYTOWN, ', ', customer.PROVINCE) AS ADDRESS From customer Where customer.CUSTOMERID = " & Val(txtCustID.Text) & ";")
    End If
    flxOrd.Rows = 1
    'Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, [PRICE]*[QTY] AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST " & _
          "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE (((SPSORDERDETAIL.REQID)=" & ReqID & "));"
    'mySQL
    
    '**Update 05/11/2010
    Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, SPSORDERDETAIL.PRICE*SPSORDERDETAIL.QTY AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST " & _
          "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE SPSORDERDETAIL.REQID=" & ReqID & ";"
    
    With rsReqDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = 1
            Do While Not .EOF
                flxOrd.AddItem vbTab & !QtyOnHand & vbTab & !Qty & vbTab & IIf(IsNull(!Unit), "", !Unit) & vbTab & !PartNumber & vbTab & !Description & vbTab & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlPrice, "###,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SPAREPARTID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & !MaxDiscount & vbTab & !Price & vbTab & !Cost
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
    If rsParts.State = adStateClosed Then Timer1.Enabled = True
    Call LockData(False)
End Sub

Private Sub GetCustomerDefaultTerm()
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String, Gtotal As Currency
    
    'Saving On Account Job Order to On Account Terms
    'Get Terms
    'Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.STATUS)='Default') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
    Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
          "WHERE ONACCTERMS.STATUS='Default' AND ONACCTERMS.CUSTOMERID=" & txtCustID.Text & ";"
    
    With rsTerm
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtTerms.Text = !TermS
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
    
    On Error GoTo LocalErr
    Select Case Index
        Case 0  'New
            Call ClearEntries
            Call LockData(False)
            'txtDate.Text = Date
            txtOrderNo.Text = "<Auto>"
            AddNew = True
            Call DisplayMainButton(1)
            cmdCancel.Visible = False
            'cmdCashier.Visible = False
            cmdWareHouse.Visible = False
            If rsParts.State = adStateOpen Then rsParts.Close
            Timer1.Enabled = True
            flxPND.Cols = 0
            flxOrd.SetFocus
            flxOrd.Col = 4
        Case 1  'Edit
            If txtSPSorderID.Text = "" Then
                MsgBox "No available Sales Order to edit", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If txtInvNo.Text <> "" Then
                MsgBox "Unable to edit Order# " & txtOrderNo.Text & " because an Invoice was already issued.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If InStr(lblRemarks.Caption, "CANCELLED") > 0 Then
                MsgBox "Unable to edit Sales Order " & txtOrderNo.Text & " because it has already been Cancelled.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If rsParts.State = adStateOpen Then rsParts.Close
            Timer1.Enabled = True
            flxPND.Cols = 0
            Call DisplayMainButton(1)
            AddNew = False
        Case 2  'Exit
            MotoRep.tbMain.Visible = True
            'ProperExit = True
            Unload Me
            'LogIn.Show
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
                'LogIn.Show
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
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
End Sub

Private Sub cmdWareHouse_Click()
    Dim i As Integer, vSkip As Boolean, eMsg As String
    
    'Validate items if all are found
    On Error GoTo LocalErr
    
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
    eMsg = "Updating Destination of Sales Order"
    'LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.DESTINATION = 'WAREHOUSE' " & _
                   "WHERE (((SPSALESORDER.SPSORDERID)=" & txtSPSorderID.Text & "));"
    LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.DESTINATION = 'WAREHOUSE' " & _
                   "WHERE SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ";"
    
    MsgBox "Sales Order can now be viewed in Warehouse!", vbInformation + vbOKOnly, "Updated in Warehouse"
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    
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
                flxBO.AddItem Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !REFERENCE & vbTab & !Order
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
                    xBal = xBal - rsRec!RECEIVE
                    'flxStock.AddItem Format(rsRec!TransDate, "mm/dd/yyyy") & vbTab & MyPO & vbTab & rsRec!Reference & vbTab & vbTab & rsRec!Receive & vbTab & xBal
                    rsRec.MoveNext
                Loop
            End If
            rsRec.Close
        Next i
    End With
    GetBackOrder = xBal
End Function

Function GetSalesQty(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String, StartDate As Date, EndDate As Date
    Dim MySales As Long, MyRet As Long
    
    StartDate = DateSerial(Year(DateAdd("m", -3, Date)), Month(DateAdd("m", -3, Date)), 1)
    EndDate = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), Day(DateAdd("d", -1, DateSerial(Year(Date), Month(Date), 1))))
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT Sum(SPITR.`OUT`) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>='" & Format(StartDate, "yyyy-mm-dd") & "' And (SPITR.TRANSDATE)<='" & Format(EndDate, "yyyy-mm-dd") & "') AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & MyID & "));"
    
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
    'Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>='" & Format(StartDate, "yyyy-mm-dd") & "' And (SPITR.TRANSDATE)<='" & Format(EndDate, "yyyy-mm-dd") & "'));"
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function

Private Sub cmdWholesale_Click()
    Dim rsWS As New ADODB.Recordset
    
    'Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID From CUSTOMER WHERE (((CUSTOMER.CTYPE)='WHOLESALE')) ORDER BY CUSTOMER.CUSTOMERNAME;"
    'mySQL
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID From CUSTOMER WHERE CUSTOMER.CTYPE='WHOLESALE' ORDER BY CUSTOMER.CUSTOMERNAME;"
    
    cboWS.Clear
    With rsWS
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboWS.AddItem !CustomerName
                cboWS.ItemData(cboWS.ListCount - 1) = !CUSTOMERID
                .MoveNext
            Loop
        End If
        .Close
    End With
    flxOrd.Enabled = False
    Frame2.Enabled = False
    cmdTrans(3).Enabled = False
    cmdTrans(4).Enabled = False
    txtRefNo.Locked = True
    With tabWS
        .Caption = ""
        .Top = (Me.Height - .Height) / 2
        .Left = (Me.Width - .Width) / 2
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub cmdWSCancel_Click()
    flxOrd.Enabled = True
    Frame2.Enabled = True
    tabWS.Visible = False
    tabWS.Top = 13000
    cmdTrans(3).Enabled = True
    cmdTrans(4).Enabled = True
End Sub

Private Sub cmdWSOK_Click()
    Dim rsDisc As New ADODB.Recordset
    Dim Qry As String, Ans As Integer
    
    If cboWS.ListIndex = -1 Then
        MsgBox "Please select a customer from the list!", vbExclamation + vbOKOnly, "Wholesale"
        Exit Sub
    End If
    txtCustName.Text = cboWS.List(cboWS.ListIndex)
    txtCustID.Text = cboWS.ItemData(cboWS.ListIndex)
    cboAddress.Text = GetFieldValue("SELECT CONCAT(customer.CITYTOWN, ', ', customer.PROVINCE) AS ADDRESS From customer Where customer.CUSTOMERID = " & Val(txtCustID.Text) & ";")
    txtCusType.Text = "WHOLESALE"
    lblCustomer.Caption = "Wholesale Cust."
    If OptCash.Value = True Then
        chkWholesaleCash.Value = 1
        txtPayScheme.Text = ""
        txtTerms.Text = ""
        txtNPays.Text = ""
        txtBillAmt.Text = ""
        txtPayStart.Text = ""
        lblTerms.Caption = ""
    Else
        chkWholesaleCash.Value = 0
    End If
    'Qry = "SELECT SETTINGS.* From SETTINGS " & _
          "WHERE (((SETTINGS.DESCRIPTION)='Wholesale Discount')) ORDER BY SETTINGS.VALUE;"
    'mySQL
    Qry = "SELECT SETTINGS.* From SETTINGS " & _
          "WHERE SETTINGS.DESCRIPTION='Wholesale Discount' ORDER BY SETTINGS.VALUE;"
    
    rsDisc.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsDisc.RecordCount > 0 Then
        rsDisc.MoveFirst
        flxWDisc.Rows = 0
        Do While Not rsDisc.EOF
            flxWDisc.AddItem rsDisc!Remarks & vbTab & rsDisc!Value
            rsDisc.MoveNext
        Loop
    Else
        flxWDisc.Rows = 0
    End If
    rsDisc.Close
    With rsDisc
        'Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='SP Disclusion')) ORDER BY SETTINGS.REMARKS;"
        'mySQL
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION='SP Disclusion' ORDER BY SETTINGS.REMARKS;"
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSPec.Rows = 0
            Do While Not .EOF
                flxSPec.AddItem !Remarks & vbTab & !Value
                .MoveNext
            Loop
        Else
            flxSPec.Rows = 0
        End If
        .Close
    End With
    cmdRemoveWS.Visible = True
    Call cmdWSCancel_Click
    If OptOnAcct.Value = True And txtPayScheme.Text = "" Then
        lblTerms.Caption = "No Terms Applied"
    ElseIf OptOnAcct.Value = True And txtPayScheme.Text <> "" Then
        Ans = MsgBox("Do you want to use existing " & lblTerms.Caption & "?", vbQuestion + vbYesNo, "Terms")
        If Ans = vbNo Then
            txtPayScheme.Text = ""
            txtTerms.Text = ""
            txtNPays.Text = ""
            txtBillAmt.Text = ""
            txtPayStart.Text = ""
            lblTerms.Caption = "No Terms Applied"
        End If
    End If
    If flxOrd.Rows = 2 And flxOrd.TextMatrix(1, 10) = "" Then
        Exit Sub
    Else
        Ans = MsgBox("Do you want the existing orders to avail the wholesale discount?", vbQuestion + vbYesNo, "Wholesale")
        If Ans = vbNo Then Exit Sub
        Call ComputeWholesaleDiscount
    End If
End Sub

Public Sub ComputeWholesaleDiscount()
    Dim PctDisc As Currency, DiscPrice As Currency, i As Integer, j As Integer, myVat As Single
    Dim myRem As String, Qry As String
    
    With flxOrd
        For j = 1 To .Rows - 1
            If .TextMatrix(j, 10) <> "" Then
                PctDisc = 0
                
                Qry = "SELECT SPAREPARTS.DLRPRICE From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(.TextMatrix(j, 10)) & "));"
                PctDisc = Val(GetFieldValue(Qry))

                'old code
                'For i = 0 To flxSPec.Rows - 1 'Get Pct Discount base on SP Exclusion list
                '    If Val(.TextMatrix(j, 10)) = Val(flxSPec.TextMatrix(i, 0)) Then
                '        PctDisc = Val(flxSPec.TextMatrix(i, 1))
                '        Exit For
                '    End If
                'Next i
                
                If PctDisc = 0 Then
                
                    rsParts.MoveFirst
                    rsParts.Find "[SPAREPARTID] = " + .TextMatrix(j, 10)
                    If Not rsParts.EOF Then
                        myRem = rsParts!Remarks
                    Else
                        myRem = ""
                    End If
                    If myRem <> "" Then
                        For i = 0 To flxWDisc.Rows - 1  'Get Pct Discount base on SP Type
                            If myRem = flxWDisc.TextMatrix(i, 0) Then
                                PctDisc = Val(flxWDisc.TextMatrix(i, 1))
                                Exit For
                            End If
                        Next i
                    End If
                    If PctDisc > 0 Then
                        myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
                        DiscPrice = (Val(.TextMatrix(j, 14)) * myVat) * PctDisc
                        If InStr(myRem, "GENUINE") > 0 Then
                            If DiscPrice < 10 Then
                                DiscPrice = DiscPrice + 3
                            ElseIf DiscPrice >= 10 And DiscPrice <= 100 Then
                                DiscPrice = DiscPrice + 2
                            Else
                                If InStr(Str(DiscPrice), ".") > 0 Then DiscPrice = DiscPrice + 1
                            End If
                            DiscPrice = Int(DiscPrice)
                        Else
                            If InStr(Str(DiscPrice), ".") > 0 Then
                                DiscPrice = DiscPrice + 1
                                DiscPrice = Int(DiscPrice)
                            End If
                        End If
                        flxOrd.TextMatrix(j, 7) = Format$(DiscPrice, "###,##0.00") 'Price
                    Else
                        flxOrd.TextMatrix(j, 7) = Format$(Val(.TextMatrix(j, 13)), "###,##0.00") 'Orig SRP
                    End If
                Else    'Display Amount from Exclusion List
                    flxOrd.TextMatrix(j, 7) = Format$(PctDisc, "###,##0.00") 'Price
                End If
            End If
        Next j
    End With
    Call ComputeTotal
End Sub

Private Sub flxOrd_Click()
    Dim MyDesc As String, PrvW As Integer, CrnW As Integer, MaxW As Long, MyBO As Long, MySQ As Long
    
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
                    txtNoDiscount.Text = "No Discount"
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
    Dim eMsg As String
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
    
    On Error GoTo LocalErr
    If ActCol = 0 And txtSPSorderID.Text <> "" Then
        If flxOrd.TextMatrix(ActRow, 0) = "" Then
            eMsg = "Updating Order Status [Item on Showroom]"
            flxOrd.Row = ActRow: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
            LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'FOUND' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(ActRow, 10) & "));"
                      
        Else
            eMsg = "Updating Order Status [Item not yet in Showroom]"
            flxOrd.TextMatrix(ActRow, 0) = ""
            LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'BLANK' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(ActRow, 10) & "));"
                       
        End If
    End If
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    
End Sub


Private Sub flxOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxOrd_Click
End Sub
                                                  
Private Sub flxOrd_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
End Sub

Private Sub flxOrd_LeaveCell()
    txtNoDiscount.Visible = False
    txtAppModel.Visible = False
    txtSQtyBck.Visible = False
End Sub

Private Sub flxOrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
    Dim PctDisc As Single, myVat As Single, i As Integer
    Dim DiscPrice As Currency, TtlPrice As Currency, Qry As String
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxOrd.TextMatrix(ActRow, 1) = .TextMatrix(.Row, 5) 'QtyOnHand
        flxOrd.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 3) 'Description
        flxOrd.TextMatrix(ActRow, 6) = ""   'Discount
        flxOrd.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 6) 'Location
        flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 7) 'Applic Model
        flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 1) 'SPid
        flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 8) 'Max Discount
        flxOrd.TextMatrix(ActRow, 14) = .TextMatrix(.Row, 9) 'COST
        If Val(.TextMatrix(.Row, 9)) = 0 Then
            flxOrd.TextMatrix(ActRow, 7) = "" 'Price
            flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
        Else
            If UCase$(txtCusType.Text) = "WHOLESALE" Or UCase$(txtCusType.Text) = "LGC MARKETING" Then
                PctDisc = 0
                
                Qry = "SELECT SPAREPARTS.DLRPRICE From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(.TextMatrix(.Row, 1)) & "));"
                PctDisc = Val(GetFieldValue(Qry))
                  
                'For i = 0 To flxSPec.Rows - 1 'Get Pct Discount base on SP Exclusion list
                '    If Val(.TextMatrix(.Row, 1)) = Val(flxSPec.TextMatrix(i, 0)) Then
                '        PctDisc = Val(flxSPec.TextMatrix(i, 1))
                '        Exit For
                '    End If
                'Next i
                
                If PctDisc = 0 Then
                    For i = 0 To flxWDisc.Rows - 1  'Get Pct Discount base on SP Type
                        If .TextMatrix(.Row, 10) = flxWDisc.TextMatrix(i, 0) Then
                            PctDisc = Val(flxWDisc.TextMatrix(i, 1))
                            Exit For
                        End If
                    Next i
                    If PctDisc > 0 Then
                        myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
                        DiscPrice = (Val(.TextMatrix(.Row, 9)) * myVat) * PctDisc
                        If InStr(.TextMatrix(.Row, 10), "GENUINE") > 0 Then
                            If DiscPrice < 10 Then
                                DiscPrice = DiscPrice + 3
                            ElseIf DiscPrice >= 10 And DiscPrice <= 100 Then
                                DiscPrice = DiscPrice + 2
                            Else
                                If InStr(Str(DiscPrice), ".") > 0 Then DiscPrice = DiscPrice + 1
                            End If
                            DiscPrice = Int(DiscPrice)
                        Else
                            If InStr(Str(DiscPrice), ".") > 0 Then
                                DiscPrice = DiscPrice + 1
                                DiscPrice = Int(DiscPrice)
                            End If
                        End If
                        flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Price
                    Else
                        flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                    End If
                Else    'Display Amount from Exclusion List
                    flxOrd.TextMatrix(ActRow, 7) = Format$(PctDisc, "###,##0.00") 'Price
                End If
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
            'ElseIf InStr(SysBranch, "PANABO") > 0 Then
            '    If Val(.TextMatrix(.Row, 4)) > 1 Then
            '        Qry = "SELECT settings.Value From settings Where settings.DESCRIPTION = 'MARK-UP PRICE';"
            '        PctDisc = Val(GetFieldValue(Qry))
            '        If PctDisc > 0 Then
            '            PctDisc = 1 + (PctDisc / 100)
            '            DiscPrice = Val(.TextMatrix(.Row, 4)) * PctDisc
            '            flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Pricegud
            ' '           flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4))   'DiscPrice 'SRP
            '        Else
            '            DiscPrice = Val(.TextMatrix(.Row, 4)) * 1.1
            '            flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Pricegud
            '            flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4))   'DiscPrice 'SRP
            '        End If
            '    End If
            Else
                flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
            End If
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
    Dim PctDisc As Single, myVat As Single, i As Integer
    Dim DiscPrice As Currency, TtlPrice As Currency, Qry As String
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxOrd.TextMatrix(ActRow, 1) = .TextMatrix(.Row, 5) 'QtyOnHand
        flxOrd.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 3) 'Description
        flxOrd.TextMatrix(ActRow, 6) = ""   'Discount
        flxOrd.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 6) 'Location
        flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 7) 'Applic Model
        flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 1) 'SPid
        flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 8) 'Max Discount
        flxOrd.TextMatrix(ActRow, 14) = .TextMatrix(.Row, 9) 'COST
        If Val(.TextMatrix(.Row, 9)) = 0 Then
            flxOrd.TextMatrix(ActRow, 7) = "" 'Price
            flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
        Else
            If UCase$(txtCusType.Text) = "WHOLESALE" Or UCase$(txtCusType.Text) = "LGC MARKETING" Then
                PctDisc = 0
                
                Qry = "SELECT SPAREPARTS.DLRPRICE From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(.TextMatrix(.Row, 1)) & "));"
                PctDisc = Val(GetFieldValue(Qry))
               
                'For i = 0 To flxSPec.Rows - 1 'Get Pct Discount base on SP Exclusion list
                '    If Val(.TextMatrix(.Row, 1)) = Val(flxSPec.TextMatrix(i, 0)) Then
                '        PctDisc = Val(flxSPec.TextMatrix(i, 1))
                '        Exit For
                '    End If
                'Next i
                If PctDisc = 0 Then
                    For i = 0 To flxWDisc.Rows - 1  'Get Pct Discount base on SP Type
                        If .TextMatrix(.Row, 10) = flxWDisc.TextMatrix(i, 0) Then
                            PctDisc = Val(flxWDisc.TextMatrix(i, 1))
                            Exit For
                        End If
                    Next i
                    If PctDisc > 0 Then
                        myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
                        DiscPrice = (Val(.TextMatrix(.Row, 9)) * myVat) * PctDisc
                        If InStr(.TextMatrix(.Row, 10), "GENUINE") > 0 Then
                            If DiscPrice < 10 Then
                                DiscPrice = DiscPrice + 3
                            ElseIf DiscPrice >= 10 And DiscPrice <= 100 Then
                                DiscPrice = DiscPrice + 2
                            Else
                                If InStr(Str(DiscPrice), ".") > 0 Then DiscPrice = DiscPrice + 1
                            End If
                            DiscPrice = Int(DiscPrice)
                        Else
                            If InStr(Str(DiscPrice), ".") > 0 Then
                                DiscPrice = DiscPrice + 1
                                DiscPrice = Int(DiscPrice)
                            End If
                        End If
                        flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Price
                    Else
                        flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                    End If
                Else    'Display Amount from Exclusion List
                    flxOrd.TextMatrix(ActRow, 7) = Format$(PctDisc, "###,##0.00") 'Price
                End If
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
            'ElseIf InStr(SysBranch, "PANABO") > 0 Then
            '    If Val(.TextMatrix(.Row, 4)) > 1 Then
            '        Qry = "SELECT settings.Value From settings Where settings.DESCRIPTION = 'MARK-UP PRICE';"
            '        PctDisc = Val(GetFieldValue(Qry))
            '        If PctDisc > 0 Then
            '            PctDisc = 1 + (PctDisc / 100)
            '            DiscPrice = Val(.TextMatrix(.Row, 4)) * PctDisc
            '            flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Pricegud
            '            flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4))       'DiscPrice 'SRP
            '        Else
            '            DiscPrice = Val(.TextMatrix(.Row, 4)) * 1.1
            '            flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Pricegud
            '            flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4))       'DiscPrice 'SRP
            '        End If
            '    End If
            Else
                flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
            End If
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
    
    'Qry = "SELECT EMPLOYEES.FIRSTNAME, EMPLOYEES.EMPID From Employees " & _
          "WHERE (((EMPLOYEES.EMPID)=" & EmpId & "));"
    'mySQL
    Qry = "SELECT EMPLOYEES.FIRSTNAME, EMPLOYEES.EMPID From Employees " & _
          "WHERE EMPLOYEES.EMPID=" & UserID & ";"
    
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
    
    'Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
          "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & EmpId & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#)) ORDER BY SPSALESORDER.ORDERNO;"
    'mySQL
    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
          "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & EmpId & ") AND ((SPSALESORDER.ORDERDATE)='" & Format(txtDate.Text, "yyyy-mm-dd") & "')) ORDER BY SPSALESORDER.ORDERNO;"
    
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
    Dim Qry As String, myRem As String
    
    Call ClearEntries
    If UserType <> "Administrator" Then
        If flxPND.Cols = 0 Then Exit Sub
        If flxPND.TextMatrix(0, flxPND.Col) = "" Then Exit Sub
        
        'Qry = "SELECT SPSALESORDER.* From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & flxPND.TextMatrix(0, flxPND.Col) & ") AND ((SPSALESORDER.ORDERDATE)=#" & Date & "#) AND ((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & EmpId & ")) ORDER BY SPSALESORDER.SPSORDERID DESC;"
        'mySQL
        Qry = "SELECT SPSALESORDER.* From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & flxPND.TextMatrix(0, flxPND.Col) & ") AND ((SPSALESORDER.ORDERDATE)='" & Format(Date, "yyyy-mm-dd") & "') AND ((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & EmpId & ")) ORDER BY SPSALESORDER.SPSORDERID DESC;"
        
    Else
        'Qry = "SELECT SPSALESORDER.* From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & Val(cboOrderNo.Text) & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#) AND ((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & txtEmpID.Text & ")) ORDER BY SPSALESORDER.SPSORDERID DESC;"
        'mySQL
        Qry = "SELECT SPSALESORDER.* From SPSALESORDER " & _
              "WHERE SPSALESORDER.ORDERNO=" & Val(cboOrderNo.Text) & " AND SPSALESORDER.ORDERDATE='" & Format(txtDate.Text, "yyyy-mm-dd") & "' AND ((SPSALESORDER.INVOICENO) Is Null) AND SPSALESORDER.EMPID=" & txtEmpID.Text & " ORDER BY SPSALESORDER.SPSORDERID DESC;"
        
    End If
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtOrderNo.Text = !ORDERNO
            txtDate.Text = Format(!ORDERDATE, "mm/dd/yyyy")
            txtFrnt.Text = !FRONTLINER
            txtSPSorderID.Text = !SPSORDERID
            txtCustID.Text = IIf(IsNull(!CUSTOMERID), "", !CUSTOMERID)
            txtRefNo.Text = IIf(IsNull(!REFERENCE), "", !REFERENCE)
            txtInvNo.Text = IIf(IsNull(!InvoiceNO), "", !InvoiceNO)
            txtReqID.Text = IIf(IsNull(!ReqID), "", !ReqID)
            txtCusType.Text = IIf(IsNull(!Ctype), "", !Ctype)
            txtCustName.Text = IIf(IsNull(!customer), "", !customer)
            cboAddress.Text = IIf(IsNull(!Address), "", !Address)
            EmpId = rsOrder!EmpId
            txtSPSorderID.Text = !SPSORDERID
            If Val(txtCustID.Text) > 0 Then
                txtPayScheme.Text = IIf(IsNull(!PayScheme), "", !PayScheme)
                txtBillAmt.Text = IIf(IsNull(!BILLAMT), "", Format(!BILLAMT, "#,##0.00"))
                txtTerms.Text = IIf(IsNull(!TermS), "", !TermS)
                txtNPays.Text = IIf(IsNull(!Npays), "", !Npays)
                txtPayStart.Text = IIf(IsNull(!PayStart), "", Format(!PayStart, "mm/dd/yyyy"))
                myRem = IIf(IsNull(!Remarks), "", !Remarks)
                If myRem = "CUSTOMER W/ DP" Then
                    lblRef.Caption = "O.R. Number"
                    lblCustomer.Caption = "Customer w/ DP"
                    lblRemarks.Caption = IIf(IsNull(!Remarks), "", !Remarks)
                    chkDP.Value = 1
                    Call GetCustomerTerm
                Else
                    Call GetCustomerTerm
                    lblRef.Caption = "P.O. Number"
                    If txtCusType.Text = "WHOLESALE" Then
                        lblCustomer.Caption = "Wholesale Cust."
                        If Not IsNull(!Remarks) Then
                            If !Remarks = "CASH" Then
                                OptCash.Value = True
                                OptOnAcct.Value = False
                            Else
                                OptCash.Value = False
                                OptOnAcct.Value = True
                                lblTerms.Caption = "Terms: " & txtTerms.Text & " Days @ " & Format(Val(txtBillAmt.Text), "#,##0.00") & " " & txtPayScheme.Text
                            End If
                        End If
                    Else
                        lblCustomer.Caption = "On-Acct. Customer"
                        lblRemarks.Caption = IIf(IsNull(!Remarks), "", !Remarks)
                    End If
                End If
            ElseIf Val(txtReqID.Text) > 0 Then
                lblRef.Caption = "Requisition No."
                txtRefNo.Text = GetReqInfo(txtReqID.Text, 2)
                txtCustName.Text = GetReqInfo(txtReqID.Text, 1)
            Else
                'txtCustName.Text = ""
                'cboAddress.Text = ""
                txtPayScheme.Text = ""
                txtBillAmt.Text = ""
                txtTerms.Text = ""
                txtNPays.Text = ""
                txtPayStart.Text = ""
            End If
        Else
            txtOrderNo.Text = ""
            'txtDate.Text = Format(Date, "mm/dd/yyyy")
            txtFrnt.Text = ""
            txtSPSorderID.Text = ""
            txtPayScheme.Text = ""
            txtBillAmt.Text = ""
            txtTerms.Text = ""
            txtNPays.Text = ""
            txtPayStart.Text = ""
            'lblRemarks.Caption = ""
            txtCustID.Text = ""
            lblRef.Caption = ""
            txtInvNo.Text = ""
            txtReqID.Text = ""
            txtPOno.Text = ""
            lblRef.Caption = ""
            'EmpId = 0
            Call DisplayPendingOrders
        End If
        .Close
    End With
    If txtSPSorderID.Text = "" Then Exit Sub
    'Qry = "SELECT SPSORDERDETAIL.STATUS, SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.DISCOUNT, SPSORDERDETAIL.PRICE, [QTY]*[PRICE] AS TTLPRICE, SPAREPARTS.LOCATION, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.STOREPRICE, SPAREPARTS.COST " & _
          "FROM SPSORDERDETAIL INNER JOIN SPAREPARTS ON SPSORDERDETAIL.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    'mySQL
    Qry = "SELECT SPSORDERDETAIL.STATUS, SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.DISCOUNT, SPSORDERDETAIL.PRICE, SPSORDERDETAIL.QTY*SPSORDERDETAIL.PRICE AS TTLPRICE, SPAREPARTS.LOCATION, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.STOREPRICE, SPAREPARTS.COST " & _
          "FROM SPSORDERDETAIL INNER JOIN SPAREPARTS ON SPSORDERDETAIL.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    
    With rsOrdDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = 1
            Do While Not .EOF
                flxOrd.AddItem IIf(IsNull(!Status), "", !Status) & vbTab & !QtyOnHand & vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(!DISCOUNT > 0, !DISCOUNT & "%", "") & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlPrice, "###,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SPAREPARTID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MaxDiscount), "", !MaxDiscount) & vbTab & !StorePrice & vbTab & !Cost
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
    cmdWholesale.Visible = False
End Sub

Function GetReqInfo(MyID As Long, MyCode As Integer) As String
    Dim rsReqCust As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT REQUISITION.CUSTOMER, REQUISITION.REQNO From REQUISITION WHERE (((REQUISITION.REQID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT REQUISITION.CUSTOMER, REQUISITION.REQNO From REQUISITION WHERE REQUISITION.REQID=" & MyID & ";"
    
    With rsReqCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If MyCode = 1 Then
                GetReqInfo = IIf(IsNull(!customer), "", !customer)
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
    Dim Qry As String
    
    'Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.OSBAL From CUSTOMER " & _
          "WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
    'mySQL
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.OSBAL From CUSTOMER " & _
          "WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
    
    With rsCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtCustName.Text = !CustomerName
            cboAddress.Text = IIf(cboAddress.Text <> "", cboAddress.Text, !Address)
            txtOSBal.Text = !OSbal
        End If
        .Close
    End With

    If txtTerms.Text = "" Then Call GetCustomerDefaultTerm
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
End Sub


Private Sub Form_Load()
    Dim rsPartsMan As New ADODB.Recordset
    Dim rsAdrs As New ADODB.Recordset
    Dim Qry As String
    
    ActRow = 0
    ActCol = 0
    JustScroll = True
    txtDate.Text = Format(Date, "mm/dd/yyyy")
    'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='VAT'));"
    'mySQL
    Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='VAT';"
    
    txtVat.Text = GetFieldValue(Qry)
    'If rsParts.State = adStateClosed Then Timer1.Enabled = True
    If UserType = "Administrator" Then
        Call DisplayMainButton(2)
        flxPND.Visible = False
        'Shape1.Visible = False
        Shape2.Visible = False
        fraAdmin.Visible = True
        'Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER Where (((PCUSER.PartsMan) = True)) ORDER BY PCUSER.USERNAME;"
        'mySQL
        'Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER Where PCUSER.PartsMan = True ORDER BY PCUSER.USERNAME;"
        Qry = "SELECT DISTINCT spsalesorder.FRONTLINER, spsalesorder.EmpId From spsalesorder Where spsalesorder.ORDERDATE = '" & Format(Date, "yyyy-mm-dd") & "' Order By spsalesorder.FRONTLINER;"
        With rsPartsMan
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                cboPartsMan.Clear
                Do While Not .EOF
                    cboPartsMan.AddItem !FRONTLINER
                    cboPartsMan.ItemData(cboPartsMan.ListCount - 1) = !EmpId
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
        Call DisplayPendingOrders
        If flxPND.Cols = 0 Then
            Call DisplayMainButton(1)
            txtOrderNo.Text = "<Auto>"
            'flxOrd.SetFocus
            flxOrd.Col = 4
            AddNew = True
        Else
            Call DisplayMainButton(2)
        End If
    End If
    txtFrnt.Text = getFrontLiner
    txtDate.Text = Format(Date, "mm/dd/yyyy")
    Qry = "SELECT concat(settings.VALUE,', ', settings.REMARKS) as Address From settings Where settings.DESCRIPTION LIKE 'REGION%' Order By settings.Value"
    With rsAdrs
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboAddress.AddItem !Address
                .MoveNext
            Loop
        End If
        .Close
    End With
    MotoRep.sBar.Panels(6).Text = "Spare Parts Sales Order"
End Sub

Private Sub DisplayMainButton(MyCode As Integer)
    If MyCode = 1 Then
        cmdTrans(3).Visible = True
        cmdTrans(4).Visible = True
        cmdTrans(0).Visible = False
        cmdTrans(1).Visible = False
        cmdTrans(2).Visible = False
        Call LockData(False)
        cmdPO.Visible = True
        cmdWholesale.Visible = True
        cmdReq.Visible = True
        cmdAdvFind.Visible = True
        cmdCancel.Visible = False
        'cmdCashier.Visible = False
        cmdWareHouse.Visible = False
        cmdExcel.Visible = False
    Else
        cmdTrans(3).Visible = False
        cmdTrans(4).Visible = False
        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = True
        cmdTrans(2).Visible = True
        Call LockData(True)
        cmdPO.Visible = False
        cmdWholesale.Visible = False
        cmdReq.Visible = False
        cmdAdvFind.Visible = False
        cmdCancel.Visible = False
        'cmdCashier.Visible = False
        cmdWareHouse.Visible = False
        cmdExcel.Visible = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ProperExit Then
        If rsParts.State = adStateOpen Then rsParts.Close
        ProperExit = False
    ElseIf JustScroll Then
        Call cmdTrans_Click(2)
        ProperExit = False
    Else
        Cancel = True
    End If
End Sub

Private Sub mnuAdd_Click()
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
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
        If Not AddNew Then lstDelete.AddItem flxOrd.TextMatrix(ActRow, 10)
        If flxOrd.Rows = 2 Then
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        Else
            flxOrd.RemoveItem ActRow
        End If
        Call ComputeTotal
    End If
End Sub
Private Sub SaveSalesOrder()
    Dim i As Integer, myTtl As Single, Ans As Integer, NewOnAcct As Boolean
    Dim eMsg As String, Qry As String, myDisc As Currency, My
    Dim ReqID As Long, myRem As String, CustID As Long, MyRef As String, MyPSkm As String, MyPStart As String
    Dim MyBillAmt As Double, MyTerm As Integer, MyNpay As Integer, Pctd As Integer, OrdersNo As Long
                                  
    'Saving Table Spare Parts Invoice
    On Error GoTo LocalErr
    myTtl = Val(Replace(txtGrandTotal.Text, ",", ""))
    If InStr(txtPctDisc.Text, "%") > 0 Then
        Pctd = Val(Left$(txtPctDisc.Text, Len(txtPctDisc.Text) - 1))
    Else
        Pctd = Val(txtPctDisc.Text)
    End If
    If Val(txtReqID.Text) > 0 Then
        ReqID = txtReqID.Text
        myRem = lblRemarks.Caption
    Else
        ReqID = 0
        myRem = ""
    End If
    
    If AddNew Then
        eMsg = "Saving New Record for SP Sales Order"
        'old txtOrderno.Text = GetOrderNo + 1
        'txtOrderno.Text = GetOrdersNo
        'LYRcon.Execute "INSERT INTO SPSalesOrder (ORDERNO, ORDERDATE, TotalPrice, LessDiscPrice, GrossPrice, InDiscAmt, Ctype, PctDisc, Frontliner, EmpId, ReqID, Remarks) " & _
        "VALUES (" & Val(txtOrderno.Text) & ", #" & txtDate.Text & "#, " & myTtl & ", " & Val(Replace(txtLessDiscAmt.Text, ",", "")) & ", " & Val(Replace(txtTotal.Text, ",", "")) & ", " & Val(txtTtlDisc.Text) & ", '" & txtCusType.Text & "', " & Pctd & ", '" & txtFrnt.Text & "', " & EmpId & ", " & ReqID & ", '" & myRem & "');"
        
        'mySQL
        'Update Max order on previous date
        Call UpdateNumOrder
        mystation = ComputerName
        mystation = Left$(mystation, Len(mystation) - 1)
        LYRcon.Execute "INSERT INTO SPSalesOrder (ORDERDATE, TotalPrice, LessDiscPrice, GrossPrice, InDiscAmt, Ctype, PctDisc, Frontliner, EmpId, ReqID, Remarks, WorkStation) " & _
        "VALUES ('" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & myTtl & ", " & Val(Replace(txtLessDiscAmt.Text, ",", "")) & ", " & Val(Replace(txtTotal.Text, ",", "")) & ", " & Val(txtTtlDisc.Text) & ", '" & txtCusType.Text & "', " & Pctd & ", '" & txtFrnt.Text & "', " & EmpId & ", " & ReqID & ", '" & myRem & "', '" & mystation & "');"
       
       'Generate Order number
       '-------------------------------------------------------------
        Qry = "SELECT MAX(spsalesorder.SPSORDERID) As CrtOrder " & _
                "From " & _
                    "spsalesorder " & _
                "Where " & _
                    "spsalesorder.ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "' AND " & _
                "spsalesorder.WorkStation = '" & mystation & "';"
            
            MyCrtOrder = Val(GetFieldValue(Qry))
            Qry = "SELECT settings.`VALUE`,settings.DESCRIPTION,settings.REMARKS,settings.UPDATEDBY," & _
                "settings.LASTUPDATE From settings Where settings.DESCRIPTION = 'SALES ORDER';"
            OrdersNo = MyCrtOrder - Val(GetFieldValue(Qry))
        LYRcon.Execute "UPDATE spsalesorder SET ORDERNO = " & OrdersNo & " Where spsalesorder.SPSORDERID = " & MyCrtOrder & ";"
        '-------------------------------------------------------------
        'Old Qry = "SELECT SPSALESORDER.SPSORDERID From SPSALESORDER " & _
              "WHERE (SPSALESORDER.ORDERNO=" & txtOrderno.Text & " AND SPSALESORDER.ORDERDATE ='" & Format(txtDate.Text, "yyyy-mm-dd") & "' SPSALESORDER.WORKSTATION='" & ComputerName & "');"
        txtOrderNo = OrdersNo
        txtSPSorderID.Text = MyCrtOrder
        
    Else
        eMsg = "Updating Record for SP Sales Order"
        'LYRcon.Execute "UPDATE SPSalesOrder SET ORDERNO = " & Val(txtOrderno.Text) & ", ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "', TotalPrice = " & myTtl & ", LessDiscPrice = " & Val(Replace(txtLessDiscAmt.Text, ",", "")) & ", GrossPrice = " & Val(Replace(txtTotal.Text, ",", "")) & ", InDiscAmt = " & Val(txtTtlDisc.Text) & ", Ctype = '" & txtCusType.Text & "', PctDisc = " & Pctd & ", Frontliner = '" & txtFrnt.Text & "', EmpId = " & EmpId & " " & _
        "WHERE SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ";"
        LYRcon.Execute "UPDATE SPSalesOrder SET ORDERNO = " & Val(txtOrderNo.Text) & ", ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "', TotalPrice = " & myTtl & ", LessDiscPrice = " & Val(Replace(txtLessDiscAmt.Text, ",", "")) & ", GrossPrice = " & Val(Replace(txtTotal.Text, ",", "")) & ", InDiscAmt = " & Val(txtTtlDisc.Text) & ", Ctype = '" & txtCusType.Text & "', PctDisc = " & Pctd & ", Frontliner = '" & txtFrnt.Text & "', EmpId = " & EmpId & ", " & _
        "CUSTOMER = '" & txtCustName.Text & "', ADDRESS = '" & cboAddress.Text & "' WHERE SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ";"
    End If
    If Not AddNew Then
         'Reset Records On-Account/Cash/Quick Service
         'Whole Sale Cash
        
        LYRcon.Execute "UPDATE SPSalesOrder SET  Remarks = '', CUSTOMERID = 0, PAYSCHEME = NULL, TERMS = 0, NPAYS = 0, PAYSTART = NULL, BILLAMT = 0 " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"
    
        'Sales Order On-Account
        
        LYRcon.Execute "UPDATE SPSalesOrder SET  ReqID = 0, Remarks = '', CUSTOMERID = 0, Reference = '', PayScheme = 0, BillAmt = 0, Terms = 0, Npays = 0, PayStart = NULL " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"

    End If
    
    If txtCusType.Text = "WHOLESALE" Then
        If OptCash.Value = True Then
            myRem = "CASH"
            CustID = txtCustID.Text
        Else
            myRem = "ONACCOUNT"
            CustID = Val(txtCustID.Text)
            MyRef = txtRefNo.Text
            MyPSkm = txtPayScheme.Text
            MyBillAmt = myTtl / Val(txtNPays.Text)
            MyTerm = Val(txtTerms.Text)
            MyNpay = Val(txtNPays.Text)
            If IsDate(txtPayStart.Text) Then
                MyPStart = txtPayStart.Text
            Else
                MyPStart = DateAdd("m", 1, DateValue(txtDate.Text))
            End If
        End If
    ElseIf Val(txtCustID.Text) > 0 Then
        CustID = txtCustID.Text
        MyPSkm = txtPayScheme.Text
        If Val(txtBillAmt.Text) > 0 Then
            MyBillAmt = Val(Replace(txtBillAmt.Text, ",", ""))
        Else
            MyBillAmt = myTtl / Val(txtNPays.Text)
        End If
        MyTerm = Val(txtTerms.Text)
        MyNpay = Val(txtNPays.Text)
        If IsDate(txtPayStart.Text) Then
            MyPStart = txtPayStart.Text
        Else
            MyPStart = DateAdd("m", 1, DateValue(txtDate.Text))
        End If
        MyRef = txtRefNo.Text
        If chkDP.Value = 1 Then
            myRem = "CUSTOMER W/ DP"
        End If
    End If
    If Val(txtCustID.Text) > 0 And OptCash.Value = True Then
        eMsg = "Updating Record for SP Sales Order Wholesale Cash"
        LYRcon.Execute "UPDATE SPSalesOrder SET  Remarks = '" & myRem & "', CUSTOMERID = " & CustID & ", PAYSCHEME = NULL, TERMS = 0, NPAYS = 0, PAYSTART = NULL, BILLAMT = 0 " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"
        
    ElseIf Val(txtCustID.Text) > 0 Then
        eMsg = "Updating Record for SP Sales Order On Account Customer"
        'LYRcon.Execute "UPDATE SPSalesOrder SET  ReqID = " & ReqID & ", Remarks = '" & myRem & "', CUSTOMERID = " & CustID & ", Reference = '" & MyRef & "', PayScheme = '" & MyPSkm & "', BillAmt = " & MyBillAmt & ", Terms = " & MyTerm & ", Npays = " & MyNpay & ", PayStart =" & IIf(IsDate(MyPStart), "#" & MyPStart & "#", Null) & " " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"
        'mySQL
        LYRcon.Execute "UPDATE SPSalesOrder SET  ReqID = " & ReqID & ", Remarks = '" & myRem & "', CUSTOMERID = " & CustID & ", Reference = '" & MyRef & "', PayScheme = '" & MyPSkm & "', BillAmt = " & MyBillAmt & ", Terms = " & MyTerm & ", Npays = " & MyNpay & ", PayStart =" & IIf(IsDate(MyPStart), "'" & Format(MyPStart, "yyyy-mm-dd") & "'", Null) & " " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"
    ElseIf txtCustName.Text <> "" Or cboAddress.Text <> "" Then
        LYRcon.Execute "UPDATE SPSalesOrder SET CUSTOMER = '" & txtCustName.Text & "', ADDRESS = '" & cboAddress.Text & "' WHERE SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ";"

    End If
    
    
    'Saving Table Spare Parts Order Details
    If Not AddNew Then
        'LYRcon.Execute "DELETE SPSORDERDETAIL.* From SPSORDERDETAIL " & _
                       "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
        'mySQL
         LYRcon.Execute "DELETE From SPSORDERDETAIL " & _
                       "WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & ";"
     
    ElseIf Val(txtReqID.Text) > 0 Then
        'LYRcon.Execute "DELETE SPSORDERDETAIL.* From SPSORDERDETAIL " & _
                       "WHERE (((SPSORDERDETAIL.REQID)=" & Val(txtReqID.Text) & "));"
        'mySQL
        LYRcon.Execute "DELETE From SPSORDERDETAIL " & _
                       "WHERE SPSORDERDETAIL.REQID=" & Val(txtReqID.Text) & ";"
    
    End If
    With flxOrd
        For i = 1 To flxOrd.Rows - 1
                If InStr(.TextMatrix(i, 6), "%") > 0 Then
                    myDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                Else
                    myDisc = Val(.TextMatrix(i, 6))
                End If
                eMsg = "Line 1136: Adding New Sales Order Details @ row " & i
                LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, QTY, UNIT, PRICE, STATUS, DISCOUNT, REQID) " & _
                               "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & myDisc & ", " & Val(txtReqID.Text) & ");"
        Next i
    End With
    
    Exit Sub
LocalErr:
    If eMsg = "Saving SP Sales Order" Or Err.Number = 3219 Then
        Call SaveFromError3219
    ElseIf InStr(eMsg, "Line 1136") > 0 Then
        Call SaveDetailsOnly
    Else
        MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub
Private Sub SaveDetailsOnly()
    Dim i As Integer, myDisc As Currency
    On Error GoTo LocalErr
    'LYRcon.Execute "DELETE SPSORDERDETAIL.*, SPSORDERDETAIL.SPSORDERID From SPSORDERDETAIL " & _
                   "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    'mySQL
    LYRcon.Execute "DELETE From SPSORDERDETAIL " & _
                   "WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & ";"
    
    With flxOrd
        For i = 1 To .Rows - 1
            If InStr(.TextMatrix(i, 6), "%") > 0 Then
                myDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
            Else
                myDisc = Val(.TextMatrix(i, 6))
            End If
            LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, QTY, UNIT, PRICE, STATUS, DISCOUNT, REQID) " & _
                           "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & myDisc & ", " & Val(txtReqID.Text) & ");"
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
    Dim eMsg As String, eOrderID As Long
                                  
    'Saving Table Spare Parts Invoice
    On Error GoTo LocalErr
    myTtl = Val(Replace(txtGrandTotal.Text, ",", ""))
    If Not AddNew Then
        AddNew = True
        eMsg = "Deleting Previous SP Sales Order ID " & txtSPSorderID.Text
        eOrderID = Val(txtSPSorderID.Text)
        'LYRcon.Execute "DELETE SPSALESORDER.* From SPSALESORDER WHERE (((SPSALESORDER.SPSORDERID)=" & eOrderID & "));"
        'mySQL
        LYRcon.Execute "DELETE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID=" & eOrderID & ";"
        
    End If
    eMsg = "Saving SP Sales Order RD2"
    
    'rsOrder.Open "SPSalesOrder", LYRcon, adOpenKeyset, adLockOptimistic
    'rsOrder.AddNew
    'txtOrderno.Text = GetOrderNo + 1
    'rsOrder!ORDERNO = Val(txtOrderno.Text)
    'rsOrder!ORDERDATE = txtDate.Text
    'rsOrder!TotalPrice = myTtl
    'rsOrder!LessDiscPrice = Val(Replace(txtLessDiscAmt.Text, ",", ""))
    'rsOrder!GrossPrice = Val(Replace(txtTotal.Text, ",", ""))
    'rsOrder!InDiscAmt = Val(txtTtlDisc.Text)
    'If InStr(txtPctDisc.Text, "%") > 0 Then
    '    rsOrder!PctDisc = Val(Left$(txtPctDisc.Text, Len(txtPctDisc.Text) - 1))
    'Else
    '    rsOrder!PctDisc = Val(txtPctDisc.Text)
    'End If
    'rsOrder!Frontliner = txtFrnt.Text
    'rsOrder!EmpId = EmpId
    'If Val(txtReqID.Text) > 0 Then
    '    rsOrder!ReqID = txtReqID.Text
    '    rsOrder!Remarks = lblRemarks.Caption
    'End If
    'If Val(txtCustID.Text) > 0 Then
    '    rsOrder!CUSTOMERID = txtCustID.Text
    '    rsOrder!PayScheme = txtPayScheme.Text
    '    rsOrder!BillAmt = Val(Replace(txtBillAmt.Text, ",", ""))
    '    rsOrder!Terms = Val(txtTerms.Text)
    '    rsOrder!Npays = Val(txtNPays.Text)
    '    If IsDate(txtPayStart.Text) Then rsOrder!PayStart = txtPayStart.Text
    '    rsOrder!Reference = txtRefNo.Text
    '    If chkDP.Value = 1 Then rsOrder!Remarks = "CUSTOMER W/ DP"
    'ElseIf Val(txtCustID.Text) = 0 And Not AddNew Then
    '    rsOrder!CUSTOMERID = 0
    '    rsOrder!PayScheme = ""
    '    rsOrder!Terms = 0
    '    rsOrder!Npays = 0
    '    rsOrder!Reference = ""
    '    rsOrder!Remarks = ""
    'End If
    'rsOrder.Update
    'txtSPSorderID.Text = rsOrder!SPSORDERID
    'rsOrder.Close
    
    'mySQl
    txtOrderNo.Text = GetOrderNo + 1
    LYRcon.Execute "INSERT INTO spsalesorder( ORDERNO, ORDERDATE, TOTALPRICE, LESSDISCPRICE, GROSSPRICE, INDISCAMT)" & _
        "VALUES( " & Val(txtOrderNo.Text) & ", '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & myTtl & ", " & _
        "" & Val(Replace(txtLessDiscAmt.Text, ",", "")) & ", " & Val(Replace(txtTotal.Text, ",", "")) & ", " & _
        "" & Val(txtTtlDisc.Text) & ");"
        
    If InStr(txtPctDisc.Text, "%") > 0 Then
        SPPctDisc = Val(Left$(txtPctDisc.Text, Len(txtPctDisc.Text) - 1))
    Else
        SPPctDisc = Val(txtPctDisc.Text)
    End If
    LYRcon.Execute "INSERT INTO spsalesorder( PCTDISC, FRONTLINER, EMPID)VALUES( " & SPPctDisc & ", '" & txtFrnt.Text & "', " & EmpId & ");"
    If Val(txtReqID.Text) > 0 Then
        LYRcon.Execute "INSERT INTO spsalesorder( REQID, REMARKS )VALUES( " & txtReqID.Text & ", '" & lblRemarks.Caption & "');"
    End If
    If Val(txtCustID.Text) > 0 Then
        
        LYRcon.Execute "INSERT INTO spsalesorder( CUSTOMERID, PAYSCHEME, BILLAMT, TERMS, NPAYS, REFERENCE)" & _
            "VALUES( " & txtCustID.Text & ", '" & txtPayScheme.Text & "', " & Val(Replace(txtBillAmt.Text, ",", "")) & ", " & _
            "" & Val(txtTerms.Text) & ", " & Val(txtNPays.Text) & ", '" & txtRefNo.Text & "');"
        
        If IsDate(txtPayStart.Text) Then LYRcon.Execute "INSERT INTO spsalesorder(PAYSTART)VALUES('" & Format(PayStart, "yyyy-mm-dd") & "');"
        If chkDP.Value = 1 Then LYRcon.Execute "INSERT INTO spsalesorder(REMARKS)VALUES('CUSTOMER W/ DP');"
    
    ElseIf Val(txtCustID.Text) = 0 And Not AddNew Then
        
        LYRcon.Execute "INSERT INTO spsalesorder(CUSTOMERID, PAYSCHEME, TERMS, NPAYS, REFERENCE, REMARKS)VALUES(0, '', 0, 0, '', '');"
    
    End If
    Qry = "SELECT spsalesorder.SPSORDERID From spsalesorder Where spsalesorder.ORDERNO = " & Val(txtOrderNo.Text) & " AND " & _
        "spsalesorder.ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "';"
    txtSPSorderID.Text = GetFieldValue(Qry)
    
    
    'Saving Table Spare Parts Order Details
    If Not AddNew Then
        'LYRcon.Execute "DELETE SPSORDERDETAIL.*, SPSORDERDETAIL.SPSORDERID From SPSORDERDETAIL " & _
                       "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & eOrderID & "));"
        'mySQL
        LYRcon.Execute "DELETE From SPSORDERDETAIL " & _
                       "WHERE SPSORDERDETAIL.SPSORDERID=" & eOrderID & ";"
        
    End If
    With flxOrd
        For i = 1 To flxOrd.Rows - 1
            If Val(txtReqID.Text) <> 0 Then
                'Qry = "SELECT SPSORDERDETAIL.SPSORDERID, SPSORDERDETAIL.SPAREPARTID From SPSORDERDETAIL " & _
                      "WHERE (((SPSORDERDETAIL.REQID)=" & Val(txtReqID.Text) & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & .TextMatrix(i, 10) & "));"
                'mySQL
                Qry = "SELECT SPSORDERDETAIL.SPSORDERID, SPSORDERDETAIL.SPAREPARTID From SPSORDERDETAIL " & _
                      "WHERE (((SPSORDERDETAIL.REQID)=" & Val(txtReqID.Text) & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & .TextMatrix(i, 10) & "));"
                              
                If CtrlNumExist(Qry) Then
                    eMsg = "Line 1110: Updating Sales Order Details"
                    LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.SPSORDERID = " & txtSPSorderID.Text & ", SPSORDERDETAIL.QTY = " & .TextMatrix(i, 2) & ", SPSORDERDETAIL.PRICE = " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", SPSORDERDETAIL.STATUS = '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "' " & _
                                   "WHERE SPSORDERDETAIL.SPAREPARTID=" & .TextMatrix(i, 10) & " AND SPSORDERDETAIL.REQID=" & txtReqID.Text & ";"
                Else
                    If InStr(.TextMatrix(i, 6), "%") > 0 Then
                        myDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                    Else
                        myDisc = Val(.TextMatrix(i, 6))
                    End If
                    eMsg = "Line 1119: Adding record to Sales Order Details"
                    LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, REQID, QTY, UNIT, PRICE, STATUS, DISCOUNT) " & _
                                   "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & Val(txtReqID.Text) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & myDisc & ");"
                End If
            Else
                If InStr(.TextMatrix(i, 6), "%") > 0 Then
                    myDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                Else
                    myDisc = Val(.TextMatrix(i, 6))
                End If
                eMsg = "Line 1136: Adding New Sales Order Details"
                LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, QTY, UNIT, PRICE, STATUS, DISCOUNT) " & _
                               "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & myDisc & ");"
            End If
        Next i
    End With
    
    Qry = "SELECT SPSALESORDER.SPSORDERID From SPSALESORDER WHERE (((SPSALESORDER.SPSORDERID)=" & eOrderID & "));"
    If CtrlNumExist(Qry) Then
        'LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = 'CANCELLED' " & _
                       "WHERE (((SPSALESORDER.SPSORDERID)=" & eOrderID & "));"
        'mySQL
        LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = 'CANCELLED' " & _
                       "WHERE SPSALESORDER.SPSORDERID=" & eOrderID & ";"
        
    End If
    MsgBox "This is Order Number " & txtOrderNo.Text, vbInformation + vbOKOnly, "New Order was made"
    Exit Sub
LocalErr:
    If InStr(eMsg, "Deleting Previous SP Sales Order ID") > 0 Then
        Resume Next
    Else
        MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub

Function GetOrderNo() As Integer
    Dim rsMyOrder As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT Max(SPSALESORDER.ORDERNO) AS MaxOfORDERNO From SPSALESORDER " & _
          "WHERE (((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#));"
    'mySQL
    Qry = "SELECT Max(SPSALESORDER.ORDERNO) AS MaxOfORDERNO From SPSALESORDER " & _
          "WHERE SPSALESORDER.ORDERDATE='" & Format(txtDate.Text, "yyyy-mm-dd") & "';"
    
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
    If chkDP.Value = 0 Then
        If Val(txtCustID.Text) > 0 And chkWholesaleCash.Value = 0 And txtPayScheme.Text = "" Then MyErr = "No Terms specified for On-Account sales order."
    End If
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
        If Val(flxOrd.TextMatrix(i, 2)) = 0 Then MyErr = MyErr & vbCrLf & "Ordered Spare Parts Sales @ Row " & i & " no Qty.!"
        'If Val(flxOrd.TextMatrix(i, 7)) = 0 Then MyErr = MyErr & vbCrLf & "Ordered Spare Parts Sales @ Row " & i & " no Price!"
    Next i
    ValidateSalesOrder = MyErr
End Function

Private Sub LoadSpareParts()
    Dim rsSet As New ADODB.Recordset
    Dim Qry As String
    Dim pPct As Single
        

    Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST " & _
          "From SPAREPARTS WHERE SPAREPARTS.DESCRIPTION NOT LIKE '%INACTIVE%' ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
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
        .ColWidth(14) = 0
        .Col = 6: .Row = 1
        txtTotal.Left = flxOrd.Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7)
        txtTotal.Width = .ColWidth(8)
        flxParts.ColWidth(2) = .ColWidth(4)
        Label3.Left = txtTotal.Left - (Label3.Width + 50)
        Label3.Top = txtTotal.Top
        Label3.FontSize = txtTotal.FontSize
    End With
    'flxGridRef.ColWidth(0) = flxGridRef.Width * 0.2
    'flxGridRef.ColWidth(1) = flxGridRef.Width * 0.8
    
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
    'Shape1.Left = Shape2.Left - 240
    fraAdmin.Left = Frame1.Left + (Frame1.Width - fraAdmin.Width)
    Exit Sub
LocalErr:
    'Unload Me
End Sub

Private Sub Timer1_Timer()
    Dim Qry As String

    Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST, SPAREPARTS.REMARKS " & _
          "From SPAREPARTS WHERE SPAREPARTS.DESCRIPTION NOT LIKE '%INACTIVE%' ORDER BY SPAREPARTS.PARTNUMBER;"
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
        .ColWidth(10) = 0
    End With
    Timer1.Enabled = False
End Sub


Private Sub tmrQtyTip_Timer()
    Dim MyCost As Currency, CostVat As Currency
    
    If tmrCount = 0 Then
        With txtQtyTip
            .Top = flxOrd.CellTop + flxOrd.Top + 290
            .Left = flxOrd.Left + flxOrd.ColWidth(0) + flxOrd.ColWidth(1) + flxOrd.ColWidth(2) + flxOrd.ColWidth(3) + flxOrd.ColWidth(4) + flxOrd.ColWidth(5) + flxOrd.ColWidth(6)
            If Val(txtVat.Text) = 0 Then
                .Text = "No VAT available"
            Else
                MyCost = Val(flxOrd.TextMatrix(ActRow, 14))
                CostVat = MyCost * (1 + (Val(txtVat.Text) / 100))
                .Text = "Cost with Vat: " & Format(CostVat, "#,##0.00")
            End If
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

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
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
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
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
        Else
            .TextMatrix(ActRow, 7) = Format$(.TextMatrix(ActRow, 13), "#,##0.00")
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
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
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
        txtOrderNo.SetFocus
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
    Dim strValid As String
    
    strValid = "0123456789"
    'if ActCol = 1
    If ActCol = 1 Or ActCol = 2 Then
        If KeyAscii > 26 Then
            If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
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
    Dim AvlQty As Long
    
    If txtFlx.Tag = "skiplostfocus" Then
        txtFlx.Tag = ""
        Exit Sub
    End If
    With flxOrd
        Call TOT_PENDINGORDER(Format(txtDate.Text, "yyyy-mm-dd"), Val(flxOrd.TextMatrix(ActRow, 10)))
        AvlQty = Val(.TextMatrix(ActRow, 1)) - Val(txtTotPen.Text)
        If Val(txtFlx.Text) > AvlQty And .TextMatrix(ActRow, 10) <> "" Then
            MsgBox "Order Qty. must not exceed Qty. On-Hand (" & AvlQty & ")!", vbCritical + vbOKOnly, "Validation"
            txtFlx.Tag = "XcdQty"
            txtFlx.SelStart = 0
            txtFlx.SelLength = Len(txtFlx.Text)
            txtFlx.SetFocus
            Exit Sub
        End If
        If InStr(.TextMatrix(ActRow, 5), "ENGINE OIL") > 0 Or InStr(.TextMatrix(ActRow, 5), "OIL 4T") > 0 Then
            .TextMatrix(ActRow, 3) = IIf(Val(.TextMatrix(ActRow, 2)) > 1, "Qrts.", "Qrt.") & IIf(InStr(.TextMatrix(ActRow, 3), "*") > 0, "*", "")
        Else
            flxOrd.TextMatrix(ActRow, 3) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Pcs.", "Pc.") & IIf(InStr(.TextMatrix(ActRow, 3), "*") > 0, "*", "")
        End If
        flxOrd.TextMatrix(ActRow, ActCol) = txtFlx.Text
    End With
    Call ComputeTotal
    txtFlx.Visible = False

End Sub

Public Sub ComputeTotal()
    Dim i As Integer
    Dim TtlPrice As Currency, GrndTtl As Currency, DiscAmt As Currency, TtlDiscAmt As Currency
    Dim Pctd As Single, myPrice As Currency
    
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
                    Pctd = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                Else
                    Pctd = Val(.TextMatrix(i, 6))
                End If
                myPrice = Val(.TextMatrix(i, 13))
                DiscAmt = myPrice * (Pctd / 100)
                TtlDiscAmt = TtlDiscAmt + DiscAmt
            End If
        Next i
        txtTtlDisc.Text = TtlDiscAmt
        txtTotal.Text = Format$(GrndTtl, "#,##0.00")
        txtGrandTotal.Text = Format$(GrndTtl - Val(Replace(txtLessDiscAmt.Text, ",", "")), "#,##0.00")
    End With
            
End Sub

Private Sub txtOrderno_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
End Sub

Private Sub txtPartNo_Change()
    Dim TtlPrice As Currency, DiscPrice As Currency
    Dim PctDisc As Single, myVat As Single, Qry As String
    
    On Error GoTo LocalErr
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
            flxOrd.TextMatrix(ActRow, 1) = .TextMatrix(.Row, 5) 'QtyOnHand
            flxOrd.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 3) 'Description
            flxOrd.TextMatrix(ActRow, 6) = ""   'Discount
            flxOrd.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 6) 'Location
            flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 7) 'Applic Model
            flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 1) 'SPid
            flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 8) 'Max Discount
            flxOrd.TextMatrix(ActRow, 14) = .TextMatrix(.Row, 9) 'COST
            If Val(.TextMatrix(.Row, 9)) = 0 Then
                flxOrd.TextMatrix(ActRow, 7) = "" 'Price
                flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
            Else
                If UCase$(txtCusType.Text) = "WHOLESALE" Or UCase$(txtCusType.Text) = "LGC MARKETING" Then
                    PctDisc = 0
                    Qry = "SELECT SPAREPARTS.DLRPRICE From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(.TextMatrix(.Row, 1)) & "));"
                    PctDisc = Val(GetFieldValue(Qry))
                   
                    'new code 03/31/2010
                    'For i = 0 To flxSPec.Rows - 1 'Get Pct Discount base on SP Exclusion list
                    '    If Val(.TextMatrix(.Row, 1)) = Val(flxSPec.TextMatrix(i, 0)) Then
                    '        PctDisc = Val(flxSPec.TextMatrix(i, 1))
                    '        Exit For
                    '    End If
                    'Next i
                    If PctDisc = 0 Then
                        For i = 0 To flxWDisc.Rows - 1  'Get Pct Discount base on SP Type
                            If .TextMatrix(.Row, 10) = flxWDisc.TextMatrix(i, 0) Then
                                PctDisc = Val(flxWDisc.TextMatrix(i, 1))
                                Exit For
                            End If
                        Next i
                        If PctDisc > 0 Then
                            myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
                            DiscPrice = (Val(.TextMatrix(.Row, 9)) * myVat) * PctDisc
                            If InStr(.TextMatrix(.Row, 10), "GENUINE") > 0 Then
                                If DiscPrice < 10 Then
                                    DiscPrice = DiscPrice + 3
                                ElseIf DiscPrice >= 10 And DiscPrice <= 100 Then
                                    DiscPrice = DiscPrice + 2
                                Else
                                    If InStr(Str(DiscPrice), ".") > 0 Then DiscPrice = DiscPrice + 1
                                End If
                                DiscPrice = Int(DiscPrice)
                            Else
                                If InStr(Str(DiscPrice), ".") > 0 Then
                                    DiscPrice = DiscPrice + 1
                                    DiscPrice = Int(DiscPrice)
                                End If
                            End If
                            flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Price
                        Else
                            flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                        End If
                    Else    'Display Amount from Exclusion List
                        flxOrd.TextMatrix(ActRow, 7) = Format$(PctDisc, "###,##0.00") 'Price
                    End If
                    flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
                'ElseIf InStr(SysBranch, "PANABO") > 0 Then
                '    If Val(.TextMatrix(.Row, 4)) > 1 Then
                '        Qry = "SELECT settings.Value From settings Where settings.DESCRIPTION = 'MARK-UP PRICE';"
                '        PctDisc = Val(GetFieldValue(Qry))
                '        If PctDisc > 0 Then
                '            PctDisc = 1 + (PctDisc / 100)
                '            DiscPrice = Val(.TextMatrix(.Row, 4)) * PctDisc
                '            flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Pricegud
                '            flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4))   'DiscPrice 'SRP
                '        Else
                '            DiscPrice = Val(.TextMatrix(.Row, 4)) * 1.1
                '            flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Pricegud
                '            flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4))   'DiscPrice 'SRP
                '        End If
                '    End If
                Else
                    flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                    flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
                End If
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
    Exit Sub
LocalErr:
    If Err.Number = 30009 Then
        flxParts.Visible = False
        txtPartNo.Visible = False
        If rsParts.State = adStateOpen Then rsParts.Close
        Timer1.Enabled = True
        MsgBox "Program is reloading Part Number" & vbCrLf & "Please wait for a few seconds...", vbInformation + vbOKOnly, "Reloading"
    Else
        flxParts.Visible = False
        txtPartNo.Visible = False
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub

Private Sub txtPartNo_GotFocus()
    If rsParts.State = adStateClosed Then Timer1.Enabled = True
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
    Dim eCol As Integer
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        'Call txtPartNo_LostFocus
        
        
        'If Val(flxOrd.TextMatrix(ActRow, 7)) = 0 And txtPartNo.Text <> "" Then
        '    MsgBox ("Sparepart Zero SRP" & vbCrLf & "Please contact your Manager"), vbCritical, "Zero SRP"
        '    For eCol = 1 To 14
        '        flxOrd.TextMatrix(ActRow, eCol) = ""
        '    Next eCol
        '    txtPartNo.Visible = False
        '    flxParts.Visible = False
        '    Exit Sub
        'End If
 
        flxOrd.SetFocus
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
    Dim i, eCol As Integer, Qry As String
    
    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    
    If Val(flxOrd.TextMatrix(ActRow, 7)) = 0 And txtPartNo.Text <> "" Then
        MsgBox ("Sparepart Zero SRP" & vbCrLf & "Please contact your Manager"), vbCritical, "Zero SRP"
        For eCol = 1 To 14
           flxOrd.TextMatrix(ActRow, eCol) = ""
        Next eCol
        txtPartNo.Visible = False
        flxParts.Visible = False
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
    
    'flxGridRef.TextMatrix(1, 0) = flxOrd.TextMatrix(ActRow, 4)
    'Qry = "SELECT spareparts.SUPERSEDINGPN From spareparts where spareparts.PARTNUMBER = '" & flxOrd.TextMatrix(ActRow, 4) & "';"
    'flxGridRef.TextMatrix(1, 1) = CStr(GetFieldValue(Qry))

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
    Dim strValid As String
    
    strValid = "0123456789."    'dati naay X
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
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
    Dim AmtDsc As Currency, LessAmt As Currency, Pctd As Double, myVat As Currency
    
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
            myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
            AmtDsc = Val(.TextMatrix(ActRow, 14)) * myVat
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
            Pctd = (Val(.TextMatrix(ActRow, 13)) - Val(Replace(txtPrice.Text, ",", ""))) / Val(.TextMatrix(ActRow, 13))
            If Val(Replace(txtPrice.Text, ",", "")) > 0 Then .TextMatrix(ActRow, 6) = Format(Pctd, "#0.000%")
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
Sub UpdateNumOrder()
    Dim MyOrder As New ADODB.Recordset, rsMyPvOrder As New ADODB.Recordset, myOrderDt As String
    Dim myPvOrder As Long, Qry As String, MyCrtOrder As Long, dts As Integer, myOrderNo As Integer
        'Query number of order today
        Qry = "SELECT " & _
            "spsalesorder.SPSORDERID " & _
        "From " & _
            "spsalesorder " & _
        "Where " & _
            "spsalesorder.ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "';"
            
            MyOrder.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If MyOrder.RecordCount > 0 Then MyOrder.Close: Exit Sub
            'Get previous max orderID and Add to the Setting
            dts = 1
            Do
                myOrderDt = Date - dts
                Qry = "SELECT " & _
                    "Max(spsalesorder.SPSORDERID) As LastSpID " & _
                "From " & _
                    "spsalesorder " & _
                "Where " & _
                    "spsalesorder.ORDERDATE = '" & Format(myOrderDt, "yyyy-mm-dd") & "';"
                 myPvOrder = Val(GetFieldValue(Qry))
                 dts = dts + 1
            Loop Until myPvOrder <> 0
            
        'Save Last max orderID
        MyOrder.Close
        LYRcon.Execute "Update settings SET `VALUE` = '" & myPvOrder & "', REMARKS = '" & Format(myOrderDt, "yyyy-mm-dd") & "', " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Date, "yyyy-mm-dd hh:mm:ss") & "' " & _
            "Where settings.DESCRIPTION = 'SALES ORDER';"
End Sub

Private Sub TOT_PENDINGORDER(SP_ORDERDATE As String, SP_ID As Long)
    Dim rsPOrder As New ADODB.Recordset
    Dim QRY_ORDER As String, PENDINGORDER As Long, S_Order As Long, s_Col As Long
    
    QRY_ORDER = ""
    If AddNew Then
        QRY_ORDER = "SELECT SUM(spsorderdetail.QTY) AS TOTALQTY, spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO " & _
            " From spsalesorder RIGHT OUTER JOIN spsorderdetail ON (spsalesorder.SPSORDERID = spsorderdetail.SPSORDERID) " & _
            " Where spsalesorder.ORDERDATE = '" & SP_ORDERDATE & "' AND spsorderdetail.SPAREPARTID = " & SP_ID & " AND spsalesorder.INVOICENO IS NULL Group By " & _
            " spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO IS NULL;"
    Else
        QRY_ORDER = "SELECT SUM(spsorderdetail.QTY) AS TOTALQTY, spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO " & _
            " From spsalesorder RIGHT OUTER JOIN spsorderdetail ON (spsalesorder.SPSORDERID = spsorderdetail.SPSORDERID) " & _
            " Where spsalesorder.ORDERDATE = '" & SP_ORDERDATE & "' AND spsorderdetail.SPAREPARTID = " & SP_ID & " AND spsalesorder.ORDERNO <> " & Val(txtOrderNo.Text) & " AND spsalesorder.INVOICENO IS NULL Group By " & _
            " spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO IS NULL;"
    End If
    
    PENDINGORDER = Val(GetFieldValue(QRY_ORDER))
    QRY_ORDER = "SELECT SUM(stocktransdtl.QTY) As TOTTRANS From stocktransfer RIGHT OUTER JOIN stocktransdtl ON (stocktransfer.TRANSFERID = stocktransdtl.TRANSFERID) " & _
                 " Where stocktransfer.`STATUS` = 'PENDING' AND stocktransdtl.SPAREPARTID = " & SP_ID & ";"
    TotPTrans = Val(GetFieldValue(QRY_ORDER))

    txtTotPen.Text = PENDINGORDER + TotPTrans
    If TotPTrans > 0 Then
        flxOrder.TextMatrix(0, 0) = "TRANS = " & TotPTrans
    End If
    With rsPOrder
    
        If AddNew Then
            QRY_ORDER = "SELECT spsorderdetail.QTY, spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO, spsalesorder.FRONTLINER, spsalesorder.ORDERNO " & _
                "From spsalesorder RIGHT OUTER JOIN spsorderdetail ON (spsalesorder.SPSORDERID = spsorderdetail.SPSORDERID) " & _
                "Where spsalesorder.ORDERDATE = '" & SP_ORDERDATE & "' AND spsorderdetail.SPAREPARTID = " & SP_ID & " AND spsalesorder.INVOICENO IS NULL;"
        Else
            QRY_ORDER = "SELECT spsorderdetail.QTY, spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO, spsalesorder.FRONTLINER, spsalesorder.ORDERNO " & _
                "From spsalesorder RIGHT OUTER JOIN spsorderdetail ON (spsalesorder.SPSORDERID = spsorderdetail.SPSORDERID) " & _
                "Where spsalesorder.ORDERDATE = '" & SP_ORDERDATE & "' AND spsorderdetail.SPAREPARTID = " & SP_ID & " AND spsalesorder.ORDERNO <> " & Val(txtOrderNo.Text) & "  AND spsalesorder.INVOICENO IS NULL;"
        End If
       
       If .State = adStateOpen Then .Close
       .Open QRY_ORDER, LYRcon, adOpenKeyset, adLockOptimistic
       
       If .RecordCount > 0 Then
       
       flxOrder.Cols = .RecordCount + 1
       s_Col = 1
       .MoveFirst
           Do Until .EOF
               flxOrder.TextMatrix(0, s_Col) = "[ " & !ORDERNO & " ] " & !FRONTLINER & " = " & !Qty
               s_Col = s_Col + 1
               flxOrder.ColWidth(s_Col - 1) = flxOrder.Width * 0.25
               .MoveNext
           Loop
       End If
    End With
End Sub


