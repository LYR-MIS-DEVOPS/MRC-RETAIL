VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SPpricingOld 
   Caption         =   "Spare Parts Price Factor"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   13770
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDPfactorID 
      Height          =   285
      Left            =   12600
      TabIndex        =   126
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDPvat 
      Height          =   285
      Left            =   11880
      TabIndex        =   125
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDPMarkUp 
      Height          =   285
      Left            =   11280
      TabIndex        =   124
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtOldDP 
      Height          =   285
      Left            =   10560
      TabIndex        =   123
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   2895
      Left            =   7320
      TabIndex        =   89
      Top             =   9120
      Visible         =   0   'False
      Width           =   6135
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         Left            =   1680
         TabIndex        =   111
         Tag             =   "XYF"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1680
         TabIndex        =   110
         Tag             =   "XYF"
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox chkAllSupp 
         Caption         =   "All"
         Height          =   255
         Left            =   4320
         TabIndex        =   109
         Tag             =   "XYF"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtPartNo 
         Height          =   285
         Left            =   840
         TabIndex        =   92
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtDPrice 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   91
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPartsD 
         Height          =   1335
         Left            =   960
         TabIndex        =   90
         Top             =   1560
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2355
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
      Begin MSFlexGridLib.MSFlexGrid flxSPec 
         Height          =   3795
         Left            =   480
         TabIndex        =   93
         Tag             =   "xyf"
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   6694
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
         FormatString    =   "      |<PARTNUMBER        |^COST (w/VAT)|^Disc. Price "
      End
      Begin lvButton.lvButtons_H cmdGo 
         Height          =   315
         Left            =   4200
         TabIndex        =   108
         Tag             =   "XYF"
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
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
         ImgAlign        =   5
         Image           =   "SPpricingold.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.Label Label9 
         Caption         =   "SUPPLIER"
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
         Left            =   600
         TabIndex        =   113
         Tag             =   "XYF"
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "DESCRIPTION"
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
         Left            =   600
         TabIndex        =   112
         Tag             =   "XYF"
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Specific Sparepart Discount"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   94
         Tag             =   "xyf"
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.TextBox txtDVat 
      Height          =   285
      Left            =   8880
      TabIndex        =   87
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtSuppCode 
      Height          =   285
      Left            =   6840
      TabIndex        =   86
      Top             =   9120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8880
      Top             =   120
   End
   Begin TabDlg.SSTab tabProg 
      Height          =   735
      Left            =   480
      TabIndex        =   78
      Top             =   9120
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1296
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPpricingold.frx":059A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPct"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLoading"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pBar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblLoading 
         Caption         =   "Updating Prices..."
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label lblPct 
         Alignment       =   1  'Right Justify
         Caption         =   "0%"
         Height          =   255
         Left            =   4920
         TabIndex        =   80
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame fraFind 
      Height          =   615
      Left            =   1320
      TabIndex        =   73
      Tag             =   "XYF"
      Top             =   400
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtFind 
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
         Left            =   2040
         TabIndex        =   72
         Tag             =   "XYF"
         Top             =   210
         Width           =   3615
      End
      Begin VB.Label lblFind 
         Alignment       =   1  'Right Justify
         Caption         =   "Find Description:"
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
         Left            =   120
         TabIndex        =   75
         Tag             =   "XYF"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label8 
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox txtPDate 
      Height          =   285
      Left            =   7920
      TabIndex        =   34
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtOldRef 
      Height          =   285
      Left            =   7200
      TabIndex        =   33
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtOldCost 
      Height          =   285
      Left            =   4560
      TabIndex        =   23
      Top             =   8640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtHist 
      Height          =   285
      Left            =   6360
      TabIndex        =   22
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtOldPrice 
      Height          =   285
      Left            =   5520
      TabIndex        =   21
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   7695
      Left            =   120
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   13573
      _Version        =   393216
      RowHeightMin    =   300
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   360
      TabIndex        =   13
      Top             =   8520
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtFParts 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "SPpricingold.frx":05B6
         Top             =   277
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Find Part Number:"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   270
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab tabSP 
      Height          =   7695
      Left            =   8880
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   13573
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Price Factor"
      TabPicture(0)   =   "SPpricingold.frx":0B40
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdDEdit"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdDcancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdDsave"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdFNew"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdFEdit"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdFCancel"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdFSave"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraFactor"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Single Pricing"
      TabPicture(1)   =   "SPpricingold.frx":0B5C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxOld"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdCancel"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdSave"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtTag"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdEdit"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtFactorID"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "FraSRP"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "FraDP"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cboHistory"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Group Pricing"
      TabPicture(2)   =   "SPpricingold.frx":0B78
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flxGroup"
      Tab(2).Control(1)=   "cmdGNew"
      Tab(2).Control(2)=   "cmdGEdit"
      Tab(2).Control(3)=   "cmdGCancel"
      Tab(2).Control(4)=   "cmdGSave"
      Tab(2).Control(5)=   "Frame3"
      Tab(2).Control(6)=   "txtGEditRow"
      Tab(2).ControlCount=   7
      Begin VB.ComboBox cboHistory 
         Height          =   315
         Left            =   -73200
         TabIndex        =   122
         Tag             =   "XYF"
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Frame FraDP 
         Caption         =   "DEALER PRICE"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   116
         Tag             =   "XYF"
         Top             =   3960
         Width           =   4215
         Begin VB.ComboBox cboDPfactor 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   118
            Tag             =   "XYF"
            Top             =   300
            Width           =   2535
         End
         Begin VB.TextBox txtDP 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   117
            Tag             =   "XYF"
            Top             =   640
            Width           =   1575
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "PRICE FACTOR:"
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
            Index           =   2
            Left            =   120
            TabIndex        =   120
            Tag             =   "XYF"
            Top             =   330
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "DEALER PRICE:"
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
            Left            =   120
            TabIndex        =   119
            Tag             =   "XYF"
            Top             =   690
            Width           =   1335
         End
      End
      Begin VB.Frame FraSRP 
         Caption         =   "RETAIL PRICE"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   103
         Tag             =   "XFY"
         Top             =   2760
         Width           =   4215
         Begin VB.TextBox txtPrice 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   105
            Tag             =   "XYF"
            Top             =   660
            Width           =   1575
         End
         Begin VB.ComboBox cboSPFactor 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   104
            Tag             =   "XYF"
            Top             =   300
            Width           =   2535
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "RETAIL PRICE:"
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
            Left            =   120
            TabIndex        =   107
            Tag             =   "XYF"
            Top             =   690
            Width           =   1335
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "PRICE FACTOR:"
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
            Left            =   120
            TabIndex        =   106
            Tag             =   "XYF"
            Top             =   330
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1815
         Left            =   120
         TabIndex        =   98
         Tag             =   "XYF"
         Top             =   5760
         Width           =   4455
         Begin VB.TextBox txtDFctr 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   100
            Top             =   960
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cboSPType 
            Height          =   315
            Left            =   480
            TabIndex        =   99
            Top             =   840
            Visible         =   0   'False
            Width           =   2055
         End
         Begin MSFlexGridLib.MSFlexGrid flxWDisc 
            Height          =   1365
            Left            =   120
            TabIndex        =   101
            Tag             =   "xyf"
            Top             =   120
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   2408
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   -2147483633
            ScrollBars      =   2
            FormatString    =   "      |<     SPAREPART   TYPE        |^Discount Factor "
         End
         Begin VB.Label Label11 
            Caption         =   "Wholesale Discount Formula = COST * Discount Factor"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Tag             =   "xyf"
            Top             =   1480
            Width           =   4095
         End
      End
      Begin VB.TextBox txtGEditRow 
         Height          =   285
         Left            =   -71280
         TabIndex        =   85
         Tag             =   "XYF"
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtFactorID 
         Height          =   285
         Left            =   -71280
         TabIndex        =   69
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   3135
         Left            =   -74880
         TabIndex        =   52
         Tag             =   "XYF"
         Top             =   900
         Width           =   4215
         Begin VB.TextBox txtGDisc 
            Height          =   375
            Left            =   600
            TabIndex        =   84
            Top             =   2640
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtGVat 
            Height          =   375
            Left            =   480
            TabIndex        =   83
            Top             =   2640
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtGmrk 
            Height          =   375
            Left            =   360
            TabIndex        =   82
            Top             =   2640
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtGFactorID 
            Height          =   375
            Left            =   120
            TabIndex        =   77
            Top             =   2640
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtTtlGMember 
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
            Left            =   3000
            TabIndex        =   64
            Tag             =   "XYF"
            Top             =   2600
            Width           =   975
         End
         Begin VB.ComboBox cboGFactor 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   59
            Tag             =   "XYF"
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtGDate 
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
            Height          =   285
            Left            =   1560
            TabIndex        =   58
            Tag             =   "XYF"
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txtGDesc 
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
            Left            =   240
            MaxLength       =   100
            TabIndex        =   57
            Tag             =   "XYF"
            Top             =   480
            Width           =   3855
         End
         Begin lvButton.lvButtons_H cmdCriteria 
            Height          =   375
            Left            =   240
            TabIndex        =   56
            Tag             =   "XYF"
            Top             =   1920
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Criteria"
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
            Image           =   "SPpricingold.frx":0B94
            cBack           =   -2147483633
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Group Members"
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
            Index           =   3
            Left            =   240
            TabIndex        =   63
            Tag             =   "XYF"
            Top             =   2625
            Width           =   2655
         End
         Begin VB.Label lblSQL 
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
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblCriteria 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1680
            TabIndex        =   61
            Tag             =   "XY"
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label lblGFctrDtls 
            Alignment       =   2  'Center
            Caption         =   "Mark-up: 0, VAT: 0, Max Discount: 0"
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
            Left            =   240
            TabIndex        =   60
            Tag             =   "XYF"
            Top             =   1560
            Width           =   3855
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Price Factor"
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
            Index           =   2
            Left            =   240
            TabIndex        =   55
            Tag             =   "XYF"
            Top             =   1230
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Date Created"
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
            Left            =   120
            TabIndex        =   54
            Tag             =   "XYF"
            Top             =   855
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Group Description"
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
            Index           =   0
            Left            =   240
            TabIndex        =   53
            Tag             =   "XYF"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraFactor 
         Height          =   4410
         Left            =   120
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   880
         Width           =   4455
         Begin VB.TextBox txtSPFid 
            Height          =   285
            Left            =   240
            TabIndex        =   47
            Top             =   1200
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtFMaxDisc 
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
            Left            =   1680
            TabIndex        =   46
            Tag             =   "XYF"
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtFVat 
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
            Left            =   1680
            TabIndex        =   45
            Tag             =   "XYF"
            Top             =   1335
            Width           =   2055
         End
         Begin VB.TextBox txtFMarkUp 
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
            Left            =   1680
            TabIndex        =   44
            Tag             =   "XYF"
            Top             =   1005
            Width           =   2055
         End
         Begin VB.TextBox txtFDate 
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
            Left            =   1680
            TabIndex        =   40
            Tag             =   "XYF"
            Top             =   660
            Width           =   2055
         End
         Begin VB.TextBox txtFDesc 
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
            MaxLength       =   100
            TabIndex        =   39
            Tag             =   "XYF"
            Top             =   240
            Width           =   2655
         End
         Begin MSFlexGridLib.MSFlexGrid flxFactors 
            Height          =   2175
            Left            =   120
            TabIndex        =   88
            Tag             =   "XYF"
            Top             =   2160
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   3836
            _Version        =   393216
            Cols            =   7
            RowHeightMin    =   300
            BackColorBkg    =   -2147483633
            SelectionMode   =   1
            FormatString    =   "     |^DATE           |< FACTOR DESCRIPTION         |^MARK-UP      |^    VAT          |^MAX DISCOUNT"
         End
         Begin VB.Label lblFactor 
            Alignment       =   1  'Right Justify
            Caption         =   "Max. Discount:"
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
            Index           =   4
            Left            =   120
            TabIndex        =   43
            Tag             =   "XYF"
            Top             =   1710
            Width           =   1455
         End
         Begin VB.Label lblFactor 
            Alignment       =   1  'Right Justify
            Caption         =   "VAT"
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
            Index           =   3
            Left            =   240
            TabIndex        =   42
            Tag             =   "XYF"
            Top             =   1365
            Width           =   1335
         End
         Begin VB.Label lblFactor 
            Alignment       =   1  'Right Justify
            Caption         =   "Mark Up"
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
            Index           =   2
            Left            =   240
            TabIndex        =   41
            Tag             =   "XYF"
            Top             =   1035
            Width           =   1335
         End
         Begin VB.Label lblFactor 
            Alignment       =   1  'Right Justify
            Caption         =   "Date Created:"
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
            Left            =   240
            TabIndex        =   38
            Tag             =   "XYF"
            Top             =   690
            Width           =   1335
         End
         Begin VB.Label lblFactor 
            Caption         =   "Factor Description:"
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
            Index           =   0
            Left            =   240
            TabIndex        =   37
            Tag             =   "XYF"
            Top             =   240
            Width           =   2415
         End
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   375
         Left            =   -74760
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Edit"
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
         Image           =   "SPpricingold.frx":112E
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtTag 
         Height          =   285
         Left            =   -73440
         TabIndex        =   20
         Top             =   6720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   920
         Width           =   4215
         Begin VB.TextBox txtMaxDisc 
            Height          =   285
            Left            =   3240
            TabIndex        =   68
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtVat 
            Height          =   285
            Left            =   3240
            TabIndex        =   67
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtMarkUp 
            Height          =   285
            Left            =   3240
            TabIndex        =   66
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtRefNo 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   26
            Tag             =   "XYF"
            Top             =   1305
            Width           =   1575
         End
         Begin VB.TextBox txtCost 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   18
            Tag             =   "XYF"
            Top             =   945
            Width           =   1575
         End
         Begin VB.TextBox txtPParts 
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
            Left            =   1560
            TabIndex        =   17
            Tag             =   "XYF"
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "REF. NO."
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
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Tag             =   "XYF"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblDescription 
            Alignment       =   1  'Right Justify
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
            Left            =   240
            TabIndex        =   19
            Tag             =   "XYF"
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "PART NUMBER:"
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
            Left            =   120
            TabIndex        =   12
            Tag             =   "XYF"
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "UNIT COST:"
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
            Left            =   120
            TabIndex        =   11
            Tag             =   "XYF"
            Top             =   975
            Width           =   1335
         End
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   375
         Left            =   -74760
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Save"
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
         Image           =   "SPpricingold.frx":16C8
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   -73200
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Cancel"
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
         Image           =   "SPpricingold.frx":1C62
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdFSave 
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Tag             =   "XYF"
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Save"
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
         Image           =   "SPpricingold.frx":21FC
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdFCancel 
         Height          =   375
         Left            =   1800
         TabIndex        =   31
         Tag             =   "XYF"
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Cancel"
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
         Image           =   "SPpricingold.frx":2796
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdFEdit 
         Height          =   375
         Left            =   1800
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Edit Factor"
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
         Image           =   "SPpricingold.frx":2D30
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdFNew 
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "New Factor"
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
         Image           =   "SPpricingold.frx":32CA
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdGSave 
         Height          =   375
         Left            =   -74880
         TabIndex        =   48
         Tag             =   "XYF"
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Save"
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
         Image           =   "SPpricingold.frx":3864
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdGCancel 
         Height          =   375
         Left            =   -73320
         TabIndex        =   49
         Tag             =   "XYF"
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Cancel"
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
         Image           =   "SPpricingold.frx":3DFE
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdGEdit 
         Height          =   375
         Left            =   -73320
         TabIndex        =   50
         Tag             =   "XYF"
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Edit Group"
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
         Image           =   "SPpricingold.frx":4398
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdGNew 
         Height          =   375
         Left            =   -74880
         TabIndex        =   51
         Tag             =   "XYF"
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "New Group"
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
         Image           =   "SPpricingold.frx":4932
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxGroup 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   65
         Tag             =   "XYF"
         Top             =   4080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5106
         _Version        =   393216
         Cols            =   7
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         FormatString    =   "     |^DATE           |< GROUP DESCRIPTION         |^PRICE FACTOR      |<   CRITERIA                   "
      End
      Begin lvButton.lvButtons_H cmdDsave 
         Height          =   375
         Left            =   240
         TabIndex        =   95
         Tag             =   "XYF"
         Top             =   5385
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Save"
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
         Image           =   "SPpricingold.frx":4ECC
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdDcancel 
         Height          =   375
         Left            =   1800
         TabIndex        =   96
         Tag             =   "XYF"
         Top             =   5385
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Cancel"
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
         Image           =   "SPpricingold.frx":5466
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdDEdit 
         Height          =   375
         Left            =   240
         TabIndex        =   97
         Tag             =   "XYF"
         Top             =   5400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Edit DP Discount"
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
         Image           =   "SPpricingold.frx":5A00
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxOld 
         Height          =   1935
         Left            =   -74760
         TabIndex        =   121
         Tag             =   "XYF"
         Top             =   5520
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   315
         BackColorBkg    =   -2147483633
         FormatString    =   "^ DATE       |<REF. NO.      |>  OLD COST  |>    OLD PRICE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPpricingold.frx":5F9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   5160
      TabIndex        =   24
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Pricing"
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
      Image           =   "SPpricingold.frx":6874
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   375
      Left            =   120
      TabIndex        =   70
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Find Description"
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
      Image           =   "SPpricingold.frx":6E0E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   375
      Left            =   1800
      TabIndex        =   71
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Refresh"
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
      Image           =   "SPpricingold.frx":73A8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmd0SRP 
      Height          =   375
      Left            =   3480
      TabIndex        =   76
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Zero SRP"
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
      Image           =   "SPpricingold.frx":7942
      cBack           =   -2147483633
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COST UPDATE HISTORY"
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
      Height          =   210
      Index           =   5
      Left            =   2610
      TabIndex        =   115
      Tag             =   "XYF"
      Top             =   9630
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COST UPDATE HISTORY"
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
      Height          =   210
      Index           =   4
      Left            =   2640
      TabIndex        =   114
      Tag             =   "XYF"
      Top             =   9600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   9960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13890
      m_MDI_Height    =   9120
   End
   Begin VB.Label lblVat 
      Height          =   255
      Left            =   9360
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblB 
      Height          =   255
      Left            =   8760
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblA 
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblSPid 
      Height          =   255
      Left            =   10200
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "View:"
      Height          =   255
      Left            =   9480
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE FACTOR"
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
      Left            =   11100
      TabIndex        =   2
      Top             =   450
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
      Left            =   210
      TabIndex        =   0
      Top             =   450
      Width           =   3000
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
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3000
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE FACTOR"
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
      Left            =   11130
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Menu mnuGrpF 
      Caption         =   "Group List"
      Visible         =   0   'False
      Begin VB.Menu mnuGDisp 
         Caption         =   "Display Members in Grid"
      End
   End
   Begin VB.Menu mnuWDisc 
      Caption         =   "Wholesale Discount"
      Visible         =   0   'False
      Begin VB.Menu mnuWhlAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuWhlDel 
         Caption         =   "Delete Row"
      End
   End
   Begin VB.Menu mnuSpec 
      Caption         =   "Specific SP Discount"
      Visible         =   0   'False
      Begin VB.Menu mnuSpcAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuSpcDel 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "SPpricingOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim rsPartsD As New ADODB.Recordset
Dim rsFactor As New ADODB.Recordset
Dim rsSupplier As New ADODB.Recordset
Dim PrtSrch As Boolean, bNone As Boolean, EditMode As Boolean, SuppSrch As Boolean, PrtBln As Boolean
Dim CostChange As Boolean, NewFactor As Boolean, GrpAddNew As Boolean, DWLock As Boolean
Dim ActRow As Integer, ActCol As Integer

Private Sub LockData(Bln As Boolean)
    Frame1.Enabled = Not Bln
    flxParts.Enabled = Not Bln
    cboView.Locked = Bln
    cboPartNo.Locked = Bln
    Frame2.Enabled = Not Bln
    flxOld.Enabled = Not Bln
End Sub

Private Sub FindPartNumber()
    Dim i As Integer
    
    If lblSPid.Caption = "" Then Exit Sub
    With flxParts
        .Col = 0: .Row = 1: .ColSel = 4: .RowSel = .Rows - 1
        .CellBackColor = vbWhite
        .CellForeColor = vbBlack
        .ColSel = 0: .RowSel = 1
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 4) = lblSPid.Caption Then
                .TopRow = i
                .Col = 0: .Row = i: .ColSel = 4: .RowSel = i
                .CellBackColor = .BackColorSel
                .CellForeColor = .ForeColorSel
                .ColSel = 0
                If CostChange Then
                    .TextMatrix(i, 2) = lblCost.Caption
                    .TextMatrix(i, 3) = lblPrice.Caption
                    CostChange = False
                End If
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub DisplayCostHistory(MyHis As String)
    Dim pHis As Variant
    Dim pRow As Variant
    Dim i As Integer
    
    pHis = Split(MyHis, "*")
    With flxOld
        .Rows = 1
        For i = LBound(pHis) To UBound(pHis)
            If pHis(i) <> "" Then
                If InStr(pHis(i), "^") > 0 Then
                    pRow = Split(pHis(i), "^")
                    .AddItem pRow(0) & vbTab & pRow(1) & vbTab & pRow(2) & vbTab & pRow(3)
                Else
                    .AddItem pHis(i)
                End If
            End If
        Next i
    End With
End Sub

Private Sub DisplayDealerPriceHistory(SPid As Integer)
    Dim rsPDhist As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    'Qry = "SELECT DPFACTOR.* From DPFACTOR WHERE (((DPFACTOR.SPAREPARTID)=" & SPid & ")) Order By SeqNo Desc;"
    'mySQL
    Qry = "SELECT DPFACTOR.* From DPFACTOR WHERE DPFACTOR.SPAREPARTID=" & SPid & " Order By SeqNo Desc;"
    
    flxOld.Rows = 1
    With rsPDhist
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                'date , ref, oldcost, old price
                flxOld.AddItem Format(!FDATE, "mm/dd/yyyy") & vbTab & IIf(IsNull(!RefNo), "", !RefNo) & vbTab & Format(!Cost, "#,##0.00") & vbTab & Format(!DP, "#,##0.00")
                .MoveNext
            Loop
        Else
            flxOld.Rows = 2
        End If
        .Close
    End With
                
End Sub

Private Sub cboDPfactor_Change()
    Dim FndFctr As Boolean
    Dim i As Integer
    Dim CostVat As Currency, DlrPrice As Currency, MyCost As Currency
    
    If cboDPfactor.Text = "" Then
        txtDPfactorID.Text = ""
        Exit Sub
    End If
    With flxFactors
        FndFctr = False
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) = cboDPfactor.Text Then
                txtDPMarkUp.Text = .TextMatrix(i, 3)
                txtDPvat.Text = .TextMatrix(i, 4)
                txtDPfactorID.Text = .TextMatrix(i, 6)
                MyCost = Val(Replace(txtCost.Text, ",", ""))
                If MyCost > 0 Then
                    CostVat = MyCost + (MyCost * (Val(txtDPvat.Text) / 100))
                    DlrPrice = CostVat + (CostVat * (Val(txtDPMarkUp.Text) / 100))
                    txtDP.Text = Format$(Int(DlrPrice + 0.6) + 1, "####0.00")
                End If
                FndFctr = True
            End If
        Next i
    End With
    If Not FndFctr Then
        txtDPMarkUp.Text = ""
        txtDPvat.Text = ""
        txtDPfactorID.Text = ""
    End If
End Sub

Private Sub cboDPfactor_Click()
    Call cboDPfactor_Change
End Sub

Private Sub cboDPfactor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txtDP.Text = ""
    End If
End Sub

Private Sub cboDPfactor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDP.SetFocus
End Sub

Private Sub cboGFactor_Change()
    Dim FndFctr As Boolean
    Dim i As Integer
    Dim CostVat As Currency, NewPrice As Currency
    
    If cboGFactor.Text = "" Then Exit Sub
    With flxFactors
        FndFctr = False
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) = cboGFactor.Text Then
                lblGFctrDtls.Caption = "Mark-up: " & .TextMatrix(i, 3) & ", VAT: " & _
                .TextMatrix(i, 4) & ", Discount: " & .TextMatrix(i, 5)
                txtGFactorID.Text = .TextMatrix(i, 6)
                txtGmrk.Text = IIf(InStr(.TextMatrix(i, 3), "%") > 0, Left$(.TextMatrix(i, 3), Len(.TextMatrix(i, 3)) - 1), .TextMatrix(i, 3))
                txtGVat.Text = IIf(InStr(.TextMatrix(i, 4), "%") > 0, Left$(.TextMatrix(i, 4), Len(.TextMatrix(i, 4)) - 1), .TextMatrix(i, 4))
                txtGDisc.Text = .TextMatrix(i, 5)
                FndFctr = True
            End If
        Next i
    End With
    If Not FndFctr Then
        lblGFctrDtls.Caption = ""
        txtGFactorID.Text = ""
        txtGmrk.Text = ""
        txtGVat.Text = ""
        txtGDisc.Text = ""
    End If
End Sub

Private Sub cboGFactor_Click()
    Call cboGFactor_Change
End Sub

Private Sub cboHistory_Click()
    If cboHistory.ListIndex = 0 Then    'Retail Price
        Call DisplayCostHistory(flxParts.TextMatrix(flxParts.Row, 6))
    ElseIf cboHistory.ListIndex = 1 Then
        Call DisplayDealerPriceHistory(Val(flxParts.TextMatrix(flxParts.Row, 5)))
    End If
End Sub

Private Sub cboSPFactor_Change()
    Dim FndFctr As Boolean
    Dim i As Integer
    Dim CostVat As Currency, NewPrice As Currency, MyCost As Currency
    
    If cboSPFactor.Text = "" Then
        txtFactorID.Text = ""
        txtMaxDisc.Text = ""
        Exit Sub
    End If
    With flxFactors
        FndFctr = False
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) = cboSPFactor.Text Then
                txtMarkUp.Text = .TextMatrix(i, 3)
                txtVat.Text = .TextMatrix(i, 4)
                txtMaxDisc.Text = .TextMatrix(i, 5)
                txtFactorID.Text = .TextMatrix(i, 6)
                MyCost = Val(Replace(txtCost.Text, ",", ""))
                If MyCost > 0 Then
                    CostVat = MyCost + (MyCost * (Val(txtVat.Text) / 100))
                    NewPrice = CostVat + (CostVat * (Val(txtMarkUp.Text) / 100))
                    txtPrice.Text = Format$(Int(NewPrice + 0.6) + 1, "####0.00")
                End If
                FndFctr = True
            End If
        Next i
    End With
    If Not FndFctr Then
        txtMarkUp.Text = ""
        txtVat.Text = ""
        txtMaxDisc.Text = ""
        txtFactorID.Text = ""
    End If
End Sub

Private Sub cboSPFactor_Click()
    Call cboSPFactor_Change
End Sub

Private Sub cboSPFactor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then txtPrice.Text = ""
End Sub

Private Sub cboSPFactor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPrice.SetFocus
End Sub

Private Sub cboSPFactor_LostFocus()
    If cboSPFactor.Text = "" Then txtFactorID.Text = 0
End Sub

Private Sub cboSPtype_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSPType.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSPtype_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cboSPType.Text = Empty
        flxWDisc.TextMatrix(ActRow, 1) = cboSPType.Tag
        cboSPType.Tag = "skiplostfocus"
        flxWDisc.SetFocus
        flxWDisc.Col = 1
        cboSPType.Visible = False
    End If
End Sub

Private Sub cboSPtype_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxWDisc.SetFocus
        If cboSPType.Text = "" Then
            flxWDisc.Col = 1
            flxWDisc.Row = ActRow
        ElseIf flxWDisc.TextMatrix(ActRow, 2) = "" Then
            flxWDisc.Col = 2
            flxWDisc.Row = ActRow
        ElseIf cboSPType.Text <> "" And flxWDisc.TextMatrix(ActRow, 2) <> "" And ActRow = flxWDisc.Rows - 1 Then
            flxWDisc.AddItem ""
            flxWDisc.Row = flxWDisc.Rows - 1
            flxWDisc.Col = 1
        Else
            flxWDisc.Row = ActRow + 1
            flxWDisc.Col = 1
        End If
    End If

End Sub

Private Sub cboSPType_LostFocus()
    If cboSPType.Tag = "skiplostfocus" Or cboSPType.Text = "" Then
        cboSPType.Tag = ""
        cboSPType.Visible = False
        Exit Sub
    End If
    If cboSPType.ListIndex = -1 Then
        MsgBox "Please select a Sparepart type from the list!", vbExclamation + vbOKOnly, "Sparepart Type"
        Exit Sub
        cboSPType.Visible = False
    End If
    flxWDisc.TextMatrix(ActRow, 1) = cboSPType.Text
    cboSPType.Visible = False
    'lstChange.AddItem "Wholesale Discount"
End Sub

Private Sub cboSupplier_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsSupplier.MoveFirst
            rsSupplier.Move i
            txtSuppCode.Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            chkAllSupp.Value = 0
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppCode.Text = ""
        End If
    End With

End Sub

Private Sub cboSupplier_Click()
    Call cboSupplier_Change
End Sub

Private Sub cboSupplier_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSupplier.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboFind.SetFocus
End Sub

Private Sub chkAllSupp_Click()
    If chkAllSupp.Value = 1 Then 'All Supplier
        cboSupplier.Text = ""
        txtSuppCode.Text = ""
    End If
End Sub

Private Sub cmd0SRP_Click()
    Dim Qry As String
    
    If rsParts.State = adStateOpen Then rsParts.Close
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST],'Standard') AS [UNIT COST], Format([STOREPRICE],'Standard') AS [SRP], SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format([DLRPRICE],'Standard') AS [DP] " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.STOREPRICE)=0)) ORDER BY SPAREPARTS.PARTNUMBER;"
    
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST,2) AS UNIT_COST, Format(SPAREPARTS.STOREPRICE,2) AS SRP, SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format(SPAREPARTS.DLRPRICE,2) AS DP " & _
          "From SPAREPARTS WHERE SPAREPARTS.STOREPRICE=0 ORDER BY SPAREPARTS.PARTNUMBER;"
    
    
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxParts.Clear
        flxParts.Rows = 2
        flxParts.FormatString = "    |<PART NUMBER  |<DESCRIPTION  |>UNIT COST |>SRP"
        Call RXResizer1_ResizeComplete
    Else
        Set flxParts.DataSource = rsParts
        flxParts.Refresh
        flxParts.ColAlignment(3) = flexAlignRightCenter
        flxParts.ColAlignment(4) = flexAlignRightCenter
    End If
    'format cur
    Call FormatCur
End Sub

Private Sub cmdCancel_Click()
    Dim Ans As Integer, Qry As String
    
    Ans = MsgBox("Do you want to Cancel Editing?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    
    'rsParts.Requery
    cmdEdit.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    cmdRefresh.Enabled = True
    cmd0SRP.Enabled = True
    tabSP.TabEnabled(0) = True
    tabSP.TabEnabled(2) = True
    'tabSP.TabEnabled(3) = True
    txtRefNo.Text = ""
    txtRefNo.Text = flxParts.TextMatrix(flxParts.Row, 7)
    txtCost.Text = flxParts.TextMatrix(flxParts.Row, 3)
    txtPrice.Text = flxParts.TextMatrix(flxParts.Row, 4)
    txtHist.Text = flxParts.TextMatrix(flxParts.Row, 6)
    If Val(flxParts.TextMatrix(flxParts.Row, 8)) > 0 Then
        FctrID = flxParts.TextMatrix(flxParts.Row, 8)
        For i = 1 To flxFactors.Rows - 1
            If FctrID = flxFactors.TextMatrix(i, 6) Then
                cboSPFactor.Text = flxFactors.TextMatrix(i, 2)
                txtMarkUp.Text = flxFactors.TextMatrix(i, 3)
                txtVat.Text = flxFactors.TextMatrix(i, 4)
                txtMaxDisc.Text = flxFactors.TextMatrix(i, 5)
                Exit For
            End If
        Next i
    Else
        cboSPFactor.Text = ""
        txtMarkUp.Text = ""
        txtVat.Text = ""
        txtMaxDisc.Text = ""
    End If
    If Val(Replace(txtDP.Text, ",", "")) > 0 Then
        'Qry = "SELECT SPFACTOR.SPFDESC FROM DPFACTOR INNER JOIN SPFACTOR ON DPFACTOR.SPFID = SPFACTOR.SPFID Where (((DPFACTOR.SPAREPARTID) = " & Val(flxParts.TextMatrix(flxParts.Row, 5)) & ")) ORDER BY DPFACTOR.FDATE;"
        'mySQL
        Qry = "SELECT SPFACTOR.SPFDESC FROM DPFACTOR INNER JOIN SPFACTOR ON DPFACTOR.SPFID = SPFACTOR.SPFID Where DPFACTOR.SPAREPARTID = " & Val(flxParts.TextMatrix(flxParts.Row, 5)) & " ORDER BY DPFACTOR.FDATE;"
        cboDPfactor.Text = GetFieldValue(Qry)
    End If
    txtDP.Text = flxParts.TextMatrix(flxParts.Row, 9)
    
    'txtCost.Text = Format(txtOldCost.Text, "###,##0.00")
    'txtPrice.Text = Format(txtOldPrice.Text, "###,##0.00")
    'txtDP.Text = Format(txtOldDP.Text, "#,##0.00")
    Call LockCost(True)
    txtPParts.Locked = False
    flxParts.Enabled = True
End Sub

Private Sub cmdCriteria_Click()
    Dim Qry As String

    SPcriteria.Show vbModal
    If lblSQL.Tag = "Cancel" Then Exit Sub
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST],'Standard') AS [UNIT COST], Format([STOREPRICE],'Standard') AS [SRP], SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format([DLRPRICE],'Standard') AS [DP] " & _
          "From SPAREPARTS " & lblSQL.Caption & " ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST,2) AS `UNIT COST`, Format(SPAREPARTS.STOREPRICE,2) AS SRP, SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format(SPAREPARTS.DLRPRICE,2) AS DP " & _
          "From SPAREPARTS " & lblSQL.Caption & " ORDER BY SPAREPARTS.PARTNUMBER;"
    
    If rsParts.State = adStateOpen Then rsParts.Close
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxParts.Clear
        flxParts.Rows = 2
        flxParts.FormatString = "   |<PART NUMBER    |<DESCRIPTION   |^ UNIT COST   |^SRP      "
        Call RXResizer1_ResizeComplete
    Else
        Set flxParts.DataSource = rsParts
        flxParts.Refresh
    End If
    flxParts.ColAlignment(3) = flexAlignRightCenter
    flxParts.ColAlignment(4) = flexAlignRightCenter
    txtTtlGMember.Text = rsParts.RecordCount
    'format cur
    Call FormatCur
End Sub

Private Sub cmdDcancel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Cancel Wholesale Discount?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    Call ReloadWholeSaleDiscount
        
    cmdDEdit.Visible = True
    cmdDsave.Visible = False
    cmdDcancel.Visible = False
    tabSP.TabEnabled(1) = True
    tabSP.TabEnabled(0) = True
    tabSP.TabEnabled(2) = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    cmdRefresh.Enabled = True
    cmd0SRP.Enabled = True
    flxGroup.Enabled = True

End Sub

Private Sub cmdDEdit_Click()
    Dim Ans As Integer
    Dim MyAcc As String, Qry As String
    
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=37));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=37;"
    
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Edit Wholesale Discount?", vbQuestion + vbYesNo, "Edit")
    If Ans = vbNo Then Exit Sub
    DWLock = False
    cmdDEdit.Visible = False
    cmdDsave.Visible = True
    cmdDcancel.Visible = True
    tabSP.TabEnabled(1) = False
    tabSP.TabEnabled(0) = False
    tabSP.TabEnabled(2) = False
    cmdExit.Enabled = False
    cmdFind.Enabled = False
    cmdRefresh.Enabled = False
    cmd0SRP.Enabled = False
    flxGroup.Enabled = False

End Sub

Private Sub cmdDsave_Click()
    Dim rsSettings As New ADODB.Recordset
    Dim Ans As Integer, i As Integer
    
    Ans = MsgBox("Do you want to Save Wholesale Discount?", vbQuestion + vbYesNo, "Save")
    If Ans = vbNo Then Exit Sub
        
    'LYRcon.Execute "DELETE SETTINGS.* From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='Wholesale Discount'));"
    'mySQL
    LYRcon.Execute "DELETE From SETTINGS WHERE SETTINGS.DESCRIPTION='Wholesale Discount';"
    
    With rsSettings
        '.Open "Settings", LYRcon, adOpenKeyset, adLockOptimistic
        For i = 1 To flxWDisc.Rows - 1
            If flxWDisc.TextMatrix(i, 1) <> "" And flxWDisc.TextMatrix(i, 2) <> "" Then
                '.AddNew
                '!Description = "Wholesale Discount"
                '!Value = flxWDisc.TextMatrix(i, 2)
                '!Remarks = flxWDisc.TextMatrix(i, 1)
                '!UpdatedBy = UserID
                '!LastUpdate = Now
                '.Update
                
                'mySQL
                LYRcon.Execute "INSERT INTO settings(DESCRIPTION, VALUE, REMARKS, UPDATEDBY, LASTUPDATE)" & _
                    "VALUES('Wholesale Discount', '" & flxWDisc.TextMatrix(i, 2) & "', '" & flxWDisc.TextMatrix(i, 1) & "', " & _
                    "" & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"

            End If
        Next i
        .Close
    End With
        
    'LYRcon.Execute "DELETE SETTINGS.* From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='SP Disclusion'));"
    'mySQL
    LYRcon.Execute "DELETE From SETTINGS WHERE SETTINGS.DESCRIPTION='SP Disclusion';"
    
    'With rsSettings
        '.Open "Settings", LYRcon, adOpenKeyset, adLockOptimistic
        For i = 1 To flxSPec.Rows - 1
            If flxSPec.TextMatrix(i, 4) <> "" And flxSPec.TextMatrix(i, 3) <> "" Then
                '.AddNew
                '!Description = "SP Disclusion"
                '!Value = flxSPec.TextMatrix(i, 3)
                '!Remarks = flxSPec.TextMatrix(i, 4)
                '!UpdatedBy = UserID
                '!LastUpdate = Now
                '.Update
                
                'mySQL
                lyrgcon.Execute "Insert Into settings(DESCRIPTION, VALUE, REMARKS, UPDATEDBY, LASTUPDATE)" & _
                    "VALUES('SP Disclusion', '" & flxSPec.TextMatrix(i, 3) & "', '" & flxSPec.TextMatrix(i, 4) & "', " & _
                    "" & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"

            End If
        Next i
        '.Close
    'End With
    
    cmdDEdit.Visible = True
    cmdDsave.Visible = False
    cmdDcancel.Visible = False
    tabSP.TabEnabled(1) = True
    tabSP.TabEnabled(0) = True
    tabSP.TabEnabled(2) = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    cmdRefresh.Enabled = True
    cmd0SRP.Enabled = True
    flxGroup.Enabled = True
    
    MsgBox "Wholesale Discount successfully saved!", vbInformation + vbOKOnly, "Save"
End Sub

Private Sub cmdEdit_Click()
    Dim MyAcc As String, Qry As String
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=37));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=37;"
    
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If

    If Not EditMode Then
        Ans = MsgBox("Do you want to edit Cost?", vbQuestion + vbYesNo, "Edit")
        If Ans = vbNo Then Exit Sub
    End If
    txtOldRef.Text = txtRefNo.Text
    txtOldCost.Text = Val(Replace(txtCost.Text, ",", ""))
    txtOldPrice.Text = Val(Replace(txtPrice.Text, ",", ""))
    If txtHist.Text = "" Then
        txtPDate.Text = "1/1/2008"
        txtOldRef.Text = "BEG. PRICE"
    End If
    cmdEdit.Visible = False
    cmdSave.Visible = True
    cmdCancel.Visible = True
    cmdExit.Enabled = False
    cmdFind.Enabled = False
    cmdRefresh.Enabled = False
    cmd0SRP.Enabled = False
    tabSP.TabEnabled(0) = False
    tabSP.TabEnabled(2) = False
    'tabSP.TabEnabled(3) = False
    txtPParts.Locked = True
    flxParts.Enabled = False
    Call LockCost(False)
    txtRefNo.SetFocus
    EditMode = True
End Sub

Private Sub LockCost(Bln As Boolean)
    txtCost.Locked = Bln
    txtRefNo.Locked = Bln
    cboSPFactor.Locked = Bln
    txtPrice.Locked = Bln
    cboDPfactor.Locked = Bln
    txtDP.Locked = Bln
    
    'Frame2.Enabled = Not Bln
    'FraSRP.Enabled = Not Bln
    'FraDP.Enabled = Not Bln

End Sub

Private Sub cmdExit_Click()
    MotoRep.tbMain.Visible = True
    Unload Me
End Sub

Private Sub cmdFCancel_Click()
    Dim Ans As Integer, i As Integer
    
    Ans = MsgBox("Do you want to Cancel Factor?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    
    If NewFactor Then
        txtFDesc.Text = ""
        txtFDate.Text = ""
        txtFMarkUp.Text = ""
        txtFVat.Text = ""
        txtFMaxDisc.Text = ""
        txtSPFid.Text = ""
    Else
        With flxFactors
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 6) = txtSPFid.Text Then
                    txtFDesc.Text = .TextMatrix(i, 2)
                    txtFDate.Text = .TextMatrix(i, 1)
                    txtFMarkUp.Text = .TextMatrix(i, 3)
                    txtFVat.Text = .TextMatrix(i, 4)
                    txtFMaxDisc.Text = .TextMatrix(i, 5)
                    txtSPFid.Text = .TextMatrix(i, 6)
                    Exit For
                End If
            Next i
        End With
    End If
    fraFactor.Enabled = False
    cmdFNew.Visible = True
    cmdFEdit.Visible = True
    cmdFSave.Visible = False
    cmdFCancel.Visible = False
    tabSP.TabEnabled(1) = True
    tabSP.TabEnabled(2) = True
    'tabSP.TabEnabled(3) = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    cmdRefresh.Enabled = True
    cmd0SRP.Enabled = True
    
End Sub

Private Sub cmdFEdit_Click()
    Dim Ans As Integer
    Dim MyAcc As String, Qry As String
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=37));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=37;"
    
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    If txtSPFid.Text = "" Then
        MsgBox "Please select a Price a Factor to edit!", vbExclamation + vbOKOnly, "Blank Price Factor"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to Edit Factor?", vbQuestion + vbYesNo, "Edit")
    If Ans = vbNo Then Exit Sub
    cmdFNew.Visible = False
    cmdFEdit.Visible = False
    cmdFSave.Visible = True
    cmdFCancel.Visible = True
    fraFactor.Enabled = True
    tabSP.TabEnabled(1) = False
    tabSP.TabEnabled(2) = False
    'tabSP.TabEnabled(3) = False
    cmdExit.Enabled = False
    cmdFind.Enabled = False
    cmdRefresh.Enabled = False
    cmd0SRP.Enabled = False
    txtFDesc.SetFocus
    NewFactor = False
End Sub

Private Sub cmdFind_Click()
    fraFind.Top = cmdFind.Height + 20
    txtFind.Text = ""
    fraFind.Visible = True
    txtFind.SetFocus
End Sub

Private Sub cmdFNew_Click()
    Dim MyAcc As String, Qry As String
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=37));"
    
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=37;"
    
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    'ClearEntries
    txtFDesc.Text = "" 'last touch here
    txtFDate.Text = ""
    txtFMarkUp.Text = ""
    txtFVat.Text = ""
    txtFMaxDisc.Text = ""
    txtSPFid.Text = ""
    cmdFNew.Visible = False
    cmdFEdit.Visible = False
    cmdFSave.Visible = True
    cmdFCancel.Visible = True
    fraFactor.Enabled = True
    tabSP.TabEnabled(1) = False
    tabSP.TabEnabled(2) = False
    'tabSP.TabEnabled(3) = False
    cmdExit.Enabled = False
    cmdFind.Enabled = False
    cmdRefresh.Enabled = False
    cmd0SRP.Enabled = False
    txtFDate.Text = Date
    txtFDesc.SetFocus
    NewFactor = True
    
End Sub

Private Sub cmdFSave_Click()
    Dim myMsg As String, Qry  As String
    Dim Ans As Integer
    
    myMsg = ValidateFactor
    If myMsg <> "" Then
        MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to save changes?", vbQuestion + vbYesNo, "Save")
    If Ans = vbNo Then Exit Sub
    'Access
    'If NewFactor Then
    '    rsFactor.Open "SPFACTOR", LYRcon, adOpenKeyset, adLockOptimistic
    '    rsFactor.AddNew
    'Else
    '    Qry = "SELECT SPFACTOR.* From SPFACTOR WHERE (((SPFACTOR.SPFID)=" & txtSPFid.Text & "));"
    '    rsFactor.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    rsFactor.MoveFirst
    'End If
    'rsFactor!SPFDesc = txtFDesc.Text
    'rsFactor!SPFDate = txtFDate.Text
    'If InStr(txtFMarkUp.Text, "%") > 0 Then
    '    rsFactor!Markup = Val(Left$(txtFMarkUp.Text, Len(txtFMarkUp.Text) - 1))
    'Else
    '    rsFactor!Markup = Val(txtFMarkUp.Text)
    '    txtFMarkUp.Text = txtFMarkUp.Text & "%"
    'End If
    'If InStr(txtFVat.Text, "%") > 0 Then
    '    rsFactor!Vat = Val(Left$(txtFVat.Text, Len(txtFVat.Text) - 1))
    'Else
    '    rsFactor!Vat = Val(txtFVat.Text)
    '    txtFVat.Text = txtFVat.Text & "%"
    'End If
    'old rsFactor!Vat = IIf(Right$(txtFVat.Text, 1) = "%", Val(Left$(txtFVat.Text, Len(txtFVat.Text) - 1)), Val(txtFVat.Text))
    'rsFactor!MaxDiscount = Val(txtFMaxDisc.Text)
    'rsFactor.Update
    
    'mySQL
    If InStr(txtFMarkUp.Text, "%") > 0 Then
        MyMarkUp = Val(Left$(txtFMarkUp.Text, Len(txtFMarkUp.Text) - 1))
    Else
        MyMarkUp = Val(txtFMarkUp.Text)
        txtFMarkUp.Text = txtFMarkUp.Text & "%"
    End If
    If InStr(txtFVat.Text, "%") > 0 Then
        myVat = Val(Left$(txtFVat.Text, Len(txtFVat.Text) - 1))
    Else
        myVat = Val(txtFVat.Text)
        txtFVat.Text = txtFVat.Text & "%"
    End If
    myMaxDiscount = Val(txtFMaxDisc.Text)

    If NewFactor Then
    
        LYRcon.Execute "Insert Into spfactor(SPFDESC, SPFDATE, MARKUP, VAT, MAXDISCOUNT) " & _
            "VALUES('" & txtFDesc.Text & "', '" & Format(txtFDate.Text, "yyyy-mm-dd") & "', " & _
            "" & MyMarkUp & ", " & myVat & ", " & myMaxDiscount & ");"
    
    Else
        LYRcon.Execute "Update spfactor SET SPFDESC = '" & txtFDesc.Text & "', " & _
            "SPFDATE = '" & Format(txtFDate.Text, "yyyy-mm-dd") & "', MARKUP = " & MyMarkUp & ", " & _
            "VAT = " & myVat & ", MAXDISCOUNT = " & myMaxDiscount & ", " & _
            " Where spfactor.SPFID = " & txtSPFid.Text & ";"
    End If
    '
    
    If NewFactor Then
        'txtSPFid.Text = rsFactor!SPFid
        'Retrieve sa last entry sa factor but this intended TO ONE USER
        Qry = "SELECT MAX(spfactor.SPFID) As LstSPFID From spfactor;"
        txtSPFid.Text = GetFieldValue(Qry)
        
        flxFactors.AddItem vbTab & txtFDate.Text & vbTab & txtFDesc.Text & vbTab & txtFMarkUp.Text & vbTab & txtFVat.Text & vbTab & txtFMaxDisc.Text & vbTab & txtSPFid.Text, 1
        cboSPFactor.AddItem txtFDesc.Text
        cboGFactor.AddItem txtFDesc.Text
    Else
        With flxFactors
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 6) = txtSPFid.Text Then
                    .TextMatrix(i, 2) = txtFDesc.Text   'last touch here
                    .TextMatrix(i, 1) = txtFDate.Text
                    .TextMatrix(i, 3) = txtFMarkUp.Text
                    .TextMatrix(i, 4) = txtFVat.Text
                    .TextMatrix(i, 5) = txtFMaxDisc.Text
                    .TextMatrix(i, 6) = txtSPFid.Text
                    Exit For
                End If
            Next i
            cboSPFactor.Clear
            cboGFactor.Clear
            For i = 1 To .Rows - 1
                cboSPFactor.AddItem .TextMatrix(i, 2)
                cboGFactor.AddItem .TextMatrix(i, 2)
            Next i
        End With
    End If
    'rsFactor.Close
    MsgBox "Price Factor successfully saved!", vbInformation + vbOKOnly, "Save"
    fraFactor.Enabled = False
    cmdFNew.Visible = True
    cmdFEdit.Visible = True
    cmdFSave.Visible = False
    cmdFCancel.Visible = False
    tabSP.TabEnabled(1) = True
    tabSP.TabEnabled(2) = True
    'tabSP.TabEnabled(3) = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    cmdRefresh.Enabled = True
    cmd0SRP.Enabled = True

End Sub

Private Sub cmdGCancel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Cancel Group Factor?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    txtGDesc.Text = ""
    txtGDate.Text = ""
    cboGFactor.Text = ""
    txtGmrk.Text = ""
    txtGVat.Text = ""
    txtGFactorID.Text = ""
    lblGFctrDtls.Caption = ""
    lblSQL.Caption = ""
    lblCriteria.Caption = ""
    txtTtlGMember.Text = ""
    txtGEditRow.Text = ""
    Call cmdRefresh_Click

    Frame3.Enabled = False
    cmdGNew.Visible = True
    cmdGEdit.Visible = True
    cmdGSave.Visible = False
    cmdGCancel.Visible = False
    tabSP.TabEnabled(1) = True
    tabSP.TabEnabled(0) = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    cmdRefresh.Enabled = True
    cmd0SRP.Enabled = True
    flxGroup.Enabled = True

End Sub

Private Sub cmdGEdit_Click()
    Dim Ans As Integer
    Dim MyAcc As String, Qry As String
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=37));"
    'mySQL
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=37;"
    
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Edit Group Factor?", vbQuestion + vbYesNo, "Edit")
    If Ans = vbNo Then Exit Sub
    GrpAddNew = False
    Frame3.Enabled = True
    cmdGNew.Visible = False
    cmdGEdit.Visible = False
    cmdGSave.Visible = True
    cmdGCancel.Visible = True
    tabSP.TabEnabled(1) = False
    tabSP.TabEnabled(0) = False
    cmdExit.Enabled = False
    cmdFind.Enabled = False
    cmdRefresh.Enabled = False
    cmd0SRP.Enabled = False
    flxGroup.Enabled = False

End Sub

Private Sub cmdGNew_Click()
    Dim MyAcc As String, Qry As String
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=37));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=37;"
    
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    'Clear objects for entries
    txtGDesc.Text = ""
    txtGDate.Text = ""
    cboGFactor.Text = ""
    txtGmrk.Text = ""
    txtGVat.Text = ""
    txtGFactorID.Text = ""
    lblGFctrDtls.Caption = ""
    lblSQL.Caption = ""
    lblCriteria.Caption = ""
    txtTtlGMember.Text = ""
    txtGEditRow.Text = ""
        
    Frame3.Enabled = True
    cmdGNew.Visible = False
    cmdGEdit.Visible = False
    cmdGSave.Visible = True
    cmdGCancel.Visible = True
    tabSP.TabEnabled(1) = False
    tabSP.TabEnabled(0) = False
    cmdExit.Enabled = False
    cmdFind.Enabled = False
    cmdRefresh.Enabled = False
    cmd0SRP.Enabled = False
    flxGroup.Enabled = False
    GrpAddNew = True
    txtGEditRow.Text = ""
    txtGDate.Text = Date
    txtGDesc.SetFocus
End Sub

Private Sub cmdGo_Click()
    If txtSuppCode.Text = "" And txtDesc.Text = "" Then
        Call cmdRefresh_Click
    Else
        Call SearchPartsDesc(0)
    End If
End Sub

Private Sub cmdGSave_Click()
    Dim rsGFctr As New ADODB.Recordset
    Dim vMsg As String, vQry As String, strHist As String
    Dim Ans As Integer, myDisc As Integer
    Dim i As Integer, TtlRec As Integer, Pos As Integer
    Dim pPct As Single, CostVat As Currency, NewPrice As Currency, MyCost As Currency
    
    'Validate Group
    vMsg = ""
    If txtGDesc.Text = "" Then
        vMsg = "Incomplete entry on field Description"
    Else
        If GrpAddNew Then
            'vQry = "SELECT SPGFACTOR.GDESC From SPGFACTOR WHERE (((SPGFACTOR.GDESC)='" & txtGDesc.Text & "'));"
            'mySQL
            vQry = "SELECT SPGFACTOR.GDESC From SPGFACTOR WHERE SPGFACTOR.GDESC='" & txtGDesc.Text & "';"
            
            If CtrlNumExist(vQry) Then
                vMsg = "A Group factor already exist with the same Description"
            End If
        End If
    End If
    If txtGFactorID.Text = "" Then vMsg = vMsg & vbCrLf & "Invalid Price Factor - Please select from list"
    If Not IsDate(txtGDate.Text) Then vMsg = vMsg & vbCrLf & "Invalid value of Date Created"
    If lblSQL.Caption = "" Then vMsg = vMsg & vbCrLf & "No Criteria available!"
    If vMsg <> "" Then
        MsgBox vMsg, vbCritical + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to Save Group Factor?", vbQuestion + vbYesNo, "Save")
    If Ans = vbNo Then Exit Sub
    'Start Saving
    'With rsGFctr
        'If GrpAddNew Then
        '    .Open "SPGFactor", LYRcon, adOpenKeyset, adLockOptimistic
        '    .AddNew
        'Else
        '    vQry = "SELECT SPGFACTOR.* From SPGFACTOR WHERE (((SPGFACTOR.GDESC)='" & flxGroup.TextMatrix(Val(txtGEditRow.Text), 2) & "'));"
        '    .Open vQry, LYRcon, adOpenKeyset, adLockOptimistic
        '    .MoveFirst
        'End If
        '!SPFid = txtGFactorID.Text
        '!GDesc = txtGDesc.Text
        '!GDate = txtGDate.Text
        '!SQLst = lblSQL.Caption
        '!GCriteria = lblCriteria.Caption
        '.Update
        '.Close
    'End With
    'update in flx
    If GrpAddNew Then
        'mySQL
        LYRcon.Execute "Insert Into spgfactor(SPFID, GDESC, GDATE, SQLST, GCRITERIA)" & _
        "VALUES(" & txtGFactorID.Text & ", '" & txtGDesc.Text & "', '" & txtGDate.Text & "', " & _
        "'" & lblSQL.Caption & "', '" & lblCriteria.Caption & "');"
    Else
        'mySQL
        LYRcon.Execute "Update spgfactor SET SPFID = " & txtGFactorID.Text & ", GDESC = '" & txtGDesc.Text & "', GDATE = " & txtGDate.Text & ", SQLST = '" & lblSQL.Caption & "', " & _
        "GCRITERIA = '" & lblCriteria.Caption & "' Where spgfactor.GDESC = '" & flxGroup.TextMatrix(Val(txtGEditRow.Text), 2) & "';"
    End If
    
    
    If GrpAddNew Then
        If flxGroup.Rows = 2 And flxGroup.TextMatrix(1, 2) = "" Then flxGroup.Rows = 1
        flxGroup.AddItem vbTab & txtGDate.Text & vbTab & txtGDesc.Text & vbTab & cboGFactor.Text & vbTab & lblCriteria.Caption & vbTab & txtGFactorID.Text & vbTab & lblSQL.Caption
        txtGEditRow.Text = flxGroup.Rows - 1
    Else
        
        flxGroup.TextMatrix(Val(txtGEditRow.Text), 1) = txtGDate.Text
        flxGroup.TextMatrix(Val(txtGEditRow.Text), 2) = txtGDesc.Text
        flxGroup.TextMatrix(Val(txtGEditRow.Text), 3) = cboGFactor.Text
        flxGroup.TextMatrix(Val(txtGEditRow.Text), 4) = lblCriteria.Caption
        flxGroup.TextMatrix(Val(txtGEditRow.Text), 5) = txtGFactorID.Text
        flxGroup.TextMatrix(Val(txtGEditRow.Text), 6) = lblSQL.Caption
    End If
    flxGroup.ColWidth(5) = 0
    flxGroup.ColWidth(6) = 0
    'Process Updating of Cost and Prices ... Start Progress bar
    With flxParts
        .Enabled = False
        tabProg.Top = (.Top + .Height) \ 2
        tabProg.Left = (.Width - tabProg.Width) \ 2
        tabProg.Caption = ""
        tabProg.Visible = True
        tabProg.ZOrder
        Pbar.Value = 0
        Pbar.Max = .Rows - 1
        For i = 1 To .Rows - 1
            '1-PARTNO, 2-DESC, 3-COST, 4-SRP, 5-SPID, 6-PRICEHIST, 7-PRICEREF, 8-SPFID
            If Val(.TextMatrix(i, 3)) > 0 Then
                MyCost = PureNum(.TextMatrix(i, 3))
                CostVat = MyCost + (MyCost * (Val(txtGVat.Text) / 100))
                NewPrice = CostVat + (CostVat * (Val(txtGmrk.Text) / 100))
                NewPrice = Int(NewPrice + 0.6) + 1
                strHist = IIf(.TextMatrix(i, 6) = "", "1/1/2008", Date) & "^" & IIf(.TextMatrix(i, 7) = "", "BEG. PRICE", .TextMatrix(i, 7)) & "^" & .TextMatrix(i, 3) & "^" & .TextMatrix(i, 4) & IIf(txtHist.Text = "", "", "*" & .TextMatrix(i, 6))
                If Len(strHist) >= 250 Then
                    Pos = InStrRev(strHist, "*")
                    strHist = Left$(strHist, Len(strHist) - Pos) '  Mid$(strHist, Pos + 1, Len(strHist) - Pos)
                End If
                If InStr(txtGDisc.Text, "%") Then
                    myDisc = Val(Left$(txtGDisc.Text, Len(txtGDisc.Text) - 1))
                Else
                    myDisc = Val(txtGDisc.Text)
                End If
                
                'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & MyCost & ", SPAREPARTS.STOREPRICE = " & NewPrice & ", SPAREPARTS.PRICEHIST = '" & strHist & "', SPAREPARTS.PRICEREF = '" & Trim(Left$(txtGDesc.Text, 20)) & "', SPAREPARTS.SPFID = '" & txtGFactorID.Text & "', SPAREPARTS.MAXDISCOUNT = " & myDisc & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
                'mySQL
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & MyCost & ", SPAREPARTS.STOREPRICE = " & NewPrice & ", SPAREPARTS.PRICEHIST = '" & strHist & "', SPAREPARTS.PRICEREF = '" & Trim(Left$(txtGDesc.Text, 20)) & "', SPAREPARTS.SPFID = '" & txtGFactorID.Text & "', SPAREPARTS.MAXDISCOUNT = " & myDisc & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
                
                'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & MyCost & ", SOURCESP.RETAILPRICE = " & NewPrice & _
                              " WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                'mySQL
                LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & MyCost & ", SOURCESP.RETAILPRICE = " & NewPrice & _
                              " WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                
                flxParts.TextMatrix(i, 4) = Format$(NewPrice, "###,##0.00")
                flxParts.TextMatrix(i, 6) = strHist
                flxParts.TextMatrix(i, 7) = Trim(Left$(txtGDesc.Text, 20))
                flxParts.TextMatrix(i, 8) = txtGFactorID.Text
            Else
                NewPrice = 0
            End If
            Pbar.Value = Pbar.Value + 1
            pPct = (Pbar.Value / Pbar.Max)
            DoEvents
            lblPct.Caption = Format$(pPct, "##0.0%")
        Next i
        tabProg.Visible = False
        DoEvents
        tabProg.Top = 12560
        .Enabled = True
    End With
    MsgBox "Group Factor successfully updated!", vbInformation + vbOKOnly, "Save"
    Frame3.Enabled = False
    cmdGNew.Visible = True
    cmdGEdit.Visible = True
    cmdGSave.Visible = False
    cmdGCancel.Visible = False
    tabSP.TabEnabled(1) = True
    tabSP.TabEnabled(0) = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    cmdRefresh.Enabled = True
    cmd0SRP.Enabled = True
    flxGroup.Enabled = True
End Sub

Private Sub cmdRefresh_Click()
    Dim Qry As String
    
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST],'Standard') AS [UNIT COST], Format([STOREPRICE],'Standard') AS [SRP], SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format([DLRPRICE],'Standard') AS [DP] " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST,2) AS `UNIT COST`, Format(SPAREPARTS.STOREPRICE,2) AS SRP, SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format(SPAREPARTS.DLRPRICE,2) AS DP " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    
    If rsParts.State = adStateOpen Then rsParts.Close
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxParts.Clear
        flxParts.Rows = 2
        flxParts.FormatString = "   |<PART NUMBER    |<DESCRIPTION   |^ UNIT COST   |^SRP      "
        Call RXResizer1_ResizeComplete
    Else
        Set flxParts.DataSource = rsParts
        flxParts.Refresh
    End If
    flxParts.ColAlignment(3) = flexAlignRightCenter
    flxParts.ColAlignment(4) = flexAlignRightCenter
    fraFind.Visible = False
    'format cur
    Call FormatCur
End Sub

Private Sub cmdSave_Click()
    Dim Ans As Integer, myDisc As Integer
    
    If txtRefNo.Text = "" Then
        MsgBox "Incomplete entry on field Reference No.", vbInformation + vbOKOnly, "Save"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to save changes?", vbQuestion + vbYesNo, "Save")
    If Ans = vbNo Then
        rsParts.Requery
        txtPParts.Locked = False
        Exit Sub
    End If
    strHist = IIf(txtHist.Text = "", txtPDate.Text, Date) & "^" & txtOldRef.Text & "^" & txtOldCost.Text & "^" & txtOldPrice.Text & IIf(txtHist.Text = "", "", "*" & txtHist.Text)
    If Len(strHist) >= 250 Then
        Pos = InStrRev(strHist, "*")
        strHist = Left$(strHist, Len(strHist) - Pos) '  Mid$(strHist, Pos + 1, Len(strHist) - Pos)
    End If
    If InStr(txtMaxDisc.Text, "%") Then
        myDisc = Val(Left$(txtMaxDisc.Text, Len(txtMaxDisc.Text) - 1))
    Else
        myDisc = Val(txtMaxDisc.Text)
    End If
    'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & Val(Replace(txtCost.Text, ",", "")) & ", SPAREPARTS.STOREPRICE = " & Val(Replace(txtPrice.Text, ",", "")) & ", SPAREPARTS.PRICEHIST = '" & strHist & "', SPAREPARTS.PRICEREF = '" & txtRefNo.Text & "', SPAREPARTS.SPFID = " & Val(txtFactorID.Text) & ", SPAREPARTS.MAXDISCOUNT = " & myDisc & ", SPAREPARTS.DLRPRICE = " & Val(Replace(txtDP.Text, ",", "")) & _
                   " WHERE (((SPAREPARTS.PARTNUMBER)='" & txtPParts.Text & "'));"
    'mySQL
    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & Val(Replace(txtCost.Text, ",", "")) & ", SPAREPARTS.STOREPRICE = " & Val(Replace(txtPrice.Text, ",", "")) & ", SPAREPARTS.PRICEHIST = '" & strHist & "', SPAREPARTS.PRICEREF = '" & txtRefNo.Text & "', SPAREPARTS.SPFID = " & Val(txtFactorID.Text) & ", SPAREPARTS.MAXDISCOUNT = " & myDisc & ", SPAREPARTS.DLRPRICE = " & Val(Replace(txtDP.Text, ",", "")) & _
                   " WHERE SPAREPARTS.PARTNUMBER='" & txtPParts.Text & "';"
    
    'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & Val(Replace(txtCost.Text, ",", "")) & ", SOURCESP.RETAILPRICE = " & Val(Replace(txtPrice.Text, ",", "")) & _
                  " WHERE (((SOURCESP.PARTNUMBER)='" & txtPParts.Text & "'));"
    'mySQL
    LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & Val(Replace(txtCost.Text, ",", "")) & ", SOURCESP.RETAILPRICE = " & Val(Replace(txtPrice.Text, ",", "")) & _
                  " WHERE SOURCESP.PARTNUMBER='" & txtPParts.Text & "';"
    
    'Do Not Save if Dealer Price same with last Dealer
    cboHistory.ListIndex = 1
    If flxOld.Rows = 1 And Val(Replace(txtCost.Text, ",", "")) > 0 Then
        LYRcon.Execute "INSERT INTO DPFACTOR (FDATE, SPFID, COST, DP, SPAREPARTID, REFNO) " & _
                   "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', " & Val(txtDPfactorID.Text) & ", " & Val(Replace(txtCost.Text, ",", "")) & ", " & Val(Replace(txtDP.Text, ",", "")) & ", " & Val(flxParts.TextMatrix(flxParts.Row, 5)) & ", '" & txtRefNo.Text & "');"
        
        flxParts.TextMatrix(flxParts.Row, 9) = Format(txtDP.Text, "#,##0.00")
        
    Else
        If Val(Replace(txtCost.Text, ",", "")) <> Val(Replace(flxOld.TextMatrix(1, 2), ",", "")) Or Val(Replace(txtDP.Text, ",", "")) <> Val(Replace(flxOld.TextMatrix(1, 3), ",", "")) Then
            'LYRcon.Execute "INSERT INTO DPFACTOR (FDATE, SPFID, COST, DP, SPAREPARTID, REFNO) " & _
                       "VALUES (#" & Date & "#, " & Val(txtDPfactorID.Text) & ", " & Val(Replace(txtCost.Text, ",", "")) & ", " & Val(Replace(txtDP.Text, ",", "")) & ", " & Val(flxParts.TextMatrix(flxParts.Row, 5)) & ", '" & txtRefNo.Text & "');"
            'mySQL
            LYRcon.Execute "INSERT INTO DPFACTOR (FDATE, SPFID, COST, DP, SPAREPARTID, REFNO) " & _
                       "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', " & Val(txtDPfactorID.Text) & ", " & Val(Replace(txtCost.Text, ",", "")) & ", " & Val(Replace(txtDP.Text, ",", "")) & ", " & Val(flxParts.TextMatrix(flxParts.Row, 5)) & ", '" & txtRefNo.Text & "');"
            
            flxParts.TextMatrix(flxParts.Row, 9) = Format(txtDP.Text, "#,##0.00")
        End If
    End If
    flxParts.TextMatrix(flxParts.Row, 3) = txtCost.Text
    flxParts.TextMatrix(flxParts.Row, 4) = txtPrice.Text
    flxParts.TextMatrix(flxParts.Row, 6) = strHist
    flxParts.TextMatrix(flxParts.Row, 7) = txtRefNo.Text
    flxParts.TextMatrix(flxParts.Row, 8) = Val(txtFactorID.Text)
    flxOld.AddItem IIf(txtPDate.Text = "", Date, txtPDate.Text) & vbTab & txtOldRef.Text & vbTab & Format$(txtOldCost.Text, "###,##0.00") & vbTab & Format$(txtOldPrice.Text, "###,##0.00"), 1
    Ans = MsgBox("New Price successfully saved!" & vbCrLf & vbCrLf & _
          "Do you want to edit another Part Number?", vbQuestion + vbYesNo, "Save")
    If Ans = vbYes Then
        txtPParts.Locked = False
        flxParts.Enabled = True
        Call LockCost(False)
        txtPParts.SetFocus
        txtPParts.SelStart = 0
        txtPParts.SelLength = Len(txtPParts.Text)
        EditMode = True
    Else
        rsParts.Requery
        cmdEdit.Visible = True
        cmdSave.Visible = False
        cmdCancel.Visible = False
        cmdExit.Enabled = True
        cmdFind.Enabled = True
        cmdRefresh.Enabled = True
        cmd0SRP.Enabled = True
        
        tabSP.TabEnabled(0) = True
        tabSP.TabEnabled(2) = True
        'tabSP.TabEnabled(3) = True
        Call LockCost(True)
        txtPParts.Locked = True
        flxParts.Enabled = True
        EditMode = False
    End If

End Sub

Private Sub flxFactors_Click()
    Dim i As Integer
    
    With flxFactors
        i = .Row
        If i = 0 Then Exit Sub
        txtFDesc.Text = .TextMatrix(i, 2)
        txtFDate.Text = .TextMatrix(i, 1)
        txtFMarkUp.Text = .TextMatrix(i, 3)
        txtFVat.Text = .TextMatrix(i, 4)
        txtFMaxDisc.Text = .TextMatrix(i, 5)
        txtSPFid.Text = .TextMatrix(i, 6)
    End With
End Sub

Private Sub flxGroup_Click()
    
    With flxGroup
        If .Row = 0 Or .TextMatrix(.Row, 2) = "" Then Exit Sub
        'Display
        txtGDesc.Text = .TextMatrix(.Row, 2)
        txtGDate.Text = .TextMatrix(.Row, 1)
        cboGFactor.Text = .TextMatrix(.Row, 3)
        lblCriteria.Caption = .TextMatrix(.Row, 4)
        txtGFactorID.Text = flxGroup.TextMatrix(.Row, 5)
        lblSQL.Caption = flxGroup.TextMatrix(.Row, 6)
        Call cboGFactor_Change
        txtGEditRow.Text = .Row
    End With
End Sub

Private Sub flxGroup_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuGrpF)
    End If
End Sub

Function GetFactorDesc(MyFid As Integer) As String
    Dim i As Integer
    
    For i = 1 To flxFactors.Rows - 1
        If MyFid = flxFactors.TextMatrix(i, 6) Then
            cboSPFactor.Text = flxFactors.TextMatrix(i, 2)
            txtMarkUp.Text = flxFactors.TextMatrix(i, 3)
            txtVat.Text = flxFactors.TextMatrix(i, 4)
            txtMaxDisc.Text = flxFactors.TextMatrix(i, 5)
            Exit For
        End If
    Next i
    
End Function

Private Sub flxParts_Click()
    Dim i As Integer, FctrID As Integer, Qry As String
    
    lblDescription.Caption = flxParts.TextMatrix(flxParts.Row, 2)
    txtRefNo.Text = flxParts.TextMatrix(flxParts.Row, 7)
    txtCost.Text = flxParts.TextMatrix(flxParts.Row, 3)
    txtPrice.Text = flxParts.TextMatrix(flxParts.Row, 4)
    txtHist.Text = flxParts.TextMatrix(flxParts.Row, 6)
    If Val(flxParts.TextMatrix(flxParts.Row, 8)) > 0 Then
        FctrID = flxParts.TextMatrix(flxParts.Row, 8)
        For i = 1 To flxFactors.Rows - 1
            If FctrID = flxFactors.TextMatrix(i, 6) Then
                cboSPFactor.Text = flxFactors.TextMatrix(i, 2)
                txtMarkUp.Text = flxFactors.TextMatrix(i, 3)
                txtVat.Text = flxFactors.TextMatrix(i, 4)
                txtMaxDisc.Text = flxFactors.TextMatrix(i, 5)
                Exit For
            End If
        Next i
    Else
        cboSPFactor.Text = ""
        txtMarkUp.Text = ""
        txtVat.Text = ""
        txtMaxDisc.Text = ""
    End If
    txtDP.Text = flxParts.TextMatrix(flxParts.Row, 9)
    If Val(Replace(txtDP.Text, ",", "")) > 0 Then
        'Qry = "SELECT SPFACTOR.SPFDESC FROM DPFACTOR INNER JOIN SPFACTOR ON DPFACTOR.SPFID = SPFACTOR.SPFID Where (((DPFACTOR.SPAREPARTID) = " & Val(flxParts.TextMatrix(flxParts.Row, 5)) & ")) ORDER BY DPFACTOR.FDATE;"
        'mySQL
        Qry = "SELECT dpfactor.SPFID From dpfactor Where dpfactor.SPAREPARTID = " & Val(flxParts.TextMatrix(flxParts.Row, 5)) & " Order By dpfactor.FDATE Desc"
        FctrID = Val(GetFieldValue(Qry))
        If FctrID > 0 Then
            Qry = "SELECT spfactor.SPFDESC From spfactor Where spfactor.SPFID = " & FctrID & ";"
            cboDPfactor.Text = GetFieldValue(Qry)
        Else
            cboDPfactor.Text = ""
        End If
    Else
        cboDPfactor.Text = ""
    End If
    Call DisplayCostHistory(flxParts.TextMatrix(flxParts.Row, 6))
    If EditMode Then
        txtOldRef.Text = txtRefNo.Text
        txtOldCost.Text = Val(Replace(txtCost.Text, ",", ""))
        txtOldPrice.Text = Val(Replace(txtPrice.Text, ",", ""))
        If txtHist.Text = "" Then
            txtPDate.Text = "1/1/2008"
            txtOldRef.Text = "BEG. PRICE"
        End If
    End If
    PrtSrch = True
    txtPParts.Text = flxParts.TextMatrix(flxParts.Row, 1)
'    cboPartNo.Text = flxParts.Text
End Sub

Private Sub flxParts_EnterCell()
    Call flxParts_Click
End Sub

Private Sub flxPartsD_Click()
    Dim Cost As Currency, CostVat As Currency
    
    With flxPartsD
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        Cost = flxPartsD.TextMatrix(flxPartsD.Row, 3)
        CostVat = Cost + (Cost * (Val(txtVat.Text) / 100))
        flxSPec.TextMatrix(ActRow, 2) = Format(CostVat, "#,##0.00") 'Cost
        flxSPec.TextMatrix(ActRow, 4) = flxPartsD.TextMatrix(flxPartsD.Row, 1) 'SPID
    End With
End Sub

Private Sub flxPartsD_EnterCell()
    Dim Cost As Currency, CostVat As Currency
    
    With flxPartsD
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        Cost = flxPartsD.TextMatrix(flxPartsD.Row, 3)
        CostVat = Cost + (Cost * (Val(txtVat.Text) / 100))
        flxSPec.TextMatrix(ActRow, 2) = Format(CostVat, "#,##0.00") 'Cost
        flxSPec.TextMatrix(ActRow, 4) = flxPartsD.TextMatrix(flxPartsD.Row, 1) 'SPID
    End With
End Sub

Private Sub flxPartsD_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        If txtPartNo.Tag <> "" Then
            Tmp = Split(txtPartNo.Tag, "|")
            flxSPec.TextMatrix(ActRow, 1) = Tmp(0)
            flxSPec.TextMatrix(ActRow, 2) = Tmp(1)
            flxSPec.TextMatrix(ActRow, 4) = Tmp(2)
        End If
        txtPartNo.Visible = False
        flxPartsD.Visible = False
        flxSPec.SetFocus
        flxSPec.Col = 1
    End If

End Sub

Private Sub flxPartsD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        If flxSPec.TextMatrix(ActRow, 4) <> "" And flxSPec.TextMatrix(ActRow, 3) <> "" Then
            If ActRow = flxSPec.Rows - 1 Then flxSPec.AddItem ""
            flxSPec.SetFocus
            flxSPec.Row = flxSPec.Rows - 1
            flxSPec.Col = 1
        ElseIf flxSPec.TextMatrix(ActRow, 3) = "" Then   'No COST
            flxSPec.SetFocus
            flxSPec.Col = 3
        End If
        flxPartsD.Visible = False
    End If
End Sub

Private Sub flxSPec_Click()
    ActRow = flxSPec.Row
    ActCol = flxSPec.Col
    
    If ActCol = 0 Or DWLock = True Then Exit Sub 'Or Not FullAccess
    Select Case ActCol
    Case 1
        With txtPartNo
            If flxSPec.TextMatrix(ActRow, 1) <> "" Then .Tag = flxSPec.TextMatrix(ActRow, 1) & "|" & flxSPec.TextMatrix(ActRow, 2) & "|" & flxSPec.TextMatrix(ActRow, 4)
            .Top = flxSPec.CellTop + flxSPec.Top
            .Left = flxSPec.CellLeft + flxSPec.Left
            .Width = flxSPec.CellWidth
            .Text = flxSPec.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 3
        With txtDPrice
            If flxSPec.TextMatrix(ActRow, 2) <> "" Then .Tag = flxSPec.TextMatrix(ActRow, 2)
            .Top = flxSPec.CellTop + flxSPec.Top
            .Left = flxSPec.CellLeft + flxSPec.Left
            .Width = flxSPec.CellWidth
            .Text = Replace(flxSPec.Text, ",", "")
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End Select
End Sub

Private Sub flxSPec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxSPec_Click
End Sub

Private Sub flxSPec_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Not DWLock Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuSpec)
        End If
    End If
End Sub

Private Sub flxSPec_Scroll()
    txtPartNo.Visible = False
    txtDPrice.Visible = False
End Sub

Private Sub flxWDisc_Click()
    ActRow = flxWDisc.Row
    ActCol = flxWDisc.Col
    
    If ActCol = 0 Or DWLock = True Then Exit Sub 'Or Not FullAccessd
    Select Case ActCol
    Case 1
        With cboSPType
            If flxWDisc.TextMatrix(ActRow, 1) <> "" Then .Tag = flxWDisc.TextMatrix(ActRow, 1)
            .Top = flxWDisc.CellTop + flxWDisc.Top
            .Left = flxWDisc.CellLeft + flxWDisc.Left
            .Width = flxWDisc.CellWidth
            .Text = flxWDisc.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 2
        With txtDFctr
            If flxWDisc.TextMatrix(ActRow, 2) <> "" Then .Tag = flxWDisc.TextMatrix(ActRow, 2)
            .Top = flxWDisc.CellTop + flxWDisc.Top
            .Left = flxWDisc.CellLeft + flxWDisc.Left
            .Width = flxWDisc.CellWidth
            .Text = Replace(flxWDisc.Text, ",", "")
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End Select

End Sub

Private Sub flxWDisc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxWDisc_Click
End Sub

Private Sub flxWDisc_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuWDisc)
    End If
End Sub


Private Sub flxWDisc_Scroll()
    txtDFctr.Visible = False
    cboSPType.Visible = False
End Sub

Private Sub Form_Load()
    Dim Qry As String
    
    '1-PARTNO, 2-DESC, 3-COST, 4-SRP, 5-SPID, 6-PRICEHIST, 7-PRICEREF, 8-SPFID
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST],'Standard') AS [UNIT COST], Format([STOREPRICE],'Standard') AS [SRP], SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format([DLRPRICE],'Standard') AS [DP] " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST,2) AS `UNIT COST`, Format(SPAREPARTS.STOREPRICE,2) AS SRP, SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format(SPAREPARTS.DLRPRICE,2) AS DP " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    flxParts.Refresh
    DWLock = True
    Timer1.Enabled = True
    cboHistory.AddItem "RETAIL PRICE HISTORY"
    cboHistory.AddItem "DEALER PRICE HISTORY"
    cboHistory.ListIndex = 0
    Call LockCost(True)
    tabSP.Tab = 0
    'format cur
    Call FormatCur
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsParts.State = adStateOpen Then rsParts.Close
    If rsFactor.State = adStateOpen Then rsFactor.Close
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
    If rsPartsD.State = adStateOpen Then rsPartsD.Close
End Sub

Private Sub mnuGDisp_Click()
    Dim Qry As String
    
    If lblSQL.Caption = "" Then
        MsgBox "No Group Factor to Display!", vbInformation + vbYesNo, "Display Member"
        Exit Sub
    End If
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST],'Standard') AS [UNIT COST], Format([STOREPRICE],'Standard') AS [SRP], SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format([DLRPRICE],'Standard') AS [DP] " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.SPFID)=" & txtGFactorID.Text & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST,2) AS `UNIT COST`, Format(SPAREPARTS.STOREPRICE,2) AS SRP, SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format(SPAREPARTS.DLRPRICE,2) AS DP " & _
          "From SPAREPARTS WHERE SPAREPARTS.SPFID=" & txtGFactorID.Text & " ORDER BY SPAREPARTS.PARTNUMBER;"
    
    If rsParts.State = adStateOpen Then rsParts.Close
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxParts.Clear
        flxParts.Rows = 2
        flxParts.FormatString = "   |<PART NUMBER    |<DESCRIPTION   |^ UNIT COST   |^SRP   |^DP     "
        Call RXResizer1_ResizeComplete
    Else
        Set flxParts.DataSource = rsParts
        flxParts.Refresh
    End If
    flxParts.ColAlignment(3) = flexAlignRightCenter
    flxParts.ColAlignment(4) = flexAlignRightCenter
    txtTtlGMember.Text = rsParts.RecordCount
    'format cur
    Call FormatCur
End Sub

Private Sub mnuSpcAdd_Click()
    If flxSPec.TextMatrix(flxSPec.Rows - 1, 4) = "" Or flxSPec.TextMatrix(flxSPec.Rows - 1, 3) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    End If
    flxSPec.AddItem ""
    flxSPec.SetFocus
    flxSPec.Row = flxSPec.Rows - 1
    flxSPec.Col = 1
End Sub

Private Sub mnuSpcDel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxSPec.Rows = 2 Then
            flxSPec.Rows = 1
            flxSPec.Rows = 2
        Else
            flxSPec.RemoveItem ActRow
        End If
        'lstChange.AddItem "SP Disclusion"
    End If
End Sub

Private Sub mnuWhlAdd_Click()
    If flxWDisc.TextMatrix(flxWDisc.Rows - 1, 1) = "" Or flxWDisc.TextMatrix(flxWDisc.Rows - 1, 2) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    End If
    flxWDisc.AddItem ""
    flxWDisc.SetFocus
    flxWDisc.Row = flxWDisc.Rows - 1
    flxWDisc.Col = 1
End Sub

Private Sub mnuWhlDel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxWDisc.Rows = 2 Then
            flxWDisc.Rows = 1
            flxWDisc.Rows = 2
        Else
            flxWDisc.RemoveItem ActRow
        End If
        'lstChange.AddItem "Wholesale Discount"
    End If
End Sub


Private Sub RXResizer1_ResizeComplete()
 Dim rw As Long
    With flxParts
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.26
        .ColWidth(2) = (.Width - 300) * 0.38
        .ColWidth(3) = (.Width - 300) * 0.11
        .ColWidth(4) = (.Width - 300) * 0.11
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = (.Width - 300) * 0.1
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
    End With
    With flxOld
        .ColWidth(0) = (.Width / 4) - 20
        .ColWidth(1) = (.Width / 4) - 20
        .ColWidth(2) = (.Width / 4) - 20
        .ColWidth(3) = (.Width / 4) - 20
    End With
    With flxWDisc
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.65
        .ColWidth(2) = (.Width - 300) * 0.3
    End With
    With flxSPec
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.45
        .ColWidth(2) = (.Width - 300) * 0.25
        .ColWidth(3) = (.Width - 300) * 0.25
        .ColWidth(4) = 0
    End With
    lblHeading(2).Left = tabSP.Left + tabSP.Width - lblHeading(2).Width
    lblHeading(3).Left = lblHeading(2).Left + 30
  
    
End Sub

Function ValidateFactor() As String
    Dim eMsg As String, Qry As String
    Dim fMup As Single, fVat As Single, fMxD As Single
    
    eMsg = ""
    If txtFDesc.Text = "" Then
        eMsg = "Incomplete entry on Factor Description"
    Else
        If NewFactor Then
            'Qry = "SELECT SPFACTOR.SPFDESC From SPFACTOR WHERE (((SPFACTOR.SPFDESC)='" & txtFDesc.Text & "'));"
            'mySQL
            Qry = "SELECT SPFACTOR.SPFDESC From SPFACTOR WHERE SPFACTOR.SPFDESC='" & txtFDesc.Text & "';"
            If CtrlNumExist(Qry) Then
                eMsg = eMsg & vbCrLf & "Factor Description already exist"
            End If
        Else
            'Qry = "SELECT SPFACTOR.SPFDESC From SPFACTOR WHERE (((SPFACTOR.SPFDESC)='" & txtFDesc.Text & "') AND ((SPFACTOR.SPFID)<>" & txtSPFid.Text & "));"
            'mySQL
            Qry = "SELECT SPFACTOR.SPFDESC From SPFACTOR WHERE SPFACTOR.SPFDESC='" & txtFDesc.Text & "' AND SPFACTOR.SPFID<>" & txtSPFid.Text & ";"
            
            If CtrlNumExist(Qry) Then
                eMsg = eMsg & vbCrLf & "Factor Description already exist"
            End If
        End If
    End If
    If Not IsDate(txtFDate.Text) Then eMsg = eMsg & vbCrLf & "Invalid entry on Date Created"
    If txtFMarkUp.Text = "" Or Val(txtFMarkUp.Text) = 0 Then eMsg = eMsg & vbCrLf & "Incomplete entry on Field Mark Up"
    If txtFVat.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on Field VAT"
    If txtFMaxDisc.Text <> "" Then
        'eMsg = eMsg & vbCrLf & "Incomplete entry on Field Maximum Discount"
        If Val(txtFMaxDisc.Text) > 100 Then eMsg = eMsg & vbCrLf & "Maximum Discount must not exceed 100"
    End If
    If txtFMarkUp.Text <> "" Then
        fMup = IIf(Right$(txtFMarkUp.Text, 1) = "%", Val(Left$(txtFMarkUp.Text, Len(txtFMarkUp.Text) - 1)), Val(txtFMarkUp.Text))
    End If
    If txtFVat.Text <> "" Then
        fVat = IIf(Right$(txtFVat.Text, 1) = "%", Val(Left$(txtFVat.Text, Len(txtFVat.Text) - 1)), Val(txtFVat.Text))
    End If
    If txtFMaxDisc.Text <> "" Then
        fMxD = IIf(Right$(txtFMaxDisc.Text, 1) = "%", Val(Left$(txtFMaxDisc.Text, Len(txtFMaxDisc.Text) - 1)), Val(txtFMaxDisc.Text))
    End If
    If NewFactor Then
        'Validate Existing Description
        'Qry = "SELECT SPFACTOR.MARKUP, SPFACTOR.VAT, SPFACTOR.MAXDISCOUNT From SPFACTOR " & _
              "WHERE (((SPFACTOR.MARKUP)=" & fMup & ") AND  ((SPFACTOR.VAT)=" & fVat & ") AND ((SPFACTOR.MAXDISCOUNT)=" & fMxD & "));"
        'mySQL
        Qry = "SELECT SPFACTOR.MARKUP, SPFACTOR.VAT, SPFACTOR.MAXDISCOUNT From SPFACTOR " & _
              "WHERE SPFACTOR.MARKUP=" & fMup & " AND  SPFACTOR.VAT=" & fVat & " AND SPFACTOR.MAXDISCOUNT=" & fMxD & ";"
        
        If CtrlNumExist(Qry) Then
            eMsg = eMsg & vbCrLf & "Factor with the same Mark-up, VAT and Max. Discount already exist"
        End If
    Else
        'Qry = "SELECT SPFACTOR.MARKUP, SPFACTOR.VAT, SPFACTOR.MAXDISCOUNT From SPFACTOR " & _
              "WHERE (((SPFACTOR.MARKUP)=" & fMup & ") AND ((SPFACTOR.SPFID)<>" & txtSPFid.Text & ") AND ((SPFACTOR.VAT)=" & fVat & ") AND ((SPFACTOR.MAXDISCOUNT)=" & fMxD & "));"
        'mySQL
        Qry = "SELECT SPFACTOR.MARKUP, SPFACTOR.VAT, SPFACTOR.MAXDISCOUNT From SPFACTOR " & _
              "WHERE SPFACTOR.MARKUP=" & fMup & " AND SPFACTOR.SPFID<>" & txtSPFid.Text & " AND SPFACTOR.VAT=" & fVat & " AND SPFACTOR.MAXDISCOUNT=" & fMxD & ";"
        
        If CtrlNumExist(Qry) Then
            eMsg = eMsg & vbCrLf & "Factor with the same Mark-up, VAT and Max. Discount already exist"
        End If
    End If

    ValidateFactor = eMsg
End Function

Private Sub Timer1_Timer()
    Dim rsGrpFact As New ADODB.Recordset
    Dim rsCType As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer, j As Integer, CostVat As Currency
    
    Qry = "SELECT SPFACTOR.* From SPFACTOR ORDER BY SPFACTOR.SPFDATE;"
    With rsFactor
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxFactors.Rows = 1
            Do While Not .EOF
                flxFactors.AddItem vbTab & !SPFDate & vbTab & !SPFDESC & vbTab & !Markup & "%" & vbTab & !Vat & "%" & vbTab & !MaxDiscount & vbTab & !SPFID
                cboSPFactor.AddItem !SPFDESC
                cboGFactor.AddItem !SPFDESC
                cboDPfactor.AddItem !SPFDESC
                .MoveNext
            Loop
        End If
        .Close
    End With
    flxFactors.ColWidth(6) = 0
    fraFactor.Enabled = False
    
    'Display Group
    Qry = "SELECT SPGFACTOR.* From SPGFACTOR ORDER BY SPGFACTOR.GDESC;"
    With rsGrpFact
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxGroup.Rows = 1
            Do While Not .EOF
                flxGroup.AddItem vbTab & !GDate & vbTab & !GDesc & vbTab & vbTab & !GCriteria & vbTab & !SPFID & vbTab & !SQLst
                .MoveNext
            Loop
        End If
        .Close
    End With
    'Display Name of Factor in FlxGroup
    For i = 1 To flxGroup.Rows - 1
        For j = 1 To flxFactors.Rows - 1
            If flxFactors.TextMatrix(j, 6) = flxGroup.TextMatrix(i, 5) Then
                flxGroup.TextMatrix(i, 3) = flxFactors.TextMatrix(j, 2)
                Exit For
            End If
        Next j
    Next i
    
    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE " & _
          "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
    With rsSupplier
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupplier.Clear
            Do While Not .EOF
                cboSupplier.AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
    End With
    With rsCType
        'Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='Wholesale Discount')) ORDER BY SETTINGS.VALUE;"
        'mySQL
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION='Wholesale Discount' ORDER BY SETTINGS.VALUE;"
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxWDisc.Rows = 1
            Do While Not .EOF
                flxWDisc.AddItem vbTab & !Remarks & vbTab & !Value
                .MoveNext
            Loop
        End If
        .Close
    End With
    With rsCType
        'Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='SP Disclusion')) ORDER BY SETTINGS.VALUE;"
        'mySQL
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION='SP Disclusion' ORDER BY SETTINGS.VALUE;"
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSPec.Rows = 1
            Do While Not .EOF
                flxSPec.AddItem vbTab & vbTab & vbTab & !Value & vbTab & !Remarks
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    With rsCType
        Qry = "SELECT DISTINCT SUPPLIER.SPTYPE From SUPPLIER ORDER BY SUPPLIER.SPTYPE;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboSPType.AddItem !SPtype
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.COST From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsPartsD.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxPartsD.DataSource = rsPartsD
    flxPartsD.Refresh
    With flxPartsD
        .Width = flxSPec.ColWidth(1)
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = flxSPec.ColWidth(1)
        .ColWidth(3) = 0
    End With
    'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='VAT'));"
    'mySQL
    Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='VAT';"
    
    txtVat.Text = GetFieldValue(Qry)
    
    With flxSPec
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 4)) > 0 Then
                rsPartsD.MoveFirst
                rsPartsD.Find "[SPAREPARTID] = " & Val(.TextMatrix(i, 4))
                If Not rsPartsD.EOF Then
                    .TextMatrix(i, 1) = rsPartsD!PartNumber
                    CostVat = rsPartsD!Cost
                    CostVat = CostVat + (CostVat * Val(txtVat.Text) / 100)
                    .TextMatrix(i, 2) = Format(CostVat, "#,##0.00")
                End If
            End If
        Next i
    End With
    
    
    Timer1.Enabled = False
End Sub

Private Sub ReloadWholeSaleDiscount()
    Dim rsCType As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer, j As Integer, CostVat As Currency
    
    With rsCType
        'Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='Wholesale Discount')) ORDER BY SETTINGS.VALUE;"
        'mySQL
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION='Wholesale Discount' ORDER BY SETTINGS.VALUE;"
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxWDisc.Rows = 1
            Do While Not .EOF
                flxWDisc.AddItem vbTab & !Remarks & vbTab & !Value
                .MoveNext
            Loop
        End If
        .Close
    End With
    With rsCType
        'Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='SP Disclusion')) ORDER BY SETTINGS.VALUE;"
        'mySQL
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION='SP Disclusion' ORDER BY SETTINGS.VALUE;"
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSPec.Rows = 1
            Do While Not .EOF
                flxSPec.AddItem vbTab & vbTab & vbTab & !Value & vbTab & !Remarks
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    With rsCType
        Qry = "SELECT DISTINCT SUPPLIER.SPTYPE From SUPPLIER ORDER BY SUPPLIER.SPTYPE;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboSPType.AddItem !SPtype
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, format(SPAREPARTS.COST,2) AS `UNIT COST` From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    If rsPartsD.State = adStateOpen Then rsPartsD.Close
    rsPartsD.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxPartsD.DataSource = rsPartsD
    flxPartsD.Refresh
    'format cur
    Call FormatCur
    With flxPartsD
        .Width = flxSPec.ColWidth(1)
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = flxSPec.ColWidth(1)
        .ColWidth(3) = 0
    End With
    'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='VAT'));"
    'mySQL
    Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='VAT';"
    
    txtVat.Text = GetFieldValue(Qry)
    
    With flxSPec
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 4)) > 0 Then
                rsPartsD.MoveFirst
                rsPartsD.Find "[SPAREPARTID] = " & Val(.TextMatrix(i, 4))
                If Not rsPartsD.EOF Then
                    .TextMatrix(i, 1) = rsPartsD!PartNumber
                    CostVat = rsPartsD!Cost
                    CostVat = CostVat + (CostVat * Val(txtVat.Text) / 100)
                    .TextMatrix(i, 2) = Format(CostVat, "#,##0.00")
                End If
            End If
        Next i
    End With

End Sub


Private Sub txtCost_Change()
    Dim CostVat As Currency, NewPrice As Currency, MyCost As Currency, DlrPrice As Currency
    
    If cboSPFactor.Text <> "" And txtMarkUp.Text <> "" And txtVat.Text <> "" Then
        MyCost = Val(Replace(txtCost.Text, ",", ""))
        CostVat = MyCost + (MyCost * (Val(txtVat.Text) / 100))
        NewPrice = CostVat + (CostVat * (Val(txtMarkUp.Text) / 100))
        txtPrice.Text = Format$(Int(NewPrice + 0.6) + 1, "####0.00")
    End If
    If cboDPfactor.Text <> "" And txtDPMarkUp.Text <> "" And txtVat.Text <> "" Then
        MyCost = Val(Replace(txtCost.Text, ",", ""))
        CostVat = MyCost + (MyCost * (Val(txtVat.Text) / 100))
        DlrPrice = CostVat + (CostVat * (Val(txtDPMarkUp.Text) / 100))
        txtDP.Text = Format$(Int(DlrPrice + 0.6) + 1, "####0.00")
    End If
    
End Sub

Private Sub txtCost_GotFocus()
    If EditMode Then
        txtCost.SelStart = 0
        txtCost.SelLength = Len(txtCost.Text)
    End If
End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtRefNo.SetFocus
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call SearchPartsDesc(0)
End Sub

Private Sub txtDFctr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtDFctr.Text = Empty
        flxWDisc.TextMatrix(ActRow, 2) = txtDFctr.Tag
        txtDFctr.Tag = "skiplostfocus"
        flxWDisc.SetFocus
        flxWDisc.Col = 1
        txtDFctr.Visible = False
    End If
End Sub

Private Sub txtDFctr_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxWDisc.SetFocus
        If flxWDisc.TextMatrix(ActRow, 1) = "" Then
            flxWDisc.Col = 1
            flxWDisc.Row = ActRow
        ElseIf txtDFctr.Text = "" Then
            flxWDisc.Col = 2
            flxWDisc.Row = ActRow
        ElseIf flxWDisc.TextMatrix(ActRow, 1) <> "" And txtDFctr.Text <> "" And ActRow = flxWDisc.Rows - 1 Then
            flxWDisc.AddItem ""
            flxWDisc.Row = flxWDisc.Rows - 1
            flxWDisc.Col = 1
        Else
            flxWDisc.Row = ActRow + 1
            flxWDisc.Col = 1
        End If
    End If
End Sub

Private Sub txtDFctr_LostFocus()
    
    If txtDFctr.Tag = "skiplostfocus" Or txtDFctr.Text = "" Then
        txtDFctr.Tag = ""
        txtDFctr.Visible = False
        Exit Sub
    End If
    flxWDisc.TextMatrix(ActRow, 2) = txtDFctr.Text
    txtDFctr.Visible = False
'    lstChange.AddItem "Wholesale Discount"
End Sub

Private Sub txtDP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cboDPfactor.Text = ""
    End If
End Sub

Private Sub txtDP_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        cboDPfactor.Text = ""
    End If
    If KeyAscii = 13 Then Call cmdSave_Click    'tbCosting_ButtonClick(tbCosting.Buttons(2))

End Sub

Private Sub txtDPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtDPrice.Text = Empty
        flxSPec.TextMatrix(ActRow, 3) = txtDPrice.Tag
        txtDPrice.Tag = "skiplostfocus"
        flxSPec.SetFocus
        flxSPec.Col = 3
        txtDPrice.Visible = False
    End If
End Sub

Private Sub txtDPrice_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxSPec.SetFocus
        If flxSPec.TextMatrix(ActRow, 1) = "" Then
            flxSPec.Col = 1
            flxSPec.Row = ActRow
        ElseIf txtDPrice.Text = "" Then
            flxSPec.Col = 3
            flxSPec.Row = ActRow
        ElseIf txtDPrice.Text <> "" And flxSPec.TextMatrix(ActRow, 1) <> "" And ActRow = flxSPec.Rows - 1 Then
            flxSPec.AddItem ""
            flxSPec.Row = flxSPec.Rows - 1
            flxSPec.Col = 1
        Else
            flxSPec.Row = ActRow + 1
            flxSPec.Col = 1
        End If
    End If
End Sub

Private Sub txtDPrice_LostFocus()
    If txtDPrice.Tag = "skiplostfocus" Or txtDPrice.Text = "" Then
        txtDPrice.Tag = ""
        txtDPrice.Visible = False
        Exit Sub
    End If
    flxSPec.TextMatrix(ActRow, 3) = Format(txtDPrice.Text, "###0.00")
    txtDPrice.Visible = False
    'lstChange.AddItem "SP Disclusion"
End Sub

Private Sub txtFDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789/-"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtFMarkUp.SetFocus
    
End Sub

Private Sub txtFDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtFDate.SetFocus
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim Qry As String
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtFind.Text <> "" Then
            Call SearchPartsDesc(1)
        Else
            Call cmdRefresh_Click
        End If
    End If
End Sub

Private Sub SearchPartsDesc(SType As Integer)
    Dim MySP() As String
    Dim myW As String, MySel As String
    
    If SType = 1 Then
        MySP = Split(txtFind.Text, " ")
    Else
        MySP = Split(txtDesc.Text, " ")
    End If
    myW = ""
    For i = LBound(MySP) To UBound(MySP)
        If myW = "" Then
            myW = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
        Else
            myW = myW & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
        End If
    Next i
    If SType = 1 Then
    '    old Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST],'Standard') AS [UNIT COST], Format([STOREPRICE],'Standard') AS [SRP], SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format([DLRPRICE],'Standard') AS [DP] " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"

        'MySel = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST],'Standard') AS [UNIT COST], Format([STOREPRICE],'Standard') AS SRP, SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format([DLRPRICE],'Standard') AS [DP] " & _
                "From SPAREPARTS WHERE ((" & myW & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
        'mySQL
        MySel = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST,2) AS `UNIT COST`, Format(SPAREPARTS.STOREPRICE,2) AS SRP, SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format(SPAREPARTS.DLRPRICE,2) AS DP " & _
                "From SPAREPARTS WHERE ((" & myW & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
    
    Else
        If txtDesc.Text = "" And txtSuppCode.Text <> "" Then
            myW = "((SPAREPARTS.SUPPLIERCODE) = '" & txtSuppCode.Text & "')"
        Else
            If txtSuppCode.Text <> "" Then
                myW = "(" & myW & ") AND ((SPAREPARTS.SUPPLIERCODE) = '" & txtSuppCode.Text & "')"
            End If
        End If
        'MySel = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST],'Standard') AS [UNIT COST], Format([STOREPRICE],'Standard') AS SRP, SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format([DLRPRICE],'Standard') AS [DP] " & _
                "From SPAREPARTS WHERE (" & myW & ") ORDER BY SPAREPARTS.PARTNUMBER;"
        'mySQL
        MySel = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST,2) AS `UNIT COST`, Format(SPAREPARTS.STOREPRICE,2) AS SRP, SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format(SPAREPARTS.DLRPRICE,2) AS DP " & _
                "From SPAREPARTS WHERE (" & myW & ") ORDER BY SPAREPARTS.PARTNUMBER;"
        
    End If
    If rsParts.State = adStateOpen Then rsParts.Close
    rsParts.Open MySel, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxParts.Clear
        flxParts.Rows = 2
        flxParts.FormatString = "    |<PART NUMBER  |<DESCRIPTION  |>UNIT COST |>SRP"
        Call RXResizer1_ResizeComplete
    Else
        Set flxParts.DataSource = rsParts
        flxParts.Refresh
        flxParts.ColAlignment(3) = flexAlignRightCenter
        flxParts.ColAlignment(4) = flexAlignRightCenter
    End If
    'format cur
    Call FormatCur
End Sub

Private Sub txtFMarkUp_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtFVat.SetFocus

End Sub


Private Sub txtFMaxDisc_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then cmdFSave.SetFocus
End Sub

Private Sub txtFVat_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtFMaxDisc.SetFocus
End Sub


Private Sub txtPParts_Change()
    Dim rsDP As New ADODB.Recordset
    Dim FctrID As Integer, DP As Currency, Qry As String, SPid As Long
    
    If PrtSrch Then
        PrtSrch = False
        Exit Sub
    End If
    'flxParts.Visible = True
    If txtPParts.Text = "" Then Exit Sub
    rsParts.Find "[PARTNUMBER] LIKE '" + txtPParts.Text + "%'", , adSearchForward, 1
    If Not rsParts.EOF Then
        If Not bNone Then
            txtTag.Text = Len(txtPParts.Text)
            bNone = True
        End If
        flxParts.TopRow = rsParts.AbsolutePosition 'IIf(rsParts.AbsolutePosition - 1 = 0, 1, rsParts.AbsolutePosition - 1)
        flxParts.Row = rsParts.AbsolutePosition
        flxParts.Col = 0: flxParts.ColSel = 5
        'flxParts.ColSel = 0
        txtOldCost.Text = rsParts.Fields(2)
        txtCost.Text = rsParts.Fields(2)
        txtOldPrice.Text = rsParts!SRP
        txtPrice.Text = rsParts!SRP
        txtRefNo.Text = IIf(IsNull(rsParts!PriceRef), "", rsParts!PriceRef)
        lblDescription.Caption = rsParts!Description
        FctrID = IIf(IsNull(rsParts!SPFID), 0, rsParts!SPFID)
        If IsNull(rsParts!DP) Then
            DP = 0
        Else
            DP = rsParts!DP
        End If
        'DP = IIf(IsNull(rsParts!DP), "", Va(rsParts!DP))
        SPid = rsParts!SparepartID
        If FctrID > 0 Then
            For i = 1 To flxFactors.Rows - 1
                If FctrID = flxFactors.TextMatrix(i, 6) Then
                    cboSPFactor.Text = flxFactors.TextMatrix(i, 2)
                    txtMarkUp.Text = flxFactors.TextMatrix(i, 3)
                    txtVat.Text = flxFactors.TextMatrix(i, 4)
                    txtMaxDisc.Text = flxFactors.TextMatrix(i, 5)
                    Exit For
                End If
            Next i
        Else
            cboSPFactor.Text = ""
            txtMarkUp.Text = ""
            txtVat.Text = ""
            txtMaxDisc.Text = ""
        End If
        If DP > 0 Then
            txtDP.Text = Format(rsParts!DP, "#,##0.00")
            Qry = "SELECT SPFACTOR.SPFDESC FROM DPFACTOR INNER JOIN SPFACTOR ON DPFACTOR.SPFID = SPFACTOR.SPFID Where (((DPFACTOR.SPAREPARTID) = " & SPid & ")) ORDER BY DPFACTOR.SEQNO DESC;"
            cboDPfactor.Text = GetFieldValue(Qry)
        Else
            txtDP.Text = ""
            txtDPMarkUp.Text = ""
            txtDPvat.Text = ""
            txtDPfactorID.Text = ""
        End If
            
        txtHist.Text = IIf(IsNull(rsParts!PriceHist), "", rsParts!PriceHist)
        Call DisplayCostHistory(txtHist.Text)
        txtPParts.Text = rsParts!PartNumber
        txtPParts.SelStart = Val(txtTag.Text)
        txtPParts.SelLength = Len(txtPParts.Text)
        'Hflexg.Tag = Text1.Text
        bNone = False
     Else
        flxParts.TopRow = 1
        flxParts.Col = 1: flxParts.Row = 1
        flxParts.ColSel = 1
        rsParts.Requery
        txtOldCost.Text = ""
        txtCost.Text = ""
        txtOldPrice.Text = ""
        txtPrice.Text = ""
        txtHist.Text = ""
        lblDescription.Caption = ""
         '.Visible = False
     End If
End Sub

Private Sub txtPParts_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            txtPParts.Text = Empty
            txtPParts.SetFocus
        Case vbKeyDelete
            PrtSrch = True
        Case vbKeyBack
            PrtSrch = True
    End Select
End Sub

Private Sub txtPParts_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboSPFactor.SetFocus
End Sub


Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cboSPFactor.Text = ""
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        cboSPFactor.Text = ""
    End If
    If KeyAscii = 13 Then cboDPfactor.SetFocus  'Call cmdSave_Click    'tbCosting_ButtonClick(tbCosting.Buttons(2))
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboSPFactor.SetFocus
End Sub

Private Sub txtPartNo_Change()
    Dim Cost As Currency, CostVat As Currency
    
    'On Error GoTo LocalErr
    With flxPartsD
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPartNo.Height + txtPartNo.Top + Frame4.Top + tabSP.Top
        .Left = txtPartNo.Left + Frame4.Left + tabSP.Left
        .Width = txtPartNo.Width
        .Visible = True
        .ZOrder
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsPartsD.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsPartsD.EOF Then
            If Not PrtBln Then
                flxPartsD.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            .TopRow = rsPartsD.AbsolutePosition - 1
            .Row = rsPartsD.AbsolutePosition - 1
            .Col = 0: .ColSel = 3
            txtPartNo.Text = rsPartsD!PartNumber
            'Insert Additional retrieval here
            Cost = flxPartsD.TextMatrix(flxPartsD.Row, 3)
            CostVat = Cost + (Cost * (Val(txtVat.Text) / 100))
            flxSPec.TextMatrix(ActRow, 2) = Format(CostVat, "#,##0.00") 'Cost
            flxSPec.TextMatrix(ActRow, 4) = flxPartsD.TextMatrix(flxPartsD.Row, 1) 'SPID
            txtPartNo.SelStart = Val(flxPartsD.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            flxSPec.TextMatrix(ActRow, 2) = ""
            flxSPec.TextMatrix(ActRow, 4) = ""
            .Visible = False
        End If
    End With
End Sub

Private Sub txtPartNo_GotFocus()
    With flxPartsD
        .Top = txtPartNo.Height + txtPartNo.Top + Frame4.Top + tabSP.Top
        .Left = txtPartNo.Left + Frame4.Left + tabSP.Left
        .Width = txtPartNo.Width
    End With
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
        
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyDown
            txtPartNo.Tag = "skiplostfocus"
            flxPartsD.Visible = True
            flxPartsD.SetFocus
        Case vbKeyEscape
            txtPartNo.Text = Empty
            If flxPartsD.Tag <> "" Then
                Tmp = Split(flxPartsD.Tag, "|")
                flxSPec.TextMatrix(ActRow, 1) = Tmp(0)
                flxSPec.TextMatrix(ActRow, 2) = Tmp(1)
                flxSPec.TextMatrix(ActRow, 4) = Tmp(2)
            End If
            txtPartNo.Tag = "skiplostfocus"
            flxSPec.SetFocus
            flxSPec.Col = 1
            txtPartNo.Visible = False
            flxPartsD.Visible = False
        Case vbKeyDelete
            PrtSrch = True
            If rsPartsD.State = adStateOpen Then rsPartsD.MoveFirst
        Case vbKeyBack
            PrtSrch = True
            If rsPartsD.State = adStateOpen Then rsPartsD.MoveFirst
    End Select
ErrHand:
    If Err.Number = 381 Then
        MsgBox ("Invalid row selection"), vbCritical
        Exit Sub
    End If
    If Err.Number > 0 Then
        MsgBox (Err.Number & " " & Err.Description), vbCritical
        Exit Sub
    End If
    
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxSPec.SetFocus
        If flxSPec.TextMatrix(ActRow, 4) <> "" And flxSPec.TextMatrix(ActRow, 3) <> "" Then
            If ActRow = flxSPec.Rows - 1 Then flxSPec.AddItem ""
            flxSPec.SetFocus
            flxSPec.Row = flxSPec.Rows - 1
            flxSPec.Col = 1
        ElseIf flxSPec.TextMatrix(ActRow, 3) = "" Then   'No Discount Price
            flxSPec.SetFocus
            flxSPec.Col = 3
        End If
        flxPartsD.Visible = False
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    Dim i As Integer
    
    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    If flxSPec.TextMatrix(ActRow, 4) = "" And txtPartNo.Text <> "" Then
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
        txtPartNo.Visible = False
        flxPartsD.Visible = False
        flxSPec.TextMatrix(ActRow, 2) = ""
        Exit Sub
    End If
    flxSPec.TextMatrix(ActRow, 1) = txtPartNo.Text
    txtPartNo.Visible = False
    flxPartsD.Visible = False
    'lstChange.AddItem "SP Disclusion"
End Sub
Sub FormatCur()
    For rw = 1 To flxParts.Rows - 1
        flxParts.TextMatrix(rw, 3) = Format(flxParts.TextMatrix(rw, 3), "##,###,###.#0")
        flxParts.TextMatrix(rw, 4) = Format(flxParts.TextMatrix(rw, 4), "##,###,###.#0")
        flxParts.TextMatrix(rw, 9) = Format(flxParts.TextMatrix(rw, 9), "##,###,###.#0")
    Next rw
End Sub


