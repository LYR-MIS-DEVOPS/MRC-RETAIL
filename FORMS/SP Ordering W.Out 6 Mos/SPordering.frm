VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form SPordering 
   Caption         =   "Spare Parts Order Form"
   ClientHeight    =   8685
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   11595
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   4200
   End
   Begin lvButton.lvButtons_H cmdInsX 
      Height          =   300
      Left            =   9720
      TabIndex        =   71
      Tag             =   "XYF"
      ToolTipText     =   "Insert Existing Parts [F12]"
      Top             =   1990
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
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
      Image           =   "SPordering.frx":0000
      cBack           =   -2147483633
   End
   Begin MSComCtl2.DTPicker SalesDateTo 
      Height          =   255
      Left            =   3000
      TabIndex        =   70
      Top             =   10920
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Format          =   54329345
      CurrentDate     =   39498
   End
   Begin MSComCtl2.DTPicker SalesDateFrom 
      Height          =   255
      Left            =   960
      TabIndex        =   69
      Top             =   10920
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Format          =   54329345
      CurrentDate     =   39498
   End
   Begin MSFlexGridLib.MSFlexGrid flxNewSP 
      Height          =   1215
      Left            =   480
      TabIndex        =   68
      Tag             =   "xyf"
      Top             =   9720
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   0
      Cols            =   11
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxEdit 
      Height          =   495
      Left            =   2280
      TabIndex        =   67
      Top             =   9600
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox txtSuppID 
      Height          =   285
      Left            =   5520
      TabIndex        =   66
      Top             =   9120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAddress 
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   58
      Tag             =   "XYF"
      Top             =   8700
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtDealerName 
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   57
      Tag             =   "XYF"
      Top             =   8385
      Width           =   3255
   End
   Begin VB.TextBox txtDealerCode 
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
      Left            =   7920
      MaxLength       =   15
      TabIndex        =   56
      Tag             =   "XYF"
      Top             =   8025
      Width           =   3255
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   555
      TabIndex        =   55
      Tag             =   "XY"
      Top             =   7275
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtOrderID 
      Height          =   285
      Left            =   10080
      TabIndex        =   53
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   4440
      TabIndex        =   52
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtFind 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "Part Number Exist @ row 5"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtConfirmNo 
      Height          =   285
      Left            =   5400
      TabIndex        =   50
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtNM 
      Height          =   285
      Left            =   6120
      MaxLength       =   1
      TabIndex        =   49
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin TabDlg.SSTab tabFind 
      Height          =   1095
      Left            =   2520
      TabIndex        =   42
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPordering.frx":059A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtSearch"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboMatch"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdNext"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Find Next"
         Height          =   375
         Left            =   3480
         TabIndex        =   46
         Top             =   195
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3480
         TabIndex        =   45
         Top             =   645
         Width           =   975
      End
      Begin VB.ComboBox cboMatch 
         Height          =   315
         ItemData        =   "SPordering.frx":05B6
         Left            =   1080
         List            =   "SPordering.frx":05B8
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   675
         Width           =   2295
      End
      Begin VB.ComboBox txtSearch 
         Height          =   315
         Left            =   1080
         TabIndex        =   43
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Find What:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Match:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   705
         Width           =   855
      End
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   300
      Left            =   10200
      TabIndex        =   41
      Tag             =   "XYF"
      ToolTipText     =   "Find Part Number"
      Top             =   1990
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
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
      Image           =   "SPordering.frx":05BA
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   7080
      TabIndex        =   40
      Top             =   9960
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid flxDel 
      Height          =   615
      Left            =   3240
      TabIndex        =   39
      Top             =   9000
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   5280
      MaxLength       =   249
      TabIndex        =   37
      Top             =   9840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      TabIndex        =   35
      Tag             =   "XYF"
      Top             =   7600
      Width           =   2055
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   8760
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   9240
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   8760
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   9240
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   8400
      TabIndex        =   23
      Top             =   9480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   5280
      TabIndex        =   22
      Top             =   9480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSrch 
      Height          =   1695
      Left            =   6960
      TabIndex        =   21
      Top             =   9480
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   0
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   10920
      Width           =   135
   End
   Begin VB.TextBox txtFlx 
      Height          =   285
      Left            =   7080
      TabIndex        =   19
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Height          =   285
      Left            =   2040
      MaxLength       =   150
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   7600
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid flxParts 
      Height          =   5295
      Left            =   480
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   2280
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "^    |<PART NUMBER   |<DESCRIPTION  |^QTY.   |>DEALER'S NET|>TOTAL PRICE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10920
      Top             =   9720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":0B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":1A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":2908
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":37E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":46BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":5596
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   480
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   700
      Width           =   10695
      Begin VB.TextBox txtTerms 
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
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   876
         Width           =   3735
      End
      Begin VB.ComboBox cboSupplier 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   192
         Width           =   3735
      End
      Begin VB.TextBox txtSaddress 
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
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   538
         Width           =   3735
      End
      Begin VB.TextBox txtScode 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   65
         Tag             =   "XYF"
         Top             =   200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   7800
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   876
         Width           =   2535
      End
      Begin VB.TextBox txtPONumber 
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
         Left            =   7800
         TabIndex        =   4
         Tag             =   "XYF"
         Text            =   "<Auto>"
         Top             =   538
         Width           =   2535
      End
      Begin VB.TextBox txtPODate 
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
         Left            =   7800
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   200
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "SUPPLIERADDRESS:"
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
         TabIndex        =   64
         Tag             =   "XYF"
         Top             =   568
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "TERMS:"
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
         TabIndex        =   63
         Tag             =   "XYF"
         Top             =   906
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "SUPPLIER NAME:"
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
         TabIndex        =   62
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER TYPE"
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
         Index           =   5
         Left            =   6600
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   906
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O. NUMBER"
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
         Left            =   6240
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   568
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O. DATE"
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
         Left            =   6600
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   230
         Width           =   1095
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   24
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
      Image           =   "SPordering.frx":6470
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   25
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&View Summary"
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
      Image           =   "SPordering.frx":6A0A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   26
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Order"
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
      Image           =   "SPordering.frx":6FA4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   27
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
      Image           =   "SPordering.frx":753E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10080
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Save"
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
      Image           =   "SPordering.frx":7AD8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   11760
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SPordering.frx":8072
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1678
      TabIndex        =   30
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
      Image           =   "SPordering.frx":860C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   38
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Update Order"
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
      Image           =   "SPordering.frx":8BA6
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtSubmit 
      Height          =   285
      Left            =   2040
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   8040
      Width           =   2775
   End
   Begin VB.TextBox txtAprv 
      Height          =   285
      Left            =   2040
      TabIndex        =   34
      Tag             =   "XYF"
      Top             =   8400
      Width           =   2775
   End
   Begin lvButton.lvButtons_H cmdXLS 
      Height          =   300
      Left            =   10680
      TabIndex        =   54
      Tag             =   "XYF"
      ToolTipText     =   "Import P.O. from Excel"
      Top             =   1990
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
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
      Image           =   "SPordering.frx":9140
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   1095
      Left            =   120
      TabIndex        =   72
      Top             =   9240
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
   Begin lvButton.lvButtons_H cmdTopSP 
      Height          =   300
      Left            =   9240
      TabIndex        =   73
      Tag             =   "XYF"
      ToolTipText     =   "Insert Top Selling Parts [F11]"
      Top             =   1990
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
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
      Image           =   "SPordering.frx":96DA
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      Caption         =   "DEALER CODE"
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
      Left            =   6360
      TabIndex        =   61
      Tag             =   "XYF"
      Top             =   8055
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "DEALER NAME"
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
      Left            =   6360
      TabIndex        =   60
      Tag             =   "XYF"
      Top             =   8415
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "DEALER ADDRESS"
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
      Left            =   6360
      TabIndex        =   59
      Tag             =   "XYF"
      Top             =   8730
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL:"
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
      Left            =   8040
      TabIndex        =   36
      Tag             =   "XYF"
      Top             =   7615
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "APPROVED BY:"
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
      Left            =   480
      TabIndex        =   32
      Tag             =   "XYF"
      Top             =   8415
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "SUBMITTED BY:"
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
      Left            =   480
      TabIndex        =   31
      Tag             =   "XYF"
      Top             =   8040
      Width           =   1335
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   0
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11715
      m_MDI_Height    =   9195
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
      Left            =   480
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPARE PARTS ORDER FORM"
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
      Height          =   240
      Index           =   2
      Left            =   8190
      TabIndex        =   15
      Tag             =   "TRF"
      Top             =   450
      Width           =   2745
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
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Tag             =   "TLF"
      Top             =   450
      Width           =   2900
   End
   Begin VB.Label lbl2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "REMARKS:"
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
      Left            =   1140
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   7637
      Width           =   795
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
      Height          =   255
      Index           =   1
      Left            =   510
      TabIndex        =   14
      Tag             =   "TLF"
      Top             =   475
      Width           =   2900
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPARE PARTS ORDER FORM"
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
      Height          =   240
      Index           =   3
      Left            =   8220
      TabIndex        =   16
      Tag             =   "TRF"
      Top             =   480
      Width           =   2745
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
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   18
      Tag             =   "XYF"
      Top             =   2070
      Width           =   2415
   End
   Begin VB.Menu mnuRows 
      Caption         =   "SpareParts"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row                Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIns 
         Caption         =   "Insert Existing SP      F12"
      End
   End
   Begin VB.Menu mnuBO 
      Caption         =   "Adjust Back Order"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Item in P.O."
      End
      Begin VB.Menu mnuChange 
         Caption         =   "Change Item Order Qty."
      End
   End
End
Attribute VB_Name = "SPordering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim rsSupplier As New ADODB.Recordset

Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean, ImportPO As Boolean
Dim ActRow As Integer, ActCol As Integer, BtnIndx As Integer
Dim AddNew As Boolean, Confirm As Boolean, SuppSrch As Boolean
Dim Qry As String
Public LockFlx As Boolean

Private Sub ClearEntries()
    cboSupplier.Text = ""
    txtScode.Text = ""
    txtSaddress.Text = ""
    txtSuppID.Text = ""
    txtDealerCode.Text = ""
    txtDealerName.Text = ""
    txtAddress.Text = ""
    txtPODate.Text = ""
    txtPONumber.Text = "<Auto>"
    cboType.Text = ""
    txtRemarks.Text = ""
    txtSubmit.Text = ""
    txtAprv.Text = ""
    txtPrnCopy.Text = ""
    txtOrderID.Text = ""
    txtTerms.Text = ""
    flxParts.Rows = 1
    flxParts.Rows = 2
    flxEdit.Rows = 0
    flxDel.Rows = 0
    flxNewSP.Rows = 0
    txtConfirmNo.Text = ""
End Sub

Private Sub LockData(Bln As Boolean)
    Frame1.Enabled = Not Bln
    txtDealerCode.Locked = Bln
    txtDealerName.Locked = Bln
    txtAddress.Locked = Bln
    txtRemarks.Locked = Bln
    txtSubmit.Locked = Bln
    txtAprv.Locked = Bln
    cmdXLS.Enabled = Not Bln
    LockFlx = Bln
End Sub

Public Sub ComputeTotal()
    Dim Ttl As Single, i As Integer
    Dim TtlCost As Single
    
    Ttl = 0
    With flxParts
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 3)) > 0 And Val(.TextMatrix(i, 4)) > 0 Then       'Val(.TextMatrix(i, 5)) = 0 And
                TtlCost = Val(.TextMatrix(i, 3)) * Val(Replace(.TextMatrix(i, 4), ",", ""))
                .TextMatrix(i, 5) = Format$(TtlCost, "###,##0.00")
            End If
            Ttl = Ttl + Val(Replace(.TextMatrix(i, 5), ",", ""))
        Next i
    End With
    txtTotal.Text = Format$(Ttl, "#,##0.00")
    
End Sub

Private Sub cboFind_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    With cboFind
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            'If ActRow = i + 1 Then Exit Sub
            txtFind.Tag = i + 1 'IIf(LastRowLoaded >= i, i + 1, i)
            txtFind.Top = (flxParts.CellTop + flxParts.Top) - 315
            txtFind.Left = flxParts.CellLeft + flxParts.Left
            txtFind.Text = "Part Number Exist @ row " & Val(txtFind.Tag)
            txtFind.Visible = True
            txtFind.ZOrder
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PrtSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtFind.Tag = ""
            txtFind.Visible = False
            txtFind.Text = ""
        End If
    End With
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
            txtSuppID.Text = rsSupplier!SUPPLIERID
            txtScode.Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            txtSaddress.Text = rsSupplier!Address
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppID.Text = ""
            txtScode.Text = ""
            txtSaddress.Text = ""
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
    If KeyAscii = 13 Then txtSaddress.SetFocus
End Sub

Private Sub cboSupplier_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub
Private Sub cboSupplier_LostFocus()
    Dim rsBranch As New ADODB.Recordset
    
    If InStr(cboSupplier.Text, "HONDA PHIL") > 0 Then
        With rsBranch
            '***Access
            'Qry = "SELECT BRANCH.* From BRANCH WHERE (((BRANCH.BRANCHID)=1));"
            'mySQL
            Qry = "SELECT BRANCH.* From BRANCH WHERE BRANCH.BRANCHID=1;"
            
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                txtDealerCode.Text = IIf(IsNull(!MCDealerCode), "", !MCDealerCode)
                txtDealerName.Text = !BranchName
                txtAddress.Text = !Address
            End If
            .Close
        End With
    Else
        txtDealerCode.Text = ""
        txtDealerName.Text = "Moto-R Republic Corp."
        txtAddress.Text = "Sobrecary St. Tagum City"
    End If
    'Transfer to Save
    'If cboSupplier.Text <> "" Then
    '    txtPONumber.Text = GetOrderNo
    'End If
End Sub
Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Call cboType_LostFocus
        flxParts.SetFocus
        flxParts.Col = 1
        flxParts.Row = flxParts.Rows - 1
    End If
End Sub

Private Sub cboType_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub

Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub

Private Sub cmdCancel_Click()
    tabFind.Visible = False
    cmdNext.Tag = 1
    cmdCancel.Tag = "hide"
    Call AddInCombo
    flxParts.SetFocus
End Sub

Private Sub AddInCombo()
    Dim X As Integer, w As Boolean
    
    If txtSearch.Text = "" Then Exit Sub
    w = False
    For X = 0 To txtSearch.ListCount - 1
        If txtSearch.Text = txtSearch.List(X) Then
            If X = 0 Then
                w = True
                Exit For
            Else
                txtSearch.RemoveItem X
                Exit For
            End If
        End If
    Next X
    If w = False Then txtSearch.AddItem txtSearch.Text, 0
End Sub

Private Sub cmdFind_Click()
    tabFind.Visible = True
    tabFind.Caption = ""
    If cboMatch.ListCount > 0 Then cboMatch.Text = cboMatch.List(1)
    cmdNext.Tag = 1
    cmdCancel.Tag = ""
    If Len(txtSearch.Text) > 0 Then
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.Text)
    End If
    txtSearch.SetFocus
End Sub

Private Sub cmdInsX_Click()
    SPInsXst.MySource = "SPORDERING"
    SPInsXst.Show vbModal
End Sub

Private Sub cmdNext_Click()
    Dim i As Integer, j As Integer
    
    With flxParts
        For i = Val(cmdNext.Tag) To .Rows - 1
            For j = 1 To IIf(Confirm, 8, 5)
                If cboMatch.Text = "Whole Word" Then
                    If UCase$(txtSearch.Text) = UCase$(.TextMatrix(i, j)) Then
                        .SetFocus: .Row = i: .Col = 0: .ColSel = 8
                        .TopRow = i: cmdNext.Tag = i + 1
                        txtSearch.SelStart = 0: txtSearch.SelLength = Len(txtSearch.Text)
                        txtSearch.SetFocus
                        Call AddInCombo
                        Exit Sub
                    End If
                Else
                    If InStr(1, UCase$(.TextMatrix(i, j)), UCase$(txtSearch.Text)) > 0 Then
                        .SetFocus: .Row = i: .Col = 0: .ColSel = 5
                        .TopRow = i: cmdNext.Tag = i + 1
                        txtSearch.SelStart = 0: txtSearch.SelLength = Len(txtSearch.Text)
                        txtSearch.SetFocus
                        Call AddInCombo
                        Exit Sub
                    End If
                End If
            Next j
        Next i
        MsgBox "End of Search!" & vbCrLf & "No more records found!", vbInformation + vbOKOnly, "Find"
        cmdNext.Tag = 1
        Call AddInCombo
        .Col = 0: .Row = .Rows - 1
        .TopRow = .Rows - 1
        txtSearch.SetFocus
    End With
End Sub


Private Sub cmdTopSP_Click()
    SPToolPO.Show vbModal
End Sub

Private Sub cmdTrans_Click(Index As Integer)

    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 21) As String
    Dim MyConf As String, MyAcc As String, Qry As String
    
    Select Case Index
        Case 0  'New
            '*Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=18));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=18;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            BtnIndx = 5
            tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            Call flxResizeNew
            txtPODate.Text = Date
            txtSubmit.Text = GetUserName
            txtAprv.Text = GetSettings("MANAGER")
            AddNew = True
            flxParts.TextMatrix(1, 0) = 1
            Confirm = False
            cboType.Locked = False
            cboSupplier.SetFocus
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=18));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=18;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            
            If Val(txtPrnCopy.Text) > 0 Then
                MsgBox "Editing is not allowed because P.O. was already printed.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            ElseIf txtPONumber.Text = "" Then
               MsgBox "No available P.O. to Edit", vbInformation + vbOKOnly, "Confirm P.O."
                Exit Sub
            End If
            
            AddNew = False
            BtnIndx = 5
            tmrFloat.Enabled = True
            AddNew = False
            Call LockData(False)
            txtDealerCode.Locked = True
            txtDealerName.Locked = True
            txtAddress.Locked = True
            txtPODate.Locked = True
            txtPONumber.Locked = True
            cboSupplier.Locked = True
            cboType.Locked = False
            cmdXLS.Enabled = False
            'txtConfirmNo.Locked = True
            Confirm = False
        Case 2  'Confirm Order
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=18));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=18;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            MyConf = ""
            If Val(txtPrnCopy.Text) = 0 Then
                MsgBox "P.O. must be finalized thru printing before you can update!", vbInformation + vbOKOnly, "Confirm P.O."
                Exit Sub
            ElseIf txtPONumber.Text = "" Then
                MsgBox "No available P.O. to confirm", vbInformation + vbOKOnly, "Confirm P.O."
                Exit Sub
            End If
            If txtConfirmNo.Text = "" Then
                MyConf = InputBox("Please type Order Confirmation No.", "Update")
                If MyConf = "" Then
                    MsgBox "No Order Confirmation.... Unable to update!", vbExclamation + vbOKOnly, "Update failed"
                    Exit Sub
                End If
                txtConfirmNo.Text = MyConf
                txtRemarks.Text = txtRemarks.Text & " Order Confirmation No. " & MyConf
            End If
            Confirm = True
            Call flxResizeEdit
            Call DisplayFlxEdit
            BtnIndx = 5
            tmrFloat.Enabled = True
            AddNew = False
            Call LockData(False)
            txtDealerCode.Locked = True
            txtDealerName.Locked = True
            txtAddress.Locked = True
            txtPODate.Locked = True
            txtPONumber.Locked = True
            cboType.Locked = True
            cmdXLS.Enabled = False
            'txtConfirmNo.Locked = False
        Case 3  'View
            SPorderSum.Show
            Unload Me
        Case 4  'Print
        
            'Add Additional Genuine supplier
            If InStr(cboSupplier.Text, "HONDA PHILS.") > 0 Or InStr(cboSupplier.Text, "YMPI") > 0 Or InStr(cboSupplier.Text, "KAWASAKI PHILS.") > 0 Or InStr(cboSupplier.Text, "SUZUKI PHILS.") > 0 Then
                Ans = MsgBox("Do you want to Print P.O. with Amount?", vbQuestion + vbYesNo, "Print")
                If Ans = vbYes Then
                    '***Access
                    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER, SPITR.PRICE " & _
                        "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.TRANSTYPE)='" & IIf(Val(txtPrnCopy.Text) > 0, "ORDER", "ORDER TEMP") & "') AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "'));"
                    'mySQL
                    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE " & _
                        "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.TRANSTYPE='" & IIf(Val(txtPrnCopy.Text) > 0, "ORDER", "ORDER TEMP") & "' AND SPITR.REFERENCE='" & txtPONumber.Text & "';"
                    
                    PRNmc.QryPrn = Qry
                    PRNmc.ParamValue1 = txtPONumber.Text
                    PRNmc.ParamValue2 = cboType.Text
                    PRNmc.ParamValue3 = txtPODate.Text
                    PRNmc.MySource = "SPPOPRICE"
                    PRNmc.Show
                Else
                    '***Access
                    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER " & _
                          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.TRANSTYPE)='" & IIf(Val(txtPrnCopy.Text) > 0, "ORDER", "ORDER TEMP") & "') AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
                    'mySQL
                    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER` " & _
                          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.TRANSTYPE='" & IIf(Val(txtPrnCopy.Text) > 0, "ORDER", "ORDER TEMP") & "' AND SPITR.REFERENCE='" & txtPONumber.Text & "' ORDER BY SPAREPARTS.PARTNUMBER;"
                    
                    PRNmc.QryPrn = Qry
                    PRNmc.ParamValue1 = txtPONumber.Text
                    PRNmc.ParamValue2 = cboType.Text
                    PRNmc.ParamValue3 = txtPODate.Text
                    PRNmc.MySource = "SPPO"
                    PRNmc.Show
                End If
            Else
                Ans = MsgBox("Do you want to Print with Part Number?", vbQuestion + vbYesNo, "Print Part Number")
                If Ans = vbYes Then
                    '***Access
                    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER, SPITR.PRICE " & _
                        "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.TRANSTYPE)='" & IIf(Val(txtPrnCopy.Text) > 0, "ORDER", "ORDER TEMP") & "') AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "'));"
                    'mySQL
                    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE " & _
                        "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.TRANSTYPE='" & IIf(Val(txtPrnCopy.Text) > 0, "ORDER", "ORDER TEMP") & "' AND SPITR.REFERENCE='" & txtPONumber.Text & "';"
                    
                    PRNmc.QryPrn = Qry
                    PRNmc.ParamValue1 = cboSupplier.Text
                    PRNmc.ParamValue2 = txtPONumber.Text
                    PRNmc.ParamValue3 = txtPODate.Text
                    PRNmc.MySource = "SPPOreplPnum"
                    PRNmc.Show
                    Exit Sub
                End If
                Ans = MsgBox("Do you want to Print P.O. with Amount?", vbQuestion + vbYesNo, "Print")
                If Ans = vbNo Then
                    '***Access
                    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER " & _
                          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.TRANSTYPE)='" & IIf(Val(txtPrnCopy.Text) > 0, "ORDER", "ORDER TEMP") & "') AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
                    'mySQL
                    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER` " & _
                          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.TRANSTYPE='" & IIf(Val(txtPrnCopy.Text) > 0, "ORDER", "ORDER TEMP") & "' AND SPITR.REFERENCE='" & txtPONumber.Text & "' ORDER BY SPAREPARTS.PARTNUMBER;"
                    
                    PRNmc.QryPrn = Qry
                    PRNmc.ParamValue1 = cboSupplier.Text
                    PRNmc.ParamValue2 = txtPONumber.Text
                    PRNmc.ParamValue3 = txtPODate.Text
                    PRNmc.MySource = "SPPOrepl"
                    PRNmc.Show
                Else
                    '***Access
                    'Qry = "SELECT SPAREPARTS.DESCRIPTION, SPITR.ORDER, SPITR.PRICE " & _
                        "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.TRANSTYPE)='" & IIf(Val(txtPrnCopy.Text) > 0, "ORDER", "ORDER TEMP") & "') AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "'));"
                    'mySQL
                    Qry = "SELECT SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE " & _
                        "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.TRANSTYPE='" & IIf(Val(txtPrnCopy.Text) > 0, "ORDER", "ORDER TEMP") & "' AND SPITR.REFERENCE='" & txtPONumber.Text & "';"
                    
                    PRNmc.QryPrn = Qry
                    PRNmc.ParamValue1 = cboSupplier.Text
                    PRNmc.ParamValue2 = txtPONumber.Text
                    PRNmc.ParamValue3 = txtPODate.Text
                    PRNmc.MySource = "SPPOreplAmt"
                    PRNmc.Show
                End If
            End If
        Case 5  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
        Case 6  'Save
            myMsg = ValidateSPOrder
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save Order?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            If AddNew Then
                cmdTrans(6).Enabled = False
                cmdTrans(7).Enabled = False
                Call SaveNewSPOrder
                cmdTrans(6).Enabled = True
                cmdTrans(7).Enabled = True
                
            ElseIf Not AddNew And Not Confirm Then
                Call SaveSPOrderEdit
                txtDealerCode.Locked = False
                txtDealerName.Locked = False
                txtAddress.Locked = False
                txtPODate.Locked = False
                txtPONumber.Locked = False
                cboSupplier.Locked = False
                cboType.Locked = False
                Call LockData(False)
                
            
            ElseIf Not AddNew And Confirm Then
                SaveSPConfirm
                txtDealerCode.Locked = False
                txtDealerName.Locked = False
                txtAddress.Locked = False
                txtPODate.Locked = False
                txtPONumber.Locked = False
                cboSupplier.Locked = False
                'txtConfirmNo.Locked = False
                Call LockData(False)
                MsgBox "Spare Parts Order confirmed successfully!", vbInformation + vbOKOnly, "Save"
            End If
                
            BtnIndx = 7
            tmrFloatRev.Enabled = True
            Call LockData(True)
            
        Case 7  'Cancel
            'flxCust.Visible = False
            flxSrch.Visible = False
            Call flxParts_Scroll
            Ans = MsgBox("Do you want to Cancel Order?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            BtnIndx = 7
            tmrFloatRev.Enabled = True
            
            If AddNew Then
                Call ClearEntries
            Else
                txtDealerCode.Locked = False
                txtDealerName.Locked = False
                txtAddress.Locked = False
                txtPODate.Locked = False
                txtPONumber.Locked = False
                txtTerms.Locked = False
                cboType.Locked = False
                cboSupplier.Locked = False
                Call flxResizeNew
                Call DisplaySPOrder(txtPONumber.Text)
            End If
            Call LockData(True)
    End Select

End Sub
Function GetOrderNo() As String
    Dim rsOrder As New ADODB.Recordset
    Dim Cntr As Integer, RefExist As Boolean
    Dim Qry As String, MyOrdNo As String, MyRef As String
    
    '***Access
    'Qry = "SELECT SPORDER.SPORDERID, SPORDER.REFERENCE, SPORDER.ORDERDATE, Month([ORDERDATE]) AS MYMONTH " & _
          "From SPORDER WHERE (((Month([ORDERDATE]))=" & Month(txtPODate.Text) & ") AND ((SPORDER.SUPPLIERID)=" & txtSuppID.Text & ")) ORDER BY SPORDER.REFERENCE DESC;"
    'mySQL
    Qry = "SELECT SPORDER.SPORDERID, SPORDER.REFERENCE, SPORDER.ORDERDATE, Month(SPORDER.ORDERDATE) AS MYMONTH " & _
          "From SPORDER WHERE Month(SPORDER.ORDERDATE)='" & Month(txtPODate.Text) & "' AND SPORDER.SUPPLIERID=" & txtSuppID.Text & " ORDER BY SPORDER.REFERENCE DESC;"
    
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        Cntr = .RecordCount
        .Close
    End With
    If InStr(cboSupplier.Text, "HONDA") > 0 Then
        MyOrdNo = Format$(txtPODate.Text, "yyyy-mm-") & Format$(Cntr + 1, "000")
    Else
        MyOrdNo = Format$(txtPODate.Text, "yy-mm-") & Format$(Cntr + 1, "00") & txtScode.Text
    End If
    RefExist = False
    Do
        '***ACCESS
        'Qry = "SELECT SPORDER.REFERENCE From SPORDER WHERE (((SPORDER.REFERENCE)='" & MyOrdNo & "'));"
        'mySQL
        Qry = "SELECT SPORDER.REFERENCE From SPORDER WHERE SPORDER.REFERENCE='" & MyOrdNo & "';"
        If CtrlNumExist(Qry) Then
            Cntr = Cntr + 1
            If InStr(cboSupplier.Text, "HONDA") > 0 Then
                MyOrdNo = Format$(txtPODate.Text, "yyyy-mm-") & Format$(Cntr + 1, "000")
            Else
                MyOrdNo = Format$(txtPODate.Text, "yy-mm-") & Format$(Cntr + 1, "00") & txtScode.Text
            End If
            RefExist = True
        Else
            RefExist = False
        End If
    Loop While RefExist
    GetOrderNo = MyOrdNo
    
End Function
Private Sub cmdTrans_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub

Private Sub cmdXLS_Click()
    POexcel.Show vbModal
End Sub


Public Sub ProcessImportedPO()
    Dim j As Integer, c As Integer, Ans As Integer
    Dim bRow As String
    
    With flxParts
        ImportPO = True
        cmdTrans(6).Enabled = False
        cmdTrans(7).Enabled = False
        LockFlx = True
        cmdFind.Enabled = False
        cmdXLS.Enabled = False
        pBar.Visible = True
        pBar.Value = 0
        pBar.Max = .Rows - 1
        For j = 1 To .Rows - 1
            ActRow = j
            txtPartNo.Text = .TextMatrix(j, 1)
            If .TextMatrix(j, 2) <> "" Then
                Call txtPartNo_LostFocus
                txtFlx.Text = .TextMatrix(j, 3)
                Call txtFlx_LostFocus
            End If
            pBar.Value = pBar.Value + 1
            DoEvents
        Next j
        ImportPO = False
        If flxParts.TextMatrix(flxParts.Rows - 1, 1) = "" And flxParts.TextMatrix(flxParts.Rows - 1, 2) = "" And flxParts.TextMatrix(flxParts.Rows - 1, 3) = "" Then flxParts.RemoveItem flxParts.Rows - 1
        cmdTrans(6).Enabled = True
        cmdTrans(7).Enabled = True
        LockFlx = False
        cmdFind.Enabled = True
        cmdXLS.Enabled = True
        
        'Identify PartNumber without Description
        bRow = ""
        c = 0
        For j = 1 To .Rows - 1
            If .TextMatrix(j, 2) = "" Then
                bRow = IIf(bRow <> "", bRow & ", " & j, "Row " & j)
                c = c + 1
            End If
        Next j
        pBar.Visible = False
        'Prompt user for invalid partnumbers!
        If c > 0 Then
            Ans = MsgBox("Invalid Part Numbers found!" & vbCrLf & bRow & vbCrLf & vbCrLf & "Do you want to remove invalid Part Number?", vbQuestion + vbYesNo, "Import P.O.")
            If Ans = vbNo Then Exit Sub
            For j = .Rows - 1 To 1 Step -1
                If .TextMatrix(j, 2) = "" Then .RemoveItem j
            Next j
            For j = 1 To .Rows - 1
                .TextMatrix(j, 0) = j
            Next j
            MsgBox "Invalid Part Number  successfully removed!", vbInformation + vbOKOnly, "Import P.O."
        End If
    End With
    
End Sub

Private Sub flxParts_Click()
    ActRow = flxParts.Row
    ActCol = flxParts.Col
    
    'FlxParts Columns: 1-PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-ID, 7-Remarks
    If ActCol = 0 Or ActRow = 0 Or LockFlx Then Exit Sub
    
    If txtSuppID.Text = "" Then
        MsgBox "Please select a Supplier from list before entering orders!", vbExclamation + vbOKOnly, "Supplier Required"
        Exit Sub
    End If
    If Not Confirm Then
      Select Case ActCol
        Case 1  'PartNumber
            With txtPartNo
                If flxParts.TextMatrix(ActRow, 1) <> "" Then
                    PrtSrch = True
                    flxSrch.Tag = flxParts.TextMatrix(ActRow, 1) & "|" & flxParts.TextMatrix(ActRow, 2) & "|" & flxParts.TextMatrix(ActRow, 3) & "|" & flxParts.TextMatrix(ActRow, 4) & "|" & flxParts.TextMatrix(ActRow, 5) & "|" & flxParts.TextMatrix(ActRow, 6) & "|" & flxParts.TextMatrix(ActRow, 7) '& "|" & flxParts.TextMatrix(ActRow, 8)
                Else
                    flxSrch.Tag = ""
                End If
                .Top = flxParts.CellTop + flxParts.Top
                .Left = flxParts.CellLeft + flxParts.Left
                .Width = flxParts.CellWidth
                .Text = flxParts.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 2
            With txtDesc
                If flxParts.TextMatrix(ActRow, 2) <> "" Then
                    txtDesc.Tag = flxParts.TextMatrix(ActRow, 2)
                Else
                    txtDesc.Tag = ""
                End If
                .Top = flxParts.CellTop + flxParts.Top
                .Left = flxParts.CellLeft + flxParts.Left
                .Width = flxParts.CellWidth
                .Text = flxParts.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 3 'Qty.
            With txtFlx
                .Top = flxParts.CellTop + flxParts.Top
                .Left = flxParts.CellLeft + flxParts.Left
                .Width = flxParts.CellWidth
                .Text = flxParts.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 4  'Unit Price     Changes must use price factor feature
            If flxParts.TextMatrix(ActRow, 1) <> "" Then
                If txtScode.Text = "" Then
                    MsgBox "Please indicate Supplier before updating price!", vbExclamation + vbOKOnly, "Price"
                    Exit Sub
                End If
                SPAddnew.sRow = ActRow
                SPAddnew.MySource = "SPORDER COST"
                SPAddnew.Show vbModal
                Call ComputeTotal
            End If
       End Select
    End If
    If Confirm Then
      'Change To: 1 - PartNumber, 2-Desc, 3-Qty,  4-Price, 5-Total, 6-NM, 7-Status, 8-Remarks, 9-ID
      Select Case ActCol
        Case 6  'Allocation, Back Order
            With txtNM
                .Top = flxParts.CellTop + flxParts.Top
                .Left = flxParts.CellLeft + flxParts.Left
                .Width = flxParts.CellWidth
                .Text = flxParts.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
       End Select
    End If
End Sub
Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxParts_Click
End Sub
Private Sub flxParts_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub
Private Sub flxParts_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If LockFlx Then
        If Val(txtOrderID.Text) = 0 Or Val(txtPrnCopy.Text) = 0 Then Exit Sub
        If Button = vbRightButton Then
            Call PopupMenu(mnuBO)
        End If
    Else
        If Button = vbRightButton Then
            Call PopupMenu(mnuRows)
        End If
    End If
End Sub

Private Sub flxParts_Scroll()
    txtFlx.Visible = False
    txtNM.Visible = False
    txtPrice.Visible = False
    txtPartNo.Visible = False
    txtDesc.Visible = False
End Sub

Private Sub flxSrch_Click()
    With flxSrch
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        FromFlx = True
        txtPartNo.Text = .TextMatrix(.Row, 1)
        'Insert Additional retrieval here
        flxParts.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 2)
        flxParts.TextMatrix(ActRow, 8) = .TextMatrix(.Row, 3) ' IIf(IsNull(rsParts!Superseding), "", rsParts!Superseding)
        flxParts.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 5) ' IIf(IsNull(rsParts!RetailPrice), 0, rsParts!RetailPrice)
        flxParts.TextMatrix(ActRow, 4) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") ' IIf(IsNull(rsParts!Cost), 0, rsParts!Cost)
        flxParts.TextMatrix(ActRow, 5) = IIf(flxParts.TextMatrix(ActRow, 3) <> "", Val(flxParts.TextMatrix(ActRow, 3)) * Val(.TextMatrix(.Row, 4)), "")
        flxParts.TextMatrix(ActRow, 5) = Format$(flxParts.TextMatrix(ActRow, 5), "#,##0.00")
        flxParts.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 6) ' IIf(IsNull(rsParts!ApplicModel), "", rsParts!ApplicModel)
        'Add ApplModel
        flxParts.TextMatrix(ActRow, 13) = IIf(IsNull(rsParts!ApplicModel), "", rsParts!ApplicModel)

    End With
End Sub

Private Sub flxSrch_EnterCell()
    With flxSrch
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        FromFlx = True
        txtPartNo.Text = .TextMatrix(.Row, 1)
        'Insert Additional retrieval here
        flxParts.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 2)
        flxParts.TextMatrix(ActRow, 8) = .TextMatrix(.Row, 3) ' IIf(IsNull(rsParts!Superseding), "", rsParts!Superseding)
        flxParts.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 5) ' IIf(IsNull(rsParts!RetailPrice), 0, rsParts!RetailPrice)
        flxParts.TextMatrix(ActRow, 4) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") ' IIf(IsNull(rsParts!Cost), 0, rsParts!Cost)
        flxParts.TextMatrix(ActRow, 5) = IIf(flxParts.TextMatrix(ActRow, 3) <> "", Val(flxParts.TextMatrix(ActRow, 3)) * Val(.TextMatrix(.Row, 4)), "")
        flxParts.TextMatrix(ActRow, 5) = Format$(flxParts.TextMatrix(ActRow, 5), "#,##0.00")
        flxParts.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 6) ' IIf(IsNull(rsParts!ApplicModel), "", rsParts!ApplicModel)
        'Add ApplModel
        flxParts.TextMatrix(ActRow, 13) = IIf(IsNull(rsParts!ApplicModel), "", rsParts!ApplicModel)

    End With
End Sub

Private Sub flxSrch_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        Tmp = Split(flxSrch.Tag, "|")
        For i = LBound(Tmp) To UBound(Tmp)
            flxParts.TextMatrix(ActRow, i + 1) = Tmp(i)
        Next i
        txtPartNo.Visible = False
        flxSrch.Visible = False
        flxParts.SetFocus
        flxParts.Col = 1
    End If
End Sub

Private Sub flxSrch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        If txtPartNo.Tag = "PN exist" Then Exit Sub
        If txtPartNo.Tag = "super" Then
            txtPartNo.SetFocus
            txtPartNo.Tag = ""
            Exit Sub
        End If
        If txtPartNo.Tag = "NxtRow" Then
            txtPartNo.Tag = ""
            Exit Sub
        End If
        flxParts.SetFocus
        flxParts.Col = 3
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub
Private Sub Form_Load()
    Dim i As Integer
    
    cmdA.Top = 15000
    cboType.AddItem "(01) - URGENT"
    cboType.AddItem "(04) - STOCK"
    cboMatch.AddItem "Whole Word"
    cboMatch.AddItem "Any Part of Word"
    Call LockData(True)
    
    SalesDateFrom.Value = "12/19/2007"
    SalesDateTo.Value = Date
    Timer1.Enabled = True
    MotoRep.sBar.Panels(6).Text = "Spare Parts Order Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
End Sub


Private Sub Label3_Click()
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch.Text)
    txtSearch.SetFocus
End Sub

Private Sub mnuAdd_Click()
    If flxParts.TextMatrix(ActRow, 3) = "" Or flxParts.TextMatrix(ActRow, 1) = "" Or flxParts.TextMatrix(ActRow, 2) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        flxParts.SetFocus
        flxParts.Col = ActCol
        Exit Sub
    Else
        If ActRow = flxParts.Rows - 1 Then flxParts.AddItem flxParts.Rows
        flxParts.SetFocus
        flxParts.Col = 1
        flxParts.Row = flxParts.Rows - 1
    End If
End Sub

Private Sub mnuCancel_Click()
    Dim Ans As Integer, MyCol As Integer, NewBO As Integer
    Dim Qry As String, MyPList As String, myRem As String
    
    'Column id
    MyCol = IIf(txtConfirmNo.Text <> "", 9, 6)

    'Validate if Packing list for P.O. already exist
    'Qry = "SELECT SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & flxParts.TextMatrix(ActRow, MyCol) & ") AND ((SPITR.PACKINGLIST)='" & txtPONumber.Text & "')) ORDER BY SPITR.TRANSDATE;"
    'mySQL
    Qry = "SELECT SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
          "WHERE SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & flxParts.TextMatrix(ActRow, MyCol) & " AND SPITR.PACKINGLIST='" & txtPONumber.Text & "' ORDER BY SPITR.TRANSDATE;"
    
    
    MyPList = GetFieldValue(Qry)
    If MyPList <> "" Then
        MsgBox flxParts.TextMatrix(ActRow, 1) & " was already received using Packing List " & MyPList & "!", vbExclamation + vbOKOnly, "Cancel Denied"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to Cancel " & flxParts.TextMatrix(ActRow, 1) & " in P.O.?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbYes Then
        '***Access
        'LYRcon.Execute "UPDATE SPITR SET SPITR.[ORDER] = 0, SPITR.REMARKS = 'CANCELLED' " & _
                       "WHERE (((SPITR.REFERENCE)='" & txtPONumber.Text & "') AND ((SPITR.SPAREPARTID)=" & flxParts.TextMatrix(ActRow, MyCol) & ") AND ((SPITR.TRANSTYPE)='ORDER'));"
        'mySQL
        LYRcon.Execute "UPDATE SPITR SET SPITR.`ORDER` = 0, SPITR.REMARKS = 'CANCELLED' " & _
                       "WHERE SPITR.REFERENCE='" & txtPONumber.Text & "' AND SPITR.SPAREPARTID=" & flxParts.TextMatrix(ActRow, MyCol) & " AND SPITR.TRANSTYPE='ORDER';"
        
        'Get Back Order
        NewBO = GetBackOrder(flxParts.TextMatrix(ActRow, MyCol))
        '***Access
        'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & NewBO & _
                    " WHERE (((SPAREPARTS.SPAREPARTID)=" & flxParts.TextMatrix(ActRow, MyCol) & "));"
        'mySQL
        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & NewBO & _
                    " WHERE SPAREPARTS.SPAREPARTID=" & flxParts.TextMatrix(ActRow, MyCol) & ";"
        
        'Display 0 qty in Flx and Update remarks both screen and table
        flxParts.TextMatrix(ActRow, 3) = 0
        txtRemarks.Text = flxParts.TextMatrix(ActRow, 1) & " cancelled by " & MotoRep.sBar.Panels(3).Text & ", " & txtRemarks.Text
        '***Access
        'Qry = "SELECT SPORDER.REMARKS From SPORDER WHERE (((SPORDER.SPORDERID)=" & txtOrderID.Text & "));"
        'mySQL
        Qry = "SELECT SPORDER.REMARKS From SPORDER WHERE SPORDER.SPORDERID=" & txtOrderID.Text & ";"
        
        myRem = GetFieldValue(Qry)
        myRem = IIf(myRem <> "", ", ", "") & flxParts.TextMatrix(ActRow, 1) & " cancelled by " & MotoRep.sBar.Panels(3).Text
        'LYRcon.Execute "UPDATE SPORDER SET SPORDER.REMARKS = '" & myRem & "' " & _
                       "WHERE (((SPORDER.SPORDERID)=" & txtOrderID.Text & "));"
        'mySQL
        LYRcon.Execute "UPDATE SPORDER SET SPORDER.REMARKS = '" & myRem & "' " & _
                       "WHERE SPORDER.SPORDERID=" & txtOrderID.Text & ";"
        MsgBox flxParts.TextMatrix(ActRow, 1) & " successfully cancelled in P.O.!", vbInformation + vbOKOnly, "Cancel Complete"
        Call ComputeTotal
    End If
End Sub

Private Sub mnuChange_Click()
    Dim Ans As Integer, MyCol As Integer, NewBO As Integer, NewQty As Integer
    Dim Qry As String, MyPList As String, myRem As String, Tmp As String
    
    'Column id
    MyCol = IIf(txtConfirmNo.Text <> "", 9, 6)

    'No Validation even if Packing list for P.O. already exist
    Ans = MsgBox("Do you want to Change Order Qty. of " & flxParts.TextMatrix(ActRow, 1) & "?", vbQuestion + vbYesNo, "Change Qty.")
    If Ans = vbYes Then
        Tmp = InputBox("Please enter New Quantity Ordered for " & flxParts.TextMatrix(ActRow, 1) & ":", "New Quantity")
        If Not IsNumeric(Tmp) Then
            MsgBox "No Valid Quantity entered!", vbExclamation + vbOKOnly, "Change Failed"
            Exit Sub
        Else
            NewQty = Val(Tmp)
            Tmp = flxParts.TextMatrix(ActRow, 3)
        End If
        '***Access
        'LYRcon.Execute "UPDATE SPITR SET SPITR.[ORDER] = " & NewQty & _
                       " WHERE (((SPITR.REFERENCE)='" & txtPONumber.Text & "') AND ((SPITR.SPAREPARTID)=" & flxParts.TextMatrix(ActRow, MyCol) & ") AND ((SPITR.TRANSTYPE)='ORDER'));"
        'mySQL
        LYRcon.Execute "UPDATE SPITR SET SPITR.`ORDER` = " & NewQty & _
                       " WHERE SPITR.REFERENCE='" & txtPONumber.Text & "' AND SPITR.SPAREPARTID=" & flxParts.TextMatrix(ActRow, MyCol) & " AND SPITR.TRANSTYPE='ORDER';"
        
        'Get Back Order
        NewBO = GetBackOrder(flxParts.TextMatrix(ActRow, MyCol))
        '***Access
        'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & NewBO & _
                    " WHERE (((SPAREPARTS.SPAREPARTID)=" & flxParts.TextMatrix(ActRow, MyCol) & "));"
        'mySQL
        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & NewBO & _
                    " WHERE SPAREPARTS.SPAREPARTID=" & flxParts.TextMatrix(ActRow, MyCol) & ";"
        
        'Display 0 qty in Flx and Update remarks both screen and table
        flxParts.TextMatrix(ActRow, 3) = NewQty
        txtRemarks.Text = flxParts.TextMatrix(ActRow, 1) & " change (" & Tmp & " to " & NewQty & ") by " & MotoRep.sBar.Panels(3).Text & ", " & txtRemarks.Text
        '***Access
        'Qry = "SELECT SPORDER.REMARKS From SPORDER WHERE (((SPORDER.SPORDERID)=" & txtOrderID.Text & "));"
        'mySQL
        Qry = "SELECT SPORDER.REMARKS From SPORDER WHERE SPORDER.SPORDERID=" & txtOrderID.Text & ";"
        myRem = GetFieldValue(Qry)
        myRem = IIf(myRem <> "", myRem & ", ", "") & flxParts.TextMatrix(ActRow, 1) & " change (" & Tmp & " to " & NewQty & ") by " & MotoRep.sBar.Panels(3).Text
        '***Access
        'LYRcon.Execute "UPDATE SPORDER SET SPORDER.REMARKS = '" & myRem & "' " & _
                       "WHERE (((SPORDER.SPORDERID)=" & txtOrderID.Text & "));"
        'mySQL
        LYRcon.Execute "UPDATE SPORDER SET SPORDER.REMARKS = '" & myRem & "' " & _
                       "WHERE SPORDER.SPORDERID=" & txtOrderID.Text & ";"
        MsgBox "Qty. of " & flxParts.TextMatrix(ActRow, 1) & " successfully changed!", vbInformation + vbOKOnly, "Changed Complete"
        Call ComputeTotal
    End If

End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer, i As Integer, nSProw As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If Not AddNew And flxParts.TextMatrix(ActRow, 6) <> "" Then  'Delete row during editing w/ previous id
            flxDel.AddItem flxParts.TextMatrix(ActRow, 6)
        ElseIf AddNew And flxParts.TextMatrix(ActRow, 7) <> "" Then
            If InStr(flxParts.TextMatrix(ActRow, 7), "NEW:") > 0 Then
                nSProw = Val(Right$(flxParts.TextMatrix(ActRow, 7), Len(flxParts.TextMatrix(ActRow, 7)) - 4))
                For i = 0 To 9
                    flxNewSP.TextMatrix(nSProw, i) = ""
                Next i
            End If
        End If
        If ActRow = 1 And flxParts.Rows = 2 Then
            flxParts.Rows = 1
            flxParts.Rows = 2
            flxParts.TextMatrix(1, 0) = 1
        Else
            flxParts.RemoveItem ActRow
            For i = 1 To flxParts.Rows - 1
                flxParts.TextMatrix(i, 0) = i
            Next i
        End If
        Call ComputeTotal
    End If
End Sub

Private Sub flxResizeEdit()
    
    'From: 1-PartNumber, 2-Desc, 3-Qty, 4-Cost, 5-Total, 6-ID, 7-Remarks, 8 - Superseding, 9-RetailPrice, 10 - SPFID, 11-MaxDisc, 12 - ApplicModel
    'Change To: 1 - PartNumber, 2-Desc, 3-Qty,4-Price, 5-Total,6-NM,7-Status,8-Remarks  9-ID
    
    With flxParts
        .Cols = 10
        .Rows = 1
        .ColWidth(0) = 800
        .ColWidth(1) = (.Width - 800) * 0.15
        .ColWidth(2) = (.Width - 800) * 0.22
        .ColWidth(3) = (.Width - 800) * 0.05
        .ColWidth(4) = (.Width - 800) * 0.1
        .ColWidth(5) = (.Width - 800) * 0.1
        .ColWidth(6) = (.Width - 800) * 0.05
        .ColWidth(7) = (.Width - 800) * 0.1
        .ColWidth(8) = (.Width - 800) * 0.21
        .ColWidth(9) = 0
        '.ColWidth(10) = 0
        '.ColWidth(11) = 0
        '.ColWidth(12) = 0
        .TextMatrix(0, 1) = "PART NUMBER"
        .TextMatrix(0, 2) = "DESCRIPTION"
        .TextMatrix(0, 3) = "QTY."
        .TextMatrix(0, 4) = "UNIT PRICE"
        .TextMatrix(0, 5) = "TOTAL PRICE"
        .TextMatrix(0, 6) = "NM"
        .TextMatrix(0, 7) = "STATUS"
        .TextMatrix(0, 8) = "REMARKS"
        .ColAlignment(3) = 4
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 4
        txtTotal.Width = .ColWidth(5)
        txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4)
        Label2(3).Left = txtTotal.Left - Label2(3).Width
        Label2(3).Visible = False
        Confirm = True
    End With
End Sub

Private Sub flxResizeNew()
    'From: 1 - PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-NM, 7-Status,8Remarks, 9-ID
    'To  : 1 - PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-ID, 7-Remarks, 8 - Superseding, 9-RetailPrice, 10 - SPFID, 11 - MaxDisc
    
    With flxParts
        
        '.Cols = 13
        .Cols = 14
        .ColWidth(0) = 800
        .ColWidth(1) = (.Width - 800) * 0.23
        .ColWidth(2) = (.Width - 800) * 0.3
        .ColWidth(3) = (.Width - 800) * 0.1
        .ColWidth(4) = (.Width - 800) * 0.17
        .ColWidth(5) = (.Width - 800) * 0.17
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        If .Cols > 10 Then .ColWidth(10) = 0
        If .Cols > 10 Then .ColWidth(11) = 0
        If .Cols > 10 Then .ColWidth(12) = 0
        .TextMatrix(0, 1) = "PART NUMBER"
        .TextMatrix(0, 2) = "DESCRIPTION"
        .TextMatrix(0, 3) = "QTY."
        .TextMatrix(0, 4) = "DEALER'S NET"
        .TextMatrix(0, 5) = "TOTAL PRICE"
        .TextMatrix(0, 6) = ""
        .TextMatrix(0, 7) = ""
        .TextMatrix(0, 8) = ""
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
    End With
    Confirm = False
End Sub

Private Sub mnuIns_Click()
    SPInsXst.MySource = "SPORDERING"
    SPInsXst.Show vbModal
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxParts
        .ColWidth(0) = 1200
        .ColWidth(1) = (.Width - 1200) * 0.23
        .ColWidth(2) = (.Width - 1200) * 0.3
        .ColWidth(3) = (.Width - 1200) * 0.1
        .ColWidth(4) = (.Width - 1200) * 0.17
        .ColWidth(5) = (.Width - 1200) * 0.17
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        If .Cols > 10 Then .ColWidth(10) = 0
        If .Cols > 10 Then .ColWidth(11) = 0
        If .Cols > 10 Then .ColWidth(12) = 0
        If .Cols > 10 Then .ColWidth(13) = 0
        pBar.Left = .Left
        pBar.Width = .Width
    End With
    txtTotal.Top = flxParts.Height + flxParts.Top
    txtTotal.Left = flxParts.Left + (flxParts.Width - (flxParts.ColWidth(5) + 400))
    txtTotal.Width = flxParts.ColWidth(5)
    Label2(3).Top = txtTotal.Top
    Label2(3).Left = txtTotal.Left - Label2(3).Width 'flxParts.Left + (flxParts.Width - (flxParts.ColWidth(4) + flxParts.ColWidth(5)))
    'Label2(3).Width = flxParts.ColWidth(4)

    lbl2.Top = txtRemarks.Top
    lbl2.Left = txtRemarks.Left - lbl2.Width
    lbl2.FontSize = txtRemarks.FontSize
    
    lblHeading(4).Top = flxParts.Top - 255
    lblHeading(5).Top = flxParts.Top - 235
    flxSrch.ColWidth(1) = flxParts.ColWidth(1)
End Sub

Private Sub SaveSPConfirm()
    Dim i As Integer
    Dim Qry As String
    
    'Update table SPORDER
    '***Access
    'LYRcon.Execute "UPDATE SPORDER SET SPORDER.CONFIRMNO = '" & txtConfirmNo.Text & "', SPORDER.REMARKS = Null WHERE (((SPORDER.REFERENCE)='" & txtPONumber.Text & "'));"
    'mySQL
    LYRcon.Execute "UPDATE SPORDER SET SPORDER.CONFIRMNO = '" & txtConfirmNo.Text & "', SPORDER.REMARKS = Null WHERE SPORDER.REFERENCE='" & txtPONumber.Text & "';"
    
    'flxparts: 1 - PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-NM, 7-Status, 8-Remarks, 9-ID
    With flxParts
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 7) = "CANCEL" Then
                'Update SPITR Cancelled parts
                '***Access
                'LYRcon.Execute "UPDATE SPITR SET SPITR.ORDER = 0, SPITR.ORDERBAL = '0', SPITR.REMARKS = '" & .TextMatrix(i, 7) & "', SPITR.PACKINGLIST = '" & txtPONumber.Text & "', SPITR.SOURCE = '" & .TextMatrix(i, 6) & "-" & .TextMatrix(i, 8) & "', SPITR.BALANCE = 0, SPITR.TRANSTYPE='ORDER TEMP', SPITR.PRICE = " & PureNum(.TextMatrix(i, 4)) & _
                           " WHERE (((SPITR.SPAREPARTID)=" & .TextMatrix(i, 9) & ") AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "'));"
                'mySQL
                LYRcon.Execute "UPDATE SPITR SET SPITR.`ORDER` = 0, SPITR.ORDERBAL = '0', SPITR.REMARKS = '" & .TextMatrix(i, 7) & "', SPITR.PACKINGLIST = '" & txtPONumber.Text & "', SPITR.SOURCE = '" & .TextMatrix(i, 6) & "-" & .TextMatrix(i, 8) & "', SPITR.BALANCE = 0, SPITR.TRANSTYPE='ORDER TEMP', SPITR.PRICE = " & Replace(.TextMatrix(i, 4), ",", "") & _
                           " WHERE SPITR.SPAREPARTID=" & .TextMatrix(i, 9) & " AND SPITR.REFERENCE='" & txtPONumber.Text & "';"
                
            Else
                'Update BackOrder in Spareparts
                '***Access
                'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.ORDEREF = '" & txtPONumber.Text & "', SPAREPARTS.BACKORDER = " & .TextMatrix(i, 3) & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 9) & "));"
                'mySQL              'SPAREPARTS.ORDEREF = '" & txtPONumber.Text & "',
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & .TextMatrix(i, 3) & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 9) & ";"
                
                '***Access
                'LYRcon.Execute "UPDATE SPITR SET SPITR.PRICE = " & PureNum(.TextMatrix(i, 4)) & ", SPITR.REMARKS = Null, SPITR.PACKINGLIST = '" & txtPONumber.Text & "', SPITR.ORDERBAL = '" & .TextMatrix(i, 3) & "',SPITR.BALANCE = " & .TextMatrix(i, 3) & ", SPITR.TRANSTYPE='ORDER'" & _
                               " WHERE (((SPITR.REFERENCE)='" & txtPONumber.Text & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 9) & "));"
                LYRcon.Execute "UPDATE SPITR SET SPITR.PRICE = " & Replace(.TextMatrix(i, 4), ",", "") & ", SPITR.REMARKS = Null, SPITR.PACKINGLIST = '" & txtPONumber.Text & "', SPITR.ORDERBAL = '" & .TextMatrix(i, 3) & "',SPITR.BALANCE = " & .TextMatrix(i, 3) & ", SPITR.TRANSTYPE='ORDER'" & _
                               " WHERE SPITR.REFERENCE='" & txtPONumber.Text & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 9) & ";"
            End If
        Next i
    End With
End Sub

Private Sub SaveSPOrderEdit()
    Dim rsOrder As New ADODB.Recordset
    Dim rsSP As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, myRem As String, spQry As String
    Dim i As Integer, MyQty As Integer, nSProw As Integer
    
    
    'Update sporder
    On Error GoTo LocalErr
    ErrMsg = "Updating Order failed @ line 916"
    '***Access
    'Qry = "SELECT SPORDER.* From SPORDER WHERE (((SPORDER.SPORDERID)=" & txtOrderID.Text & "));"
    'rsOrder.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    'If rsOrder.RecordCount > 0 Then
    '    rsOrder.MoveFirst
    '    rsOrder!ORDERTYPE = cboType.Text
    '    rsOrder!SUBMITTEDBY = txtSubmit.Text
    '    rsOrder!APPROVEDBY = txtAprv.Text
    '    rsOrder!UPDATEDBY = UserID
    '    rsOrder!LASTUPDATE = Date
    '    rsOrder!REMARKS = txtRemarks.Text
    '    rsOrder!TERMS = txtTerms.Text
    '    rsOrder.Update
    '    rsOrder.Close
    'End If
    
    'mySQL
    SpLastUpt = Format(Now, "yyyy-mm-dd hh:mm:ss")
    LYRcon.Execute "Update sporder SET ORDERTYPE = '" & cboType.Text & "', SUBMITTEDBY = '" & txtSubmit.Text & "', " & _
        "APPROVEDBY = '" & txtAprv.Text & "', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & SpLastUpt & "', " & _
        "REMARKS = '" & txtRemarks.Text & "', TERMS = '" & txtTerms.Text & "' Where sporder.SPORDERID = " & txtOrderID.Text & ";"
    
    'Delete SPITR
    ErrMsg = "Removing Previous Transaction Register @ Line 954"
    '***Access
    'LYRcon.Execute "DELETE SPITR.* From SPITR  WHERE (((SPITR.TRANSTYPE)='ORDER TEMP') AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "'));"
    
    'mySQL
    LYRcon.Execute "DELETE From SPITR  WHERE SPITR.TRANSTYPE='ORDER TEMP' AND SPITR.REFERENCE='" & txtPONumber.Text & "';"
    
    'Saving edited SP details
    With flxParts
        pBar.Visible = True
        pBar.Value = 0
        pBar.Max = .Rows - 1
        For i = 1 To .Rows - 1
            'check new spareparts
            myRem = .TextMatrix(i, 7)
            If .TextMatrix(i, 6) = "" Then
              If InStr(UCase$(myRem), "NEW") > 0 Then 'Save new SparePart
                nSProw = Val(Right$(myRem, Len(myRem) - 4))
                If flxNewSP.TextMatrix(nSProw, 0) <> "" Then
                    ErrMsg = "Updating New SP failed @ line 940 Moto-r>Spareparts, Row " & i
                    '***Access
                    'rsSP.Open "SpareParts", LYRcon, adOpenKeyset, adLockOptimistic
                    'rsSP.AddNew
                    'rsSP!PartNumber = flxNewSP.TextMatrix(nSProw, 0)
                    'rsSP!Description = flxNewSP.TextMatrix(nSProw, 1)
                    'rsSP!QtyOnHand = 0
                    'rsSP!Cost = PureNum(flxNewSP.TextMatrix(nSProw, 4))
                    'rsSP!StorePrice = Val(flxNewSP.TextMatrix(nSProw, 5))
                    'rsSP!BegBal = 0
                    'rsSP!REFERENCE = txtPONumber.Text
                    'If IsDate(txtPODate.Text) Then rsSP!DateBegBal = txtPODate.Text
                    'rsSP!Remarks = flxNewSP.TextMatrix(nSProw, 3)
                    'If Val(flxNewSP.TextMatrix(nSProw, 6)) > 0 Then rsSP!SPFid = Val(flxNewSP.TextMatrix(nSProw, 6))
                    'rsSP!MaxDiscount = Val(flxNewSP.TextMatrix(nSProw, 7))
                    'rsSP!SUPPLIERCODE = txtScode.Text
                    'rsSP!UPDATEDBY = UserID
                    'rsSP!LASTUPDATE = Date
                    'rsSP.Update
                    '.TextMatrix(i, 6) = rsSP!SparepartID
                    'rsSP.Close
                    
                    'mySQL observer updating of record by milliseconds
                    'Add Applicable model
                    SpLastUpt = Format(Now, "yyyy-mm-dd hh:mm:ss")
                    LYRcon.Execute "Insert Into spareparts( PARTNUMBER, DESCRIPTION, QTYONHAND, COST, STOREPRICE, BEGBAL, " & _
                        "REFERENCE, REMARKS, MAXDISCOUNT, SUPPLIERCODE, UPDATEDBY, LASTUPDATE, APPLICMODEL)VALUES( '" & flxNewSP.TextMatrix(nSProw, 0) & "', " & _
                        "'" & flxNewSP.TextMatrix(nSProw, 1) & "', 0, " & Val(Replace(flxNewSP.TextMatrix(nSProw, 4), ",", "")) & ", " & Val(flxNewSP.TextMatrix(nSProw, 5)) & ", " & _
                        "0, '" & txtPONumber.Text & "', '" & flxNewSP.TextMatrix(nSProw, 3) & "', '" & Val(flxNewSP.TextMatrix(nSProw, 7)) & "', " & _
                        "'" & txtScode.Text & "', " & UserID & ", '" & SpLastUpt & "', '" & flxNewSP.TextMatrix(nSProw, 10) & "');"
                          
                    Qry = "select spareparts.SparepartID from spareparts where spareparts.PARTNUMBER = '" & flxNewSP.TextMatrix(nSProw, 0) & "' And spareparts.UPDATEDBY = " & UserID & ";"
                    .TextMatrix(i, 6) = GetFieldValue(Qry)
                    
                    If IsDate(txtPODate.Text) Then LYRcon.Execute "UPDATE spareparts SET DATEBEGBAL = '" & Format(txtPODate.Text, "yyyy-mm-dd") & "' Where spareparts.SPAREPARTID = " & .TextMatrix(i, 6) & ";"
                    If Val(flxNewSP.TextMatrix(nSProw, 6)) > 0 Then LYRcon.Execute "UPDATE spareparts SET SPFID = " & Val(flxNewSP.TextMatrix(nSProw, 6)) & " Where spareparts.SPAREPARTID = " & .TextMatrix(i, 6) & ";"
                    
                    'Add new SP in Source SP
                    ErrMsg = "Updating New SP failed @ line 960 SPcon>Spareparts, Row " & i
                    LYRcon.Execute "INSERT INTO SOURCESP ( PARTNUMBER, DESCRIPTION, COST, RETAILPRICE ) " & _
                              "VALUES ('" & flxNewSP.TextMatrix(nSProw, 0) & "', '" & flxNewSP.TextMatrix(nSProw, 1) & "', " & Val(flxNewSP.TextMatrix(nSProw, 4)) & ", " & Val(flxNewSP.TextMatrix(nSProw, 5)) & ");"
                End If
              Else    'SP exist in Source SP
                
                'Check existence in Mtreptgm
                '***Access
                'Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS " & _
                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                
                'mySQL
                Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS " & _
                      "WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                
                If Not CtrlNumExist(Qry) Then
                    ErrMsg = "Updating SP failed @ line 969 Moto-r>Spareparts, Row " & i
                    
                    'rsSP.Open "SpareParts", LYRcon, adOpenKeyset, adLockOptimistic
                    'rsSP.AddNew
                    'rsSP!PartNumber = .TextMatrix(i, 1)
                    'rsSP!Description = .TextMatrix(i, 2)
                    'rsSP!QtyOnHand = 0
                    'rsSP!Cost = PureNum(.TextMatrix(i, 4))
                    'rsSP!StorePrice = Val(.TextMatrix(i, 9))
                    'rsSP!BEGBAL = 0
                    'rsSP!REFERENCE = txtPONumber.Text
                    'If IsDate(txtPODate.Text) Then rsSP!DATEBEGBAL = txtPODate.Text
                    'If Val(.TextMatrix(i, 10)) > 0 Then rsSP!SPFid = Val(.TextMatrix(i, 10))
                    'rsSP!MaxDiscount = Val(.TextMatrix(i, 11))
                    'rsSP!SUPPLIERCODE = txtScode.Text
                    'spQry = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE (((SUPPLIER.SUPPLIERCODE)='" & txtScode.Text & "'));"
                    'rsSP!Remarks = GetFieldValue(spQry)
                    'old
                    'rsSP!Remarks = "GENUINE"
                    'rsSP!UPDATEDBY = UserID
                    'rsSP!LASTUPDATE = Date
                    'rsSP.Update
                    '.TextMatrix(i, 6) = rsSP!SparepartID
                    'rsSP.Close
                    
                    'mySQL
                    'Remarks query from sptype field record
                    spQry = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE SUPPLIER.SUPPLIERCODE='" & txtScode.Text & "';"
                    LYRcon.Execute "INSERT INTO spareparts( PARTNUMBER, DESCRIPTION, QTYONHAND, COST, STOREPRICE, BEGBAL, REFERENCE, MAXDISCOUNT, " & _
                        "SUPPLIERCODE, REMARKS, UPDATEDBY, LASTUPDATE)VALUES('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', 0, " & Replace(.TextMatrix(i, 4), ",", "") & ", " & _
                        "" & Val(.TextMatrix(i, 9)) & ", 0, '" & Format(txtPODate.Text, "yyyy-mm-dd hh:mm:ss") & "', " & Val(.TextMatrix(i, 11)) & ", " & _
                        "'" & txtScode.Text & "', '" & GetFieldValue(spQry) & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
                    
                    Qry = "select spareparts.SPAREPARTID from spareparts where spareparts.PARTNUMBER = '" & .TextMatrix(i, 1) & "';"
                    .TextMatrix(i, 6) = GetFieldValue(Qry)
                    
                    If IsDate(txtPODate.Text) Then LYRcon.Execute "UPDATE spareparts SET DATEBEGBAL = '" & Format(txtPODate.Text, "yyyy-mm-dd") & "' Where spareparts.SPAREPARTID = " & .TextMatrix(i, 6) & ";"
                    If Val(.TextMatrix(i, 10)) > 0 Then LYRcon.Execute "UPDATE spareparts SET SPFID = " & Val(.TextMatrix(i, 10)) & " Where spareparts.SPAREPARTID = " & .TextMatrix(i, 6) & ";"
                      
                    'Update Source SP Prices
                    ErrMsg = "Updating SP Prices failed @ line 989 SPcon>Spareparts, Row " & i
                    '***Access
                    'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & PureNum(.TextMatrix(i, 4)) & ", SOURCESP.RETAILPRICE = " & Val(.TextMatrix(i, 9)) & _
                                  " WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                    'mySQL Partnumber reference of Update criteria
                    LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & Replace(.TextMatrix(i, 4), ",", "") & ", SOURCESP.RETAILPRICE = " & Val(.TextMatrix(i, 9)) & _
                                  " WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                    
                    If InStr(myRem, "Edesc") > 0 Then
                        'Update Description in Source SP
                        ErrMsg = "Updating SP Description failed @ line 995 SPcon>Spareparts, Row " & i
                        '****Access
                        'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                        
                    End If
                Else
                    .TextMatrix(i, 6) = GetSPid(.TextMatrix(i, 1))
                    If InStr(UCase$(myRem), "EDESC") > 0 Then
                        'Update Description in LYRcon
                        ErrMsg = "Updating SP Description failed @ line 1003 Moto-r>Spareparts, Row " & i
                        '***Access
                        'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                        
                        'Update Description in Source SP
                        ErrMsg = "Updating SP Description failed @ line 1007 SPcon>Spareparts, Row " & i
                        '****Access
                        'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                        
                    End If
                    If InStr(UCase$(myRem), "ECOST") > 0 Then   'Update Pricing
                        'Update cost in Source SP
                        ErrMsg = "Updating SP Prices failed @ line 1013 SPcon>Spareparts, Row " & i
                        '***Access
                        'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & PureNum(.TextMatrix(i, 4)) & ", SOURCESP.RETAILPRICE = " & Val(.TextMatrix(i, 9)) & _
                                  " WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & Replace(.TextMatrix(i, 4), ",", "") & ", SOURCESP.RETAILPRICE = " & Val(.TextMatrix(i, 9)) & _
                                  " WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                        
                        'Update cost in LYRcon
                        ErrMsg = "Updating SP Prices failed @ line 1017 Moto-r>Spareparts, Row " & i
                        '***Access
                        'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & PureNum(.TextMatrix(i, 4)) & ", SPAREPARTS.STOREPRICE = " & Val(.TextMatrix(i, 9)) & ", SPAREPARTS.SPFID = " & Val(.TextMatrix(i, 10)) & ", SPAREPARTS.MAXDISCOUNT = " & Val(.TextMatrix(i, 11)) & _
                                  " WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & Replace(.TextMatrix(i, 4), ",", "") & ", SPAREPARTS.STOREPRICE = " & Val(.TextMatrix(i, 9)) & ", SPAREPARTS.SPFID = " & Val(.TextMatrix(i, 10)) & ", SPAREPARTS.MAXDISCOUNT = " & Val(.TextMatrix(i, 11)) & _
                                  " WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                        
                    End If
                    
                End If
              End If
              'Save SPITR
              ErrMsg = "Updating transaction register failed @ line 1025 Moto-r>SPITR, Row " & i
              
              '***Access
              'rsITR.Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
              'rsITR.AddNew
              'rsITR!REFERENCE = txtPONumber.Text
              'rsITR!SPAREPARTID = flxParts.TextMatrix(i, 6)
              'rsITR!TRANSDATE = txtPODate.Text
              'rsITR!TRANSTYPE = "ORDER TEMP"
              'rsITR!Order = Val(flxParts.TextMatrix(i, 3))
              'rsITR!Price = Val(Replace(flxParts.TextMatrix(i, 4), ",", ""))
              'rsITR!Source = cboSupplier.Text
              'rsITR!UPDATEDBY = UserID
              'rsITR!LASTUPDATE = Now
              'rsITR.Update
              'rsITR.Close
              
              'mySQL
              LYRcon.Execute "Insert Into spitr(REFERENCE, SPAREPARTID, TRANSDATE, TRANSTYPE, " & _
                    "`ORDER`, PRICE, SOURCE, UPDATEDBY, LASTUPDATE)VALUES('" & txtPONumber.Text & "', " & _
                    "" & flxParts.TextMatrix(i, 6) & ", '" & Format(txtPODate.Text, "yyyy-mm-dd") & "', 'ORDER TEMP', " & _
                    "" & Val(flxParts.TextMatrix(i, 3)) & ", " & Val(Replace(flxParts.TextMatrix(i, 4), ",", "")) & ", " & _
                    "'" & cboSupplier.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
            Else
              'Save edit SparePart Desc
              If InStr(UCase$(myRem), "EDESC") > 0 Then
                  'Update Description in LYRcon
                  ErrMsg = "Updating SP Description failed @ line 1043 Moto-r>Spareparts, Row " & i
                  '***Access
                  'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                              "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                  'mySQL
                  LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                              "WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                  
                  'Update Description in Source SP
                  ErrMsg = "Updating SP Description failed @ line 1047 SPcon>Spareparts, Row " & i
                  '***Access
                  'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                              "WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                  'mySQL
                  LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                           " WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                  
              End If
              If InStr(UCase$(myRem), "ECOST") > 0 Then   'Update Pricing
                  'Update cost in Source SP
                  ErrMsg = "Updating SP Prices failed @ line 1013 SPcon>Spareparts, Row " & i
                  '***Access
                  'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & PureNum(.TextMatrix(i, 4)) & ", SOURCESP.RETAILPRICE = " & Val(.TextMatrix(i, 9)) & _
                          " WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                  'mySQL
                  LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & Val(Replace(.TextMatrix(i, 4), ",", "")) & ", SOURCESP.RETAILPRICE = " & Val(.TextMatrix(i, 9)) & _
                          " WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                  
                  'Update cost in LYRcon
                  ErrMsg = "Updating SP Prices failed @ line 1017 Moto-r>Spareparts, Row " & i
                  '***Access
                  'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & PureNum(.TextMatrix(i, 4)) & ", SPAREPARTS.STOREPRICE = " & Val(.TextMatrix(i, 9)) & ", SPAREPARTS.SPFID = " & Val(.TextMatrix(i, 10)) & ", SPAREPARTS.MAXDISCOUNT = " & Val(.TextMatrix(i, 11)) & _
                           " WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                  'mySQL
                  'Add Applicmodel
                  LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & Val(Replace(.TextMatrix(i, 4), ",", "")) & ", SPAREPARTS.STOREPRICE = " & Val(.TextMatrix(i, 9)) & ", SPAREPARTS.SPFID = " & Val(.TextMatrix(i, 10)) & ", SPAREPARTS.MAXDISCOUNT = " & Val(.TextMatrix(i, 11)) & ", SPAREPARTS.APPLICMODEL = '" & .TextMatrix(i, 13) & "'" & _
                           " WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                  
              End If
              'Save SPITR
              ErrMsg = "Updating transaction register failed @ line 1025 Moto-r>SPITR, Row " & i
              
              '***Access
              'rsITR.Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
              'rsITR.AddNew
              'rsITR!REFERENCE = txtPONumber.Text
              'rsITR!SPAREPARTID = flxParts.TextMatrix(i, 6)
              'rsITR!TRANSDATE = txtPODate.Text
              'rsITR!TRANSTYPE = "ORDER TEMP"
              'rsITR!Order = Val(flxParts.TextMatrix(i, 3))
              'rsITR!Price = Val(Replace(flxParts.TextMatrix(i, 4), ",", ""))
              'rsITR!Source = cboSupplier.Text
              'rsITR!UPDATEDBY = UserID
              'rsITR!LASTUPDATE = Now
              'rsITR.Update
              'rsITR.Close
              
              'mySQL
              LYRcon.Execute "INSERT INTO spitr( REFERENCE, SPAREPARTID, TRANSDATE, TRANSTYPE, `ORDER`, PRICE, " & _
                  "SOURCE, UPDATEDBY, LASTUPDATE)VALUES('" & txtPONumber.Text & "', " & flxParts.TextMatrix(i, 6) & ", " & _
                  "'" & Format(txtPODate.Text, "yyyy-mm-dd") & "', 'ORDER TEMP', " & Val(flxParts.TextMatrix(i, 3)) & ", " & _
                  "" & Val(Replace(flxParts.TextMatrix(i, 4), ",", "")) & ", '" & cboSupplier.Text & "', " & UserID & ", " & _
                  "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        
            End If
            pBar.Value = pBar.Value + 1
            DoEvents
        Next i
    End With
    pBar.Visible = False
    MsgBox "Spare Parts Order successfully saved!", vbInformation + vbOKOnly, "Save"
    Exit Sub
LocalErr:
    MsgBox ErrMsg & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please inform M.I.S. Department", vbCritical + vbOKOnly, "Error"
    pBar.Visible = False
    
End Sub

Private Sub SaveNewSPOrder()
    Dim rsOrder As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim rsSP As New ADODB.Recordset
    Dim Qry As String, myRem As String, ErrMsg As String, spQry As String
    Dim i As Integer, MyQty As Integer, nSProw As Integer
    'mySQL variable
    Dim mySUPPLIERID As Integer
    
    On Error GoTo LocalErr
    'Saving SPOrder
    ErrMsg = "Saving failed while processing line 1081 SPORDER"
    
    '***Access
    'rsOrder.Open "SPOrder", LYRcon, adOpenKeyset, adLockOptimistic
    'rsOrder.AddNew
    'rsOrder!DEALERCODE = txtDealerCode.Text
    'rsOrder!DEALERNAME = txtDealerName.Text
    'rsOrder!ADDRESS = txtAddress.Text
    'rsOrder!REFERENCE = txtPONumber.Text
    'rsOrder!ORDERTYPE = cboType.Text
    'rsOrder!ORDERDATE = txtPODate.Text
    'rsOrder!SUBMITTEDBY = txtSubmit.Text
    'rsOrder!APPROVEDBY = txtAprv.Text
    'old
    'rsOrder!ConfirmNo = txtConfirmNo.Text
    'rsOrder!UPDATEDBY = UserID
    'rsOrder!LASTUPDATE = Date
    'rsOrder!REMARKS = txtRemarks.Text
    'rsOrder!TERMS = txtTerms.Text
    'rsOrder!SUPPLIERID = IIf(txtSuppID.Text <> "", Val(txtSuppID.Text), 0)
    'rsOrder.Update
    'txtOrderID.Text = rsOrder!SPORDERID
    'rsOrder.Close
    
    'mySQL
    
    If cboSupplier.Text <> "" Then
        txtPONumber.Text = GetOrderNo
    End If
    
    mySUPPLIERID = IIf(txtSuppID.Text <> "", Val(txtSuppID.Text), 0)
    myLastUpdate = Format(Now, "yyyy-mm-dd hh:mm:ss")
    LYRcon.Execute "Insert Into sporder( DEALERCODE, DEALERNAME, ADDRESS, REFERENCE, ORDERTYPE, ORDERDATE, " & _
        "SUBMITTEDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE, REMARKS, TERMS, SUPPLIERID)" & _
        "VALUES('" & txtDealerCode.Text & "', '" & txtDealerName.Text & "', '" & txtAddress.Text & "', " & _
        "'" & txtPONumber.Text & "', '" & cboType.Text & "', '" & Format(txtPODate.Text, "yyyy-mm-dd") & "', " & _
        "'" & txtSubmit.Text & "', '" & txtAprv.Text & "', " & UserID & ", '" & myLastUpdate & "', " & _
        "'" & txtRemarks.Text & "', '" & txtTerms.Text & "', " & mySUPPLIERID & ");"
    Qry = "select sporder.SPORDERID from sporder where sporder.DEALERCODE='" & txtDealerCode.Text & "' AND sporder.LASTUPDATE='" & myLastUpdate & "';"
    txtOrderID.Text = GetFieldValue(Qry)
    
    
    
    'Saving SP details
    With flxParts
        pBar.Visible = True
        pBar.Value = 0
        pBar.Max = .Rows - 1
        For i = 1 To .Rows - 1
            'check new spareparts
            myRem = .TextMatrix(i, 7)
            If InStr(UCase$(myRem), "NEW") > 0 Then 'Save new SparePart
                nSProw = Val(Right$(myRem, Len(myRem) - 4))
                If flxNewSP.TextMatrix(nSProw, 0) <> "" Then
                    ErrMsg = "Saving New SP failed @ line 1110 Moto-r>Spareparts, Row " & i
                    
                    '***Access
                    'rsSP.Open "SpareParts", LYRcon, adOpenKeyset, adLockOptimistic
                    'rsSP.AddNew
                    'rsSP!PARTNUMBER = flxNewSP.TextMatrix(nSProw, 0)
                    'rsSP!Description = flxNewSP.TextMatrix(nSProw, 1)
                    'rsSP!QTYONHAND = 0
                    'rsSP!Cost = PureNum(flxNewSP.TextMatrix(nSProw, 4))
                    'rsSP!STOREPRICE = Val(flxNewSP.TextMatrix(nSProw, 5))
                    'rsSP!BEGBAL = 0
                    'rsSP!REFERENCE = txtPONumber.Text
                    'If IsDate(txtPODate.Text) Then rsSP!DATEBEGBAL = txtPODate.Text
                    'rsSP!REMARKS = flxNewSP.TextMatrix(nSProw, 3)
                    'If Val(flxNewSP.TextMatrix(nSProw, 6)) > 0 Then rsSP!SPFid = Val(flxNewSP.TextMatrix(nSProw, 6))
                    'rsSP!MAXDISCOUNT = Val(flxNewSP.TextMatrix(nSProw, 7))
                    'rsSP!SUPPLIERCODE = txtScode.Text
                    'rsSP!UPDATEDBY = UserID
                    'rsSP!LASTUPDATE = Date
                    'rsSP.Update
                    '.TextMatrix(i, 6) = rsSP!SparepartID
                    'rsSP.Close
                    
                    'mySQL
                    
                    LYRcon.Execute "INSERT INTO spareparts( PARTNUMBER, DESCRIPTION, QTYONHAND, COST, STOREPRICE, " & _
                        "BEGBAL, REFERENCE, REMARKS, MAXDISCOUNT, SUPPLIERCODE, UPDATEDBY, LASTUPDATE, APPLICMODEL) " & _
                        "VALUES('" & flxNewSP.TextMatrix(nSProw, 0) & "', '" & flxNewSP.TextMatrix(nSProw, 1) & "', " & _
                        "0, " & Val(Replace(flxNewSP.TextMatrix(nSProw, 4), ",", "")) & ", " & Val(flxNewSP.TextMatrix(nSProw, 5)) & ", " & _
                        "0, '" & txtPONumber.Text & "', '" & flxNewSP.TextMatrix(nSProw, 3) & "', " & Val(flxNewSP.TextMatrix(nSProw, 7)) & ", " & _
                        "'" & txtScode.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', '" & flxNewSP.TextMatrix(nSProw, 10) & "');"
                    
                    Qry = "select spareparts.SPAREPARTID from spareparts where spareparts.PARTNUMBER = '" & flxNewSP.TextMatrix(nSProw, 0) & "';"
                    .TextMatrix(i, 6) = GetFieldValue(Qry)
                    
                    If IsDate(txtPODate.Text) Then LYRcon.Execute "UPDATE spareparts SET DATEBEGBAL = '" & Format(txtPODate.Text, "yyyy-mm-dd") & "' Where spareparts.SPAREPARTID = '" & .TextMatrix(i, 6) & "';"
                    If Val(flxNewSP.TextMatrix(nSProw, 6)) > 0 Then LYRcon.Execute "UPDATE spareparts SET SPFID = " & Val(flxNewSP.TextMatrix(nSProw, 6)) & " Where spareparts.SPAREPARTID = " & .TextMatrix(i, 6) & ";"
                     
                    'Add new SP in Source SP
                    ErrMsg = "Saving New SP execution failed @ line 1130 SPcon>Spareparts, Row " & i
                    LYRcon.Execute "INSERT INTO SOURCESP ( PARTNUMBER, DESCRIPTION, COST, RETAILPRICE ) " & _
                              "VALUES ('" & flxNewSP.TextMatrix(nSProw, 0) & "', '" & flxNewSP.TextMatrix(nSProw, 1) & "', " & Val(flxNewSP.TextMatrix(nSProw, 4)) & ", " & Val(flxNewSP.TextMatrix(nSProw, 5)) & ");"
                End If
            Else    'SP exist in Source SP
                'Check existence in Mtreptgm
                'Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS " & _
                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                'mySQL
                Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS " & _
                      "WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                
                If Not CtrlNumExist(Qry) Then
                    ErrMsg = "Saving existing SP failed @ line 1139 Moto-r>Spareparts, Row " & i
                    'rsSP.Open "SpareParts", LYRcon, adOpenKeyset, adLockOptimistic
                    'rsSP.AddNew
                    'rsSP!PARTNUMBER = .TextMatrix(i, 1)
                    'rsSP!Description = .TextMatrix(i, 2)
                    'rsSP!QTYONHAND = 0
                    'rsSP!Cost = PureNum(.TextMatrix(i, 4))
                    'rsSP!STOREPRICE = Val(.TextMatrix(i, 9))
                    'rsSP!BEGBAL = 0
                    'rsSP!REFERENCE = txtPONumber.Text
                    'If Val(.TextMatrix(i, 10)) > 0 Then rsSP!SPFid = Val(.TextMatrix(i, 10))
                    'rsSP!MAXDISCOUNT = Val(.TextMatrix(i, 11))
                    'If IsDate(txtPODate.Text) Then rsSP!DATEBEGBAL = txtPODate.Text
                    'rsSP!APPLICMODEL = .TextMatrix(i, 12)
                    'rsSP!SUPPLIERCODE = txtScode.Text
                    'spQry = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE (((SUPPLIER.SUPPLIERCODE)='" & txtScode.Text & "'));"
                    'rsSP!REMARKS = GetFieldValue(spQry)
                    'rsSP!UPDATEDBY = UserID
                    'rsSP!LASTUPDATE = Date
                    'rsSP.Update
                    '.TextMatrix(i, 6) = rsSP!SparepartID
                    'rsSP.Close
                    'mySQL
                    'Query for remarks
                    spQry = "SELECT SUPPLIER.SPTYPE, SUPPLIER.SUPPLIERCODE From SUPPLIER WHERE SUPPLIER.SUPPLIERCODE='" & txtScode.Text & "';"
                    
                    'LYRcon.Execute "INSERT INTO spareparts( PARTNUMBER, DESCRIPTION, QTYONHAND, COST, " & _
                        "STOREPRICE, BEGBAL, REFERENCE, MAXDISCOUNT, DATEBEGBAL, APPLICMODEL, " & _
                        "SUPPLIERCODE, REMARKS, UPDATEDBY, LASTUPDATE)VALUES('" & .TextMatrix(i, 1) & "', " & _
                        "'" & .TextMatrix(i, 2) & "', 0, " & Val(Replace(.TextMatrix(i, 4), ",", "")) & ", " & _
                        "" & Val(.TextMatrix(i, 9)) & ", 0, '" & txtPONumber.Text & "', " & Val(.TextMatrix(i, 11)) & ", " & _
                        "'" & Format(txtPODate.Text, "yyyy-mm-dd hh:mm:ss") & "', '" & .TextMatrix(i, 12) & "', " & _
                        "'" & txtScode.Text & "', '" & GetFieldValue(spQry) & "', " & UserID & ", " & _
                        "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
                    'Change Applicable model '" & .TextMatrix(i, 12=13) & "'
                    LYRcon.Execute "INSERT INTO spareparts( PARTNUMBER, DESCRIPTION, QTYONHAND, COST, " & _
                        "STOREPRICE, BEGBAL, REFERENCE, MAXDISCOUNT, DATEBEGBAL, APPLICMODEL, " & _
                        "SUPPLIERCODE, REMARKS, UPDATEDBY, LASTUPDATE)VALUES('" & .TextMatrix(i, 1) & "', " & _
                        "'" & .TextMatrix(i, 2) & "', 0, " & Val(Replace(.TextMatrix(i, 4), ",", "")) & ", " & _
                        "" & Val(.TextMatrix(i, 9)) & ", 0, '" & txtPONumber.Text & "', " & Val(.TextMatrix(i, 11)) & ", " & _
                        "'" & Format(txtPODate.Text, "yyyy-mm-dd hh:mm:ss") & "', '" & .TextMatrix(i, 13) & "', " & _
                        "'" & txtScode.Text & "', '" & GetFieldValue(spQry) & "', " & UserID & ", " & _
                        "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
               
                    
                    Qry = "select spareparts.SPAREPARTID from spareparts where spareparts.PARTNUMBER = '" & .TextMatrix(i, 1) & "';"
                    .TextMatrix(i, 6) = GetFieldValue(Qry)
                    If Val(.TextMatrix(i, 10)) > 0 Then
                        LYRcon.Execute "UPDATE spareparts SET SPFID = " & Val(.TextMatrix(i, 10)) & " Where spareparts.SPAREPARTID = " & .TextMatrix(i, 6) & ";"
                    End If
                
                    'Update Source SP Prices
                    ErrMsg = "Saving existing SP Prices failed @ line 1159 SPcon>Spareparts, Row " & i
                    '***Access
                    'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & PureNum(.TextMatrix(i, 4)) & ", SOURCESP.RETAILPRICE = " & Val(.TextMatrix(i, 9)) & _
                                  " WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                    'mySQL
                    LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & Val(Replace(.TextMatrix(i, 4), ",", "")) & ", SOURCESP.RETAILPRICE = " & Val(.TextMatrix(i, 9)) & _
                                  " WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                    
                    If InStr(UCase$(myRem), "EDESC") > 0 Then
                        'Update Description in Source SP
                        ErrMsg = "Saving existing SP Description failed @ line 1164 SPcon>Spareparts, Row " & i
                        '***Access
                        'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                    End If
                Else
                    .TextMatrix(i, 6) = GetSPid(.TextMatrix(i, 1))
                    If InStr(UCase$(myRem), "EDESC") > 0 Then
                        'Update Description in LYRcon
                        ErrMsg = "Saving existing SP Description failed @ line 1172 Moto-r>Spareparts, Row " & i
                        '***Access
                        'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                        
                        'Update Description in Source SP
                        ErrMsg = "Saving existing SP Description failed @ line 1176 SPcon>Spareparts, Row " & i
                        '***Access
                        'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                        
                    End If
                    If InStr(UCase$(myRem), "ECOST") > 0 Then   'Update Pricing
                        'Update cost in Source SP
                        ErrMsg = "Saving existing SP Prices failed @ line 1182 SPcon>Spareparts, Row " & i
                        '***Access
                        'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & PureNum(.TextMatrix(i, 4)) & ", SOURCESP.RETAILPRICE = " & Val(.TextMatrix(i, 9)) & _
                                  " WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.COST = " & Val(Replace(.TextMatrix(i, 4), ",", "")) & ", SOURCESP.RETAILPRICE = " & Val(.TextMatrix(i, 9)) & _
                                  " WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                        
                        'Update cost in LYRcon
                        ErrMsg = "Saving existing SP Description failed @ line 1186 Moto-r>Spareparts, Row " & i
                        '***Access
                        'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & PureNum(.TextMatrix(i, 4)) & ", SPAREPARTS.STOREPRICE = " & Val(.TextMatrix(i, 9)) & ", SPAREPARTS.SPFID = " & Val(.TextMatrix(i, 10)) & ", SPAREPARTS.MAXDISCOUNT = " & Val(.TextMatrix(i, 11)) & _
                                  " WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'mySQL
                        'Add Applicable model
                        'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & Val(Replace(.TextMatrix(i, 4), ",", "")) & ", SPAREPARTS.STOREPRICE = " & Val(.TextMatrix(i, 9)) & ", SPAREPARTS.SPFID = " & Val(.TextMatrix(i, 10)) & ", SPAREPARTS.MAXDISCOUNT = " & Val(.TextMatrix(i, 11)),  & _
                                  " SPAREPARTS.COST = '" & .TextMatrix(1, 13) & "' WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                        
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.COST = " & Val(Replace(.TextMatrix(i, 4), ",", "")) & ", SPAREPARTS.STOREPRICE = " & Val(.TextMatrix(i, 9)) & ", SPAREPARTS.SPFID = " & Val(.TextMatrix(i, 10)) & ", SPAREPARTS.MAXDISCOUNT = " & Val(.TextMatrix(i, 11)) & ", SPAREPARTS.APPLICMODEL = '" & .TextMatrix(i, 13) & "'" & _
                                  " SPAREPARTS.COST = '" & .TextMatrix(1, 13) & "' WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"

                    End If
                End If
            End If
            'Save SPITR
            ErrMsg = "Saving transaction register failed @ line 1193, Row " & i
            
            '***Access
            'rsITR.Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
            'rsITR.AddNew
            'rsITR!REFERENCE = txtPONumber.Text
            'rsITR!SPAREPARTID = flxParts.TextMatrix(i, 6)
            'sITR!TRANSDATE = txtPODate.Text
            'rsITR!TRANSTYPE = "ORDER TEMP"
            'rsITR!Order = Val(flxParts.TextMatrix(i, 3))
            'rsITR!Price = Val(Replace(flxParts.TextMatrix(i, 4), ",", ""))
            'rsITR!Source = cboSupplier.Text
            'rsITR!UPDATEDBY = UserID
            'rsITR!LASTUPDATE = Now
            'rsITR.Update
            'rsITR.Close
            
            pBar.Value = pBar.Value + 1
            DoEvents
           
            'mySQL
            LYRcon.Execute "Insert Into spitr( REFERENCE, SPAREPARTID, TRANSDATE, TRANSTYPE, `ORDER`, " & _
                "PRICE, SOURCE, UPDATEDBY, LASTUPDATE)VALUES( '" & txtPONumber.Text & "', " & _
                "" & flxParts.TextMatrix(i, 6) & ", '" & Format(txtPODate.Text, "yyyy-mm-dd") & "', " & _
                "'ORDER TEMP', " & Val(flxParts.TextMatrix(i, 3)) & ", " & Val(Replace(flxParts.TextMatrix(i, 4), ",", "")) & ", " & _
                "'" & cboSupplier.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
            
        Next i
    End With
    pBar.Visible = False
    MsgBox "Spare Parts Order successfully saved!", vbInformation + vbOKOnly, "Save"
    Exit Sub
LocalErr:
    MsgBox ErrMsg & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please inform M.I.S. Department", vbCritical + vbOKOnly, "Error"
    pBar.Visible = False
End Sub

Function GetSPid(PartNo As String) As Integer
    Dim rsPNum As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER From SPAREPARTS " & _
          "WHERE (((SPAREPARTS.PARTNUMBER)='" & PartNo & "'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER From SPAREPARTS " & _
          "WHERE SPAREPARTS.PARTNUMBER='" & PartNo & "';"
    
    With rsPNum
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetSPid = !SPAREPARTID
        Else
            GetSPid = 0
        End If
        .Close
    End With
End Function
Private Sub LoadSpareParts()
    Dim rsSettings As New ADODB.Recordset
    Dim rsShipper As New ADODB.Recordset
    Dim Qry As String
    Dim TtlRec As Integer
    Dim pPct As Single
        
    
    Qry = "SELECT SOURCESP.SC, SOURCESP.PARTNUMBER, SOURCESP.DESCRIPTION, SOURCESP.SUPERSEDING, SOURCESP.COST, SOURCESP.RETAILPRICE, SOURCESP.APPLICMODEL " & _
          "From SOURCESP ORDER BY SOURCESP.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSrch.DataSource = rsParts
    With flxSrch
        .Width = flxParts.ColWidth(1)
        .ColWidth(0) = 0
        .ColWidth(1) = flxParts.ColWidth(1) - 300
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .Row = 0
        .Col = 3
        .CellAlignment = 4
        .Refresh
        '.Col = 4
        '.CellAlignment = 4
    End With
    
    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE, SUPPLIER.ADDRESS " & _
          "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
    With rsSupplier
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            'cboSupplier.Clear
            Do While Not .EOF
                cboSupplier.AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
    End With
    
End Sub

Function ValidateSPOrder() As String
    Dim vMsg As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If Not IsDate(txtPODate.Text) Then vMsg = "Invalid P.O. Date"
    If txtSuppID.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Supplier"
    If txtPONumber.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field P.O. Number"
    If cboType.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Order Type"
    'If txtDealerCode.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Dealer Code"
    If txtDealerName.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Dealer Name"
    If txtAddress.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Dealer Address"
    If txtSubmit.Text = "" Then vMsg = vMsg & vbCrLf & "The field Submitted By must not be empty"
    If txtAprv.Text = "" Then vMsg = vMsg & vbCrLf & "The field Approved By must not be empty"
    'If Confirm And txtConfirmNo.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Confirmation No."
    If flxParts.Rows > 2 Then
        If flxParts.TextMatrix(flxParts.Rows - 1, 1) = "" And flxParts.TextMatrix(flxParts.Rows - 1, 2) = "" And flxParts.TextMatrix(flxParts.Rows - 1, 3) = "" Then flxParts.RemoveItem flxParts.Rows - 1
        Cntr = 0
        For i = 1 To 3
            If flxParts.TextMatrix(flxParts.Rows - 1, i) = "" Then Cntr = Cntr + 1
        Next i
    Else
        If flxParts.TextMatrix(1, 1) = "" Then vMsg = vMsg & vbCrLf & "No Spare Parts ordered!"
    End If
    If Cntr > 0 Then vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Spare Parts Ordered Units"

    'Validate Qty for 0 order
    If Not Confirm Then
        For i = 1 To flxParts.Rows - 1
            If Val(flxParts.TextMatrix(i, 3)) = 0 Then vMsg = vMsg & vbCrLf & "Zero (0) Quantity at  row " & i
        Next i
    End If
    ValidateSPOrder = vMsg
    
End Function

Public Sub DisplaySPOrder(MyRef As String)
    Dim rsOrder As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    '***Access
    'Qry = "SELECT SPORDER.* From SPORDER " & _
          "WHERE (((SPORDER.REFERENCE)='" & MyRef & "'));"
    'mySQL
    Qry = "SELECT SPORDER.* From SPORDER " & _
          "WHERE SPORDER.REFERENCE='" & MyRef & "';"
    
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Call ClearEntries
            .MoveFirst
            '***Access
            'txtDealerCode.Text = !DEALERCODE
            'mySQL
            txtDealerCode.Text = IIf(IsNull(!DEALERCODE), "", !DEALERCODE)
            
            txtDealerName.Text = !DEALERNAME
            txtAddress.Text = !Address
            txtPODate.Text = !ORDERDATE
            txtPONumber.Text = !REFERENCE
            cboType.Text = !OrderType
            txtConfirmNo.Text = IIf(IsNull(!ConfirmNo), "", !ConfirmNo)
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks) & IIf(txtConfirmNo.Text <> "", " Confirmation No. " & txtConfirmNo.Text, "")
            txtTerms.Text = IIf(IsNull(!Terms), "", !Terms)
            txtSubmit.Text = !SubmittedBy
            txtAprv.Text = !ApprovedBy
            txtPrnCopy.Text = IIf(IsNull(!PRINTCOPY), 0, !PRINTCOPY)
            txtSuppID.Text = IIf(IsNull(!SUPPLIERID), "", !SUPPLIERID)
            txtOrderID.Text = !SPORDERID
        End If
        .Close
    End With
    If txtSuppID.Text <> "" Then Call DisplaySupplier
    
    If txtConfirmNo.Text <> "" Then 'Display order with allocation
        Call flxResizeEdit
        Call DisplayFlxEdit
        Call LockData(True)
        Call ComputeTotal
        cmdTrans(1).Enabled = False
        cmdTrans(2).Enabled = False
        Exit Sub
    End If
    For i = 0 To 5
        cmdTrans(i).Enabled = False
    Next i
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER, SPITR.PRICE, [ORDER]*[PRICE] AS TOTAL, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & MyRef & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE, SPITR.`ORDER`*SPITR.PRICE AS TOTAL, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & MyRef & "' ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE, SPITR.`ORDER`*SPITR.PRICE AS TOTAL, SPITR.SPAREPARTID, SPAREPARTS.APPLICMODEL " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & MyRef & "' ORDER BY SPAREPARTS.PARTNUMBER;"
    
    
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            .MoveFirst
            pBar.Visible = True
            pBar.Value = 0
            pBar.Max = .RecordCount
            Do While Not .EOF
                'flxParts.AddItem flxParts.Rows & vbTab & .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2) & vbTab & Format$(.Fields(3), "#,##0.00") & vbTab & Format$(.Fields(4), "#,##0.00") & vbTab & .Fields(5)
                flxParts.AddItem flxParts.Rows & vbTab & .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2) & vbTab & Format$(.Fields(3), "#,##0.00") & vbTab & Format$(.Fields(4), "#,##0.00") & vbTab & .Fields(5) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & !ApplicModel
                
                .MoveNext
                pBar.Value = pBar.Value + 1
                DoEvents
            Loop
        End If
        .Close
        pBar.Visible = False
    End With
    For i = 0 To 5
        cmdTrans(i).Enabled = True
    Next i
    
    Call LockData(True)
    Call ComputeTotal
End Sub

Private Sub DisplaySupplier()
    Dim rsSupp As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE, SUPPLIER.ADDRESS " & _
          "From SUPPLIER WHERE (((SUPPLIER.SUPPLIERID)=" & txtSuppID.Text & "));"
    'mySQL
    Qry = "SELECT SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE, SUPPLIER.ADDRESS " & _
          "From SUPPLIER WHERE SUPPLIER.SUPPLIERID=" & txtSuppID.Text & ";"
    
    With rsSupp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            SuppSrch = True
            cboSupplier.Text = !SUPPLIERNAME
            txtSaddress.Text = !Address
            txtScode.Text = !SUPPLIERCODE
        End If
        .Close
    End With
    
End Sub

Private Sub DisplayFlxEdit()
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, MyNM As String, myRem As String, MyStat As String
    Dim i As Integer
    
    'old Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER, SPITR.ALLOCATION, SPITR.BACKORDER, SPITR.PRICE, [ORDER]*[PRICE] AS TOTAL, SPITR.SPAREPARTID " & _
           "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & txtPONumber.Text & "'));"
    For i = 0 To 5
        cmdTrans(i).Enabled = False
    Next i
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER, SPITR.PRICE, [ORDER]*[PRICE] AS TOTAL, SPITR.SPAREPARTID, SPITR.ORDERBAL, SPITR.REMARKS, SPITR.PACKINGLIST,SPITR.SOURCE, SPITR.UPDATEDBY " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & txtPONumber.Text & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE, SPITR.`ORDER`*SPITR.PRICE AS TOTAL, SPITR.SPAREPARTID, SPITR.ORDERBAL, SPITR.REMARKS, SPITR.PACKINGLIST,SPITR.SOURCE, SPITR.UPDATEDBY " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & txtPONumber.Text & "' ORDER BY SPAREPARTS.PARTNUMBER;"
    
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            .MoveFirst
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount
            Do While Not .EOF
                If Not IsNull(!Remarks) Then
                    If Len(!PACKINGLIST) > 0 Then
                        MyNM = IIf(IsNull(!Source), "", Left$(!Source, 1))
                        MyStat = IIf(IsNull(!Remarks), "", !Remarks)
                        myRem = IIf(IsNull(!Source), "", Mid$(!Source, 3, 50))
                    Else
                        MyNM = ""
                        MyStat = ""
                        myRem = ""
                    End If
                Else
                    MyNM = ""
                    MyStat = ""
                    myRem = ""
                End If
                flxParts.AddItem flxParts.Rows & vbTab & .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2) & vbTab & Format$(!Price, "#,##0.00") & vbTab & Format$(!Total, "#,##0.00") & vbTab & MyNM & vbTab & MyStat & vbTab & myRem & vbTab & !SPAREPARTID
                pBar.Value = pBar.Value + 1
                .MoveNext
                DoEvents
            Loop
        End If
        .Close
        pBar.Visible = False
    End With
    For i = 0 To 5
        cmdTrans(i).Enabled = True
    Next i
    
End Sub



Private Sub Timer1_Timer()
    Call LoadSpareParts
    Timer1.Enabled = False
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 8 Then
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 6 Then
            tmrDropRev.Enabled = False
            'cmdTrans(0).SetFocus
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 6 'Index of Button Save
            cmdTrans(6).Top = -375
            cmdTrans(6).Left = 0
            cmdTrans(6).Visible = True
            cmdTrans(7).Top = -375
            cmdTrans(7).Left = cmdTrans(4).Width
            cmdTrans(7).Visible = True
            tmrDrop.Enabled = True
        End If
    End If
End Sub

Private Sub tmrFloatRev_Timer()
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 5 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 5
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub txtAprv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDealerCode.SetFocus
End Sub

Private Sub txtDealerCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDealerName.SetFocus
End Sub

Private Sub txtDealerCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If txtDesc.Tag <> "" Then
            flxParts.TextMatrix(ActRow, 2) = txtDesc.Tag
            txtDesc.Text = ""
        End If
        txtDesc.Tag = "skiplostfocus"
        txtDesc.Visible = False
        flxParts.SetFocus
        flxParts.Col = 2
    End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call txtDesc_LostFocus
        If txtDesc.Tag = "NxtRow" Then
            txtDesc.Tag = ""
            Exit Sub
        End If
        flxSrch.Visible = False
        flxParts.SetFocus
        flxParts.Col = 3
    End If
End Sub

Private Sub txtDesc_LostFocus()
    Dim xRow As Integer
    
    If txtDesc.Tag = "skiplostfocus" Then
        txtDesc.Tag = ""
        Exit Sub
    End If
    If flxParts.TextMatrix(ActRow, 7) <> "" Then
        If Left$(flxParts.TextMatrix(ActRow, 7), 4) = "NEW:" Then     'New SP already stored in FlxNewSP
            xRow = Val(Right$(flxParts.TextMatrix(ActRow, 7), Len(flxParts.TextMatrix(ActRow, 7)) - 4))
            flxNewSP.TextMatrix(xRow, 1) = txtDesc.Text
        Else
            If flxParts.TextMatrix(ActRow, 2) <> txtDesc.Text And txtDesc.Text <> "" Then
                flxParts.TextMatrix(ActRow, 7) = "Edesc"
            End If
        End If
    Else
        If flxParts.TextMatrix(ActRow, 2) <> txtDesc.Text And txtDesc.Text <> "" Then
            flxParts.TextMatrix(ActRow, 7) = "Edesc"
        End If
    End If
    flxParts.TextMatrix(ActRow, 2) = txtDesc.Text
    txtDesc.Visible = False
    
    If flxParts.TextMatrix(ActRow, 3) <> "" And flxParts.TextMatrix(ActRow, 1) <> "" Then
        txtDesc.Tag = "NxtRow"
        If ActRow = flxParts.Rows - 1 Then
            flxParts.AddItem flxParts.Rows
            flxParts.SetFocus
            flxParts.Col = 1
            flxParts.Row = flxParts.Rows - 1
        End If
    End If
End Sub

Private Sub txtFlx_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Call txtFlx_LostFocus
        If txtFlx.Tag = "NxtRow" Then
            txtFlx.Tag = ""
            Exit Sub
        End If
        flxParts.SetFocus
        flxParts.Col = ActCol + 1
    End If
End Sub

Private Sub txtFlx_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtFlx.Tag = "skiplostfocus"
        txtFlx.Visible = False
        flxParts.SetFocus
        flxParts.Col = ActCol
    End If
End Sub

Private Sub txtFlx_LostFocus()
    Dim Price As Single
    
    flxParts.TextMatrix(ActRow, 3) = txtFlx.Text
    txtFlx.Visible = False
    If Not Confirm And flxParts.TextMatrix(ActRow, 4) <> "" Then
        Price = PureNum(flxParts.TextMatrix(ActRow, 4))
        flxParts.TextMatrix(ActRow, 5) = Format$(Val(txtFlx.Text) * Price, "#,##0.00")
        Call ComputeTotal
    End If
    If flxParts.TextMatrix(ActRow, 1) <> "" And flxParts.TextMatrix(ActRow, 2) <> "" And Val(flxParts.TextMatrix(ActRow, 4)) > 0 Then
        'Call AddNewPartNumber
        txtFlx.Tag = "NxtRow"
        If ActRow = flxParts.Rows - 1 Then
            flxParts.AddItem flxParts.Rows
            flxParts.SetFocus
            flxParts.Col = 1
            flxParts.Row = flxParts.Rows - 1
        End If
    End If
End Sub
                                             
Private Sub txtNM_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call txtNM_LostFocus
        flxParts.SetFocus
        flxParts.Col = ActCol + 1
    End If
End Sub
       
Private Sub txtNM_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtNM.Tag = "skiplostfocus"
        txtNM.Visible = False
        flxParts.SetFocus
        flxParts.Col = ActCol
    End If
End Sub

Private Sub txtNM_LostFocus()
    
    If txtNM.Text = "" Then
        txtNM.Visible = False
        flxParts.TextMatrix(ActRow, 6) = ""
        flxParts.TextMatrix(ActRow, 7) = ""
        flxParts.TextMatrix(ActRow, 8) = ""
        Exit Sub
    End If
    If InStr("NSLHBPKUTE", txtNM.Text) = 0 Then
        MsgBox "Invalid Notice Mark!", vbExclamation + vbOKOnly, "Notice Mark (NM)"
        txtNM.Visible = False
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 6) = txtNM.Text
    txtNM.Visible = False
        
    If InStr("NSLHB", txtNM.Text) > 0 Then
        flxParts.TextMatrix(ActRow, 7) = "CANCEL"
        flxParts.TextMatrix(ActRow, 3) = "0"
        flxParts.TextMatrix(ActRow, 4) = "0.00"
        flxParts.TextMatrix(ActRow, 5) = "0.00"
        Call ComputeTotal
        If txtNM.Text = "N" Then flxParts.TextMatrix(ActRow, 8) = "No parts number Spare Parts Master"
        If txtNM.Text = "S" Then flxParts.TextMatrix(ActRow, 8) = "Not for sale"
        If txtNM.Text = "L" Then flxParts.TextMatrix(ActRow, 8) = "Parts subject to consultation"
        If txtNM.Text = "H" Then flxParts.TextMatrix(ActRow, 8) = "ABNORMAL DEMAND"
        If txtNM.Text = "B" Then flxParts.TextMatrix(ActRow, 8) = "New purchasing is prohibited"
    ElseIf InStr("PKUT", txtNM.Text) > 0 Then
        flxParts.TextMatrix(ActRow, 7) = "PENDING"
        If txtNM.Text = "P" Then flxParts.TextMatrix(ActRow, 8) = "Parts under ""Stock Taking"""
        If txtNM.Text = "K" Then flxParts.TextMatrix(ActRow, 8) = "Waiting for Pricing"
        If txtNM.Text = "U" Then flxParts.TextMatrix(ActRow, 8) = "Order was hold due to payment problem"
        If txtNM.Text = "T" Then flxParts.TextMatrix(ActRow, 8) = "Superceding link error"
    ElseIf txtNM.Text = "E" Then
        flxParts.TextMatrix(ActRow, 7) = "Just Notice"
        flxParts.TextMatrix(ActRow, 8) = "Change parts number"
    End If
End Sub

Private Sub txtPODate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPONumber.SetFocus
End Sub

Private Sub txtPODate_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub
Private Sub txtPODate_LostFocus()
    'If Not IsDate(txtPODate) Then
    '    MsgBox ("Invalid date"), vbCritical
    '   txtPODate.SetFocus
    '   txtPODate.SelStart = 0
    '   txtPODate.SelLength = Len(txtPODate.Text)
    'End If
End Sub
Private Sub txtPONumber_GotFocus()
    If txtPONumber.Text <> "" Then
        txtPONumber.SelStart = 0
        txtPONumber.SelLength = Len(txtPONumber.Text)
    End If
End Sub

Private Sub txtPONumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboType.SetFocus
End Sub

Private Sub txtPONumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub

Private Sub txtPONumber_LostFocus()
    Dim Qry As String
    
    Qry = "SELECT SPORDER.REFERENCE From SPORDER WHERE (((SPORDER.REFERENCE)='" & txtPONumber.Text & "'));"
    If AddNew Then
        If CtrlNumExist(Qry) Then
            MsgBox "P.O. Number already exist!", vbExclamation + vbOKOnly, "Validation Error"
            Exit Sub
        End If
    End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Call txtPrice_LostFocus
        If ActRow = flxParts.Rows - 1 Then
            flxParts.AddItem flxParts.Rows
            flxParts.SetFocus
            flxParts.Col = 1
            flxParts.Row = flxParts.Rows - 1
        ElseIf ActRow < flxParts.Rows - 1 Then
            flxParts.SetFocus
            flxParts.Col = 1
            flxParts.Row = ActRow + 1
        End If
    End If
End Sub

Private Sub txtPrice_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtPrice.Tag = "skiplostfocus"
        txtPrice.Visible = False
        flxParts.SetFocus
        flxParts.Col = 4
    End If
End Sub

Private Sub txtPrice_LostFocus()
    If txtPrice.Tag = "skiplostfocus" Then
        txtPrice.Tag = ""
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, ActCol) = Format$(txtPrice.Text, "#,###0.00")
    If flxParts.TextMatrix(ActRow, 4) <> "" Then
        flxParts.TextMatrix(ActRow, 5) = Format$(Val(flxParts.TextMatrix(ActRow, 3)) * PureNum(txtPrice.Text), "#,##0.00")
        Call ComputeTotal
    End If
    
    txtPrice.Visible = False
End Sub


Private Sub txtPartNo_Change()
    Dim i As Integer
    
    On Error GoTo LocalErr
    With flxSrch
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
        If Not ImportPO Then .Visible = True
        .ZOrder
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsParts.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsParts.EOF Then
            If Not PrtBln Then
                txtPartNo.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            If Not FromFlx Then     'top row only if user type @ txtpartno
                .TopRow = rsParts.AbsolutePosition - 1
                FromFlx = False
            End If
            .Row = rsParts.AbsolutePosition - 1
            .Col = 0: .ColSel = 3
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            flxParts.TextMatrix(ActRow, 2) = rsParts!Description
            flxParts.TextMatrix(ActRow, 8) = IIf(IsNull(rsParts!SUPERSEDING), "", rsParts!SUPERSEDING)
            flxParts.TextMatrix(ActRow, 9) = IIf(IsNull(rsParts!RETAILPRICE), 0, rsParts!RETAILPRICE)
            flxParts.TextMatrix(ActRow, 4) = IIf(IsNull(rsParts!Cost), 0, rsParts!Cost)
            flxParts.TextMatrix(ActRow, 5) = IIf(flxParts.TextMatrix(ActRow, 3) <> "", Val(flxParts.TextMatrix(ActRow, 3)) * Val(flxParts.TextMatrix(ActRow, 4)), "")
            flxParts.TextMatrix(ActRow, 4) = Format$(flxParts.TextMatrix(ActRow, 4), "#,##0.00")
            flxParts.TextMatrix(ActRow, 5) = Format$(flxParts.TextMatrix(ActRow, 5), "#,##0.00")
            flxParts.TextMatrix(ActRow, 12) = flxParts.TextMatrix(ActRow, 6)    'IIf(IsNull(rsParts!ApplicModel), "", rsParts!ApplicModel)
            'Add ApplModel
            flxParts.TextMatrix(ActRow, 13) = IIf(IsNull(rsParts!ApplicModel), "", rsParts!ApplicModel)
            
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            If flxParts.TextMatrix(ActRow, 1) <> txtPartNo.Text Then
                flxParts.TextMatrix(ActRow, 2) = ""
                flxParts.TextMatrix(ActRow, 4) = ""
                flxParts.TextMatrix(ActRow, 5) = ""
                flxParts.TextMatrix(ActRow, 6) = ""
                flxParts.TextMatrix(ActRow, 8) = ""
                flxParts.TextMatrix(ActRow, 9) = ""
            End If
            .Visible = False
        End If
    End With
    'cboFind.Text = txtPartNo.Text
    Exit Sub
LocalErr:
    MsgBox "[" & CStr(Err.Number) & "] " & Err.Description, vbCritical + vbOKOnly
    
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            txtPartNo.Tag = "skiplostfocus"
            flxSrch.Visible = True
            flxSrch.SetFocus
        Case vbKeyEscape
            txtPartNo.Text = Empty
            Tmp = Split(flxSrch.Tag, "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxParts.TextMatrix(ActRow, i + 1) = Tmp(i)
            Next i
            flxSrch.Tag = ""
            txtPartNo.Tag = "skiplostfocus"
            flxParts.SetFocus
            flxParts.Col = 1
            txtPartNo.Visible = False
            flxSrch.Visible = False
        Case vbKeyDelete
            PrtSrch = True
            'cboFind.Text = ""
            'cboFind.Text = txtPartNo.Text
        Case vbKeyBack
            PrtSrch = True
            'cboFind.Text = ""
            'cboFind.Text = txtPartNo.Text
        Case Else
            FromFlx = False
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        If txtPartNo.Tag = "PN exist" Then Exit Sub
        If txtPartNo.Tag = "super" Then
            txtPartNo.SetFocus
            txtPartNo.Tag = ""
            Exit Sub
        End If
        If txtPartNo.Tag = "NxtRow" Then
            txtPartNo.Tag = ""
            Exit Sub
        End If
        flxSrch.Visible = False
        flxParts.SetFocus
        If flxParts.TextMatrix(ActRow, 2) = "" Then
            flxParts.Col = 2
        Else
            flxParts.Col = 3
        End If
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    Dim Ans As Integer, i As Integer, xRow As Integer, j As Integer
    Dim Tmp() As String
    
    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    'Validate existing PartNo
    For i = 1 To flxParts.Rows - 1
        If i <> ActRow Then
            If flxParts.TextMatrix(i, 1) = txtPartNo.Text Then
                flxParts.TextMatrix(ActRow, 1) = txtPartNo.Text
                txtPartNo.Visible = False
                flxSrch.Visible = False
                Ans = MsgBox("Part Number " & txtPartNo.Text & " already exist in row " & i & vbCrLf & _
                       "Do you want to locate existing part number?", vbQuestion + vbYesNo, "Duplicate Part Number")
                If Ans = vbYes Then
                    If Not ImportPO Then
                        If flxSrch.Tag <> "" Then
                            Tmp = Split(flxSrch.Tag, "|")
                            For j = LBound(Tmp) To UBound(Tmp)
                                flxParts.TextMatrix(ActRow, j + 1) = Tmp(j)
                            Next j
                        Else
                            For j = 1 To flxParts.Cols - 1
                                flxParts.TextMatrix(ActRow, j) = ""
                            Next j
                        End If
                        flxParts.Row = i
                        flxParts.Col = flxParts.Cols - 1
                        flxParts.RowSel = i
                        flxParts.ColSel = 0
                        flxParts.TopRow = i
                        txtPartNo.Tag = "PN exist"
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next i
    'If New SP load AddNewSP form
    If flxParts.TextMatrix(ActRow, 2) = "" And txtPartNo.Text <> "" Then
        Ans = MsgBox("Unrecognized Part Number!" & vbCrLf & "Do you want to add this Part No.?", vbQuestion + vbYesNo, "New Spare Part")
        flxParts.TextMatrix(ActRow, 1) = txtPartNo.Text
        txtPartNo.Visible = False
        flxSrch.Visible = False
        If Ans = vbYes Then
        
            'Add Supplier Code
            'If not Yamaha,Honda,Kawasaki,Suzuki
            If txtScode.Text <> "BZ" And txtScode.Text <> "AA" And txtScode.Text <> "SP" And txtScode.Text <> "CA" Then
                flxParts.TextMatrix(ActRow, 1) = flxParts.TextMatrix(ActRow, 1) + txtScode.Text
            End If
            
            SPAddnew.sRow = ActRow
            SPAddnew.MySource = ""
            SPAddnew.Show vbModal
            Call ComputeTotal
        Else
            flxParts.TextMatrix(ActRow, 1) = ""
        End If
        Exit Sub
    ElseIf flxParts.TextMatrix(ActRow, 2) <> "" And txtPartNo.Text <> "" Then 'NAAY New SP encoded sa row dapat e-remove iyang record sa FlxNewSP
        If flxParts.TextMatrix(ActRow, 7) <> "" Then
            If Left$(flxParts.TextMatrix(ActRow, 7), 4) = "NEW:" Then     'New SP already stored in FlxNewSP
                xRow = Val(Right$(flxParts.TextMatrix(ActRow, 7), Len(flxParts.TextMatrix(ActRow, 7)) - 4))
                If flxNewSP.TextMatrix(xRow, 0) <> flxParts.TextMatrix(ActRow, 1) Then
                    For i = 0 To 9
                        flxNewSP.TextMatrix(xRow, i) = ""
                    Next i
                End If
            End If
        End If
        flxParts.TextMatrix(ActRow, 7) = ""
        If Not AddNew And Val(flxParts.TextMatrix(ActRow, 6)) > 0 Then
            flxDel.AddItem flxParts.TextMatrix(ActRow, 6)
            flxParts.TextMatrix(ActRow, 6) = ""
        End If
    End If
    If flxParts.TextMatrix(ActRow, 8) <> "" Then
        Ans = MsgBox("Part Number is superseded by " & flxParts.TextMatrix(ActRow, 8) & vbCrLf & "Do you want to use the Superseding Part Number?", vbQuestion + vbYesNo, "Superseded Part Number")
        If Ans = vbYes Then
            txtPartNo.Text = flxParts.TextMatrix(ActRow, 8)
            txtPartNo.Tag = "super"
            Exit Sub
        End If
    End If
    flxParts.TextMatrix(ActRow, 1) = txtPartNo.Text
    txtPartNo.Visible = False
    flxSrch.Visible = False
    If flxParts.TextMatrix(ActRow, 3) <> "" And flxParts.TextMatrix(ActRow, 4) <> "" Then
        Call ComputeTotal
        txtPartNo.Tag = "NxtRow"
        If ActRow = flxParts.Rows - 1 Then
            flxParts.AddItem flxParts.Rows
            flxParts.SetFocus
            flxParts.Col = 1
            flxParts.Row = flxParts.Rows - 1
        End If
    End If
End Sub

Function PartNoExist(MyPnum As String) As Integer
    Dim i As Integer
    
    For i = 1 To flxParts.Rows - 1
        If i <> ActRow Then
            If MyPnum = flxParts.TextMatrix(i, 1) Then
                PartNoExist = i
                Exit Function
            End If
        End If
    Next i
    PartNoExist = 0
End Function


Function PartsInSPcon(MyPartNo As String) As Boolean
    Dim rsFindSP As New ADODB.Recordset
    Dim Qry As String
    '***Access
    'Qry = "SELECT SOURCESP.PARTNUMBER From SOURCESP WHERE (((SOURCESP.PARTNUMBER)='" & MyPartNo & "'));"
    'mySQL
    Qry = "SELECT SOURCESP.PARTNUMBER From SOURCESP WHERE SOURCESP.PARTNUMBER='" & MyPartNo & "';"
    
    With rsFindSP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            PartsInSPcon = IIf(IsNull(!PartNumber), False, True)
        Else
            PartsInSPcon = False
        End If
        .Close
    End With
End Function

Function GetLocation(MyPartNo As String) As String
    Dim rsLoc As New ADODB.Recordset
    Dim Qry As String
    '***Access
    'Qry = "SELECT SPAREPARTS.LOCATION, SPAREPARTS.PARTNUMBER " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & MyPartNo & "'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.LOCATION, SPAREPARTS.PARTNUMBER " & _
          "From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & MyPartNo & "';"
    
    With rsLoc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetLocation = IIf(IsNull(!Location), "", !Location)
        Else
            GetLocation = ""
        End If
        .Close
    End With
End Function

Private Sub txtRemarks_KeyUp(KeyCode As Integer, Shift As Integer)
    SPInsXst.MySource = "SPORDERING"
    If KeyCode = vbKeyF12 Then SPInsXst.Show vbModal
End Sub

Private Sub txtSaddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTerms.SetFocus
End Sub

Private Sub txtSaddress_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub

Private Sub txtSCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSaddress.SetFocus
End Sub

Private Sub txtSearch_Change()
    cmdNext.Tag = 1
End Sub

Private Sub txtSearch_Click()
    cmdNext.Tag = 1
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdNext_Click
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub

Private Sub txtSubmit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAprv.SetFocus
End Sub

Private Sub txtSubmit_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub

Private Sub txtTerms_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPODate.SetFocus
End Sub

Private Sub txtTerms_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        SPInsXst.MySource = "SPORDERING"
        SPInsXst.Show vbModal
    End If
End Sub

Public Sub FinalSPorder()
    Dim i As Integer, BckOrd As Integer, oBal As Integer
    
    If Val(txtPrnCopy.Text) > 0 Then Exit Sub
    'UPDATE SPITR TransType to ORDER
    '***Access
    'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE='ORDER' WHERE (((SPITR.REFERENCE)='" & txtPONumber.Text & "') AND ((SPITR.TRANSTYPE)='ORDER TEMP'));"
    'mySQL
    LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE='ORDER' WHERE SPITR.REFERENCE='" & txtPONumber.Text & "' AND SPITR.TRANSTYPE='ORDER TEMP';"
    
    'Update BackOrder of Ordered SpareParts
    With flxParts
        For i = 1 To .Rows - 1
            BckOrd = GetBackOrder(.TextMatrix(i, 6))
            '***Access
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & BckOrd & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 6) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & BckOrd & _
                           " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 6) & ";"
            
        Next i
    End With
    If txtConfirmNo.Text = "" Then cmdTrans(2).Enabled = True
    'MsgBox "Spare Parts Qty-On-Hand successfully updated!"
End Sub

Function GetBackOrder(MyID As Integer) As Integer
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String
    Dim xRow As Integer, xBal As Integer, i As Integer
    
    'Clear FlxStock
    '***Access
    'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & MyId & ")) ORDER BY SPITR.TRANSDATE;"
    'mySQL
    Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.`ORDER` From SPITR " & _
          "WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.SPAREPARTID=" & MyID & " ORDER BY SPITR.TRANSDATE;"
    
    flxOrd.Rows = 1
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxOrd.AddItem Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !REFERENCE & vbTab & !Order
                .MoveNext
            Loop
        End If
        .Close
    End With
    xBal = 0
    With flxOrd
        For i = 1 To .Rows - 1
            'Display PO in Stock Card and obtain Balance
            xBal = xBal + Val(.TextMatrix(i, 2))
            'Acquire Packing list for P.O.
            MyPO = .TextMatrix(i, 1)
            'Access
            'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyId & ") AND ((SPITR.PACKINGLIST)='" & MyPO & "')) ORDER BY SPITR.TRANSDATE;"
            'mySQL
            Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & MyID & " AND SPITR.PACKINGLIST='" & MyPO & "' ORDER BY SPITR.TRANSDATE;"
        
            rsRec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsRec.RecordCount > 0 Then
                rsRec.MoveFirst
                Do While Not rsRec.EOF
                    xBal = xBal - rsRec!RECEIVE
                    rsRec.MoveNext
                Loop
            End If
            rsRec.Close
        Next i
    End With
    GetBackOrder = xBal
End Function

