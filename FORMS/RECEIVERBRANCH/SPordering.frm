VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SPordering 
   Caption         =   "Spare Parts Order Form"
   ClientHeight    =   8775
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   11400
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxNewSP 
      Height          =   1215
      Left            =   360
      TabIndex        =   68
      Tag             =   "xyf"
      Top             =   8760
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   0
      Cols            =   10
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
      Left            =   2280
      TabIndex        =   42
      Top             =   840
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPordering.frx":0000
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
         Caption         =   "Find &Next"
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
         ItemData        =   "SPordering.frx":001C
         Left            =   1080
         List            =   "SPordering.frx":001E
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
         Caption         =   "&Find What:"
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
      Image           =   "SPordering.frx":0020
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
      Left            =   1440
      TabIndex        =   39
      Top             =   8880
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
      Cols            =   10
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
            Picture         =   "SPordering.frx":05BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":1494
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":236E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":3248
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":4122
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPordering.frx":4FFC
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
      Image           =   "SPordering.frx":5ED6
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
      Image           =   "SPordering.frx":6470
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
      Image           =   "SPordering.frx":6A0A
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
      Image           =   "SPordering.frx":6FA4
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
      Image           =   "SPordering.frx":753E
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
      Image           =   "SPordering.frx":7AD8
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
      Image           =   "SPordering.frx":8072
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
      Image           =   "SPordering.frx":860C
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
      Image           =   "SPordering.frx":8BA6
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
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11520
      m_MDI_Height    =   9180
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
         Caption         =   "Add Row        Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "SPordering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SPcon As New ADODB.Connection
Dim rsParts As New ADODB.Recordset
Dim rsSupplier As New ADODB.Recordset

Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean, ImportPO As Boolean
Dim ActRow As Integer, ActCol As Integer, BtnIndx As Integer
Dim AddNew As Boolean, LockFlx As Boolean, Confirm As Boolean, SuppSrch As Boolean
Dim Qry As String

Private Sub ClearEntries()
    cboSupplier.Text = ""
    txtScode.Text = ""
    txtSaddress.Text = ""
    txtSuppID.Text = ""
    txtDealerCode.Text = ""
    txtDealerName.Text = ""
    txtAddress.Text = ""
    txtPODate.Text = ""
    txtPONumber.Text = ""
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
            If Val(.TextMatrix(i, 5)) = 0 And Val(.TextMatrix(i, 3)) > 0 And Val(.TextMatrix(i, 4)) > 0 Then
                TtlCost = Val(.TextMatrix(i, 3)) * PureNum(.TextMatrix(i, 4))
                .TextMatrix(i, 5) = Format$(TtlCost, "###,##0.00")
            End If
            Ttl = Ttl + PureNum(.TextMatrix(i, 5))
        Next i
    End With
    txtTotal.Text = Format$(Ttl, "#,##0.00")
    
End Sub

Private Sub cboFind_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
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
    Dim x As Integer
    
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
            txtScode.Text = IIf(IsNull(rsSupplier!SupplierCode), "", rsSupplier!SupplierCode)
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

Private Sub cboSupplier_LostFocus()
    Dim rsBranch As New ADODB.Recordset
    
    If InStr(cboSupplier.Text, "HONDA PHIL") > 0 Then
        With rsBranch
            Qry = "SELECT BRANCH.* From BRANCH WHERE (((BRANCH.BRANCHID)=1));"
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
    If cboSupplier.Text <> "" Then
        txtPONumber.Text = GetOrderNo
    End If
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Call cboType_LostFocus
        flxParts.SetFocus
        flxParts.Col = 1
        flxParts.Row = flxParts.Rows - 1
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
    Dim x As Integer, w As Boolean
    
    If txtSearch.Text = "" Then Exit Sub
    w = False
    For x = 0 To txtSearch.ListCount - 1
        If txtSearch.Text = txtSearch.List(x) Then
            If x = 0 Then
                w = True
                Exit For
            Else
                txtSearch.RemoveItem x
                Exit For
            End If
        End If
    Next x
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


Private Sub cmdTrans_Click(Index As Integer)

    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 21) As String
    Dim MyConf As String
    
    Select Case Index
        Case 0  'New
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
            cboSupplier.SetFocus
        Case 1  'Edit
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
            Confirm = True
        Case 3  'View
            SPorderSum.Show
            Unload Me
        Case 4  'Print
            If InStr(cboSupplier.Text, "HONDA PHILS.") > 0 Then
                Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER " & _
                      "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = txtPONumber.Text
                PRNmc.ParamValue2 = cboType.Text
                PRNmc.ParamValue3 = txtPODate.Text
                PRNmc.MySource = "SPPO"
                PRNmc.Show
            Else
                Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER " & _
                      "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = cboSupplier.Text
                PRNmc.ParamValue2 = txtPONumber.Text
                PRNmc.ParamValue3 = txtPODate.Text
                PRNmc.MySource = "SPPOrepl"
                PRNmc.Show
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
                MsgBox "Spare Parts Order successfully saved!", vbInformation + vbOKOnly, "Save"
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
                MsgBox "Spare Parts Order successfully saved!", vbInformation + vbOKOnly, "Save"
            
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
    
    Qry = "SELECT SPORDER.SPORDERID, SPORDER.REFERENCE, SPORDER.ORDERDATE, Month([ORDERDATE]) AS MYMONTH " & _
          "From SPORDER WHERE (((Month([ORDERDATE]))=" & Month(txtPODate.Text) & ") AND ((SPORDER.SUPPLIERID)=" & txtSuppID.Text & ")) ORDER BY SPORDER.REFERENCE DESC;"
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
        Qry = "SELECT SPORDER.REFERENCE From SPORDER WHERE (((SPORDER.REFERENCE)='" & MyOrdNo & "'));"
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
            With txtPrice
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
        Case 4  'Price
            With txtPrice
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

Private Sub flxParts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LockFlx Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuRows)
    End If
End Sub

Private Sub flxParts_Scroll()
    txtFlx.Visible = False
    txtNM.Visible = False
    txtPrice.Visible = False
    txtPartNo.Visible = False
    txtDesc.Visible = False
End Sub

Private Sub flxSrch_EnterCell()
    With flxSrch
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        FromFlx = True
        txtPartNo.Text = .TextMatrix(.Row, 1)
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

Private Sub Form_Load()
    Dim i As Integer
    
    cmdA.Top = 15000
    Call LockData(True)
    Call LoadSpareParts
    cboType.AddItem "(01) - URGENT"
    cboType.AddItem "(04) - STOCK"
    cboMatch.AddItem "Whole Word"
    cboMatch.AddItem "Any Part of Word"
    
    MotoRep.sBar.Panels(6).Text = "Spare Parts Order Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
    Set SPcon = Nothing
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
    
    'From: 1-PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-ID, 7-Remarks
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
        
        Confirm = True
    End With
End Sub

Private Sub flxResizeNew()
    'From: 1 - PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-NM, 7-Status,8Remarks, 9-ID
    'To  : 1 - PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-ID, 7-Remarks
    
    With flxParts
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
    LYRcon.Execute "UPDATE SPORDER SET SPORDER.CONFIRMNO = '" & txtConfirmNo.Text & "', SPORDER.REMARKS = Null WHERE (((SPORDER.REFERENCE)='" & txtPONumber.Text & "'));"
    
    'flxparts: 1 - PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-NM, 7-Status, 8-Remarks, 9-ID
    With flxParts
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 7) = "CANCEL" Then
                'Update SPITR Cancelled parts
                LYRcon.Execute "UPDATE SPITR SET SPITR.ORDERBAL = '0', SPITR.REMARKS = '" & .TextMatrix(i, 7) & "', SPITR.PACKINGLIST = '" & txtPONumber.Text & "', SPITR.SOURCE = '" & .TextMatrix(i, 6) & "-" & .TextMatrix(i, 8) & "', SPITR.BALANCE = 0, SPITR.PRICE = " & PureNum(.TextMatrix(i, 4)) & _
                           " WHERE (((SPITR.SPAREPARTID)=" & .TextMatrix(i, 9) & ") AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "'));"
            Else
                'Update BackOrder in Spareparts
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.ORDEREF = '" & txtPONumber.Text & "', SPAREPARTS.BACKORDER = " & .TextMatrix(i, 3) & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 9) & "));"
                LYRcon.Execute "UPDATE SPITR SET SPITR.PRICE = " & PureNum(.TextMatrix(i, 4)) & ", SPITR.REMARKS = Null, SPITR.PACKINGLIST = '" & txtPONumber.Text & "', SPITR.ORDERBAL = '" & .TextMatrix(i, 3) & "',SPITR.BALANCE = " & .TextMatrix(i, 3) & _
                               " WHERE (((SPITR.REFERENCE)='" & txtPONumber.Text & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 9) & "));"
            End If
        Next i
    End With
End Sub

Private Sub SaveSPOrderEdit()
    Dim rsOrder As New ADODB.Recordset
    Dim rsSP As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, myRem As String
    Dim i As Integer, myQty As Integer, nSProw As Integer
    
    
    'Update sporder
    Qry = "SELECT SPORDER.* From SPORDER WHERE (((SPORDER.SPORDERID)=" & txtOrderID.Text & "));"
    rsOrder.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsOrder.RecordCount > 0 Then
        rsOrder.MoveFirst
        rsOrder!OrderType = cboType.Text
        rsOrder!SubmittedBy = txtSubmit.Text
        rsOrder!ApprovedBy = txtAprv.Text
        rsOrder!UpdatedBy = UserID
        rsOrder!LastUpdate = Date
        rsOrder!REMARKS = txtRemarks.Text
        rsOrder!Terms = txtTerms.Text
        rsOrder.Update
        rsOrder.Close
    End If
    'Saving edited SP details
    With flxParts
        For i = 1 To .Rows - 1
            'check new spareparts
            myRem = .TextMatrix(i, 7)
            If .TextMatrix(i, 6) = "" Then
              If InStr(myRem, "NEW") > 0 Then 'Save new SparePart
                nSProw = Val(Right$(myRem, Len(myRem) - 4))
                If flxNewSP.TextMatrix(nSProw, 0) <> "" Then
                    rsSP.Open "SpareParts", LYRcon, adOpenKeyset, adLockOptimistic
                    rsSP.AddNew
                    rsSP!PartNumber = flxNewSP.TextMatrix(nSProw, 0)
                    rsSP!Description = flxNewSP.TextMatrix(nSProw, 1)
                    rsSP!QtyOnHand = 0
                    rsSP!Cost = PureNum(flxNewSP.TextMatrix(nSProw, 4))
                    rsSP!StorePrice = Val(flxNewSP.TextMatrix(nSProw, 5))
                    rsSP!BegBal = 0
                    rsSP!REFERENCE = txtPONumber.Text
                    If IsDate(txtPODate.Text) Then rsSP!DateBegBal = txtPODate.Text
                    rsSP!REMARKS = flxNewSP.TextMatrix(nSProw, 3)
                    If Val(flxNewSP.TextMatrix(nSProw, 6)) > 0 Then rsSP!SPFid = Val(flxNewSP.TextMatrix(nSProw, 6))
                    rsSP!MaxDiscount = Val(flxNewSP.TextMatrix(nSProw, 7))
                    rsSP!UpdatedBy = UserID
                    rsSP!LastUpdate = Date
                    rsSP.Update
                    .TextMatrix(i, 6) = rsSP!SparePartID
                    rsSP.Close
                    'Add new SP in SPcon
                    SPcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, COST, RETAILPRICE ) " & _
                              "VALUES ('" & flxNewSP.TextMatrix(nSProw, 0) & "', '" & flxNewSP.TextMatrix(nSProw, 1) & "', " & Val(flxNewSP.TextMatrix(nSProw, 4)) & ", " & Val(flxNewSP.TextMatrix(nSProw, 5)) & ");"
                End If
              Else    'SP exist in SPcon
                'Check existence in Mtreptgm
                Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS " & _
                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                If Not CtrlNumExist(Qry) Then
                    rsSP.Open "SpareParts", LYRcon, adOpenKeyset, adLockOptimistic
                    rsSP.AddNew
                    rsSP!PartNumber = .TextMatrix(i, 1)
                    rsSP!Description = .TextMatrix(i, 2)
                    rsSP!QtyOnHand = 0
                    rsSP!Cost = PureNum(.TextMatrix(i, 4))
                    rsSP!StorePrice = Val(.TextMatrix(i, 9))
                    rsSP!BegBal = 0
                    rsSP!REFERENCE = txtPONumber.Text
                    If IsDate(txtPODate.Text) Then rsSP!DateBegBal = txtPODate.Text
                    rsSP!REMARKS = "GENUINE"
                    rsSP!UpdatedBy = UserID
                    rsSP!LastUpdate = Date
                    rsSP.Update
                    .TextMatrix(i, 6) = rsSP!SparePartID
                    rsSP.Close
                    If InStr(myRem, "Edesc") > 0 Then
                        'Update Description in SPcon
                        SPcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                    End If
                Else
                    .TextMatrix(i, 6) = GetSPid(.TextMatrix(i, 1))
                    If InStr(myRem, "Edesc") > 0 Then
                        'Update Description in LYRcon
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'Update Description in SPcon
                        SPcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                    End If
                End If
              End If
              'Save SPITR
              rsITR.Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
              rsITR.AddNew
              rsITR!REFERENCE = txtPONumber.Text
              rsITR!SparePartID = flxParts.TextMatrix(i, 6)
              rsITR!TransDate = txtPODate.Text
              rsITR!TransType = "ORDER"
              rsITR!Order = Val(flxParts.TextMatrix(i, 3))
              rsITR!Price = PureNum(flxParts.TextMatrix(i, 4))
              rsITR!Source = cboSupplier.Text
              rsITR!UpdatedBy = UserID
              rsITR!LastUpdate = Now
              rsITR.Update
              rsITR.Close
            Else
              'Save edit SparePart Desc
              If InStr(myRem, "Edesc") > 0 Then
                  'Update Description in LYRcon
                  LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                              "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                  'Update Description in SPcon
                  SPcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                              "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
              End If
            End If
            
        Next i
        
    End With
    
    If flxDel.Rows > 0 Then
        For i = 0 To flxDel.Rows - 1
            LYRcon.Execute "DELETE SPITR.* From SPITR  WHERE (((SPITR.SPAREPARTID)=" & flxDel.TextMatrix(i, 0) & ") AND ((SPITR.REFERENCE)='" & txtPONumber.Text & "'));"
        Next i
    End If
    
End Sub

Private Sub SaveNewSPOrder()
    Dim rsOrder As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim rsSP As New ADODB.Recordset
    Dim Qry As String, myRem As String
    Dim i As Integer, myQty As Integer, nSProw As Integer
    
    'Saving SPOrder
    rsOrder.Open "SPOrder", LYRcon, adOpenKeyset, adLockOptimistic
    rsOrder.AddNew
    rsOrder!DealerCode = txtDealerCode.Text
    rsOrder!DealerName = txtDealerName.Text
    rsOrder!Address = txtAddress.Text
    rsOrder!REFERENCE = txtPONumber.Text
    rsOrder!OrderType = cboType.Text
    rsOrder!ORDERDATE = txtPODate.Text
    rsOrder!SubmittedBy = txtSubmit.Text
    rsOrder!ApprovedBy = txtAprv.Text
    'rsOrder!ConfirmNo = txtConfirmNo.Text
    rsOrder!UpdatedBy = UserID
    rsOrder!LastUpdate = Date
    rsOrder!REMARKS = txtRemarks.Text
    rsOrder!Terms = txtTerms.Text
    rsOrder!SUPPLIERID = IIf(txtSuppID.Text <> "", Val(txtSuppID.Text), 0)
    rsOrder.Update
    txtOrderID.Text = rsOrder!SPORDERID
    rsOrder.Close
    
    'Saving SP details
    With flxParts
        For i = 1 To .Rows - 1
            'check new spareparts
            myRem = .TextMatrix(i, 7)
            If InStr(myRem, "NEW") > 0 Then 'Save new SparePart
                nSProw = Val(Right$(myRem, Len(myRem) - 4))
                If flxNewSP.TextMatrix(nSProw, 0) <> "" Then
                    rsSP.Open "SpareParts", LYRcon, adOpenKeyset, adLockOptimistic
                    rsSP.AddNew
                    rsSP!PartNumber = flxNewSP.TextMatrix(nSProw, 0)
                    rsSP!Description = flxNewSP.TextMatrix(nSProw, 1)
                    rsSP!QtyOnHand = 0
                    rsSP!Cost = PureNum(flxNewSP.TextMatrix(nSProw, 4))
                    rsSP!StorePrice = Val(flxNewSP.TextMatrix(nSProw, 5))
                    rsSP!BegBal = 0
                    rsSP!REFERENCE = txtPONumber.Text
                    If IsDate(txtPODate.Text) Then rsSP!DateBegBal = txtPODate.Text
                    rsSP!REMARKS = flxNewSP.TextMatrix(nSProw, 3)
                    If Val(flxNewSP.TextMatrix(nSProw, 6)) > 0 Then rsSP!SPFid = Val(flxNewSP.TextMatrix(nSProw, 6))
                    rsSP!MaxDiscount = Val(flxNewSP.TextMatrix(nSProw, 7))
                    rsSP!UpdatedBy = UserID
                    rsSP!LastUpdate = Date
                    rsSP.Update
                    .TextMatrix(i, 6) = rsSP!SparePartID
                    rsSP.Close
                    'Add new SP in SPcon
                    SPcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, COST, RETAILPRICE ) " & _
                              "VALUES ('" & flxNewSP.TextMatrix(nSProw, 0) & "', '" & flxNewSP.TextMatrix(nSProw, 1) & "', " & Val(flxNewSP.TextMatrix(nSProw, 4)) & ", " & Val(flxNewSP.TextMatrix(nSProw, 5)) & ");"
                End If
            Else    'SP exist in SPcon
                'Check existence in Mtreptgm
                Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS " & _
                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                If Not CtrlNumExist(Qry) Then
                    rsSP.Open "SpareParts", LYRcon, adOpenKeyset, adLockOptimistic
                    rsSP.AddNew
                    rsSP!PartNumber = .TextMatrix(i, 1)
                    rsSP!Description = .TextMatrix(i, 2)
                    rsSP!QtyOnHand = 0
                    rsSP!Cost = PureNum(.TextMatrix(i, 4))
                    rsSP!StorePrice = Val(.TextMatrix(i, 9))
                    rsSP!BegBal = 0
                    rsSP!REFERENCE = txtPONumber.Text
                    If IsDate(txtPODate.Text) Then rsSP!DateBegBal = txtPODate.Text
                    rsSP!REMARKS = "GENUINE"
                    rsSP!UpdatedBy = UserID
                    rsSP!LastUpdate = Date
                    rsSP.Update
                    .TextMatrix(i, 6) = rsSP!SparePartID
                    rsSP.Close
                    If InStr(myRem, "Edesc") > 0 Then
                        'Update Description in SPcon
                        SPcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                    End If
                Else
                    .TextMatrix(i, 6) = GetSPid(.TextMatrix(i, 1))
                    If InStr(myRem, "Edesc") > 0 Then
                        'Update Description in LYRcon
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                        'Update Description in SPcon
                        SPcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & .TextMatrix(i, 2) & "' " & _
                                      "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                    End If
                End If
            End If
            'Save SPITR
            rsITR.Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
            rsITR.AddNew
            rsITR!REFERENCE = txtPONumber.Text
            rsITR!SparePartID = flxParts.TextMatrix(i, 6)
            rsITR!TransDate = txtPODate.Text
            rsITR!TransType = "ORDER"
            rsITR!Order = Val(flxParts.TextMatrix(i, 3))
            rsITR!Price = PureNum(flxParts.TextMatrix(i, 4))
            rsITR!Source = cboSupplier.Text
            rsITR!UpdatedBy = UserID
            rsITR!LastUpdate = Now
            rsITR.Update
            rsITR.Close
        Next i
        
    End With
        
End Sub

Function GetSPid(PartNo As String) As Integer
    Dim rsPNum As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER From SPAREPARTS " & _
          "WHERE (((SPAREPARTS.PARTNUMBER)='" & PartNo & "'));"
    With rsPNum
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetSPid = !SparePartID
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
        
    If SPcon.State = adStateOpen Then Exit Sub
    SPcon.CursorLocation = adUseClient
    SPcon.Open "dsn=SPAREPARTS;uid=;pwd=;database=SPAREPARTS"
    
    Qry = "SELECT SPAREPARTS.SC, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.SUPERSEDING, SPAREPARTS.COST, SPAREPARTS.RETAILPRICE " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, SPcon, adOpenKeyset, adLockOptimistic
    Set flxSrch.DataSource = rsParts
    With flxSrch
        .Width = flxParts.ColWidth(1)
        .ColWidth(0) = 0
        .ColWidth(1) = flxParts.ColWidth(1) - 300
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .Row = 0
        .Col = 3
        .CellAlignment = 4
        '.Col = 4
        '.CellAlignment = 4
    End With
    
    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE, SUPPLIER.ADDRESS " & _
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
    ValidateSPOrder = vMsg
    
End Function

Public Sub DisplaySPOrder(MyRef As String)
    Dim rsOrder As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String
    
    
    Qry = "SELECT SPORDER.* From SPORDER " & _
          "WHERE (((SPORDER.REFERENCE)='" & MyRef & "'));"
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Call ClearEntries
            .MoveFirst
            txtDealerCode.Text = !DealerCode
            txtDealerName.Text = !DealerName
            txtAddress.Text = !Address
            txtPODate.Text = !ORDERDATE
            txtPONumber.Text = !REFERENCE
            cboType.Text = !OrderType
            txtConfirmNo.Text = IIf(IsNull(!ConfirmNo), "", !ConfirmNo)
            txtRemarks.Text = IIf(IsNull(!REMARKS), "", !REMARKS)
            txtTerms.Text = IIf(IsNull(!Terms), "", !Terms)
            txtSubmit.Text = !SubmittedBy
            txtAprv.Text = !ApprovedBy
            txtPrnCopy.Text = !PrintCopy
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
                
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER, SPITR.PRICE, [ORDER]*[PRICE] AS TOTAL, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & MyRef & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            .MoveFirst
            Do While Not .EOF
                flxParts.AddItem flxParts.Rows & vbTab & .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2) & vbTab & Format$(.Fields(3), "#,##0.00") & vbTab & Format$(.Fields(4), "#,##0.00") & vbTab & .Fields(5)
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    Call LockData(True)
    Call ComputeTotal
End Sub

Private Sub DisplaySupplier()
    Dim rsSupp As New ADODB.Recordset
    Dim Qry As String
    
    
    Qry = "SELECT SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE, SUPPLIER.ADDRESS " & _
          "From SUPPLIER WHERE (((SUPPLIER.SUPPLIERID)=" & txtSuppID.Text & "));"
    With rsSupp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupplier.Text = !SUPPLIERNAME
        End If
        .Close
    End With
    
End Sub

Private Sub DisplayFlxEdit()
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, MyNM As String, myRem As String, myStat As String
    
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER, SPITR.ALLOCATION, SPITR.BACKORDER, SPITR.PRICE, [ORDER]*[PRICE] AS TOTAL, SPITR.SPAREPARTID " & _
           "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & txtPONumber.Text & "'));"

    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER, SPITR.PRICE, [ORDER]*[PRICE] AS TOTAL, SPITR.SPAREPARTID, SPITR.ORDERBAL, SPITR.REMARKS, SPITR.PACKINGLIST,SPITR.SOURCE, SPITR.UPDATEDBY " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & txtPONumber.Text & "')) ORDER BY SPAREPARTS.PARTNUMBER;"

    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            .MoveFirst
            Do While Not .EOF
                If Not IsNull(!REMARKS) Then
                    If Len(!PACKINGLIST) > 0 Then
                        MyNM = IIf(IsNull(!Source), "", Left$(!Source, 1))
                        myStat = IIf(IsNull(!REMARKS), "", !REMARKS)
                        myRem = IIf(IsNull(!Source), "", Mid$(!Source, 3, 50))
                    Else
                        MyNM = ""
                        myStat = ""
                        myRem = ""
                    End If
                Else
                    MyNM = ""
                    myStat = ""
                    myRem = ""
                End If
                flxParts.AddItem flxParts.Rows & vbTab & .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2) & vbTab & Format$(!Price, "#,##0.00") & vbTab & Format$(!Total, "#,##0.00") & vbTab & MyNM & vbTab & myStat & vbTab & myRem & vbTab & !SparePartID
                .MoveNext
            Loop
        End If
        .Close
    End With
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
    If flxParts.TextMatrix(ActRow, 1) <> "" And flxParts.TextMatrix(ActRow, 2) <> "" Then
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

Private Sub txtPONumber_GotFocus()
    If txtPONumber.Text <> "" Then
        txtPONumber.SelStart = 0
        txtPONumber.SelLength = Len(txtPONumber.Text)
    End If
End Sub

Private Sub txtPONumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboType.SetFocus
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
            flxParts.TextMatrix(ActRow, 8) = IIf(IsNull(rsParts!Superseding), "", rsParts!Superseding)
            flxParts.TextMatrix(ActRow, 9) = IIf(IsNull(rsParts!RetailPrice), 0, rsParts!RetailPrice)
            flxParts.TextMatrix(ActRow, 4) = IIf(IsNull(rsParts!Cost), 0, rsParts!Cost)
            flxParts.TextMatrix(ActRow, 5) = IIf(flxParts.TextMatrix(ActRow, 3) <> "", Val(flxParts.TextMatrix(ActRow, 3)) * Val(flxParts.TextMatrix(ActRow, 4)), "")
            flxParts.TextMatrix(ActRow, 4) = Format$(flxParts.TextMatrix(ActRow, 4), "#,##0.00")
            flxParts.TextMatrix(ActRow, 5) = Format$(flxParts.TextMatrix(ActRow, 5), "#,##0.00")
             
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
            SPAddnew.sRow = ActRow
            SPAddnew.Show vbModal
        Else
            flxParts.TextMatrix(ActRow, 1) = ""
        End If
        Exit Sub
    ElseIf flxParts.TextMatrix(ActRow, 2) <> "" And txtPartNo.Text <> "" Then
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
    
    Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & MyPartNo & "'));"
    With rsFindSP
        .Open Qry, SPcon, adOpenKeyset, adLockOptimistic
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
    
    Qry = "SELECT SPAREPARTS.LOCATION, SPAREPARTS.PARTNUMBER " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & MyPartNo & "'));"
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

Private Sub txtSaddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTerms.SetFocus
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

Private Sub txtSubmit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAprv.SetFocus
End Sub


Private Sub txtTerms_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPODate.SetFocus
End Sub

