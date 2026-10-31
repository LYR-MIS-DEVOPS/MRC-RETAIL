VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form SPbranchIncoming 
   Caption         =   "Incoming Spare Parts"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11520
   ControlBox      =   0   'False
   Icon            =   "SPbranchIncoming.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   11520
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdPO 
      Height          =   300
      Left            =   10170
      TabIndex        =   51
      Tag             =   "XYF"
      Top             =   2205
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      Caption         =   "P.O."
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
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   7680
   End
   Begin TabDlg.SSTab tabFind 
      Height          =   1095
      Left            =   5160
      TabIndex        =   43
      Top             =   3480
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPbranchIncoming.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdNext"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboMatch"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtSearch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.ComboBox txtSearch 
         Height          =   315
         Left            =   1080
         TabIndex        =   47
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cboMatch 
         Height          =   315
         ItemData        =   "SPbranchIncoming.frx":0028
         Left            =   1080
         List            =   "SPbranchIncoming.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   675
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3480
         TabIndex        =   45
         Top             =   645
         Width           =   975
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Find Next"
         Height          =   375
         Left            =   3480
         TabIndex        =   44
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Match:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   705
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Find What:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.ListBox lstDelParts 
      Height          =   450
      ItemData        =   "SPbranchIncoming.frx":002C
      Left            =   10320
      List            =   "SPbranchIncoming.frx":002E
      TabIndex        =   41
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   7920
      TabIndex        =   40
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   35
      Tag             =   "XYF"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   720
      TabIndex        =   34
      Top             =   9000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxNewSP 
      Height          =   735
      Left            =   720
      TabIndex        =   33
      Top             =   9000
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox txtRRno 
      Height          =   285
      Left            =   2640
      TabIndex        =   32
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboOrderNo 
      Height          =   315
      Left            =   7440
      TabIndex        =   31
      Top             =   8880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9840
      Top             =   9000
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7200
      Top             =   9240
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9240
      Top             =   9000
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6600
      Top             =   9240
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6360
      TabIndex        =   26
      Top             =   8880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   3240
      TabIndex        =   25
      Top             =   8880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSrch 
      Height          =   1695
      Left            =   3360
      TabIndex        =   24
      Top             =   9240
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2990
      _Version        =   393216
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   375
      Left            =   9600
      TabIndex        =   23
      Top             =   9240
      Width           =   135
   End
   Begin VB.TextBox txtRem 
      Height          =   285
      Left            =   5040
      TabIndex        =   22
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxParts 
      Height          =   3840
      Left            =   450
      TabIndex        =   10
      Tag             =   "XY"
      Top             =   2535
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   6773
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "    |<PART NUMBER   |<DESCRIPTION |<ORDER NO. |^QTY.   |<LOCATION|<REMARKS"
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
      Top             =   8400
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
            Picture         =   "SPbranchIncoming.frx":0030
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPbranchIncoming.frx":0F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPbranchIncoming.frx":1DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPbranchIncoming.frx":2CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPbranchIncoming.frx":3B98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPbranchIncoming.frx":4A72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&New Packing List"
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
      Image           =   "SPbranchIncoming.frx":594C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   28
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "SPbranchIncoming.frx":5EE6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   29
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Print Preview"
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
      Image           =   "SPbranchIncoming.frx":6480
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   7200
      TabIndex        =   30
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Packing List"
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
      Image           =   "SPbranchIncoming.frx":6A1A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   12600
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "SPbranchIncoming.frx":6FB4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   14400
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "SPbranchIncoming.frx":754E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdLoadPrtl 
      Height          =   375
      Left            =   10560
      TabIndex        =   37
      Tag             =   "XYF"
      ToolTipText     =   "Load Partial Saved"
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
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
      Image           =   "SPbranchIncoming.frx":7AE8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrtlSave 
      Height          =   375
      Left            =   9960
      TabIndex        =   38
      Tag             =   "XYF"
      ToolTipText     =   "Partial Saving"
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
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
      Image           =   "SPbranchIncoming.frx":8082
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   39
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Edit Packing List"
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
      Image           =   "SPbranchIncoming.frx":861C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   300
      Left            =   10665
      TabIndex        =   42
      Tag             =   "XYF"
      ToolTipText     =   "Find Part Number"
      Top             =   2205
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
      Image           =   "SPbranchIncoming.frx":8BB6
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   1095
      Left            =   240
      TabIndex        =   50
      Top             =   9720
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
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   450
      TabIndex        =   53
      Tag             =   "xyf"
      Top             =   6720
      Width           =   10695
      Begin VB.TextBox txtChkBy 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   600
         Width           =   9015
      End
      Begin VB.TextBox txtRemarks 
         Height          =   285
         Left            =   1440
         MaxLength       =   250
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   240
         Width           =   9015
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   55
         Tag             =   "XYF"
         Top             =   255
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "CHECKED BY:"
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
         TabIndex        =   54
         Tag             =   "XYF"
         Top             =   615
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   450
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   720
      Width           =   10695
      Begin VB.TextBox txtDateRec 
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
         Left            =   8640
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox cboShipper 
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
         Left            =   1530
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   585
         Width           =   4215
      End
      Begin VB.ComboBox cboBranch 
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
         Left            =   1530
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   165
         Width           =   4215
      End
      Begin VB.TextBox txtPackingNo 
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
         Left            =   8640
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   170
         Width           =   1695
      End
      Begin VB.TextBox txtWayBill 
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
         Left            =   1530
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   1425
         Width           =   1575
      End
      Begin VB.TextBox txtPLDate 
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
         Left            =   8640
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   538
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE RECEIVED:"
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
         Left            =   7200
         TabIndex        =   56
         Tag             =   "XYF"
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "RECEIVED FROM:"
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
         Left            =   225
         TabIndex        =   52
         Tag             =   "XYF"
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "PACKING LIST PRINT DATE"
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
         Left            =   6120
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   570
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "PACKING NO."
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
         Left            =   7440
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   230
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "WAYBILL NO."
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
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "DELIVERED BY:"
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
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   630
         Width           =   1215
      End
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Qty."
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
      Left            =   5160
      TabIndex        =   36
      Tag             =   "XYF"
      Top             =   6375
      Width           =   1215
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11760
      m_MDI_Height    =   8505
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIVED SPARE PARTS"
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
      Left            =   450
      TabIndex        =   20
      Tag             =   "XYF"
      Top             =   2205
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPARE PARTS PACKING LIST"
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
      Left            =   8115
      TabIndex        =   18
      Tag             =   "TRF"
      Top             =   450
      Width           =   2820
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MRC STO.TOMAS"
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
      TabIndex        =   16
      Tag             =   "TLF"
      Top             =   450
      Width           =   3200
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MRC STO.TOMAS"
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
      TabIndex        =   17
      Tag             =   "TLF"
      Top             =   475
      Width           =   3200
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPARE PARTS PACKING LIST"
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
      Left            =   8145
      TabIndex        =   19
      Tag             =   "TRF"
      Top             =   480
      Width           =   2820
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIVED SPARE PARTS"
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
      Left            =   510
      TabIndex        =   21
      Tag             =   "XYF"
      Top             =   2235
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
Attribute VB_Name = "SPbranchIncoming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset

Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean, OrdSrch As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim BtnIndx As Integer, FrstBtnIndx As Integer
Dim ViewOnly As Boolean, AddNew As Boolean, Confirm As Boolean, ComboSrch As Boolean
Private Sub BtnSaveCancel(bln As Boolean)
    Dim i As Integer
    
    If bln Then
        For i = 0 To 6
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(5).Left = 0
        cmdTrans(6).Left = cmdTrans(6).Width
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
        For i = 0 To 6
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(5).Left = 0
        cmdTrans(6).Left = cmdTrans(6).Width
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
    End If
End Sub

Private Sub ClearEntries()
    'cboBranch.Text = ""
    cboShipper.Text = ""
    txtWayBill.Text = ""
    txtChkBy.Text = ""
    txtRRno.Text = ""
    txtPackingNo.Text = ""
    txtPLDate.Text = ""
    txtRemarks.Text = ""
    txtDateRec.Text = ""
    txtTotal.Text = ""
    flxParts.Rows = 1
    flxParts.Rows = 2
    flxNewSP.Rows = 0
    txtPrnCopy.Text = ""
End Sub

Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    Frame2.Enabled = Not bln
    'flxParts.Enabled = Not Bln
    txtRemarks.Locked = bln
    'txtChkBy.Locked = bln
    txtDateRec.Locked = bln
    ViewOnly = bln
    cmdPO.Enabled = Not bln
    'cmdPrtlSave.Visible = Not Bln
    'cmdLoadPrtl.Visible = Not Bln
End Sub

Private Sub cboOrderNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cboOrderNo.Text = ""
        cboOrderNo.Tag = "skiplostfocus"
        flxParts.SetFocus
        flxParts.Col = 3
        cboOrderNo.Visible = False
    End If
End Sub

Private Sub cboOrderNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cboOrderNo_LostFocus
        flxParts.SetFocus
        flxParts.Col = 4
    End If
End Sub

Private Sub cboOrderNo_LostFocus()
    Dim myOrder As Integer, MyPOrder As Long, RExc As Long, vRS As String, i As Long
    
    If cboOrderNo.Tag = "skiplostfocus" Then
        cboOrderNo.Tag = ""
        Exit Sub
    End If
    
    'MODULE CREATED 09/08/2011
    'Check if the Item Already added
     For RExc = 1 To flxParts.Rows - 1
        vRS = flxParts.TextMatrix(RExc, 1)
        If RExc <> ActRow Then
            If vRS = flxParts.TextMatrix(ActRow, 1) And cboOrderNo.Text = flxParts.TextMatrix(RExc, 3) Then
                MsgBox ("P.O Number of this Spareparts Already encoded @ Row " & RExc), vbExclamation
                    For i = 1 To 9      'Clear old values of flxParts
                        flxParts.TextMatrix(ActRow, i) = ""
                    Next i
                flxParts.SetFocus
                flxParts.Col = 1
                flxParts.Row = ActRow
                Exit Sub
            End If
        End If
     Next RExc
     
     
    flxParts.TextMatrix(ActRow, 3) = cboOrderNo.Text
    cboOrderNo.Visible = False
    If cboOrderNo.ListIndex >= 0 And flxParts.TextMatrix(ActRow, 8) <> "" Then   'flxParts.TextMatrix(ActRow, 4) = "" And
        MyPOrder = GetOrderRECTEMP(cboOrderNo.Text, flxParts.TextMatrix(ActRow, 8))
        myOrder = GetOrderBalance(cboOrderNo.Text, flxParts.TextMatrix(ActRow, 8)) - MyPOrder
        
        If myOrder > 0 Then
            flxParts.TextMatrix(ActRow, 4) = myOrder
            flxParts.TextMatrix(ActRow, 11) = myOrder
        Else
            flxParts.TextMatrix(ActRow, 4) = 0
            flxParts.TextMatrix(ActRow, 11) = 0
        End If
        Call ComputeTotalQty
    End If
    
    
    
 End Sub
 Function GetOrderRECTEMP(MyRef As String, MyID As Integer) As Long
    Dim QryTemp As String, QtyRT As Long
    Dim rsPOrder As New ADODB.Recordset
    
    QryTemp = "SELECT Sum(SPITR.RECEIVE) AS TOTRECTEMP From SPITR " & _
          "WHERE SPITR.TRANSTYPE='RECTEMP' AND SPITR.SPAREPARTID=" & MyID & " AND SPITR.PACKINGLIST='" & MyRef & "';"
    
    With rsPOrder
        .Open QryTemp, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            QtyRT = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            QtyRT = 0
        End If
        .Close
    End With
    GetOrderRECTEMP = QtyRT
End Function
Private Sub cboShipper_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
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

Private Sub cmdLoadPrtl_Click()
    Dim My1, My2, My3, My4, My5, My6, My7, My8, My9

    On Error GoTo LocalErr
    Open App.Path & "\TEMP\FLXPARTS" For Input As #1   ' Open file for input.
    flxParts.Rows = 1
    Do While Not EOF(1)   ' Loop until end of file.
        Input #1, My1, My2, My3, My4, My5, My6, My7, My8, My9  ' Read data into two variables.
        flxParts.AddItem vbTab & My1 & vbTab & My2 & vbTab & My3 & vbTab & My4 & vbTab & My5 & vbTab & My6 & vbTab & My7 & vbTab & My8 & vbTab & My9
    Loop
    Close #1   ' Close file.
    
    Open App.Path & "\TEMP\FLXNEWSP" For Input As #2   ' Open file for input.
    flxNewSP.Rows = 0
    Do While Not EOF(2)   ' Loop until end of file.
        Input #2, My1, My2, My3, My4, My5
        flxNewSP.AddItem vbTab & My1 & vbTab & My2 & vbTab & My3 & vbTab & My4 & vbTab & My5
    Loop
    Close #2   ' Close file.
    Call ComputeTotalQty
    Exit Sub
LocalErr:
    MsgBox Err.Description, vbExclamation + vbOKOnly, "Load Partial"
End Sub

Private Sub cmdNext_Click()
    Dim i As Integer, j As Integer
    
    With flxParts
        For i = Val(cmdNext.Tag) To .Rows - 1
            For j = 1 To 5
                If cboMatch.Text = "Whole Word" Then
                    If UCase$(txtSearch.Text) = UCase$(.TextMatrix(i, j)) Then
                        .SetFocus: .Row = i: .Col = 0: .ColSel = 10
                        .TopRow = i: cmdNext.Tag = i + 1
                        txtSearch.SelStart = 0: txtSearch.SelLength = Len(txtSearch.Text)
                        txtSearch.SetFocus
                        Call AddInCombo
                        Exit Sub
                    End If
                Else
                    If InStr(1, UCase$(.TextMatrix(i, j)), UCase$(txtSearch.Text)) > 0 Then
                        .SetFocus: .Row = i: .Col = 0: .ColSel = 10
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

Private Sub cmdPO_Click()
    Dim rsPO As New ADODB.Recordset
    Dim Qry As String, MyPO As String, Ans As Integer, POrdQ As Long
    Dim i As Integer, OrdQ As Integer, mySupID As Integer
    
    MyPO = InputBox("Please enter your P.O. Number:", "P.O.")
    If MyPO = "" Then Exit Sub
    'Qry = "SELECT SPORDER.* From SPORDER WHERE (((SPORDER.REFERENCE)='" & MyPO & "'));"
    'mySQL
    Qry = "SELECT SPORDER.* From SPORDER WHERE SPORDER.REFERENCE='" & MyPO & "';"
    
    If Not CtrlNumExist(Qry) Then
        MsgBox "P.O. Number " & MyPO & " not found!", vbExclamation + vbOKOnly, "P.O."
        Exit Sub
    End If
    'Qry = "SELECT SPORDER.* From SPORDER WHERE (((SPORDER.REFERENCE)='" & MyPO & "') AND ((SPORDER.CONFIRMNO) Is Null));"
    'mySQL
    Qry = "SELECT SPORDER.* From SPORDER WHERE SPORDER.REFERENCE='" & MyPO & "' AND (SPORDER.CONFIRMNO) Is Null;"
    
    If CtrlNumExist(Qry) Then
        MsgBox "P.O. Number " & MyPO & " not yet confirmed!", vbExclamation + vbOKOnly, "P.O."
        Exit Sub
    End If
    
    Qry = "SELECT SPORDER.SUPPLIERID From SPORDER WHERE SPORDER.REFERENCE='" & MyPO & "';"
    mySupID = Val(GetFieldValue(Qry))
    
    Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE SUPPLIER.SUPPLIERID=" & mySupID & ";"
    lblSupPO.Caption = GetFieldValue(Qry)
    
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.REFERENCE, SPITR.ORDER, SPAREPARTS.LOCATION, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & MyPO & "'));"
    Ans = MsgBox("Supplier: " & lblSupPO.Caption & vbCrLf & "You want Load All items from P.O number: " & MyPO, vbYesNo + vbQuestion)
    If Ans = vbNo Then Exit Sub
    
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.REFERENCE, SPITR.ORDER, SPAREPARTS.LOCATION, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & MyPO & "';"
    
    With rsPO
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxParts.Rows = 1
            Do While Not .EOF
                flxParts.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !REFERENCE & vbTab & !Order & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & vbTab & vbTab & !SparepartID
                .MoveNext
            Loop
        End If
        .Close
    End With
    With flxParts
        For i = .Rows - 1 To 1 Step -1
            POrdQ = GetOrderRECTEMP(.TextMatrix(i, 3), .TextMatrix(i, 8))
            OrdQ = GetOrderBalance(.TextMatrix(i, 3), .TextMatrix(i, 8)) - POrdQ
            If OrdQ = 0 Then
                If i = 1 And .Rows = 2 Then
                    .Rows = 1
                    .Rows = 2
                Else
                    .RemoveItem i
                End If
            Else
                .TextMatrix(i, 4) = OrdQ
                .TextMatrix(i, 11) = OrdQ
            End If
        Next i
        Call ComputeTotalQty
        If .Rows = 2 And .TextMatrix(1, 1) = "" Then
            MsgBox "All Items in P.O. " & MyPO & " was already served!", vbInformation + vbOKOnly, "P.O."
        Else
            MsgBox "P.O. " & MyPO & " successflly loaded!", vbInformation + vbOKOnly, "P.O."
        End If
    End With
    
            
                 
End Sub

Private Sub cmdPrtlSave_Click()
    Dim i As Integer
    Dim MyDir As String
    
    'Create folder Temp if not available
    MyDir = Right$(App.Path, 4)
    If MyDir = "" Then  'Temp Dir not Found
        MkDir "Temp"
    End If
    Open App.Path & "\TEMP\FLXPARTS" For Output As #1   ' Open file for output.
    With flxParts
        For i = 1 To .Rows - 1
            Write #1, .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 8), .TextMatrix(i, 9) ' Write comma-delimited data.
        Next i
    End With
    Close #1
    
    Open App.Path & "\TEMP\FLXNEWSP" For Output As #2   ' Open file for output.
    With flxNewSP
        For i = 0 To .Rows - 1
            Write #2, .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 5)
        Next i
    End With
    Close #2

End Sub
Private Sub cmdTrans_Click(Index As Integer)
    Dim Qry As String, myMsg As String
    Dim Ans As Integer, i As Integer
    
    Select Case Index
        Case 0      '"new"
            Call BtnSaveCancel(True)
            'BtnIndx = 3
            'tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            txtChkBy.Text = UserName 'GetSettings("STOCK CLERK")
            txtDateRec.Text = Format(ServerTime, "mm/dd/yyyy")
            ViewOnly = False
            AddNew = True
            cboBranch.Clear
            Call cboBranchList
            'lblSupPO.Caption = ""
            'cboSupplier.SetFocus
        Case 1      'edit
            'Validate before allow edit
            If txtRRno.Text = "" Then
                MsgBox "Unable to Edit blank Packing List!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            If Val(txtPrnCopy.Text) > 0 Then
                MsgBox "Unable to Edit because Packing List already Printed!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Edit Packing List?", vbQuestion + vbYesNo, "Edit")
            If Ans = vbNo Then Exit Sub
            txtRRno.Tag = txtPackingNo.Text
            Call LockData(False)
            cmdPO.Enabled = False
            ViewOnly = False
            AddNew = False
            For i = 1 To flxParts.Rows - 1
                flxParts.TextMatrix(i, 10) = flxParts.TextMatrix(i, 8)
            Next i
            Call BtnSaveCancel(True)
            flxNewSP.Rows = 0
        
        Case 2      '"view"
            'SPRecSummary.Show
            FrmSummary.mySrc = "Received Spare Parts"
            FrmSummary.Show
            'SPbranchRecSummary.Show
            Unload Me
        
        Case 3      '"print"
            If txtRRno.Text = "" Then
                MsgBox "No Available Packing List to Print", vbExclamation + vbOKOnly, "Print"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Print Packing List?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            
            PRNmc.MySource = "SPRECEIVE"
            'PRNmc.QryPrn = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.RECEIVE, SPAREPARTS.LOCATION, SPITR.REMARKS " & _
                           "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & txtPackingNo.Text & "'));"
            'mySQL
            PRNmc.QryPrn = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.RECEIVE, SPAREPARTS.LOCATION, SPITR.REMARKS " & _
                           "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='" & txtPackingNo.Text & "';"
            
            'PRNmc.ParamValue1 = IIf((ChkrC.Value = 1), cboSupplier.Text, lblSupPO.Caption)
            PRNmc.ParamValue1 = cboBranch.Text
            PRNmc.ParamValue2 = cboShipper.Text
            PRNmc.ParamValue3 = txtPackingNo.Text
            PRNmc.ParamValue4 = txtPLDate.Text
            PRNmc.ParamValue5 = txtChkBy.Text
            PRNmc.ParamValue6 = txtDateRec.Text
            PRNmc.ParamValue7 = SysBranch
            PRNmc.ParamValue8 = SysAddress
            PRNmc.Show
            'Me.Visible = False
        Case 4      '"quit"
            'MotoRep.tbMain.Visible = True
            
            Unload Me
            'LogIn.Show vbModal
            MotoRep.toolBarButton.Visible = True
            'MotoRep.imgBG.Visible = True
        Case 5      '"save"
            myMsg = ValidateSPRec
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            txtPartNo.Visible = False
            cboOrderNo.Visible = False
            txtLoc.Visible = False
            txtQty.Visible = False
            txtRem.Visible = False
            flxSrch.Visible = False
            Ans = MsgBox("Do you want to Save Packing List?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'BtnIndx = 5
            'tmrFloatRev.Enabled = True
            If AddNew Then
                Call SaveSpareParts
            Else
                Call SaveEditedSP
            End If
            Call LockData(True)
            MsgBox "Received Spare Parts successfully saved!", vbInformation + vbOKOnly, "Save"
            Call BtnSaveCancel(False)
        
        Case 6      '"cancel"
            Ans = MsgBox("Do you want to Cancel Packing List?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            'BtnIndx = 5
            'tmrFloatRev.Enabled = True
            Call BtnSaveCancel(False)
            txtPartNo.Visible = False
            cboOrderNo.Visible = False
            txtLoc.Visible = False
            txtQty.Visible = False
            txtRem.Visible = False
            flxSrch.Visible = False
            Call ClearEntries
            Call LockData(True)
            
            'Add Checkbox Default value
            'ChkrC.Caption = "SUPPLIER"
            'lblSupPO.Visible = True
            'cboSupplier.Visible = False
            'ChkrC.BackColor = &H8080FF
            'ChkrC.Value = 0
            'cmdPO.Enabled = True
            'lblSupPO.Caption = ""
        
    End Select

End Sub

Public Sub FinalSPrec()
    Dim i As Integer, BckOrd As Integer, oBal As Integer
    Dim Qry As String, SuppCode As String, SuppFinal As String
    
    If Val(txtPrnCopy.Text) > 0 Then Exit Sub
    
    'Update SPITR field TransType to RECEIVE
    'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'RECEIVE' " & _
                   "WHERE (((SPITR.REFERENCE)='" & txtPackingNo.Text & "') AND ((SPITR.TRANSTYPE)='RECTEMP'));"
    'mySQL
    LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'RECEIVE' " & _
                   "WHERE SPITR.REFERENCE='" & txtPackingNo.Text & "' AND SPITR.TRANSTYPE='RECTEMP';"
    
    'Update Print Copy
    txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
    'LYRcon.Execute "UPDATE SPRECEIVE SET SPRECEIVE.PRINTCOPY = " & Val(txtPrnCopy.Text) & _
                   " WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
    'mySQL
    LYRcon.Execute "UPDATE SPRECEIVE SET SPRECEIVE.PRINTCOPY = " & Val(txtPrnCopy.Text) & _
                   " WHERE SPRECEIVE.PACKINGNO='" & txtPackingNo.Text & "';"
    
    'Qry = "SELECT SUPPLIER.SUPPLIERCODE From SUPPLIER WHERE (((SUPPLIER.SUPPLIERNAME)='" & cboSupplier.Text & "'));"
    'mySQL
'    If ChkrC.Value = 0 Then
'        SuppFinal = lblSupPO.Caption
'    Else
'        SuppFinal = cboSupplier.Text
'    End If
    
'    Qry = "SELECT SUPPLIER.SUPPLIERCODE From SUPPLIER WHERE SUPPLIER.SUPPLIERNAME='" & SuppFinal & "';"
'
'    SuppCode = GetFieldValue(Qry)
    'Update Qty on Hand of all received SpareParts
    With flxParts
        For i = 1 To .Rows - 1
            BckOrd = GetBackOrder(Val(.TextMatrix(i, 8)))
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+ " & Val(.TextMatrix(i, 4)) & ", SPAREPARTS.REFERENCE = '" & txtPackingNo.Text & "', SPAREPARTS.SUPPLIERCODE = '" & SuppCode & "', SPAREPARTS.BACKORDER = " & BckOrd & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
            
            'mySQL
            'Update with supplier code if Not stock transfer from Moto-r Dvo, Tagum, Mohurhub
            
            'If SuppCode <> "CW" And SuppCode <> "MT" And SuppCode <> "MH" Then
            'from supplier
            'If ChkrC.Value = 0 Then
            '    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+ " & Val(.TextMatrix(i, 4)) & ", SPAREPARTS.REFERENCE = '" & txtPackingNo.Text & "', SPAREPARTS.SUPPLIERCODE = '" & SuppCode & "', SPAREPARTS.BACKORDER = " & BckOrd & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
            'from Branch
            'Else
            '    Update w/out Supplier Code
            '    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+ " & Val(.TextMatrix(i, 4)) & ", SPAREPARTS.REFERENCE = '" & txtPackingNo.Text & "', SPAREPARTS.BACKORDER = " & BckOrd & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
            'End If
             
             LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+ " & Val(.TextMatrix(i, 4)) & ", SPAREPARTS.LOCATION = '" & .TextMatrix(i, 5) & "', SPAREPARTS.REFERENCE = '" & txtPackingNo.Text & "', SPAREPARTS.BACKORDER = " & BckOrd & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
        
            oBal = GetOrderBalance(.TextMatrix(i, 3), .TextMatrix(i, 8))
            'LYRcon.Execute "UPDATE SPITR SET SPITR.[BALANCE] = " & oBal & _
            " WHERE (((SPITR.SPAREPARTID)=" & .TextMatrix(i, 8) & ") AND ((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & .TextMatrix(i, 3) & "'));"
            'mySQL
            LYRcon.Execute "UPDATE SPITR SET SPITR.BALANCE = " & oBal & _
            " WHERE SPITR.SPAREPARTID=" & .TextMatrix(i, 8) & " AND SPITR.TRANSTYPE='ORDER' AND SPITR.REFERENCE='" & .TextMatrix(i, 3) & "';"
            
        Next i
    End With
    
    MsgBox "Spare Parts Qty-On-Hand successfully updated!"
End Sub


Function GetBackOrder(MyID As Long) As Long
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String
    Dim xRow As Integer, xBal As Integer, i As Integer
    
    'Clear FlxStock
    'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & MyID & ")) ORDER BY SPITR.TRANSDATE;"
    'mySQL
    Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
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
            'flxStock.AddItem .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & vbTab & .TextMatrix(i, 2) & vbTab & vbTab & xBal   '.TextMatrix(i, 2)
            'Acquire Packing list for P.O.
            MyPO = .TextMatrix(i, 1)
            'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.PACKINGLIST)='" & MyPO & "')) ORDER BY SPITR.TRANSDATE;"
            'mySQL
            Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & MyID & " AND SPITR.PACKINGLIST='" & MyPO & "' ORDER BY SPITR.TRANSDATE;"
            
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

Private Sub flxParts_Click()
    ActRow = flxParts.Row
    ActCol = flxParts.Col
    
    If ActCol = 0 Or ActRow = 0 Or ViewOnly Then Exit Sub
    Select Case ActCol
        Case 1  'PartNumber
            With txtPartNo
                If flxParts.TextMatrix(ActRow, 1) <> "" Then flxSrch.Tag = flxParts.TextMatrix(ActRow, 1) & "|" & flxParts.TextMatrix(ActRow, 2) & "|" & flxParts.TextMatrix(ActRow, 3) & "|" & flxParts.TextMatrix(ActRow, 4) & "|" & flxParts.TextMatrix(ActRow, 5) & "|" & flxParts.TextMatrix(ActRow, 6) & "|" & flxParts.TextMatrix(ActRow, 7) & "|" & flxParts.TextMatrix(ActRow, 8) & "|" & flxParts.TextMatrix(ActRow, 9)
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
        Case 3 'Order No.
            With cboOrderNo
                If flxParts.TextMatrix(ActRow, 3) <> "" Then .Tag = flxParts.TextMatrix(ActRow, 3)
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
        Case 4  'Qty.
            With txtQty
                If flxParts.TextMatrix(ActRow, 4) <> "" Then .Tag = flxParts.TextMatrix(ActRow, 4)
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
        Case 5  'Location
            With txtLoc
                If flxParts.TextMatrix(ActRow, 5) <> "" And flxParts.TextMatrix(ActRow, 8) <> "" Then Exit Sub
                If flxParts.TextMatrix(ActRow, 5) <> "" And flxParts.TextMatrix(ActRow, 8) = "" Then .Tag = flxParts.TextMatrix(ActRow, 5)
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
        Case 6  'Remarks
            With txtRem
                If flxParts.TextMatrix(ActRow, 6) <> "" Then .Tag = flxParts.TextMatrix(ActRow, 6)
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
End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxParts_Click
End Sub
Private Sub flxParts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ViewOnly Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuRows)
    End If
End Sub

Private Sub flxParts_RowColChange()
    With flxParts
        txtTotal.Top = .Top + .Height + 50
        txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3)
        txtTotal.Width = .ColWidth(4)
        lblTotal.Top = txtTotal.Top
        lblTotal.Left = txtTotal.Left - lblTotal.Width + 50
        lblTotal.FontSize = txtTotal.FontSize
    End With
End Sub

Private Sub flxSrch_Click()
    FromFlx = True
End Sub

Private Sub flxSrch_EnterCell()
    With flxSrch
        FromFlx = True
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        txtPartNo.Text = .TextMatrix(.Row, 2)
    End With
End Sub

Private Sub flxSrch_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        If flxSrch.Tag <> "" Then
            Tmp = Split(flxSrch.Tag, "|")
            flxParts.TextMatrix(ActRow, 1) = Tmp(0)
            flxParts.TextMatrix(ActRow, 2) = Tmp(1)
            flxParts.TextMatrix(ActRow, 6) = Tmp(2)
            flxParts.TextMatrix(ActRow, 7) = Tmp(3)
            flxParts.TextMatrix(ActRow, 8) = Tmp(4)
        End If
        txtPartNo.Visible = False
        flxSrch.Visible = False
        flxParts.SetFocus
        flxParts.Col = 1
    End If
End Sub

Private Sub flxSrch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        flxParts.SetFocus
        flxParts.Col = 3
    End If
End Sub

Private Sub flxSrch_LostFocus()
Dim i As Long, j As Long
    'Validate existing PartNo
    For i = 1 To flxParts.Rows - 1
        If i <> ActRow Then
            If flxParts.TextMatrix(i, 1) = txtPartNo.Text Then
                flxParts.TextMatrix(ActRow, 1) = txtPartNo.Text
                txtPartNo.Visible = False
                flxSrch.Visible = False
                Ans = MsgBox("Part Number " & txtPartNo.Text & " already exist in row " & i, vbInformation + vbOKOnly, "Duplicate Part Number")
                For j = 1 To flxParts.Cols - 1
                    flxParts.TextMatrix(ActRow, j) = ""
                Next j
                flxParts.Row = i
                flxParts.Col = 1
                flxParts.SetFocus
                txtPartNo.Tag = "PN exist"
                Exit Sub
            End If
        End If
    Next i
End Sub

Private Sub Form_Load()
    cmdA.Top = 15000
    Call LockData(True)
    
    Timer1.Enabled = True
    cboMatch.AddItem "Whole Word"
    cboMatch.AddItem "Any Part of Word"

    MotoRep.sBar.Panels(6).Text = Me.Caption
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
'    If rsOrder.State = adStateOpen Then rsOrder.Close

End Sub

Private Sub mnuAdd_Click()
    If flxParts.TextMatrix(ActRow, 1) = "" Or flxParts.TextMatrix(ActRow, 2) = "" Or flxParts.TextMatrix(ActRow, 4) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        flxParts.SetFocus
        flxParts.Col = ActCol
        Exit Sub
    Else
        If ActRow = flxParts.Rows - 1 Then flxParts.AddItem ""
        flxParts.SetFocus
        flxParts.Col = 1
        flxParts.Row = flxParts.Rows - 1
        flxParts.TopRow = flxParts.Rows - 1
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If AddNew Then
            If flxParts.TextMatrix(ActRow, 9) <> "" Then
                If flxNewSP.TextMatrix(Val(flxParts.TextMatrix(ActRow, 9)), 0) = flxParts.TextMatrix(ActRow, 1) Then flxNewSP.RemoveItem Val(flxParts.TextMatrix(ActRow, 9))
            End If
        Else
            If flxParts.TextMatrix(ActRow, 10) <> "" Then
                lstDelParts.AddItem flxParts.TextMatrix(ActRow, 10)
            End If
        End If
        If ActRow = 1 And flxParts.Rows = 2 Then
            flxParts.Rows = 1
            flxParts.Rows = 2
        Else
            flxParts.RemoveItem ActRow
            ActRow = ActRow - 1
        End If
    End If
End Sub
Private Sub ChkRC_Click()
    
    If ChkrC.Value = 1 Then
        ChkrC.Caption = "BRANCH"
        cboSupplier.Visible = True
        lblSupPO.Visible = False
        cmdPO.Enabled = False
        ChkrC.BackColor = &HFF8080
        
    Else
        
        ChkrC.Caption = "SUPPLIER"
        lblSupPO.Visible = True
        cboSupplier.Visible = False
        cmdPO.Enabled = True
        ChkrC.BackColor = &H8080FF
    End If
    Call LoadSpareParts
End Sub

Private Sub RXResizer1_ResizeComplete()
        
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width + 30)
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    With flxParts
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 380) * 0.2
        .ColWidth(2) = (.Width - 380) * 0.25
        .ColWidth(3) = (.Width - 380) * 0.12
        .ColWidth(4) = (.Width - 380) * 0.1
        .ColWidth(5) = (.Width - 380) * 0.16
        .ColWidth(6) = (.Width - 380) * 0.16
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        txtTotal.Top = .Top + .Height + 50
        txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3)
        txtTotal.Width = .ColWidth(4)
        lblTotal.Top = txtTotal.Top
        lblTotal.Left = txtTotal.Left - lblTotal.Width + 50
        lblTotal.FontSize = txtTotal.FontSize
    End With
    With flxSrch
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = .Width
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
    End With
    Label2.Top = txtRemarks.Top
    Label2.Left = flxParts.Left
    Label2.FontSize = txtRemarks.FontSize
    lblHeading(4).Top = flxParts.Top - 255
    lblHeading(5).Top = flxParts.Top - 235
    tabFind.Top = flxParts.Top - tabFind.Height
    tabFind.Left = cmdFind.Left - (tabFind.Width + 500)
End Sub

Private Sub ComputeTotalQty()
    Dim i As Integer, Ttl As Integer
    
    Ttl = 0
    With flxParts
        For i = 1 To .Rows - 1
            Ttl = Ttl + Val(.TextMatrix(i, 4))
        Next i
    End With
    txtTotal.Text = Ttl
End Sub
                           
Private Sub SaveEditedSP()
    Dim rsSPrec As New ADODB.Recordset
    Dim i As Integer, Rnew As Integer
    Dim Qry As String, SupNameEdit As String
    
    'Saving SPReceive
    'Qry = "SELECT SPRECEIVE.*, SPRECEIVE.SPRRNO From SPRECEIVE WHERE (((SPRECEIVE.SPRRNO)=" & txtRRno.Text & "));"
    
    'rsSPrec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    'If rsSPrec.RecordCount > 0 Then
    '    rsSPrec.MoveFirst
    '    rsSPrec!DATEREC = txtDateRec.Text
    '    rsSPrec!SUPPLIER = cboSupplier.Text
    '    rsSPrec!SHIPPER = cboShipper.Text
    '    rsSPrec!WAYBILLNO = txtWayBill.Text
    '    rsSPrec!PACKINGNO = txtPackingNo.Text
    '    If IsDate(txtPLDate.Text) Then rsSPrec!PACKDATE = txtPLDate.Text
    '    rsSPrec!CHECKEDBY = txtChkBy.Text
    '    rsSPrec!REMARKS = txtRemarks.Text
    '    rsSPrec!PRINTCOPY = 0
    '    rsSPrec!UPDATEDBY = UserID
    '    rsSPrec!LASTUPDATE = Date
    '    rsSPrec.Update
    '    rsSPrec.Close
    'End If
    
    'mySQL
'    If ChkrC.Value = 1 Then
'        SupNameEdit = cboSupplier.Text
'    Else
'        SupNameEdit = lblSupPO.Caption
'    End If
    LYRcon.Execute "UPDATE spreceive SET DATEREC = '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', " & _
        "SUPPLIER = '" & cboBranch.Text & "', SHIPPER = '" & cboShipper.Text & "', WAYBILLNO = '" & txtWayBill.Text & "', " & _
        "PACKINGNO = '" & txtPackingNo.Text & "', PACKDATE = " & IIf(IsDate(txtPLDate.Text), " '" & Format(txtPLDate.Text, "yyyy-mm-dd") & "'", Null) & ", " & _
        "CHECKEDBY = '" & txtChkBy.Text & "', REMARKS = '" & txtRemarks.Text & "', UPDATEDBY = " & UserID & ", " & _
        "LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' Where spreceive.SPRRNO = '" & txtRRno.Text & "';"
    
    
    'Delete SPITR reference then add new
    'LYRcon.Execute "DELETE SPITR.* From SPITR WHERE (((SPITR.REFERENCE)='" & txtRRno.Tag & "') AND ((SPITR.TRANSTYPE)='RECTEMP'));"
    'mySQL
    LYRcon.Execute "DELETE From SPITR WHERE SPITR.REFERENCE='" & txtRRno.Tag & "' AND SPITR.TRANSTYPE='RECTEMP';"
    
    With flxParts
        For i = 1 To .Rows - 1
            'Update table SpareParts field Qty., if SP does not exist add to new record
            If .TextMatrix(i, 8) <> "" Then     'SP already available in MDB
                'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+ " & Val(.TextMatrix(i, 4)) & ", SPAREPARTS.REFERENCE = '" & txtPackingNo.Text & "', SPAREPARTS.LOCATION = '" & .TextMatrix(i, 5) & "' " & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
                If .TextMatrix(i, 9) = "New in SP only" Then
                    'Update SparePart Database
                    LYRcon.Execute "INSERT INTO SOURCESP ( PARTNUMBER, DESCRIPTION) " & _
                                  "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "');"
                End If
            Else        'New SP
                Rnew = Val(.TextMatrix(i, 9))
                'LYRcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, REFERENCE, LOCATION, COST, STOREPRICE, BEGBAL, REMARKS, DATEBEGBAL, SUPERSEDINGPN, APPLICMODEL, UPDATEDBY, LASTUPDATE) " & _
                               "VALUES ('" & .TextMatrix(i, 1) & "', '" & Trim(.TextMatrix(i, 2)) & "', '" & txtPackingNo.Text & "', '" & .TextMatrix(i, 5) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ", " & Val(.TextMatrix(i, 4)) & ", '" & flxNewSP.TextMatrix(Rnew, 5) & "', #" & Date & "#, '" & flxNewSP.TextMatrix(Rnew, 4) & "', '" & flxNewSP.TextMatrix(Rnew, 6) & "', " & UserID & ", #" & Date & "#);"
                'mySQL
                LYRcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, REFERENCE, LOCATION, COST, STOREPRICE, BEGBAL, REMARKS, DATEBEGBAL, SUPERSEDINGPN, APPLICMODEL, UPDATEDBY, LASTUPDATE) " & _
                               "VALUES ('" & .TextMatrix(i, 1) & "', '" & Trim(.TextMatrix(i, 2)) & "', '" & txtPackingNo.Text & "', '" & .TextMatrix(i, 5) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ", " & Val(.TextMatrix(i, 4)) & ", '" & flxNewSP.TextMatrix(Rnew, 5) & "', '" & Format(Date, "yyyy-mm-dd") & "', '" & flxNewSP.TextMatrix(Rnew, 4) & "', '" & flxNewSP.TextMatrix(Rnew, 6) & "', " & UserID & ", '" & Format(Date, "yyyy-mm-dd") & "');"
                
                'If flxNewSP.TextMatrix(Rnew, 3) <> "" Then  'Add Description2
                'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '" & flxNewSP.TextMatrix(Rnew, 3) & "' " & _
                                   "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                'mySQL
                If flxNewSP.TextMatrix(Rnew, 3) <> "" Then  'Add Description2
                    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '" & flxNewSP.TextMatrix(Rnew, 3) & "' " & _
                                   "WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                
                End If
                'Update SparePart Database
                If Not SPExist(flxParts.TextMatrix(i, 1)) Then
                    LYRcon.Execute "INSERT INTO SOURCESP ( PARTNUMBER, DESCRIPTION, SUPERSEDING, COST, RETAILPRICE) " & _
                                  "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', '" & flxNewSP.TextMatrix(Rnew, 4) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ");"
                End If
            End If
            'Update table SPITR
            'if no ID ... get SP ID
            If .TextMatrix(i, 8) = "" Then .TextMatrix(i, 8) = GetSPid(i)
            'get order balance
            'Qry = "SELECT SPITR.* From SPITR WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & .TextMatrix(i, 3) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
            'mySQL
            Qry = "SELECT SPITR.* From SPITR WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.REFERENCE='" & .TextMatrix(i, 3) & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
            
            If CtrlNumExist(Qry) Then
                Qry = "Exist"
                myOrdBal = GetOrderBalance(.TextMatrix(i, 3), .TextMatrix(i, 8))
            Else
                Qry = ""
                myOrdBal = 0
            End If
            'Save SPITR
            'LYRcon.Execute "INSERT INTO SPITR ( REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, ORDERBAL, PACKINGLIST, REMARKS, SOURCE, UPDATEDBY, LASTUPDATE ) " & _
                           "VALUES ('" & txtPackingNo.Text & "', 'RECTEMP', #" & txtDateRec.Text & "#, " & flxParts.TextMatrix(i, 8) & ", " & Val(flxParts.TextMatrix(i, 4)) & ", '" & IIf(Qry = "Exist", myOrdBal, "") & "', '" & flxParts.TextMatrix(i, 3) & "', '" & flxParts.TextMatrix(i, 6) & "', '" & cboSupplier.Text & "', " & UserID & ", #" & Now & "# );"
            'mySQL
            LYRcon.Execute "INSERT INTO SPITR ( REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, ORDERBAL, PACKINGLIST, REMARKS, SOURCE, UPDATEDBY, LASTUPDATE ) " & _
                           "VALUES ('" & txtPackingNo.Text & "', 'RECTEMP', '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', " & flxParts.TextMatrix(i, 8) & ", " & Val(flxParts.TextMatrix(i, 4)) & ", '" & IIf(Qry = "Exist", myOrdBal, "") & "', '" & flxParts.TextMatrix(i, 3) & "', '" & flxParts.TextMatrix(i, 6) & "', '" & SupNameEdit & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' );"
        Next i
    End With
End Sub

Private Sub SaveSpareParts()
    Dim rsReceive As New ADODB.Recordset
    Dim i As Integer, myOrdBal As Integer, Rnew As Integer, myPrice As Currency
    Dim Qry As String, Supname As String
    
    'Saving SPReceive
    'rsReceive.Open "SPReceive", LYRcon, adOpenKeyset, adLockOptimistic
    'rsReceive.AddNew
    'rsReceive!DATEREC = txtDateRec.Text
    'rsReceive!SUPPLIER = cboSupplier.Text
    'rsReceive!SHIPPER = cboShipper.Text
    'rsReceive!WAYBILLNO = txtWayBill.Text
    'rsReceive!PACKINGNO = txtPackingNo.Text
    'If IsDate(txtPLDate.Text) Then rsReceive!PACKDATE = txtPLDate.Text
    'rsReceive!CHECKEDBY = txtChkBy.Text
    'rsReceive!REMARKS = txtRemarks.Text
    'rsReceive!PRINTCOPY = 0
    'rsReceive!UPDATEDBY = UserID
    'rsReceive!LASTUPDATE = Date
    'rsReceive.Update
    
    
    'rsReceive.Close
    'txtRRno.Text = rsReceive!SPRRNO
        
    'mySQL here
'    If ChkrC.Value = 1 Then
'        Supname = cboSupplier.Text
'    Else
'        Supname = lblSupPO.Caption
'    End If
    
    SPrlastUp = Format(Now, "yyyy-mm-dd hh:mm:ss")
    LYRcon.Execute "INSERT INTO spreceive(DATEREC, SUPPLIER, SHIPPER, WAYBILLNO, PACKINGNO, CHECKEDBY, REMARKS, UPDATEDBY, LASTUPDATE, AUTOLOAD) " & _
            "VALUES('" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', '" & cboBranch.Text & "', '" & cboShipper.Text & "', '" & txtWayBill.Text & "', '" & txtPackingNo.Text & "', '" & txtChkBy.Text & "', '" & txtRemarks.Text & "', " & UserID & ", '" & SPrlastUp & "', 1);"
    
    Qry = "SELECT spreceive.SPRRNO From spreceive Where spreceive.LASTUPDATE = '" & SPrlastUp & "' AND spreceive.PACKINGNO = '" & txtPackingNo.Text & "';"
    txtRRno.Text = GetFieldValue(Qry)
    If IsDate(txtPLDate.Text) Then LYRcon.Execute "UPDATE spreceive SET PACKDATE = '" & Format(DateValue(txtPLDate.Text), "yyyy-mm-dd") & "' Where spreceive.SPRRNO = '" & txtRRno.Text & "';"
    
    
    With flxParts
        For i = 1 To .Rows - 1
            'Update table SpareParts field Qty., if SP does not exist add to new record
            If .TextMatrix(i, 8) <> "" Then     'SP already available in MDB
                'old LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+ " & Val(.TextMatrix(i, 4)) & ", SPAREPARTS.REFERENCE = '" & txtPackingNo.Text & "', SPAREPARTS.LOCATION = '" & .TextMatrix(i, 5) & "' " & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
                If .TextMatrix(i, 9) = "New in SP only" Then
                    'Update SparePart Database
                    'LYRcon.Execute "INSERT INTO SOURCESP ( PARTNUMBER, DESCRIPTION) " & _
                                  "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "');"
                    'mySQL
                    LYRcon.Execute "INSERT INTO SOURCESP ( PARTNUMBER, DESCRIPTION) " & _
                                  "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "');"
                    
                End If
            Else        'New SP
                Rnew = Val(.TextMatrix(i, 9))
                'LYRcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, REFERENCE, LOCATION, COST, STOREPRICE, BEGBAL, REMARKS, DATEBEGBAL, SUPERSEDINGPN, APPLICMODEL, UPDATEDBY, LASTUPDATE) " & _
                               "VALUES ('" & .TextMatrix(i, 1) & "', '" & Trim(.TextMatrix(i, 2)) & "', '" & txtPackingNo.Text & "', '" & .TextMatrix(i, 5) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ", " & Val(.TextMatrix(i, 4)) & ", '" & flxNewSP.TextMatrix(Rnew, 5) & "', #" & Date & "#, '" & flxNewSP.TextMatrix(Rnew, 4) & "', '" & flxNewSP.TextMatrix(Rnew, 6) & "', " & UserID & ", #" & Date & "#);"
                'mySQL
                LYRcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, REFERENCE, LOCATION, COST, STOREPRICE, BEGBAL, REMARKS, DATEBEGBAL, SUPERSEDINGPN, APPLICMODEL, UPDATEDBY, LASTUPDATE) " & _
                               "VALUES ('" & .TextMatrix(i, 1) & "', '" & Trim(.TextMatrix(i, 2)) & "', '" & txtPackingNo.Text & "', '" & .TextMatrix(i, 5) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ", " & Val(.TextMatrix(i, 4)) & ", '" & flxNewSP.TextMatrix(Rnew, 5) & "', '" & Format(Date, "yyyy-mm-dd") & "', '" & flxNewSP.TextMatrix(Rnew, 4) & "', '" & flxNewSP.TextMatrix(Rnew, 6) & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd") & "');"
                
                If flxNewSP.TextMatrix(Rnew, 3) <> "" Then  'Add Description2
                    'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '" & flxNewSP.TextMatrix(Rnew, 3) & "' " & _
                                   "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                    'mySQL
                    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '" & flxNewSP.TextMatrix(Rnew, 3) & "' " & _
                                   "WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(i, 1) & "';"
                End If
                'Update SparePart Database
                If Not SPExist(flxParts.TextMatrix(i, 1)) Then
                    LYRcon.Execute "INSERT INTO SOURCESP ( PARTNUMBER, DESCRIPTION, SUPERSEDING, COST, RETAILPRICE) " & _
                                  "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', '" & flxNewSP.TextMatrix(Rnew, 4) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ");"
                End If
            End If
            'Update table SPITR
            'if no ID ... get SP ID
            If .TextMatrix(i, 8) = "" Then .TextMatrix(i, 8) = GetSPid(i)
            'get order balance
            'Qry = "SELECT SPITR.* From SPITR WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & .TextMatrix(i, 3) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
            'mySQL
            Qry = "SELECT SPITR.* From SPITR WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.REFERENCE='" & .TextMatrix(i, 3) & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
            
            If CtrlNumExist(Qry) Then
                Qry = "Exist"
                myOrdBal = GetOrderBalance(.TextMatrix(i, 3), .TextMatrix(i, 8))
            Else
                Qry = ""
                myOrdBal = 0
            End If
            'Save SPITR
            'Qry = "SELECT SPAREPARTS.COST From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & flxParts.TextMatrix(i, 8) & "));"
            'mySQL
            Qry = "SELECT SPAREPARTS.COST From SPAREPARTS WHERE SPAREPARTS.SPAREPARTID=" & flxParts.TextMatrix(i, 8) & ";"
            
            myPrice = Val(GetFieldValue(Qry))
            'LYRcon.Execute "INSERT INTO SPITR ( REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, ORDERBAL, PACKINGLIST, REMARKS, SOURCE, PRICE, UPDATEDBY, LASTUPDATE ) " & _
                           "VALUES ('" & txtPackingNo.Text & "', 'RECTEMP', #" & txtDateRec.Text & "#, " & flxParts.TextMatrix(i, 8) & ", " & Val(flxParts.TextMatrix(i, 4)) & ", '" & IIf(Qry = "Exist", myOrdBal, "") & "', '" & flxParts.TextMatrix(i, 3) & "', '" & flxParts.TextMatrix(i, 6) & "', '" & cboSupplier.Text & "', " & MyPrice & ", " & UserID & ", #" & Now & "# );"
            'mySQL
            LYRcon.Execute "INSERT INTO SPITR ( REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, ORDERBAL, PACKINGLIST, REMARKS, SOURCE, PRICE, UPDATEDBY, LASTUPDATE ) " & _
                           "VALUES ('" & txtPackingNo.Text & "', 'RECTEMP', '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', " & flxParts.TextMatrix(i, 8) & ", " & Val(flxParts.TextMatrix(i, 4)) & ", '" & IIf(Qry = "Exist", myOrdBal, "") & "', '" & flxParts.TextMatrix(i, 3) & "', '" & flxParts.TextMatrix(i, 6) & "', '" & cboBranch.Text & "', " & myPrice & ", " & UserID & ", '" & Format(Now, "yyyy-mm-dd") & "' );"
            
            'old Update SpareParts BackOrder and OrderRef
            'If Qry = "Exist" Then
            '    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & myOrdBal & _
            '                   " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
            '     LYRcon.Execute "UPDATE SPITR SET SPITR.BALANCE = '" & myOrdBal & "' WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 8) & ") AND ((SPITR.REFERENCE)='" & .TextMatrix(i, 3) & "'));               "
            'End If
        Next i
    End With
End Sub

Function SPExist(myPN As String) As Boolean
    Dim rsSPall As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SOURCESP.PARTNUMBER From SOURCESP WHERE (((SOURCESP.PARTNUMBER)='" & myPN & "'));"
    'mySQL
    Qry = "SELECT SOURCESP.PARTNUMBER From SOURCESP WHERE SOURCESP.PARTNUMBER='" & myPN & "';"
    
    With rsSPall
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        SPExist = IIf(.RecordCount > 0, True, False)
        .Close
    End With
End Function

Function GetOrderBalance(MyRef As String, MyID As Long) As Long
    Dim rsOBal As New ADODB.Recordset
    Dim rsROrder As New ADODB.Recordset
    Dim Qry As String, Alloc As Integer, Recv As Integer
    
    'Qry = "SELECT SPITR.ORDER From SPITR WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & MyRef & "') AND ((SPITR.SPAREPARTID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT SPITR.ORDER From SPITR WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.REFERENCE='" & MyRef & "' AND SPITR.SPAREPARTID=" & MyID & ";"
    
    With rsOBal
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Alloc = IIf(IsNull(!Order), 0, !Order)
        Else
            Alloc = 0
        End If
        .Close
    End With
    'Qry = "SELECT Sum(SPITR.[RECEIVE]) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.PACKINGLIST)='" & MyRef & "'));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & MyID & " AND SPITR.PACKINGLIST='" & MyRef & "';"
    
    With rsROrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Recv = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            Recv = 0
        End If
        .Close
    End With            'Order Allocation - [Previous Receive Parts + Current Receive Parts]
    GetOrderBalance = Alloc - Recv '+ Val(flxParts.TextMatrix(MyRow, 4)))
End Function

Function GetSPid(MyRow As Integer) As Integer
    Dim rsGetid As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.SPAREPARTID From SPAREPARTS " & _
          "WHERE (((SPAREPARTS.PARTNUMBER)='" & flxParts.TextMatrix(MyRow, 1) & "'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.SPAREPARTID From SPAREPARTS " & _
          "WHERE SPAREPARTS.PARTNUMBER='" & flxParts.TextMatrix(MyRow, 1) & "';"
    
    With rsGetid
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetSPid = !SparepartID
        Else
            GetSPid = 0
        End If
        .Close
    End With
End Function

Private Sub LoadSpareParts()
    Dim rsSupplier As New ADODB.Recordset
    Dim rsShipper As New ADODB.Recordset
    Dim Qry As String, MySupp As String
    
    If rsParts.State = adStateOpen Then Exit Sub
    'Qry = "SELECT SOURCESP.* From SOURCESP ORDER BY SOURCESP.PARTNUMBER;"
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.COST, SPAREPARTS.STOREPRICE, SPAREPARTS.APPLICMODEL " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"

    
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSrch.DataSource = rsParts
    
    
'    Qry = "SELECT supplier.SUPPLIERNAME, supplier.SUPPLIERID From supplier Where supplier.UPDATEDBY = 999 Order By supplier.SUPPLIERNAME;"
'    With rsSupplier
'        If .State = adStateOpen Then .Close
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'                .MoveFirst
'                If cboSupplier.Text <> "" Then MySupp = cboSupplier.Text
'                cboSupplier.Clear
'                Do While Not .EOF
'                    cboSupplier.AddItem !SUPPLIERNAME
'                    cboSupplier.ItemData(cboSupplier.ListCount - 1) = !SUPPLIERID
''                    A = cboSupplier.ItemData(0)
'                    .MoveNext
'                Loop
'        End If
'        .Close
'        If MySupp <> "" Then cboSupplier.Text = MySupp
'    End With
    
'    With cboSupplier
'        If .Text <> "" Then MySupp = .Text
'                .Clear
'                .AddItem "MC PARTS DEPOT CORP."
'                .ItemData(cboSupplier.ListCount - 1) = 102
'                .AddItem "MOTO-R TAGUM"
'                .ItemData(cboSupplier.ListCount - 1) = 100
'                .AddItem "MOTO-R VAN SELLING"
'                .ItemData(cboSupplier.ListCount - 1) = 109
'                .AddItem "MOTO-R ULAS"
'                .ItemData(cboSupplier.ListCount - 1) = 81
'                .AddItem "YAMAHA MOTOR-HUB"
'                .ItemData(cboSupplier.ListCount - 1) = 99
'
'        If MySupp <> "" Then cboSupplier.Text = MySupp
'    End With
    
    With rsShipper
        If .State = adStateClosed Then
            .Open "Shipper", LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                'cboShipper.Clear
                Do While Not .EOF
                    cboShipper.AddItem .Fields(1)
                    .MoveNext
                Loop
            End If
            .Close
        End If
    End With
    
    DoEvents
    Call cboBranchList
    
End Sub

Private Sub cboBranchList()

    Qry = "SELECT BRANCHID, BRANCHNAME, BRANCHCODE FROM branch where branchid >= '12';"
    If rsAnyUse.State <> 0 Then rsAnyUse.Close
    Set rsAnyUse = LYRcon.Execute(Qry)
    If rsAnyUse.RecordCount > 0 Then
        
        'cboBranch.Clear
        rsAnyUse.MoveFirst
        Do While Not rsAnyUse.EOF
            cboBranch.AddItem rsAnyUse!branchName
            rsAnyUse.MoveNext
        Loop
           
    End If
    If rsAnyUse.State <> 0 Then rsAnyUse.Close
    

End Sub

Function ValidateSPRec() As String
    Dim vMsg As String, Qry As String
    Dim cnt As Integer
    Dim i As Long, j As Long
    
    vMsg = ""
    
'    If ChkrC.Value = 1 Then
'        If cboSupplier.Text = "" Then
'            vMsg = "The field Branch must not be empty."
'        Else
'            'Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE (((SUPPLIER.SUPPLIERNAME)='" & cboSupplier.Text & "'));"
'            'mySQL
'            Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE SUPPLIER.SUPPLIERNAME='" & cboSupplier.Text & "';"
'
'            If Not CtrlNumExist(Qry) Then vMsg = "Unknown BRANCH Name, please select from the list."
'        End If
'
'    Else
'
''        If lblSupPO.Caption = "" Then
''            vMsg = "The field SUPPLIER must not be empty."
''        Else
'            'Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE (((SUPPLIER.SUPPLIERNAME)='" & cboSupplier.Text & "'));"
'            'mySQL
'            Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE SUPPLIER.SUPPLIERNAME='" & lblSupPO.Caption & "';"
'
'            If Not CtrlNumExist(Qry) Then vMsg = "Unknown SUPPLIER Name, please select from the list."
'        End If
'    End If
    
    If cboBranch.Text = "" Then
            vMsg = "The field Branch must not be empty."
    End If
    If txtPackingNo.Text = "" Then
        vMsg = vMsg & vbCrLf & "The field PACKING NO. must not be empty."
    Else
        If AddNew Then
            'Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
                  "WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
            'mySQL
            Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
                  "WHERE SPRECEIVE.PACKINGNO='" & txtPackingNo.Text & "';"
            
            If CtrlNumExist(Qry) Then vMsg = vMsg & vbCrLf & "The field PACKING NO. already exist."
        Else
            If txtRRno.Tag <> txtPackingNo.Text Then
                'Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
                      "WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
                'mySQL
                Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
                      "WHERE SPRECEIVE.PACKINGNO='" & txtPackingNo.Text & "';"
                
                If CtrlNumExist(Qry) Then vMsg = vMsg & vbCrLf & "The field PACKING NO. already exist."
            End If
        End If
    End If
    If txtPLDate.Text <> "" Then
        If Not IsDate(txtPLDate.Text) Then vMsg = vMsg & vbCrLf & "The field Packing List Print Date must not be empty."
    End If
    If txtChkBy.Text = "" Then vMsg = vMsg & vbCrLf & "The field CHECKED BY must not be empty."
    If Not IsDate(txtDateRec.Text) Then vMsg = vMsg & vbCrLf & "Invalid entry on field Date."
    With flxParts
        If .Rows > 2 Then
            If .TextMatrix(.Rows - 1, 1) = "" And .TextMatrix(.Rows - 1, 2) = "" And .TextMatrix(.Rows - 1, 4) = "" And (.TextMatrix(.Rows - 1, 8) = "" Or .TextMatrix(.Rows - 1, 9) = "") Then .RemoveItem .Rows - 1
            If .TextMatrix(.Rows - 1, 1) = "" Or .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 4) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Spare Parts Received Units"
        End If
    End With
    
    For i = 1 To flxParts.Rows - 1
        For j = 1 To flxParts.Rows - 1
            If i <> j Then
                If flxParts.TextMatrix(i, 1) = flxParts.TextMatrix(j, 1) Then
                    'MsgBox "Part Number " & txtPartNo.Text & " already exist in row " & i, vbInformation + vbOKOnly, "Duplicate Part Number"
                    vMsg = vMsg & vbCrLf & "Part Number " & flxParts.TextMatrix(i, 1) & " already exist in row " & i
                End If
            End If
        Next j
    Next i
    
    
    ValidateSPRec = vMsg
End Function

Public Sub ShowRecord(MyRef As String)
    Dim rsSPrr As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SPRECEIVE.* From SPRECEIVE WHERE (((SPRECEIVE.PACKINGNO)='" & MyRef & "'));"

    'mySQL
    Qry = "SELECT SPRECEIVE.* From SPRECEIVE WHERE SPRECEIVE.PACKINGNO='" & MyRef & "';"

    With rsSPrr
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Call ClearEntries
            .MoveFirst
            cboShipper.Text = IIf(IsNull(.Fields("Shipper")), "", .Fields("Shipper"))
            
'            If InStr(.Fields("Supplier"), "MOTO-R") > 0 Or InStr(.Fields("Supplier"), "MOTOR-HUB") > 0 Then
'                ChkrC.Value = 1
'                cboSupplier.Text = IIf(IsNull(.Fields("Supplier")), "", .Fields("Supplier"))
'            Else
'                ChkrC.Value = 0
'                lblSupPO.Caption = IIf(IsNull(.Fields("Supplier")), "", .Fields("Supplier"))
'            End If
'
                cboBranch.Text = !Supplier
                txtWayBill.Text = IIf(IsNull(.Fields("WayBillNo")), "", .Fields("WayBillNO"))
                txtPackingNo.Text = !packingNO
                txtPLDate.Text = IIf(IsNull(!PACKDATE), "", Format(!PACKDATE, "mm/dd/yyyy"))
                txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
                txtPrnCopy.Text = !PRINTCOPY
                txtChkBy.Text = IIf(IsNull(!CHECKEDBY), "", !CHECKEDBY)
                txtRRno.Text = !SPRRNO
                txtDateRec.Text = Format(!DATEREC, "mm/dd/yyyy")
                
                
            
        End If
        .Close
    End With
                '                  0                    1                     2              3                     4             5                  6                 7
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.RECEIVE, SPAREPARTS.LOCATION, SPITR.REMARKS, SPAREPARTS.QTYONHAND, SPITR.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & MyRef & "'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.RECEIVE, SPAREPARTS.LOCATION, SPITR.REMARKS, SPAREPARTS.QTYONHAND, SPITR.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='" & MyRef & "';"

    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            .MoveFirst
            Do While Not .EOF       '       1                     2                                                 3                     4                                                5                                                  6                     7                   8
                flxParts.AddItem vbTab & .Fields(0) & vbTab & .Fields(1) & vbTab & IIf(IsNull(.Fields(2)), "", .Fields(2)) & vbTab & .Fields(3) & vbTab & IIf(IsNull(.Fields(4)), "", .Fields(4)) & vbTab & IIf(IsNull(.Fields(5)), "", .Fields(5)) & vbTab & .Fields(6) & vbTab & .Fields(7)
                .MoveNext
            Loop
        End If
        .Close
    End With
    Call ComputeTotalQty
    ViewOnly = True
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
        If BtnIndx = 6 Then
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 4 Then
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
            BtnIndx = 4
            cmdTrans(4).Top = -375
            cmdTrans(4).Left = 0
            cmdTrans(4).Visible = True
            cmdTrans(5).Top = -375
            cmdTrans(5).Left = cmdTrans(4).Width
            cmdTrans(5).Visible = True
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
        If BtnIndx = 3 Then
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 3
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub txtLoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtLoc.Text = Empty
        flxParts.TextMatrix(ActRow, 5) = txtLoc.Tag
        txtLoc.Tag = "skiplostfocus"
        Call txtLoc_LostFocus
        txtLoc.Visible = False
        flxParts.SetFocus
        flxParts.Col = 5
    End If
End Sub

Private Sub txtLoc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call txtLoc_LostFocus
        flxParts.SetFocus
        flxParts.Col = 6
    End If
End Sub

Private Sub txtLoc_LostFocus()
    If txtLoc.Tag = "skiplostfocus" Then
        txtLoc.Tag = ""
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 5) = txtLoc.Text
    txtLoc.Visible = False
End Sub

Private Sub txtPackingNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPackingNo_LostFocus()
    Dim Qry As String
    
    If AddNew Then
        'Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
              "WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
        'mySQL
        Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
              "WHERE SPRECEIVE.PACKINGNO='" & txtPackingNo.Text & "' and SPRECEIVE.SUPPLIER='" & cboBranch.Text & "';"
        
        If CtrlNumExist(Qry) Then
            MsgBox "Packing No. already exist!", vbExclamation + vbOKOnly, "Validation Error"
        End If
    End If
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
        .Visible = True
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
            .Col = 0: .ColSel = 5
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            flxParts.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 3) 'Description
            flxParts.TextMatrix(ActRow, 8) = .TextMatrix(.Row, 1) 'SparePart ID
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            flxParts.TextMatrix(ActRow, 2) = ""     'Description
            .Visible = False
        End If
    End With
    For i = 3 To 9      'Clear old values of flxParts
        flxParts.TextMatrix(ActRow, i) = ""
    Next i

End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            flxSrch.Tag = "skiplostfocus"
            flxSrch.Visible = True
            flxSrch.SetFocus
        Case vbKeyEscape
            txtPartNo.Text = Empty
            If flxSrch.Tag <> "" Then
                Tmp = Split(flxSrch.Tag, "|")
                If UBound(Tmp) > 1 Then
                    For i = 1 To 9
                        flxParts.TextMatrix(ActRow, i) = Tmp(i - 1)
                    Next i
                End If
            End If
            flxSrch.Tag = "skiplostfocus"
            flxParts.SetFocus
            flxParts.Col = 1
            txtPartNo.Visible = False
            flxSrch.Visible = False
        Case vbKeyDelete
            PrtSrch = True
        Case vbKeyBack
            PrtSrch = True
        Case Else
            FromFlx = False
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        flxSrch.Visible = False
        flxParts.SetFocus
        flxParts.Col = 3
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    Dim Ans As Integer
    Dim Qry As String
    Dim i As Long, j As Long
    
    If flxSrch.Tag = "skiplostfocus" Then
        flxSrch.Tag = ""
        Exit Sub
    End If
    
    'Validate existing PartNo
    For i = 1 To flxParts.Rows - 1
        If i <> ActRow Then
            If flxParts.TextMatrix(i, 1) = txtPartNo.Text Then
                flxParts.TextMatrix(ActRow, 1) = txtPartNo.Text
                txtPartNo.Visible = False
                flxSrch.Visible = False
                Ans = MsgBox("Part Number " & txtPartNo.Text & " already exist in row " & i, vbInformation + vbOKOnly, "Duplicate Part Number")
                For j = 1 To flxParts.Cols - 1
                    flxParts.TextMatrix(ActRow, j) = ""
                Next j
                flxParts.Row = i
                flxParts.Col = 1
                flxParts.SetFocus
                txtPartNo.Tag = "PN exist"
                Exit Sub
            End If
        End If
    Next i
    
    If txtPartNo.Text = "" Then
        For i = 1 To 9
            flxParts.TextMatrix(ActRow, i) = ""
        Next i
        txtPartNo.Visible = False
        flxSrch.Visible = False
        Exit Sub
    End If
        
        
        
    If flxParts.TextMatrix(ActRow, 2) = "" And txtPartNo.Text <> "" Then
        'Find Part in Current MDB if exist then Update Sparepart MDB
        'If CtrlNumExist("SELECT SPAREPARTS.* From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & txtPartNo.Text & "'));") Then
        'mySQL
        If CtrlNumExist("SELECT SPAREPARTS.* From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & txtPartNo.Text & "';") Then
            
            flxParts.TextMatrix(ActRow, 2) = GetSPDesc(txtPartNo.Text)
            flxParts.TextMatrix(ActRow, 9) = "New in SP only"
            Call txtPartNo_LostFocus
            Exit Sub
        End If
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please inform your Manager about this Part Number", vbQuestion + vbYesNo, "New Spare Part"
        flxParts.TextMatrix(ActRow, 1) = "" 'txtPartNo.Text
        txtPartNo.Visible = False
        flxSrch.Visible = False
        'If Ans = vbYes Then SPnew.Show vbModal
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 1) = txtPartNo.Text
    txtPartNo.Visible = False
    flxSrch.Visible = False
    Call GetOrderNo(ActRow)
   End Sub
Function GetSPDesc(myPN As String) As String
    Dim rsDesc As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SPAREPARTS.DESCRIPTION From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & myPN & "'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.DESCRIPTION From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & myPN & "';"
    
    With rsDesc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetSPDesc = IIf(IsNull(!Description), "", !Description)
        Else
            GetSPDesc = ""
        End If
        .Close
    End With
End Function

Private Sub GetOrderNo(MyRow As Integer)
    Dim rsSP As New ADODB.Recordset
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String, Tmp As String, SPtype As String
    Dim MyOrd As Integer, MyPOrd As Long
    
    Tmp = RTrim(flxParts.TextMatrix(MyRow, 1))
    If Tmp = "" Then Exit Sub
    'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPAREPARTS.QTYONHAND " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER) = '" & Tmp & "'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPAREPARTS.QTYONHAND " & _
          "From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER = '" & Tmp & "';"
    
'old    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPAREPARTS.QTYONHAND " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER) = '" & Tmp & "'));"
    With rsSP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst  'SP exist in MDB
            flxParts.TextMatrix(ActRow, 5) = IIf(IsNull(!Location), "", !Location)
            flxParts.TextMatrix(ActRow, 7) = !QtyOnHand
            flxParts.TextMatrix(ActRow, 8) = !SparepartID
        Else    'PartNumber new in MDB
            flxParts.TextMatrix(MyRow, 6) = "New"
            flxParts.TextMatrix(MyRow, 9) = flxNewSP.Rows
            flxParts.TextMatrix(MyRow, 9) = flxNewSP.Rows
            If cboBranch.ListIndex >= 0 Then
                'SPtype = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE (((SUPPLIER.SUPPLIERID)=" & Val(cboSupplier.ItemData(cboSupplier.ListIndex)) & "));"
                'mySQL
                SPtype = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE SUPPLIER.SUPPLIERID=" & Val(cboBranch.ItemData(cboBranch.ListIndex)) & ";"
                
                SPtype = GetFieldValue(SPtype)
            End If
            flxNewSP.AddItem flxParts.TextMatrix(ActRow, 1) & vbTab & rsParts!Cost & vbTab & rsParts!RETAILPRICE & vbTab & vbTab & IIf(IsNull(rsParts!SUPERSEDING), "", rsParts!SUPERSEDING) & vbTab & IIf(SPtype <> "", SPtype, "REPLACEMENT") & vbTab & IIf(IsNull(rsParts!ApplicModel), "", rsParts!ApplicModel)
        End If
        .Close
    End With
    If flxParts.TextMatrix(MyRow, 8) = "" Then
        cboOrderNo.Clear
        Exit Sub
    End If
    'Qry = "SELECT SPITR.REFERENCE, SPITR.BALANCE FROM SPITR " & _
          "WHERE (((SPITR.BALANCE)>0) AND ((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & flxParts.TextMatrix(MyRow, 8) & "));"
    'mySQL
    Qry = "SELECT SPITR.REFERENCE, SPITR.BALANCE FROM SPITR " & _
          "WHERE SPITR.BALANCE>0 AND SPITR.TRANSTYPE='ORDER' AND SPITR.SPAREPARTID=" & flxParts.TextMatrix(MyRow, 8) & ";"
    
    With rsOrder
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboOrderNo.Clear
            Do While Not .EOF
                MyOrd = GetOrderBalance(!REFERENCE, flxParts.TextMatrix(MyRow, 8))
                MyPOrd = GetOrderRECTEMP(!REFERENCE, flxParts.TextMatrix(MyRow, 8))
                If MyOrd - MyPOrd > 0 Then cboOrderNo.AddItem !REFERENCE
                .MoveNext
            Loop
        Else
            cboOrderNo.Clear
        End If
        .Close
    End With
End Sub


Private Sub txtPLDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890/"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxParts.SetFocus
        flxParts.Col = 1
        flxParts.Row = 1
    End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtQty.Text = Empty
        flxParts.TextMatrix(ActRow, 4) = txtQty.Tag
        txtQty.Tag = "skiplostfocus"
        txtQty.Visible = False
        flxParts.SetFocus
        flxParts.Col = 4
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Call txtQty_LostFocus
        flxParts.SetFocus
        flxParts.Col = 5
    End If
End Sub

Private Sub txtQty_LostFocus()
    If txtQty.Tag = "skiplostfocus" Then
        txtQty.Tag = ""
        Exit Sub
    End If
    If flxParts.TextMatrix(ActRow, 3) <> "" And Val(flxParts.TextMatrix(ActRow, 11)) > 0 Then
        If Val(txtQty.Text) > Val(flxParts.TextMatrix(ActRow, 11)) Then
            MsgBox "Qty. Received must not be greater than Qty. Ordered!" & vbCrLf & "Please record all excess quantity in separate file.", vbExclamation + vbOKOnly, "Validation Error"
            flxParts.TextMatrix(ActRow, 4) = flxParts.TextMatrix(ActRow, 11)
            txtQty.Visible = False
            Call ComputeTotalQty
            Exit Sub
        End If
    End If
       
    flxParts.TextMatrix(ActRow, 4) = txtQty.Text
    txtQty.Visible = False
    Call ComputeTotalQty
End Sub


Private Sub txtRem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtRem.Text = Empty
        flxParts.TextMatrix(ActRow, 6) = txtRem.Tag
        txtRem.Tag = "skiplostfocus"
        Call txtRem_LostFocus
        txtRem.Visible = False
        flxParts.SetFocus
        flxParts.Col = 6
    End If
End Sub

Private Sub txtRem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtRem_LostFocus
        flxParts.SetFocus
        flxParts.Col = 0
    End If
End Sub

Private Sub txtRem_LostFocus()
    If txtRem.Tag = "skiplostfocus" Then
        txtRem.Tag = ""
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 6) = txtRem.Text
    txtRem.Visible = False
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

Private Sub txtWayBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Function GetLocation(MyPartNo As String) As String
    Dim rsLoc As New ADODB.Recordset
    Dim Qry As String
    
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
