VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SPincoming 
   Caption         =   "Incoming Spare Parts"
   ClientHeight    =   8835
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11505
   ControlBox      =   0   'False
   Icon            =   "SPincomingNEW.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   11505
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtBrnSPOID 
      Height          =   285
      Left            =   1680
      TabIndex        =   51
      Top             =   9000
      Visible         =   0   'False
      Width           =   735
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
      Left            =   12360
      MaxLength       =   15
      TabIndex        =   49
      Tag             =   "XYF"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdImpVan 
      Caption         =   "Import from Van"
      Height          =   255
      Left            =   6120
      TabIndex        =   48
      Tag             =   "xyf"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdExportVan 
      Caption         =   "Export to Mobile Van"
      Height          =   255
      Left            =   7800
      TabIndex        =   47
      Tag             =   "xyf"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdImpStkTrns 
      Caption         =   "MRC2 Stock Transfer"
      Height          =   255
      Left            =   9495
      TabIndex        =   46
      Tag             =   "xyf"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdlgVan 
      Left            =   1920
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   7440
   End
   Begin VB.ListBox lstDelParts 
      Height          =   450
      ItemData        =   "SPincomingNEW.frx":000C
      Left            =   10320
      List            =   "SPincomingNEW.frx":000E
      TabIndex        =   44
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   7920
      TabIndex        =   43
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      TabIndex        =   38
      Tag             =   "XYF"
      Top             =   7320
      Width           =   1335
   End
   Begin VB.TextBox txtLoc 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   37
      Top             =   9000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxNewSP 
      Height          =   735
      Left            =   6960
      TabIndex        =   36
      Top             =   9840
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
      TabIndex        =   35
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboOrderNo 
      Height          =   315
      Left            =   7440
      TabIndex        =   34
      Top             =   8880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtDateRec 
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
      Left            =   6000
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   8280
      Width           =   2775
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
      TabIndex        =   25
      Top             =   8880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   3240
      TabIndex        =   24
      Top             =   8880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSrch 
      Height          =   1695
      Left            =   3360
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   9240
      Width           =   135
   End
   Begin VB.TextBox txtRem 
      Height          =   285
      Left            =   5040
      TabIndex        =   21
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Height          =   285
      Left            =   1680
      MaxLength       =   250
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   7920
      Width           =   7095
   End
   Begin MSFlexGridLib.MSFlexGrid flxParts 
      Height          =   4935
      Left            =   480
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   2400
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   "    |<PART NUMBER   |<DESCRIPTION |<ORDER NO. |^QTY.   |^GOOD|<REMARKS|^B.O.   "
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
            Picture         =   "SPincomingNEW.frx":0010
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPincomingNEW.frx":0EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPincomingNEW.frx":1DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPincomingNEW.frx":2C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPincomingNEW.frx":3B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPincomingNEW.frx":4A52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   480
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   720
      Width           =   10695
      Begin VB.TextBox txtPOno 
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   50
         Tag             =   "XYF"
         Top             =   896
         Width           =   1575
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
         Left            =   1800
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   538
         Width           =   4215
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
         Left            =   1800
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   200
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
         Top             =   208
         Width           =   1695
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
         TabIndex        =   13
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
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   230
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "P.O. NO."
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
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   926
         Width           =   1575
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
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   576
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "RECEIVED FROM"
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
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   238
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   26
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
      Image           =   "SPincomingNEW.frx":592C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   27
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
      Image           =   "SPincomingNEW.frx":5EC6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   28
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
      Image           =   "SPincomingNEW.frx":6460
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   7200
      TabIndex        =   29
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
      Image           =   "SPincomingNEW.frx":69FA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   9000
      TabIndex        =   30
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
      Image           =   "SPincomingNEW.frx":6F94
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10800
      TabIndex        =   31
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
      Image           =   "SPincomingNEW.frx":752E
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtChkBy 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   8280
      Width           =   2775
   End
   Begin lvButton.lvButtons_H cmdLoadPrtl 
      Height          =   375
      Left            =   10560
      TabIndex        =   40
      Tag             =   "XYF"
      ToolTipText     =   "Load Partial Saved"
      Top             =   7320
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
      Image           =   "SPincomingNEW.frx":7AC8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrtlSave 
      Height          =   375
      Left            =   9960
      TabIndex        =   41
      Tag             =   "XYF"
      ToolTipText     =   "Partial Saving"
      Top             =   7320
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
      Image           =   "SPincomingNEW.frx":8062
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   42
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
      Image           =   "SPincomingNEW.frx":85FC
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   1095
      Left            =   240
      TabIndex        =   45
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
      TabIndex        =   39
      Tag             =   "XYF"
      Top             =   7335
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   4680
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   8280
      Width           =   1335
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
      Left            =   480
      TabIndex        =   32
      Tag             =   "XYF"
      Top             =   8295
      Width           =   1095
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   0
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11745
      m_MDI_Height    =   9405
      OriginalBorder  =   10
      OriginalCaption =   31
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
      Left            =   480
      TabIndex        =   19
      Tag             =   "XYF"
      Top             =   2160
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
      TabIndex        =   17
      Tag             =   "TRF"
      Top             =   450
      Width           =   2820
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO-R REPUBLIC"
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
      TabIndex        =   15
      Tag             =   "TLF"
      Top             =   450
      Width           =   2415
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
      Left            =   720
      TabIndex        =   14
      Tag             =   "XYF"
      Top             =   7935
      Width           =   795
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO-R REPUBLIC"
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
      TabIndex        =   16
      Tag             =   "TLF"
      Top             =   475
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   8145
      TabIndex        =   18
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
      TabIndex        =   20
      Tag             =   "XYF"
      Top             =   2190
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
Attribute VB_Name = "SPincoming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset

Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean, OrdSrch As Boolean
Dim ActRow As Integer, ActCol As Integer, Appmodel As String
Dim BtnIndx As Integer, FrstBtnIndx As Integer
Dim ViewOnly As Boolean, AddNew As Boolean, ComboSrch As Boolean, LoadFromWarehouse As Boolean

Private Sub BtnSaveCancel(Bln As Boolean)
    Dim i As Integer
    
    If Bln Then
        For i = 0 To 4
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(5).Left = 0
        cmdTrans(6).Left = cmdTrans(5).Width
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
        For i = 0 To 4
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(5).Left = 0
        cmdTrans(6).Left = cmdTrans(5).Width
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
    End If
End Sub

Private Sub ClearEntries()
    cboSupplier.Text = ""
    cboShipper.Text = ""
    txtWayBill.Text = ""
    txtPOno.Text = ""
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

Private Sub LockData(Bln As Boolean)
    Frame1.Enabled = Not Bln
    'flxParts.Enabled = Not Bln
    txtRemarks.Locked = Bln
    txtChkBy.Locked = Bln
    txtDateRec.Locked = Bln
    ViewOnly = Bln
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
    Dim MyOrder As Integer
    
    If cboOrderNo.Tag = "skiplostfocus" Then
        cboOrderNo.Tag = ""
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 3) = cboOrderNo.Text
    cboOrderNo.Visible = False
    If cboOrderNo.ListIndex >= 0 And flxParts.TextMatrix(ActRow, 8) <> "" Then  'And flxParts.TextMatrix(ActRow, 4) = ""
        MyOrder = GetOrderBalance(ActRow)
        If MyOrder > 0 Then
            flxParts.TextMatrix(ActRow, 4) = MyOrder
            flxParts.TextMatrix(ActRow, 5) = MyOrder
            flxParts.TextMatrix(ActRow, 14) = MyOrder
        Else
            flxParts.TextMatrix(ActRow, 4) = 0
            flxParts.TextMatrix(ActRow, 11) = 0
        End If
        Call ComputeTotalQty
    End If
End Sub

Private Sub cboShipper_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboSupplier_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If ComboSrch Then
        ComboSrch = False
        Exit Sub
    End If
    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                ComboSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
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
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then ComboSrch = True
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub

Private Sub cmdExportVan_Click()
    Dim myFile As String, MyExt As String
    Dim i As Integer
    
    On Error GoTo LocalErr
    If UCase$(SysBranch) = "MRC VAN-2 [LHC403]" Then
        MyExt = ".VN1"
    ElseIf UCase$(SysBranch) = "MRC VAN-56 [AAH1233]" Then
        MyExt = ".VN2"
    ElseIf UCase$(SysBranch) = "MRC VAN-55 [AAH1250]" Then
        MyExt = ".VN3"
    Else
        MsgBox "Unable to Identify name of VAN!" & vbCrLf & "Please contact your System Administrator!", vbCritical + vbOKOnly, "Warning"
        Exit Sub
    End If
    
    myFile = txtPackingNo.Text & MyExt
    With cdlgVan
        .CancelError = True
        .Filter = "Van Files (*" & MyExt & ")|*." & MyExt
        .InitDir = App.Path
        .filename = myFile
        .ShowSave
        myFile = .filename
    End With
    If myFile <> "" Then    ' And cdlgVan.FileTitle = txtPackingNo.Text & MyExt Then
        Open myFile For Output As #1
        Write #1, cboSupplier.Text, cboShipper.Text, txtPOno.Text, txtPackingNo.Text, txtPLDate.Text, txtRemarks.Text, txtChkBy.Text, txtDateRec.Text, ""
        With flxParts
            For i = 1 To .Rows - 1
                Write #1, .TextMatrix(i, 0), .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 8) ' Write comma-delimited data.
            Next i
        End With
        Close #1
    Else
        MsgBox "Invalid File Name!", vbExclamation + vbOKOnly, "Warning"
    End If
    
    
    Exit Sub
LocalErr:
    'Cancelled Dialog box
End Sub

Private Sub cmdImpStkTrns_Click()
    Dim rsST As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim MyRef As String, Qry As String, y As String, x As Integer
    
    On Error GoTo LocalErr
    conMRC2.CursorLocation = adUseClient
    conMRC2.ConnectionTimeout = 60
    conMRC2.Open "dsn=SQLMRC2TGM"
    
    MyRef = InputBox("Enter Stock Transfer Reference No.", "Stock Transfer")
    If MyRef = "" Then
        conMRC2.Close
        Exit Sub
    End If
    
    Qry = "SELECT * From stocktransfer Where stocktransfer.RECBRANCH = '" & SysBranch & "' AND " & _
          "stocktransfer.TRANSFERID = " & Val(MyRef) & " AND stocktransfer.`STATUS` = 'FINAL';"
    With rsST
        .Open Qry, conMRC2, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Call ClearEntries
            .MoveFirst
            cboShipper.Text = "IMPORT STOCK TRANSFER " & Format(Date, "MM/DD/YY")
            cboSupplier.Text = "MRC 2 TAGUM"
            txtPackingNo.Text = !TRANSFERID & "M2"
            txtPLDate.Text = Format(Date, "mm/dd/yyyy")
            txtPOno.Text = !PoNo
            Label1(2).Caption = "P.O. NO."
            txtRemarks.Text = "TR#" & !TRANSFERID & " " & Format(!TRANSDATE, "MM/DD/YYYY") & " Php" & Format(!TotalCost, "#,##0.00")
            cboShipper.Locked = True
            cboSupplier.Locked = True
            txtPackingNo.Locked = True
            txtRemarks.Locked = True
        Else
            MsgBox "Stock Transfer Ref.# " & MyRef & " not found!", vbExclamation + vbOKOnly, SysBranch
            .Close
            conMRC2.Close
            Exit Sub
        End If
        .Close
    End With
    'Check if transfer already exist in Van DB
    Qry = "SELECT spreceive.PACKINGNO From spreceive Where spreceive.SUPPLIER = 'MRC 2 TAGUM' AND SPRECEIVE.REMARKS LIKE 'TR#" & MyRef & "%' AND " & _
          "spreceive.SHIPPER LIKE 'IMPORT STOCK TRANSFER%'"
    y = GetFieldValue(Qry)
    If y <> "" Then
        MsgBox "Stock Transfer Ref#:" & MyRef & " already received!" & vbCrLf & "Please check Packing No. " & y & " for details.", vbExclamation + vbOKOnly, "Pack#:" & y
        conMRC2.Close
        Exit Sub
    End If
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.OUT, SPAREPARTS.LOCATION, SPITR.REMARKS, SPAREPARTS.QTYONHAND, SPITR.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='TR#" & MyRef & "'  AND SPITR.REMARKS = 'TRANSFER' ORDER BY SPITR.`ORDER`;"

    With rsITR
        .Open Qry, conMRC2, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            .MoveFirst
            Do While Not .EOF       '       1                     2                                                 3                     4                                                5                                                  6                     7                   8
                flxParts.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(IsNull(!PACKINGLIST), "", !PACKINGLIST) & vbTab & !OUT & vbTab & !OUT & vbTab & vbTab & "0" & vbTab & !SPAREPARTID
                .MoveNext
            Loop
        End If
        .Close
    End With
    conMRC2.Close
    Call ComputeTotalQty
    
    'CHECK IF SPAREPART ID IS UP TO DATE
    With flxParts
        For x = 1 To .Rows - 1
            Qry = "SELECT spareparts.PARTNUMBER From spareparts Where spareparts.SPAREPARTID = " & Val(.TextMatrix(x, 8)) & ";"
            y = GetFieldValue(Qry)
            If y = "" Then
                'Partnumbers not updated
                MsgBox "Part Number " & .TextMatrix(x, 1) & " is not found in your Database!" & vbCrLf & "Please UPDATE your database or call System Administrator for assistance.", vbExclamation + vbOKOnly, "Update Required"
                Exit Sub
            ElseIf y <> .TextMatrix(x, 1) Then
                'Partnumber not equal
                MsgBox "Part Number " & .TextMatrix(x, 1) & " has conflict in your Database!" & vbCrLf & "Please call System Administrator for assistance.", vbExclamation + vbOKOnly, "Update Required"
                Exit Sub
            End If
        Next x
    End With
    'Allow saving of received parts
    Call BtnSaveCancel(True)
    txtChkBy.Text = GetSettings("STOCK CLERK")
    txtDateRec.Text = Date
    If SysBranch = "BINOS BABAK SAMAL" Then
        ViewOnly = False
        AddNew = True
        Frame1.Enabled = False
    Else
        ViewOnly = True
        AddNew = True
        Frame1.Enabled = True
    End If
    MsgBox "Stock Transfer No. " & MyRef & " successfully loaded!", vbInformation + vbOKOnly, "Done"
    Exit Sub

LocalErr:
    MsgBox "Error connecting to MRC2 server..." & vbCrLf & "Make sure you are connected to the Server before importing Stock Transfer", vbExclamation + vbOKOnly, "Connection Failed"
    If conMRC2.State = adStateOpen Then conMRC2.Close
End Sub

Private Sub cmdImpVan_Click()
    Dim myFile As String, MyExt As String
    Dim i As Integer
    Dim My1, My2, My3, My4, My5, My6, My7, My8, My9
    
    On Error GoTo LocalErr
    If UCase$(SysBranch) = "MOBILE VAN-2 [LHC403]" Then
        MyExt = ".VN1"
    ElseIf UCase$(SysBranch) = "MOBILE VAN-56 [AAH1233]" Then
        MyExt = ".VN2"
    ElseIf UCase$(SysBranch) = "MOBILE VAN-55 [AAH1250]" Then
        MyExt = ".VN3"
    Else
        MsgBox "Unable to Identify name of VAN!" & vbCrLf & "Please contact your System Administrator!", vbCritical + vbOKOnly, "Warning"
        Exit Sub
    End If
    
    With cdlgVan
        .CancelError = True
        .Filter = "(*" & MyExt & ")|*" & MyExt
        .ShowOpen
        myFile = .filename
    End With
    If myFile <> "" Then
        i = 1
        flxParts.Rows = 1
        Open myFile For Input As #1
        Do While Not EOF(1)   ' Loop until end of file.
            Input #1, My1, My2, My3, My4, My5, My6, My7, My8, My9
            If i = 1 Then
                cboSupplier.Text = My1
                cboShipper.Text = My2
                txtPOno.Text = My3
                txtPackingNo.Text = My4
                txtPLDate.Text = My5
                txtRemarks.Text = My6
                txtChkBy.Text = My7
                txtDateRec.Text = My8
            Else
                '                 0             1             2             3             4             5             6             7
                flxParts.AddItem My1 & vbTab & My2 & vbTab & My3 & vbTab & My4 & vbTab & My5 & vbTab & My6 & vbTab & My7 & vbTab & My8 & vbTab & My9
            End If
            i = i + 1
        Loop
        Close #1
        Call ComputeTotalQty
        Call LockData(True)
    Else
        MsgBox "Invalid File Name!", vbExclamation + vbOKOnly, "Warning"
    End If
    
    Exit Sub
LocalErr:
    'MsgBox Err.Description, vbExclamation + vbOKOnly, "Import VAN"
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
    Dim Qry As String, myMsg As String, MyAcc As String
    Dim Ans As Integer, i As Integer
    
    Select Case Index
        Case 0      '"new"
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=21));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=21;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            Call BtnSaveCancel(True)
            'BtnIndx = 3
            'tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            txtChkBy.Text = GetSettings("STOCK CLERK")
            txtDateRec.Text = Date
            ViewOnly = False
            AddNew = True
            LoadFromWarehouse = False
            cboSupplier.SetFocus
            If InStr(SysBranch, "MRC VAN") >= 1 Then
                cmdImpStkTrns.Visible = True
                cmdImpVan.Visible = False
            ElseIf InStr(SysBranch, "MOBILE") >= 1 Then
                cmdImpStkTrns.Visible = False
                cmdImpVan.Visible = True
                cmdImpVan.Left = cmdImpStkTrns.Left
            ElseIf SysBranch = "BINOS BABAK SAMAL" Then
                cmdImpStkTrns.Visible = True
            Else
                cmdImpStkTrns.Visible = False
                cmdImpVan.Visible = False
            End If
            
        Case 1      'edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=21));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=21;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
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
            ViewOnly = False
            AddNew = False
            'For i = 1 To flxParts.Rows - 1
            '    flxParts.TextMatrix(i, 10) = flxParts.TextMatrix(i, 8)
            'Next i
            Call BtnSaveCancel(True)
            flxNewSP.Rows = 0
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
            If InStr(SysBranch, "MRC VAN") >= 1 Then
                cmdImpStkTrns.Visible = True
                If InStr(cboShipper.Text, "IMPORT STOCK TRANSFER") > 0 Then
                    cmdExportVan.Visible = True
                Else
                    cmdExportVan.Visible = False
                End If
            ElseIf InStr(SysBranch, "MOBILE") >= 1 Then
                cmdImpVan.Visible = False
            ElseIf SysBranch = "BINOS BABAK SAMAL" Then
                cmdImpStkTrns.Visible = True
            Else
                cmdImpStkTrns.Visible = False
                cmdExportVan.Visible = False
            End If

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
            Ans = MsgBox("Do you want to Save Receiving Receipt?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'BtnIndx = 5
            'tmrFloatRev.Enabled = True
            'If AddNew Then
                Call SaveSpareParts
            'Else
            '    Call SaveEditedSP
            'End If
            Call LockData(True)
            MsgBox "Received Spare Parts successfully saved!", vbInformation + vbOKOnly, "Save"
            Call BtnSaveCancel(False)
        
        Case 2      '"view"
            FrmSummary.mySrc = "Received Spare Parts"
            FrmSummary.Show
        
'            SPRecSummary.Show
            Unload Me
        Case 3      '"print"
            If txtRRno.Text = "" Then
                MsgBox "No Available Receiving Receipt to Print", vbExclamation + vbOKOnly, "Print"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Print Receiving Receipt?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            
            If LoadFromWarehouse Then
                PRNmc.MySource = "RECEIVING RECEIPT"
                If Val(txtPrnCopy.Text) = 0 Then
                    'mySQL
                    PRNmc.QryPrn = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spitr.POREF, spitr.BALANCE, spitr.RECEIVE, spitr.REMARKS, spitr.ORDERBAL " & _
                        "From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) Where spitr.REFERENCE = '" & txtPackingNo.Text & "' AND spitr.SOURCE = '" & cboSupplier.Text & "' AND spitr.TRANSTYPE = 'RECTEMP';"
                Else
                    PRNmc.QryPrn = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spitr.POREF, spitr.BALANCE, spitr.RECEIVE, spitr.REMARKS, spitr.ORDERBAL " & _
                        "From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) Where spitr.REFERENCE = '" & txtPackingNo.Text & "' AND spitr.SOURCE = '" & cboSupplier.Text & "' AND spitr.TRANSTYPE = 'RECEIVE';"
                End If
                PRNmc.ParamValue1 = cboSupplier.Text
                PRNmc.ParamValue2 = cboShipper.Text
                PRNmc.ParamValue3 = txtPackingNo.Text
                PRNmc.ParamValue4 = txtPLDate.Text
                PRNmc.ParamValue5 = txtChkBy.Text
                PRNmc.ParamValue6 = txtDateRec.Text
                PRNmc.ParamValue7 = SysBranch
                PRNmc.ParamValue8 = SysAddress
                PRNmc.ParamValue9 = SysTel
                PRNmc.Show
            Else
                PRNmc.MySource = "SPRECEIVE"
                'mySQL
                PRNmc.QryPrn = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.RECEIVE, SPAREPARTS.LOCATION, SPITR.REMARKS " & _
                               "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='" & txtPackingNo.Text & "';"
                
                PRNmc.ParamValue1 = cboSupplier.Text
                PRNmc.ParamValue2 = cboShipper.Text
                PRNmc.ParamValue3 = txtPackingNo.Text
                PRNmc.ParamValue4 = txtPLDate.Text
                PRNmc.ParamValue5 = txtChkBy.Text
                PRNmc.ParamValue6 = txtDateRec.Text
                PRNmc.Show
            End If
        Case 4      '"quit"
            MotoRep.tbMain.Visible = True
            Unload Me
    End Select

End Sub

Public Sub FinalSPrec()
    Dim i As Integer, BckOrd As Integer, oBal As Integer
    
    If Val(txtPrnCopy.Text) > 0 Then Exit Sub
    
    LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'RECEIVE' " & _
                   "WHERE SPITR.REFERENCE='" & txtPackingNo.Text & "' AND SPITR.TRANSTYPE='RECTEMP';"
    
    'Update Print Copy
    txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
        
    LYRcon.Execute "UPDATE SPRECEIVE SET SPRECEIVE.PRINTCOPY = " & Val(txtPrnCopy.Text) & _
                   " WHERE SPRECEIVE.PACKINGNO='" & txtPackingNo.Text & "';"
    
    'Update Qty on Hand of all received SpareParts
    With flxParts
        For i = 1 To .Rows - 1
                        
            If LoadFromWarehouse Then
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+ " & Val(.TextMatrix(i, 4)) & ", SPAREPARTS.REFERENCE = '" & txtPackingNo.Text & "', SPAREPARTS.BACKORDER = " & BckOrd & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
            Else
                        
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+ " & Val(.TextMatrix(i, 4)) & ", SPAREPARTS.REFERENCE = '" & txtPackingNo.Text & "', SPAREPARTS.BACKORDER = " & BckOrd & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
            End If
        Next i
    End With
    
    MsgBox "Spare Parts Qty-On-Hand successfully updated!"
    If InStr(SysBranch, "MRC VAN") >= 1 Then
        cmdImpStkTrns.Visible = True
        If InStr(cboShipper.Text, "IMPORT STOCK TRANSFER") > 0 Then
            cmdExportVan.Visible = True
        Else
            cmdExportVan.Visible = False
        End If
    ElseIf SysBranch = "BINOS BABAK SAMAL" Then
        cmdImpStkTrns.Visible = True
    Else
        cmdImpStkTrns.Visible = False
        cmdExportVan.Visible = False
    End If

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
            'flxStock.AddItem .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & vbTab & .TextMatrix(i, 2) & vbTab & vbTab & xBal   '.TextMatrix(i, 2)
            'Acquire Packing list for P.O.
            MyPO = .TextMatrix(i, 1)
            '***access
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
                    'old flxStock.AddItem Format(rsRec!TransDate, "mm/dd/yyyy") & vbTab & MyPO & vbTab & rsRec!Reference & vbTab & vbTab & rsRec!Receive & vbTab & xBal
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
    If LoadFromWarehouse And ActCol <> 5 Then Exit Sub
    Select Case ActCol
        Case 1  'PartNumber
            With txtPartNo
                If flxParts.TextMatrix(ActRow, 1) <> "" Then flxSrch.Tag = flxParts.TextMatrix(ActRow, 1) & "|" & flxParts.TextMatrix(ActRow, 2) & "|" & flxParts.TextMatrix(ActRow, 3) & "|" & flxParts.TextMatrix(ActRow, 4) & "|" & flxParts.TextMatrix(ActRow, 5) & "|" & flxParts.TextMatrix(ActRow, 6) & "|" & flxParts.TextMatrix(ActRow, 7) & "|" & flxParts.TextMatrix(ActRow, 8) & "|" & flxParts.TextMatrix(ActRow, 9) & "|" & flxParts.TextMatrix(ActRow, 10) & "|" & flxParts.TextMatrix(ActRow, 11) & "|" & flxParts.TextMatrix(ActRow, 12) & "|" & flxParts.TextMatrix(ActRow, 13) & "|" & flxParts.TextMatrix(ActRow, 14)
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
            'If SysBranch = "BINOS BABAK SAMAL" And InStr(cboShipper.Text, "IMPORT STOCK TRANSFER") >= 1 Then Exit Sub

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
        Case 5  'GOOD
            With txtLoc
                'If flxParts.TextMatrix(ActRow, 5) <> "" And flxParts.TextMatrix(ActRow, 8) <> "" Then Exit Sub
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
    If ViewOnly Or LoadFromWarehouse Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuRows)
    End If
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

Private Sub Form_Load()
    cmdA.Top = 15000
    Call LockData(True)
    
    Timer1.Enabled = True
    'cboMatch.AddItem "Whole Word"
    'cboMatch.AddItem "Any Part of Word"
    MotoRep.sBar.Panels(6).Text = "Spare Parts Packing List"
    flxParts.ToolTipText = "v15.7.11 wCst"
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
        'If AddNew Then
        '    If flxParts.TextMatrix(ActRow, 9) <> "" Then
        '        If flxNewSP.TextMatrix(Val(flxParts.TextMatrix(ActRow, 9)), 0) = flxParts.TextMatrix(ActRow, 1) Then flxNewSP.RemoveItem Val(flxParts.TextMatrix(ActRow, 9))
        '    End If
        'Else
        '    If flxParts.TextMatrix(ActRow, 10) <> "" Then
        '        lstDelParts.AddItem flxParts.TextMatrix(ActRow, 10)
        '    End If
        'End If
        If ActRow = 1 And flxParts.Rows = 2 Then
            flxParts.Rows = 1
            flxParts.Rows = 2
        Else
            flxParts.RemoveItem ActRow
        End If
        Call ComputeTotalQty
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
        
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width + 30)
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    With flxParts
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 380) * 0.2
        .ColWidth(2) = (.Width - 380) * 0.23
        .ColWidth(3) = (.Width - 380) * 0.13
        .ColWidth(4) = (.Width - 380) * 0.08
        .ColWidth(5) = (.Width - 380) * 0.1
        .ColWidth(6) = (.Width - 380) * 0.15
        .ColWidth(7) = (.Width - 380) * 0.08
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
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
    Dim Qry As String
    
    'Saving SPReceive
    '***Access
    'Qry = "SELECT SPRECEIVE.*, SPRECEIVE.SPRRNO From SPRECEIVE WHERE (((SPRECEIVE.SPRRNO)=" & txtRRno.Text & "));"
    
    'mySQL
    Qry = "SELECT SPRECEIVE.*, SPRECEIVE.SPRRNO From SPRECEIVE WHERE SPRECEIVE.SPRRNO=" & txtRRno.Text & ";"
    
    rsSPrec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSPrec.RecordCount > 0 Then
        '***Access
        'rsSPrec.MoveFirst
        'rsSPrec!DATEREC = txtDateRec.Text
        'rsSPrec!Supplier = cboSupplier.Text
        'rsSPrec!Shipper = cboShipper.Text
        'rsSPrec!WayBillno = txtWayBill.Text
        'rsSPrec!PACKINGNO = txtPackingNo.Text
        'If IsDate(txtPLDate.Text) Then rsSPrec!PACKDATE = txtPLDate.Text
        'rsSPrec!CHECKEDBY = txtChkBy.Text
        'rsSPrec!Remarks = txtRemarks.Text
        'rsSPrec!PrintCopy = 0
        'rsSPrec!UPDATEDBY = UserID
        'rsSPrec!LASTUPDATE = Date
        'rsSPrec.Update
        'rsSPrec.Close
        
        'mySQL
        If IsDate(txtPLDate.Text) Then myPACKDATE = txtPLDate.Text
        LYRcon.Execute "UPDATE spreceive SET DATEREC = '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', SUPPLIER = '" & cboSupplier.Text & "', " & _
            "SHIPPER = '" & cboShipper.Text & "', WAYBILLNO = '" & txtWayBill.Text & "', PACKINGNO = '" & txtPackingNo.Text & "', " & _
            "PACKDATE = '" & Format(DateValue(myPACKDATE), "yyyy-mm-dd") & "', CHECKEDBY = '" & txtChkBy.Text & "', " & _
            "REMARKS = '" & txtRemarks.Text & "', PRINTCOPY = 0, UPDATEDBY = " & UserID & ", " & _
            "LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' Where spreceive.SPRRNO = " & txtRRno.Text & ";"
    End If
    'Update SPITR reference if txtpackingno was change
    If txtRRno.Tag <> txtPackingNo.Text Then
        '***access
        'LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = '" & txtPackingNo.Text & "' " & _
                       "WHERE (((SPITR.REFERENCE)='" & txtRRno.Tag & "') AND ((SPITR.TRANSTYPE)='RECTEMP'));"
        'mySQL
        LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = '" & txtPackingNo.Text & "' " & _
                       "WHERE SPITR.REFERENCE='" & txtRRno.Tag & "' AND SPITR.TRANSTYPE='RECTEMP';"
    
    End If
    With flxParts
        For i = 1 To flxParts.Rows - 1
            If .TextMatrix(i, 10) = "" Then
                If .TextMatrix(i, 8) <> "" Then     'SP already available in MDB
                    If .TextMatrix(i, 9) = "New in SP only" Then
                        'Update SparePart Database
                        LYRcon.Execute "INSERT INTO SOURCESP ( PARTNUMBER, DESCRIPTION) " & _
                                      "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "');"
                End If
                Else        'New SP
                    Rnew = Val(.TextMatrix(i, 9))
                    '***Access
                    'LYRcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, REFERENCE, LOCATION, COST, STOREPRICE, BEGBAL, REMARKS, DATEBEGBAL, SUPERSEDINGPN, APPLICMODEL, UPDATEDBY, LASTUPDATE) " & _
                                   "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', '" & txtPackingNo.Text & "', '" & .TextMatrix(i, 5) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ", " & Val(.TextMatrix(i, 4)) & ", '" & flxNewSP.TextMatrix(Rnew, 5) & "', #" & Date & "#, '" & flxNewSP.TextMatrix(Rnew, 4) & "', '" & flxNewSP.TextMatrix(Rnew, 6) & "', " & UserID & ", #" & Date & "#);"
                    'mySQL
                    LYRcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, REFERENCE, LOCATION, COST, STOREPRICE, BEGBAL, REMARKS, DATEBEGBAL, SUPERSEDINGPN, APPLICMODEL, UPDATEDBY, LASTUPDATE) " & _
                                   "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', '" & txtPackingNo.Text & "', '" & .TextMatrix(i, 5) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ", " & Val(.TextMatrix(i, 4)) & ", '" & flxNewSP.TextMatrix(Rnew, 5) & "', '" & Format(Date, "yyyy-mm-dd") & "', '" & flxNewSP.TextMatrix(Rnew, 4) & "', '" & flxNewSP.TextMatrix(Rnew, 6) & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
                    
                    
                    If flxNewSP.TextMatrix(Rnew, 3) <> "" Then  'Add Description2
                    '***access
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
                'Update table SPITR ... if no ID ... get SP ID
                If .TextMatrix(i, 8) = "" Then .TextMatrix(i, 8) = GetSPid(i)
                'get order balance
                '***Access
                'Qry = "SELECT SPITR.* From SPITR WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & .TextMatrix(i, 3) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
                'mySQL
                Qry = "SELECT SPITR.* From SPITR WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.REFERENCE='" & .TextMatrix(i, 3) & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
                
                If CtrlNumExist(Qry) Then
                    Qry = "Exist"
                    myOrdBal = GetOrderBalance(i)
                Else
                    Qry = ""
                    myOrdBal = 0
                End If
                'Save SPITR
                '***Access
                'LYRcon.Execute "INSERT INTO SPITR ( REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, ORDERBAL, PACKINGLIST, REMARKS, SOURCE, UPDATEDBY, LASTUPDATE ) " & _
                               "VALUES ('" & txtPackingNo.Text & "', 'RECTEMP', #" & txtDateRec.Text & "#, " & flxParts.TextMatrix(i, 8) & ", " & Val(flxParts.TextMatrix(i, 4)) & ", '" & IIf(Qry = "Exist", myOrdBal, "") & "', '" & flxParts.TextMatrix(i, 3) & "', '" & flxParts.TextMatrix(i, 6) & "', '" & cboSupplier.Text & "', " & UserID & ", #" & Now & "# );"
                'mySQL
                LYRcon.Execute "INSERT INTO SPITR ( REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, ORDERBAL, PACKINGLIST, REMARKS, SOURCE, UPDATEDBY, LASTUPDATE ) " & _
                               "VALUES ('" & txtPackingNo.Text & "', 'RECTEMP', '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', " & flxParts.TextMatrix(i, 8) & ", " & Val(flxParts.TextMatrix(i, 4)) & ", '" & IIf(Qry = "Exist", myOrdBal, "") & "', '" & flxParts.TextMatrix(i, 3) & "', '" & flxParts.TextMatrix(i, 6) & "', '" & cboSupplier.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
                
            Else
                'get order balance
                '***Access
                'Qry = "SELECT SPITR.* From SPITR WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & .TextMatrix(i, 3) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
                'mySQL
                Qry = "SELECT SPITR.* From SPITR WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.REFERENCE='" & .TextMatrix(i, 3) & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
                   
                If CtrlNumExist(Qry) Then
                    Qry = "Exist"
                    myOrdBal = GetOrderBalance(i)
                Else
                    Qry = ""
                    myOrdBal = 0
                End If
                '***Access
                'LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = '" & txtPackingNo.Text & "', SPITR.TRANSTYPE = 'RECTEMP', SPITR.TRANSDATE = #" & txtDateRec.Text & "#, SPITR.SPAREPARTID = " & .TextMatrix(i, 8) & ", SPITR.RECEIVE = " & Val(.TextMatrix(i, 4)) & ", SPITR.ORDERBAL = '" & IIf(Qry = "Exist", myOrdBal, "") & "', SPITR.PACKINGLIST = '" & .TextMatrix(i, 3) & "', " & _
                               "SPITR.REMARKS = '" & .TextMatrix(i, 6) & "', SPITR.SOURCE = '" & cboSupplier.Text & "', SPITR.UPDATEDBY = " & UserID & ", SPITR.LASTUPDATE = #" & Now & "# WHERE (((SPITR.REFERENCE)='" & txtPackingNo.Text & "') AND ((SPITR.TRANSTYPE)='RECTEMP') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 10) & "));"
                'mySQL
                LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = '" & txtPackingNo.Text & "', SPITR.TRANSTYPE = 'RECTEMP', SPITR.TRANSDATE = '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', SPITR.SPAREPARTID = " & .TextMatrix(i, 8) & ", SPITR.RECEIVE = " & Val(.TextMatrix(i, 4)) & ", SPITR.ORDERBAL = '" & IIf(Qry = "Exist", myOrdBal, "") & "', SPITR.PACKINGLIST = '" & .TextMatrix(i, 3) & "', " & _
                               "SPITR.REMARKS = '" & .TextMatrix(i, 6) & "', SPITR.SOURCE = '" & cboSupplier.Text & "', SPITR.UPDATEDBY = " & UserID & ", SPITR.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' WHERE SPITR.REFERENCE='" & txtPackingNo.Text & "' AND SPITR.TRANSTYPE='RECTEMP' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 10) & ";"
            End If
        Next i
    End With
    'Update SPITR for Deleted Records
    If lstDelParts.ListCount > 0 Then
        For i = 0 To lstDelParts.ListCount - 1
            '***Access
            'LYRcon.Execute "DELETE  SPITR.* From SPITR WHERE (((SPITR.REFERENCE)='" & txtPackingNo.Text & "') AND ((SPITR.TRANSTYPE)='RECTEMP') AND ((SPITR.SPAREPARTID)=" & lstDelParts.List(i) & "));"
            'mySQL
            LYRcon.Execute "DELETE From SPITR WHERE SPITR.REFERENCE='" & txtPackingNo.Text & "' AND SPITR.TRANSTYPE='RECTEMP' AND SPITR.SPAREPARTID=" & lstDelParts.List(i) & ";"
            
        Next i
    End If
    
    
End Sub
'----------------------------------------- W/COST
Private Sub SaveSpareParts()
    Dim rsReceive As New ADODB.Recordset
    Dim i As Integer, myOrdBal As Integer, Rnew As Integer, myPrice As Currency
    Dim Qry As String
    
    SPrlastUp = Format(Now, "yyyy-mm-dd hh:mm:ss")
    If AddNew Then
        If LoadFromWarehouse Then txtRemarks.Text = "(ONLINE ENTRY) " & txtRemarks.Text
        LYRcon.Execute "INSERT INTO spreceive(DATEREC, PACKDATE, SUPPLIER, SHIPPER, POREF, PACKINGNO, CHECKEDBY, REMARKS, UPDATEDBY, LASTUPDATE) " & _
                "VALUES('" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', '" & Format(txtPLDate.Text, "yyyy-mm-dd") & "', '" & cboSupplier.Text & "', '" & cboShipper.Text & "', '" & txtPOno.Text & "', '" & txtPackingNo.Text & "', '" & txtChkBy.Text & "', '" & txtRemarks.Text & "', " & UserID & ", '" & SPrlastUp & "');"
        Qry = "SELECT spreceive.SPRRNO From spreceive Where spreceive.LASTUPDATE = '" & SPrlastUp & "' AND spreceive.PACKINGNO = '" & txtPackingNo.Text & "';"
        txtRRno.Text = GetFieldValue(Qry)
    Else
            LYRcon.Execute "UPDATE spreceive SET DATEREC = '" & Format(DateValue(txtDateRec.Text), "yyyy-mm-dd") & "', SUPPLIER = '" & cboSupplier.Text & "', " & _
            "SHIPPER = '" & cboShipper.Text & "', POREF = '" & txtPOno.Text & "', PACKINGNO = '" & txtPackingNo.Text & "', " & _
            "PACKDATE = '" & Format(DateValue(txtPLDate.Text), "yyyy-mm-dd") & "', CHECKEDBY = '" & txtChkBy.Text & "', " & _
            "REMARKS = '" & txtRemarks.Text & "', PRINTCOPY = 0, UPDATEDBY = " & UserID & ", " & _
            "LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' Where spreceive.SPRRNO = " & Val(txtRRno.Text) & ";"
    
        'REMOVE SPITR AND DISCREPANCY RECORDS
        LYRcon.Execute "DELETE From spitr Where spitr.REFERENCE = '" & txtPackingNo.Text & "' AND spitr.TRANSTYPE = 'RECTEMP' AND spitr.TRANSDATE = '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "';"
        LYRcon.Execute "DELETE From sprecdiscrepancy where sprecdiscrepancy.SPRRNO = " & Val(txtRRno.Text) & " AND sprecdiscrepancy.PACKINGNO = '" & txtPackingNo.Text & "';"
    End If
    With flxParts
        For i = 1 To .Rows - 1
            LYRcon.Execute "INSERT INTO spitr(REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, BALANCE, ORDERBAL, `ORDER`, COST, REMARKS, SOURCE, UPDATEDBY, LASTUPDATE, POREF, SEQNO) " & _
                "VALUES('" & txtPackingNo.Text & "', 'RECTEMP', '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', " & Val(flxParts.TextMatrix(i, 8)) & ", " & Val(flxParts.TextMatrix(i, 5)) & ", " & Val(flxParts.TextMatrix(i, 4)) & ", " & _
                Val(flxParts.TextMatrix(i, 7)) & ", " & Val(flxParts.TextMatrix(i, 14)) & ", " & Val(flxParts.TextMatrix(i, 13)) & ", '" & flxParts.TextMatrix(i, 6) & "', '" & cboSupplier.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', '" & txtPOno.Text & "', " & i & ");"
            
            If .TextMatrix(i, 6) <> "" Then 'Save Discrepancy
                LYRcon.Execute "INSERT INTO sprecdiscrepancy(SPRRNO, PACKINGNO, DATEREC, SPAREPARTID, QTY, GOOD, DAMAGE, LACKING, EXCESS, WRONGPN, REMARKS) " & _
                    "VALUES(" & Val(txtRRno.Text) & ", '" & txtPackingNo.Text & "', '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', " & Val(flxParts.TextMatrix(i, 8)) & ", " & Val(flxParts.TextMatrix(i, 4)) & "," & _
                    " " & Val(flxParts.TextMatrix(i, 5)) & ", " & Val(flxParts.TextMatrix(i, 9)) & ", " & Val(flxParts.TextMatrix(i, 10)) & ", " & Val(flxParts.TextMatrix(i, 11)) & ", " & Val(flxParts.TextMatrix(i, 12)) & ", '" & flxParts.TextMatrix(i, 12) & "');"
            End If
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

Function GetOrderBalance(MyRow As Integer) As Integer
    Dim rsOBal As New ADODB.Recordset
    Dim rsROrder As New ADODB.Recordset
    Dim Qry As String, Alloc As Integer, Recv As Integer
    
    '***Access
    'Qry = "SELECT SPITR.ORDER From SPITR WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & flxParts.TextMatrix(MyRow, 3) & "') AND ((SPITR.SPAREPARTID)=" & flxParts.TextMatrix(MyRow, 8) & "));"
    
    'mySQL
    Qry = "SELECT SPITR.`ORDER` From SPITR WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.REFERENCE='" & flxParts.TextMatrix(MyRow, 3) & "' AND SPITR.SPAREPARTID=" & flxParts.TextMatrix(MyRow, 8) & ";"
    
    With rsOBal
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Alloc = IIf(IsNull(!Order), 0, !Order)
        Else
            Alloc = 0
        End If
        .Close
    End With
    '***Access
    'Qry = "SELECT Sum(SPITR.[RECEIVE]) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & flxParts.TextMatrix(MyRow, 8) & ") AND ((SPITR.PACKINGLIST)='" & flxParts.TextMatrix(MyRow, 3) & "'));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & flxParts.TextMatrix(MyRow, 8) & " AND SPITR.POREF='" & flxParts.TextMatrix(MyRow, 3) & "';"
    
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
    
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.SPAREPARTID From SPAREPARTS " & _
          "WHERE (((SPAREPARTS.PARTNUMBER)='" & flxParts.TextMatrix(MyRow, 1) & "'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.SPAREPARTID From SPAREPARTS " & _
          "WHERE SPAREPARTS.PARTNUMBER='" & flxParts.TextMatrix(MyRow, 1) & "';"
    
    With rsGetid
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
    Dim rsSupplier As New ADODB.Recordset
    Dim rsShipper As New ADODB.Recordset
    Dim Qry As String, MySupp As String
    
    If rsParts.State = adStateOpen Then Exit Sub
    
    
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.COST, SPAREPARTS.STOREPRICE, SPAREPARTS.APPLICMODEL " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSrch.DataSource = rsParts
    
    With rsSupplier
        If .State = adStateClosed Then
            Qry = "SELECT supplier.SUPPLIERNAME, supplier.SUPPLIERID From supplier Order By supplier.SUPPLIERNAME"
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                If cboSupplier.Text <> "" Then MySupp = cboSupplier.Text
                cboSupplier.Clear
                Do While Not .EOF
                    cboSupplier.AddItem !SUPPLIERNAME
                    cboSupplier.ItemData(cboSupplier.ListCount - 1) = !SUPPLIERID
                    .MoveNext
                Loop
            End If
            .Close
        End If
        If MySupp <> "" Then cboSupplier.Text = MySupp
    End With
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
    If InStr(SysBranch, "MRC VAN") >= 1 Then
        cmdImpStkTrns.Visible = True
    ElseIf SysBranch = "BINOS BABAK SAMAL" Then
        cmdImpStkTrns.Visible = True
    Else
        cmdImpStkTrns.Visible = False
    End If
    
End Sub

Function ValidateSPRec() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, cnt As Integer
    
    vMsg = ""
    If cboSupplier.Text = "" Then vMsg = "The field SUPPLIER must not be empty."
    If txtPackingNo.Text = "" Then
        vMsg = vMsg & vbCrLf & "The field PACKING NO. must not be empty."
    Else
        If AddNew Then
            '***Access
            'Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
                  "WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
            'mySQL
            Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
                  "WHERE SPRECEIVE.PACKINGNO='" & txtPackingNo.Text & "';"
            
            If CtrlNumExist(Qry) Then vMsg = vMsg & vbCrLf & "The field PACKING NO. already exist."
        Else
            If txtRRno.Tag <> txtPackingNo.Text Then
                '***Access
                'Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
                      "WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
                
                'mySQL
                Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
                      "WHERE SPRECEIVE.PACKINGNO='" & txtPackingNo.Text & "';"
                
                If CtrlNumExist(Qry) Then vMsg = vMsg & vbCrLf & "The field PACKING NO. already exist."
            End If
        End If
    End If
    'If txtPLDate.Text <> "" Then
    If Not IsDate(txtPLDate.Text) Then vMsg = vMsg & vbCrLf & "The field Packing List Print Date must not be empty."
    'End If
    If txtChkBy.Text = "" Then vMsg = vMsg & vbCrLf & "The field CHECKED BY must not be empty."
    If Not IsDate(txtDateRec.Text) Then vMsg = vMsg & vbCrLf & "Invalid entry on field Date."
    With flxParts
        If .TextMatrix(.Rows - 1, 1) = "" And .TextMatrix(.Rows - 1, 2) = "" And .TextMatrix(.Rows - 1, 4) = "" And (.TextMatrix(.Rows - 1, 8) = "" Or .TextMatrix(.Rows - 1, 9) = "") Then .RemoveItem .Rows - 1
        If .TextMatrix(.Rows - 1, 1) = "" Or .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 4) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Spare Parts Received Units"
    End With
    ValidateSPRec = vMsg
End Function

Private Sub ShowWarehouseTransferInfo()
    Dim rsTransfer As New ADODB.Recordset
    Dim rsTransDtl As New ADODB.Recordset
    Dim Qry As String, MyITRef As String
    Dim i As Integer, TtlRec As Integer, Ans As Integer
    
    On Error GoTo LocalErr
    
    If MainCon.State = adStateClosed Then Call ConnectMain
'    Qry = "SELECT STOCKTRANSFER.* From STOCKTRANSFER WHERE STOCKTRANSFER.TRANSFERID=" & Val(txtPackingNo.Text) & ";"
    Qry = "SELECT stocktransfer.TRANSFERID, sporderbranch.BRANCHNAME, sporderbranch.BRANCHID, sporderbranch.REFERENCE, sporderbranch.ORDERDATE, stocktransfer.DRIVERNAME, stocktransfer.CHECKER, sporderbranch.BRANCHSPOID " & _
          "From stocktransfer INNER JOIN sporderbranch ON (stocktransfer.BRANCHSPOID = sporderbranch.BRANCHSPOID) " & _
          "Where stocktransfer.TRANSFERID = " & Val(txtPackingNo.Text) & " and sporderbranch.BRANCHID = " & SysBranchID & " Order By stocktransfer.TRANSFERID Desc;"
    With rsTransfer
        .Open Qry, MainCon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            txtPOno.Text = IIf(IsNull(!REFERENCE), "", !REFERENCE)
            cboShipper.Text = IIf(IsNull(!DriverName), "", !DriverName)
            txtPLDate.Text = Format(Date, "mm/dd/yyyy")
            txtBrnSPOID.Text = IIf(IsNull(!BRANCHSPOID), "", !BRANCHSPOID)
        Else
            MsgBox "Packing Number is not available!", vbCritical, "Receive Packing Number"
            Exit Sub
        End If
        .Close
    End With
    If txtPOno.Text <> "" Then
        'Qry = "SELECT sporder.SPORDERID From sporder Where sporder.REFERENCE = '" & txtPONO.Text & "';"
        'txtBrnSPOID.Tag = GetFieldValue(Qry)
        'If Val(txtBrnSPOID.Text) <> Val(txtBrnSPOID.Tag) Then
        '    txtPONO.Text = ""
        '    cboShipper.Text = ""
        '    txtPLDate.Text = ""
        '    txtBrnSPOID.Text = ""
        '    Exit Sub
        'Else
            Ans = MsgBox("Do you want to load P.O.#" & txtPOno.Text & "?", vbQuestion + vbYesNo, "PO from Warehouse")
            If Ans = vbNo Then
                txtPOno.Text = ""
                cboShipper.Text = ""
                txtPLDate.Text = ""
                txtBrnSPOID.Text = ""
                txtPackingNo.Text = ""
                Exit Sub
            End If
        'End If
    End If
    Qry = "SELECT Spitr.OUT, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spitr.SPAREPARTID, spitr.Price, spitr.Cost, spitr.POREF " & _
          "From spareparts INNER JOIN spitr ON (spareparts.SPAREPARTID = spitr.SPAREPARTID) Where spitr.REFERENCE = 'TR#" & txtPackingNo.Text & "' Order By spitr.SEQNO;"
    With rsTransDtl
        .Open Qry, MainCon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            .MoveFirst
            Do While Not .EOF       '       1                     2                                        3                     4              5 'good                                 5                                                  6                     7                   8
                flxParts.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(IsNull(!POref), "", !POref) & vbTab & !OUT & vbTab & !OUT
                flxParts.TextMatrix(flxParts.Rows - 1, 8) = !SPAREPARTID
                'flxParts.TextMatrix(flxParts.Rows - 1, 13) = !Price
 '---------------- W/COST
                flxParts.TextMatrix(flxParts.Rows - 1, 13) = !Cost
                .MoveNext
            Loop
        End If
        .Close
    End With
    With flxParts
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 3) <> "" Then
                Qry = "SELECT SPITR.`ORDER` From SPITR WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.REFERENCE='" & .TextMatrix(i, 3) & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
                .TextMatrix(i, 14) = GetFieldValue(Qry) 'Order Qty
                If Val(.TextMatrix(i, 14)) > 0 Then
                    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR WHERE SPITR.TRANSTYPE='RECEIVE' AND " & _
                          "SPITR.SPAREPARTID=" & flxParts.TextMatrix(i, 8) & " AND SPITR.POREF='" & flxParts.TextMatrix(i, 3) & "';"
                    TtlRec = Val(GetFieldValue(Qry)) + Val(.TextMatrix(i, 4))
                    .TextMatrix(i, 7) = Val(.TextMatrix(i, 14)) - TtlRec    'Back Order
                Else
                    .TextMatrix(i, 7) = 0
                End If
            End If
        Next i
    End With
    Call ComputeTotalQty
    MainCon.Close
    LoadFromWarehouse = True
    Exit Sub
LocalErr:
    MsgBox ErrMsg & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & "Please inform M.I.S. Department", vbCritical + vbOKOnly, "Error"

End Sub


Public Sub ShowRecord(MyRef As String)
    Dim rsSPrr As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim rsREM As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPRECEIVE.* From SPRECEIVE WHERE (((SPRECEIVE.PACKINGNO)='" & MyRef & "'));"
    
    'mySQl
    Qry = "SELECT SPRECEIVE.* From SPRECEIVE WHERE SPRECEIVE.PACKINGNO='" & MyRef & "';"
    
    With rsSPrr
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Call ClearEntries
            .MoveFirst
            cboShipper.Text = IIf(IsNull(.Fields("Shipper")), "", .Fields("Shipper"))
            cboSupplier.Text = .Fields("Supplier")
           ' txtWayBill.Text = IIf(IsNull(.Fields("WayBillNo")), "", .Fields("WayBillNO"))
            txtPOno.Text = IIf(IsNull(!POref), "", !POref)
            txtPackingNo.Text = !PACKINGNO
            txtPLDate.Text = IIf(IsNull(!PACKDATE), "", Format(!PACKDATE, "mm/dd/yyyy"))
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtPrnCopy.Text = !PRINTCOPY
            txtChkBy.Text = .Fields("CheckedBy")
            txtRRno.Text = .Fields("SPRRno")
            txtDateRec.Text = Format(!DATEREC, "mm/dd/yyyy")
        End If
        .Close
    End With
                '                  0                    1                     2              3                     4             5                  6                 7
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.RECEIVE, SPAREPARTS.LOCATION, SPITR.REMARKS, SPAREPARTS.QTYONHAND, SPITR.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & MyRef & "'));"

    'mySQL
    
    If InStr(txtRemarks.Text, "ONLINE ENTRY") > 0 Then  'REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, BALANCE, ORDERBAL, `ORDER`, PRICE, REMARKS, SOURCE, UPDATEDBY, LASTUPDATE, POREF
        LoadFromWarehouse = True
                '                  0                    1                     2              3                     4             5                  6                 7
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.POREF, SPITR.RECEIVE, SPITR.BALANCE, SPITR.REMARKS, SPITR.ORDERBAL, SPITR.SPAREPARTID, SPITR.PRICE " & _
              "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='" & MyRef & "';"
    
        With rsITR
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                flxParts.Rows = 1
                .MoveFirst
                Do While Not .EOF       '       1                     2                                       3                        4                 5                   6                   7                   8
                    flxParts.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(IsNull(!POref), "", !POref) & vbTab & !Balance & vbTab & !RECEIVE & vbTab & !Remarks & vbTab & !ORDERBAL & vbTab & !SPAREPARTID
                    flxParts.TextMatrix(flxParts.Rows - 1, 13) = !Price
                    Qry = "SELECT SPITR.`ORDER` From SPITR WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.REFERENCE= '" & flxParts.TextMatrix(flxParts.Rows - 1, 3) & "' AND SPITR.SPAREPARTID=" & flxParts.TextMatrix(flxParts.Rows - 1, 8) & ";"
                    flxParts.TextMatrix(flxParts.Rows - 1, 14) = GetFieldValue(Qry)
                    .MoveNext
                Loop
            End If
            .Close
        End With
        For i = 1 To flxParts.Rows - 1
            If flxParts.TextMatrix(i, 6) <> "" Then
                With rsREM
                    Qry = "SELECT * From sprecdiscrepancy Where sprecdiscrepancy.PACKINGNO = '" & txtPackingNo.Text & "' AND sprecdiscrepancy.SPAREPARTID = " & flxParts.TextMatrix(i, 8) & ";"
                    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                    If .RecordCount > 0 Then
                        .MoveFirst
                        flxParts.TextMatrix(i, 9) = !DAMAGE
                        flxParts.TextMatrix(i, 10) = !LACKING
                        flxParts.TextMatrix(i, 11) = !Excess
                        flxParts.TextMatrix(i, 12) = !WRONGPN
                        flxParts.TextMatrix(i, 15) = IIf(IsNull(!Remarks), "", !Remarks)
                    End If
                    .Close
                End With
            End If
        Next i
    Else                '                  0                    1                     2              3                     4             5                  6                 7
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
    End If
    Call ComputeTotalQty
    ViewOnly = True
    If InStr(SysBranch, "MRC VAN") >= 1 Then
        cmdImpStkTrns.Visible = True
        If InStr(cboShipper.Text, "IMPORT STOCK TRANSFER") > 0 Or InStr(txtRemarks.Text, "ONLINE ENTRY") > 0 Then
            cmdExportVan.Visible = True
        Else
            cmdExportVan.Visible = False
        End If
    ElseIf SysBranch = "BINOS BABAK SAMAL" Then
        cmdImpStkTrns.Visible = True
    Else
        cmdImpStkTrns.Visible = False
        cmdExportVan.Visible = False
    End If
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
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If

    If KeyAscii = 13 Then
        Call txtLoc_LostFocus
        flxParts.SetFocus
        flxParts.Col = 6
    End If
End Sub

Private Sub txtLoc_LostFocus()
    Dim Ans As Integer, Excess As Integer
    
    If txtLoc.Tag = "skiplostfocus" Then
        txtLoc.Tag = ""
        Exit Sub
    End If
    If Val(txtLoc.Text) > Val(flxParts.TextMatrix(ActRow, 4)) Then
        Excess = Val(txtLoc.Text) - Val(flxParts.TextMatrix(ActRow, 4))
        Ans = MsgBox("Do you consider " & Excess & IIf(Excess > 1, " pieces ", " piece ") & "variance as EXCESS?", vbQuestion + vbYesNo, "Excess")
        If Ans = vbYes Then
            flxParts.TextMatrix(ActRow, 5) = txtLoc.Text
            flxParts.TextMatrix(ActRow, 6) = "E(" & Excess & ")"
            flxParts.TextMatrix(ActRow, 11) = Excess
            txtLoc.Visible = False
            Exit Sub
        Else
            MsgBox "Invalid Entry! Good must not be more than " & flxParts.TextMatrix(ActRow, 4) & "!", vbExclamation + vbOKOnly, "Invalid Entry"
            Exit Sub
        End If
    End If
    flxParts.TextMatrix(ActRow, 5) = txtLoc.Text
    txtLoc.Visible = False
    If Val(txtLoc.Text) < Val(flxParts.TextMatrix(ActRow, 4)) Then 'Good less than received qty
        SPIncDisc.MyRow = ActRow
        SPIncDisc.Show
    ElseIf Val(txtLoc.Text) = Val(flxParts.TextMatrix(ActRow, 4)) Then
        flxParts.TextMatrix(ActRow, 6) = ""
    End If
End Sub

Private Sub txtPackingNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPackingNo_LostFocus()
    Dim Qry As String
    
    If AddNew Then
        '***Access
        'Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
              "WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
        'qry
        Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
              "WHERE SPRECEIVE.PACKINGNO='" & txtPackingNo.Text & "';"
              
        If CtrlNumExist(Qry) Then
            MsgBox "Packing No. already exist!", vbExclamation + vbOKOnly, "Validation Error"
        Else
            If cboSupplier.Text = "MRC WAREHOUSE" And Not LoadFromWarehouse Then
                Call ShowWarehouseTransferInfo
            End If
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
    
    If flxSrch.Tag = "skiplostfocus" Or txtPartNo.Text = "" Then
        flxSrch.Tag = ""
        txtPartNo.Visible = False
        Exit Sub
    End If
    If flxParts.TextMatrix(ActRow, 2) = "" And txtPartNo.Text <> "" Then
        'Find Part in Current MDB if exist then Update Sparepart MDB
        '***Access
        'If CtrlNumExist("SELECT SPAREPARTS.* From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & txtPartNo.Text & "'));") Then
            'mySQL
        If CtrlNumExist("SELECT SPAREPARTS.* From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & txtPartNo.Text & "';") Then
            flxParts.TextMatrix(ActRow, 2) = GetSPDesc(txtPartNo.Text)
            'flxParts.TextMatrix(ActRow, 9) = "New in SP only"
            Call txtPartNo_LostFocus
            Exit Sub
        End If
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please inform your Manager about this Part Number", vbQuestion + vbYesNo, "New Spare Part"
        flxParts.TextMatrix(ActRow, 1) = "" 'txtPartNo.Text
        txtPartNo.Visible = False
        flxSrch.Visible = False
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
    
    '***Access
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
    
    Tmp = RTrim(flxParts.TextMatrix(MyRow, 1))
    '***Access
    'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPAREPARTS.QTYONHAND, SPAREPARTS.REMARKS " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER) = '" & Tmp & "'));"
    'mySQl
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPAREPARTS.QTYONHAND, SPAREPARTS.REMARKS " & _
          "From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER = '" & Tmp & "';"
    
    ' old   Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPAREPARTS.QTYONHAND " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER) = '" & Tmp & "'));"
    With rsSP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst  'SP exist in MDB
            'flxParts.TextMatrix(ActRow, 5) = IIf(IsNull(!Location), "", !Location)
            'flxParts.TextMatrix(ActRow, 7) = !QtyOnHand
            flxParts.TextMatrix(ActRow, 8) = !SPAREPARTID
        Else    'PartNumber new in MDB
            'flxParts.TextMatrix(MyRow, 6) = "New"
            'flxParts.TextMatrix(MyRow, 9) = flxNewSP.Rows
            If cboSupplier.ListIndex >= 0 Then
                '***Access
                'SPtype = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE (((SUPPLIER.SUPPLIERID)=" & Val(cboSupplier.ItemData(cboSupplier.ListIndex)) & "));"
                 
                'mySQl
                SPtype = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE SUPPLIER.SUPPLIERID=" & Val(cboSupplier.ItemData(cboSupplier.ListIndex)) & ";"
         
                SPtype = GetFieldValue(SPtype)
            End If
            'flxNewSP.AddItem flxParts.TextMatrix(ActRow, 1) & vbTab & rsParts!Cost & vbTab & rsParts!RETAILPRICE & vbTab & vbTab & IIf(IsNull(rsParts!SUPERSEDING), "", rsParts!SUPERSEDING) & vbTab & IIf(SPtype <> "", SPtype, "REPLACEMENT") & vbTab & IIf(IsNull(rsParts!ApplicModel), "", rsParts!ApplicModel)
        End If
        .Close
    End With
    If flxParts.TextMatrix(MyRow, 8) = "" Then
        cboOrderNo.Clear
        Exit Sub
    End If
    '***Access
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
                cboOrderNo.AddItem !REFERENCE
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
    If flxParts.TextMatrix(ActRow, 3) <> "" And Val(flxParts.TextMatrix(ActRow, 14)) > 0 Then
        If Val(txtQty.Text) > Val(flxParts.TextMatrix(ActRow, 14)) Then
            MsgBox "Qty. Received must not be greater than Qty. Ordered!" & vbCrLf & "Please record all excess quantity in separate file.", vbExclamation + vbOKOnly, "Validation Error"
            flxParts.TextMatrix(ActRow, 4) = flxParts.TextMatrix(ActRow, 14)
            txtQty.Visible = False
            Call ComputeTotalQty
            Exit Sub
        End If
    End If
    flxParts.TextMatrix(ActRow, 4) = txtQty.Text
    flxParts.TextMatrix(ActRow, 5) = txtQty.Text
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

Private Sub txtWayBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Function GetLocation(MyPartNo As String) As String
    Dim rsLoc As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPAREPARTS.LOCATION, SPAREPARTS.PARTNUMBER " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & MyPartNo & "'));"
    'mySQl
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
