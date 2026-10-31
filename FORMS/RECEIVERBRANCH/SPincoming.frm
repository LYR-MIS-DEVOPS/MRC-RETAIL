VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SPincoming 
   Caption         =   "Incoming Spare Parts"
   ClientHeight    =   8730
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11595
   ClipControls    =   0   'False
   Icon            =   "SPincoming.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstDelParts 
      Height          =   450
      ItemData        =   "SPincoming.frx":17D2A
      Left            =   10320
      List            =   "SPincoming.frx":17D2C
      TabIndex        =   45
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   7920
      TabIndex        =   44
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      TabIndex        =   39
      Tag             =   "XYF"
      Top             =   7320
      Width           =   1335
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   720
      TabIndex        =   38
      Top             =   9000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxNewSP 
      Height          =   735
      Left            =   720
      TabIndex        =   37
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
      TabIndex        =   36
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboOrderNo 
      Height          =   315
      Left            =   7440
      TabIndex        =   35
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
      TabIndex        =   8
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
   Begin VB.TextBox txtRemarks 
      Height          =   285
      Left            =   1680
      MaxLength       =   250
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   7920
      Width           =   7095
   End
   Begin MSFlexGridLib.MSFlexGrid flxParts 
      Height          =   4935
      Left            =   480
      TabIndex        =   5
      Tag             =   "XY"
      Top             =   2400
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
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
            Picture         =   "SPincoming.frx":17D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPincoming.frx":18C08
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPincoming.frx":19AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPincoming.frx":1A9BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPincoming.frx":1B896
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPincoming.frx":1C770
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   480
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   720
      Width           =   10695
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
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   208
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   896
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
         TabIndex        =   4
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   906
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
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   568
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
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   230
         Width           =   1455
      End
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
      Image           =   "SPincoming.frx":1D64A
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
      Image           =   "SPincoming.frx":1DBE4
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
      Caption         =   "&Print Packing List"
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
      Image           =   "SPincoming.frx":1E17E
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
      Image           =   "SPincoming.frx":1E718
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   9000
      TabIndex        =   31
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
      Image           =   "SPincoming.frx":1ECB2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10800
      TabIndex        =   32
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
      Image           =   "SPincoming.frx":1F24C
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtChkBy 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   8280
      Width           =   2775
   End
   Begin lvButton.lvButtons_H cmdLoadPrtl 
      Height          =   375
      Left            =   10560
      TabIndex        =   41
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
      Image           =   "SPincoming.frx":1F7E6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrtlSave 
      Height          =   375
      Left            =   9960
      TabIndex        =   42
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
      Image           =   "SPincoming.frx":1FD80
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   43
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
      Image           =   "SPincoming.frx":2031A
      cBack           =   -2147483633
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
      TabIndex        =   40
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      m_MDI_Width     =   11715
      m_MDI_Height    =   9135
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
      TabIndex        =   20
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
      TabIndex        =   18
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   17
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
Dim SPcon As New ADODB.Connection
Dim rsParts As New ADODB.Recordset

Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean, OrdSrch As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim BtnIndx As Integer, FrstBtnIndx As Integer
Dim ViewOnly As Boolean, AddNew As Boolean

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
    'txtChkBy.Text = ""
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
    'txtChkBy.Locked = Bln
    txtDateRec.Locked = Bln
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
    If cboOrderNo.ListIndex >= 0 And flxParts.TextMatrix(ActRow, 4) = "" And flxParts.TextMatrix(ActRow, 8) <> "" Then
        MyOrder = GetOrderBalance(ActRow)
        If MyOrder > 0 Then
            flxParts.TextMatrix(ActRow, 4) = MyOrder
            Call ComputeTotalQty
        End If
    End If
End Sub

Private Sub cboShipper_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdA_Click()
    Call mnuAdd_Click
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
    Dim Qry As String, myMsg As String
    Dim Ans As Integer, i As Integer
    
    Select Case Index
        Case 0      '"new"
            Call BtnSaveCancel(True)
            'BtnIndx = 3
            'tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            'txtChkBy.Text = GetSettings("STOCK CLERK")
            txtDateRec.Text = Date
            ViewOnly = False
            AddNew = True
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
            ViewOnly = False
            AddNew = False
            For i = 1 To flxParts.Rows - 1
                flxParts.TextMatrix(i, 10) = flxParts.TextMatrix(i, 8)
            Next i
            Call BtnSaveCancel(True)
            flxNewSP.Rows = 0
        
        Case 2      '"view"
            SPRecSummary.Show
            Unload Me
        Case 3      '"print"
            If txtRRno.Text = "" Then
                MsgBox "No Available Packing List to Print", vbExclamation + vbOKOnly, "Print"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Print Packing List?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            
            PRNmc.MySource = "SPRECEIVE"
            PRNmc.QryPrn = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.RECEIVE, SPAREPARTS.LOCATION, SPITR.REMARKS " & _
                           "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & txtPackingNo.Text & "'));"
            PRNmc.ParamValue1 = cboSupplier.Text
            PRNmc.ParamValue2 = cboShipper.Text
            PRNmc.ParamValue3 = txtPackingNo.Text
            PRNmc.ParamValue4 = txtPLDate.Text
            PRNmc.ParamValue5 = txtChkBy.Text
            PRNmc.ParamValue6 = txtDateRec.Text
            PRNmc.Show
        Case 4      '"quit"
            'MotoRep.tbMain.Visible = True
            Unload Me
            LogIn.Show vbModal
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
    End Select

End Sub

Public Sub FinalSPrec()
    Dim i As Integer
    
    If Val(txtPrnCopy.Text) > 0 Then Exit Sub
    
    'Update SPITR field TransType to RECEIVE
    LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'RECEIVE' " & _
                   "WHERE (((SPITR.REFERENCE)='" & txtPackingNo.Text & "') AND ((SPITR.TRANSTYPE)='RECTEMP'));"
    'Update Print Copy
    txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
    LYRcon.Execute "UPDATE SPRECEIVE SET SPRECEIVE.PRINTCOPY = " & Val(txtPrnCopy.Text) & _
                   " WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
    'Update Qty on Hand of all received SpareParts
    With flxParts
        For i = 1 To .Rows - 1
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+ " & Val(.TextMatrix(i, 4)) & ", SPAREPARTS.REFERENCE = '" & txtPackingNo.Text & "' " & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
        Next i
    End With
    
    MsgBox "Spare Parts Qty-On-Hand successfully updated!"
    
End Sub

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
    
    SPcon.CursorLocation = adUseClient
    SPcon.Open "dsn=SPAREPARTS;uid=;pwd=;database=SPAREPARTS"
    Call LoadSpareParts
    txtChkBy.Text = GetUserName
    'MotoRep.sBar.Panels(6).Text = "Spare Parts Packing List"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'MotoRep.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
'    If rsOrder.State = adStateOpen Then rsOrder.Close

    Set SPcon = Nothing
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
        End If
    End If
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
    Qry = "SELECT SPRECEIVE.*, SPRECEIVE.SPRRNO From SPRECEIVE WHERE (((SPRECEIVE.SPRRNO)=" & txtRRno.Text & "));"
    
    rsSPrec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSPrec.RecordCount > 0 Then
        rsSPrec.MoveFirst
        rsSPrec!DATEREC = txtDateRec.Text
        rsSPrec!Supplier = cboSupplier.Text
        rsSPrec!Shipper = cboShipper.Text
        rsSPrec!WayBillno = txtWayBill.Text
        rsSPrec!PACKINGNO = txtPackingNo.Text
        If IsDate(txtPLDate.Text) Then rsSPrec!PackDate = txtPLDate.Text
        rsSPrec!CheckedBy = txtChkBy.Text
        rsSPrec!REMARKS = txtRemarks.Text
        rsSPrec!PrintCopy = 0
        rsSPrec!UpdatedBy = UserID
        rsSPrec!LastUpdate = Date
        rsSPrec.Update
        rsSPrec.Close
    End If
    'Update SPITR reference if txtpackingno was change
    If txtRRno.Tag <> txtPackingNo.Text Then
        LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = '" & txtPackingNo.Text & "' " & _
                       "WHERE (((SPITR.REFERENCE)='" & txtRRno.Tag & "') AND ((SPITR.TRANSTYPE)='RECTEMP'));"
    End If
    With flxParts
        For i = 1 To flxParts.Rows - 1
            If .TextMatrix(i, 10) = "" Then
                If .TextMatrix(i, 8) <> "" Then     'SP already available in MDB
                    If .TextMatrix(i, 9) = "New in SP only" Then
                        'Update SparePart Database
                        SPcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION) " & _
                                      "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "');"
                    End If
                Else        'New SP
                    Rnew = Val(.TextMatrix(i, 9))
                    LYRcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, REFERENCE, LOCATION, COST, STOREPRICE, BEGBAL, REMARKS, DATEBEGBAL, SUPERSEDINGPN, APPLICMODEL, UPDATEDBY, LASTUPDATE) " & _
                                   "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', '" & txtPackingNo.Text & "', '" & .TextMatrix(i, 5) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ", " & Val(.TextMatrix(i, 4)) & ", '" & flxNewSP.TextMatrix(Rnew, 5) & "', #" & Date & "#, '" & flxNewSP.TextMatrix(Rnew, 4) & "', '" & flxNewSP.TextMatrix(Rnew, 6) & "', " & UserID & ", #" & Date & "#);"
                    If flxNewSP.TextMatrix(Rnew, 3) <> "" Then  'Add Description2
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '" & flxNewSP.TextMatrix(Rnew, 3) & "' " & _
                                       "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                    End If
                    'Update SparePart Database
                    If Not SPExist(flxParts.TextMatrix(i, 1)) Then
                        SPcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, SUPERSEDING, COST, RETAILPRICE) " & _
                                      "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', '" & flxNewSP.TextMatrix(Rnew, 4) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ");"
                    End If
                End If
                'Update table SPITR ... if no ID ... get SP ID
                If .TextMatrix(i, 8) = "" Then .TextMatrix(i, 8) = GetSPid(i)
                'get order balance
                Qry = "SELECT SPITR.* From SPITR WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & .TextMatrix(i, 3) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
                If CtrlNumExist(Qry) Then
                    Qry = "Exist"
                    myOrdBal = GetOrderBalance(i)
                Else
                    Qry = ""
                    myOrdBal = 0
                End If
                'Save SPITR
                LYRcon.Execute "INSERT INTO SPITR ( REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, ORDERBAL, PACKINGLIST, REMARKS, SOURCE, UPDATEDBY, LASTUPDATE ) " & _
                               "VALUES ('" & txtPackingNo.Text & "', 'RECTEMP', #" & txtDateRec.Text & "#, " & flxParts.TextMatrix(i, 8) & ", " & Val(flxParts.TextMatrix(i, 4)) & ", '" & IIf(Qry = "Exist", myOrdBal, "") & "', '" & flxParts.TextMatrix(i, 3) & "', '" & flxParts.TextMatrix(i, 6) & "', '" & cboSupplier.Text & "', " & UserID & ", #" & Now & "# );"
            Else
                'get order balance
                Qry = "SELECT SPITR.* From SPITR WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & .TextMatrix(i, 3) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
                If CtrlNumExist(Qry) Then
                    Qry = "Exist"
                    myOrdBal = GetOrderBalance(i)
                Else
                    Qry = ""
                    myOrdBal = 0
                End If
                LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = '" & txtPackingNo.Text & "', SPITR.TRANSTYPE = 'RECTEMP', SPITR.TRANSDATE = #" & txtDateRec.Text & "#, SPITR.SPAREPARTID = " & .TextMatrix(i, 8) & ", SPITR.RECEIVE = " & Val(.TextMatrix(i, 4)) & ", SPITR.ORDERBAL = '" & IIf(Qry = "Exist", myOrdBal, "") & "', SPITR.PACKINGLIST = '" & .TextMatrix(i, 3) & "', " & _
                               "SPITR.REMARKS = '" & .TextMatrix(i, 6) & "', SPITR.SOURCE = '" & cboSupplier.Text & "', SPITR.UPDATEDBY = " & UserID & ", SPITR.LASTUPDATE = #" & Now & "# WHERE (((SPITR.REFERENCE)='" & txtPackingNo.Text & "') AND ((SPITR.TRANSTYPE)='RECTEMP') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 10) & "));"
            End If
        Next i
    End With
    'Update SPITR for Deleted Records
    If lstDelParts.ListCount > 0 Then
        For i = 0 To lstDelParts.ListCount - 1
            LYRcon.Execute "DELETE  SPITR.* From SPITR WHERE (((SPITR.REFERENCE)='" & txtPackingNo.Text & "') AND ((SPITR.TRANSTYPE)='RECTEMP') AND ((SPITR.SPAREPARTID)=" & lstDelParts.List(i) & "));"
        Next i
    End If
    
    
End Sub

Private Sub SaveSpareParts()
    Dim rsReceive As New ADODB.Recordset
    Dim i As Integer, myOrdBal As Integer, Rnew As Integer
    Dim Qry As String
    
    'Saving SPReceive
    rsReceive.Open "SPReceive", LYRcon, adOpenKeyset, adLockOptimistic
    rsReceive.AddNew
    rsReceive!DATEREC = txtDateRec.Text
    rsReceive!Supplier = cboSupplier.Text
    rsReceive!Shipper = cboShipper.Text
    rsReceive!WayBillno = txtWayBill.Text
    rsReceive!PACKINGNO = txtPackingNo.Text
    If IsDate(txtPLDate.Text) Then rsReceive!PackDate = txtPLDate.Text
    rsReceive!CheckedBy = txtChkBy.Text
    rsReceive!REMARKS = txtRemarks.Text
    rsReceive!PrintCopy = 0
    rsReceive!UpdatedBy = UserID
    rsReceive!LastUpdate = Date
    rsReceive.Update
    txtRRno.Text = rsReceive!SPRRNO
    rsReceive.Close
        
    With flxParts
        For i = 1 To .Rows - 1
            'Update table SpareParts field Qty., if SP does not exist add to new record
            If .TextMatrix(i, 8) <> "" Then     'SP already available in MDB
                'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+ " & Val(.TextMatrix(i, 4)) & ", SPAREPARTS.REFERENCE = '" & txtPackingNo.Text & "', SPAREPARTS.LOCATION = '" & .TextMatrix(i, 5) & "' " & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
                If .TextMatrix(i, 9) = "New in SP only" Then
                    'Update SparePart Database
                    SPcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION) " & _
                                  "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "');"
                End If
            Else        'New SP
                Rnew = Val(.TextMatrix(i, 9))
                LYRcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, REFERENCE, LOCATION, COST, STOREPRICE, BEGBAL, REMARKS, DATEBEGBAL, SUPERSEDINGPN, APPLICMODEL, UPDATEDBY, LASTUPDATE) " & _
                               "VALUES ('" & .TextMatrix(i, 1) & "', '" & Trim(.TextMatrix(i, 2)) & "', '" & txtPackingNo.Text & "', '" & .TextMatrix(i, 5) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ", " & Val(.TextMatrix(i, 4)) & ", '" & flxNewSP.TextMatrix(Rnew, 5) & "', #" & Date & "#, '" & flxNewSP.TextMatrix(Rnew, 4) & "', '" & flxNewSP.TextMatrix(Rnew, 6) & "', " & UserID & ", #" & Date & "#);"
                If flxNewSP.TextMatrix(Rnew, 3) <> "" Then  'Add Description2
                    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '" & flxNewSP.TextMatrix(Rnew, 3) & "' " & _
                                   "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(i, 1) & "'));"
                End If
                'Update SparePart Database
                If Not SPExist(flxParts.TextMatrix(i, 1)) Then
                    SPcon.Execute "INSERT INTO SPAREPARTS ( PARTNUMBER, DESCRIPTION, SUPERSEDING, COST, RETAILPRICE) " & _
                                  "VALUES ('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', '" & flxNewSP.TextMatrix(Rnew, 4) & "', " & Val(flxNewSP.TextMatrix(Rnew, 1)) & ", " & Val(flxNewSP.TextMatrix(Rnew, 2)) & ");"
                End If
            End If
            'Update table SPITR
            'if no ID ... get SP ID
            If .TextMatrix(i, 8) = "" Then .TextMatrix(i, 8) = GetSPid(i)
            'get order balance
            Qry = "SELECT SPITR.* From SPITR WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & .TextMatrix(i, 3) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
            If CtrlNumExist(Qry) Then
                Qry = "Exist"
                myOrdBal = GetOrderBalance(i)
            Else
                Qry = ""
                myOrdBal = 0
            End If
            'Save SPITR
            LYRcon.Execute "INSERT INTO SPITR ( REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, ORDERBAL, PACKINGLIST, REMARKS, SOURCE, UPDATEDBY, LASTUPDATE ) " & _
                           "VALUES ('" & txtPackingNo.Text & "', 'RECTEMP', #" & txtDateRec.Text & "#, " & flxParts.TextMatrix(i, 8) & ", " & Val(flxParts.TextMatrix(i, 4)) & ", '" & IIf(Qry = "Exist", myOrdBal, "") & "', '" & flxParts.TextMatrix(i, 3) & "', '" & flxParts.TextMatrix(i, 6) & "', '" & cboSupplier.Text & "', " & UserID & ", #" & Now & "# );"
            'Update SpareParts BackOrder and OrderRef
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
    
    Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & myPN & "'));"
    With rsSPall
        .Open Qry, SPcon, adOpenKeyset, adLockOptimistic
        SPExist = IIf(.RecordCount > 0, True, False)
        .Close
    End With
End Function

Function GetOrderBalance(MyRow As Integer) As Integer
    Dim rsOBal As New ADODB.Recordset
    Dim rsROrder As New ADODB.Recordset
    Dim Qry As String, Alloc As Integer, Recv As Integer
    
    Qry = "SELECT SPITR.ORDER From SPITR WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.REFERENCE)='" & flxParts.TextMatrix(MyRow, 3) & "') AND ((SPITR.SPAREPARTID)=" & flxParts.TextMatrix(MyRow, 8) & "));"
   
    With rsOBal
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Alloc = IIf(IsNull(!Order), 0, !Order)
        Else
            Alloc = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.[RECEIVE]) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & flxParts.TextMatrix(MyRow, 8) & ") AND ((SPITR.PACKINGLIST)='" & flxParts.TextMatrix(MyRow, 3) & "'));"
    With rsROrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Recv = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            Recv = 0
        End If
        .Close
    End With            'Order Allocation - [Previous Receive Parts + Current Receive Parts]
    GetOrderBalance = Alloc - (Recv + Val(flxParts.TextMatrix(MyRow, 4)))
End Function

Function GetSPid(MyRow As Integer) As Integer
    Dim rsGetid As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.SPAREPARTID From SPAREPARTS " & _
          "WHERE (((SPAREPARTS.PARTNUMBER)='" & flxParts.TextMatrix(MyRow, 1) & "'));"
    With rsGetid
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
    Dim rsSupplier As New ADODB.Recordset
    Dim rsShipper As New ADODB.Recordset
    Dim Qry As String
    
    If rsParts.State = adStateOpen Then Exit Sub
    Qry = "SELECT SPAREPARTS.* From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"

    
    rsParts.Open Qry, SPcon, adOpenKeyset, adLockOptimistic
    Set flxSrch.DataSource = rsParts
    
    With rsSupplier
        .Open "Supplier", LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupplier.Clear
            Do While Not .EOF
                cboSupplier.AddItem .Fields(1)
                .MoveNext
            Loop
        End If
        .Close
    End With
    With rsShipper
        .Open "Shipper", LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboShipper.Clear
            Do While Not .EOF
                cboShipper.AddItem .Fields(1)
                .MoveNext
            Loop
        End If
        .Close
    End With
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
            Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
                  "WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
            If CtrlNumExist(Qry) Then vMsg = vMsg & vbCrLf & "The field PACKING NO. already exist."
        Else
            If txtRRno.Tag <> txtPackingNo.Text Then
                Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
                      "WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
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
        If .TextMatrix(.Rows - 1, 1) = "" And .TextMatrix(.Rows - 1, 2) = "" And .TextMatrix(.Rows - 1, 4) = "" And (.TextMatrix(.Rows - 1, 8) = "" Or .TextMatrix(.Rows - 1, 9) = "") Then .RemoveItem .Rows - 1
        If .TextMatrix(.Rows - 1, 1) = "" Or .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 4) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Spare Parts Received Units"
    End With
    ValidateSPRec = vMsg
End Function

Public Sub ShowRecord(MyRef As String)
    Dim rsSPrr As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT SPRECEIVE.* From SPRECEIVE WHERE (((SPRECEIVE.PACKINGNO)='" & MyRef & "'));"

    With rsSPrr
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Call ClearEntries
            .MoveFirst
            cboSupplier.Text = .Fields("Supplier")
            cboShipper.Text = IIf(IsNull(.Fields("Shipper")), "", .Fields("Shipper"))
            txtWayBill.Text = IIf(IsNull(.Fields("WayBillNo")), "", .Fields("WayBillNO"))
            txtPackingNo.Text = !PACKINGNO
            txtPLDate.Text = IIf(IsNull(!PackDate), "", Format(!PackDate, "mm/dd/yyyy"))
            txtRemarks.Text = IIf(IsNull(!REMARKS), "", !REMARKS)
            txtPrnCopy.Text = !PrintCopy
            txtChkBy.Text = .Fields("CheckedBy")
            txtRRno.Text = .Fields("SPRRno")
            txtDateRec.Text = Format(!DATEREC, "mm/dd/yyyy")
        End If
        .Close
    End With
                '                  0                    1                     2              3                     4             5                  6                 7
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.RECEIVE, SPAREPARTS.LOCATION, SPITR.REMARKS, SPAREPARTS.QTYONHAND, SPITR.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & MyRef & "'));"

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
        Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE " & _
              "WHERE (((SPRECEIVE.PACKINGNO)='" & txtPackingNo.Text & "'));"
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
            'flxParts.TextMatrix(ActRow, 8) = .TextMatrix(.Row, 1) 'SparePart ID
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
    
    If flxSrch.Tag = "skiplostfocus" Then
        flxSrch.Tag = ""
        Exit Sub
    End If
    If flxParts.TextMatrix(ActRow, 2) = "" And txtPartNo.Text <> "" Then
        'Find Part in Current MDB if exist then Update Sparepart MDB
        If CtrlNumExist("SELECT SPAREPARTS.* From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & txtPartNo.Text & "'));") Then
            flxParts.TextMatrix(ActRow, 2) = GetSPDesc(txtPartNo.Text)
            flxParts.TextMatrix(ActRow, 9) = "New in SP only"
            Call txtPartNo_LostFocus
            Exit Sub
        End If
        Ans = MsgBox("Unrecognized Part Number!" & vbCrLf & "Do you want to add this Part No.?", vbQuestion + vbYesNo, "New Spare Part")
        flxParts.TextMatrix(ActRow, 1) = txtPartNo.Text
        txtPartNo.Visible = False
        flxSrch.Visible = False
        If Ans = vbYes Then SPnew.Show vbModal
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
    
    Qry = "SELECT SPAREPARTS.DESCRIPTION From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & myPN & "'));"
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
    Dim Qry As String, Tmp As String
    
    Tmp = RTrim(flxParts.TextMatrix(MyRow, 1))
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPAREPARTS.QTYONHAND " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER) = '" & Tmp & "'));"
'    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPAREPARTS.QTYONHAND " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER) = '" & Tmp & "'));"
    With rsSP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst  'SP exist in MDB
            flxParts.TextMatrix(ActRow, 5) = IIf(IsNull(!Location), "", !Location)
            flxParts.TextMatrix(ActRow, 7) = !QtyOnHand
            flxParts.TextMatrix(ActRow, 8) = !SparePartID
        Else    'PartNumber new in MDB
            flxParts.TextMatrix(MyRow, 6) = "New"
            flxParts.TextMatrix(MyRow, 9) = flxNewSP.Rows
            flxNewSP.AddItem flxParts.TextMatrix(ActRow, 1) & vbTab & rsParts!Cost & vbTab & rsParts!RetailPrice & vbTab & vbTab & IIf(IsNull(rsParts!Superseding), "", rsParts!Superseding) & vbTab & "GENUINE" & vbTab & IIf(IsNull(rsParts!ApplicModel), "", rsParts!ApplicModel)
        End If
        .Close
    End With
    If flxParts.TextMatrix(MyRow, 8) = "" Then
        cboOrderNo.Clear
        Exit Sub
    End If
    Qry = "SELECT SPITR.REFERENCE, SPITR.BALANCE FROM SPITR " & _
          "WHERE (((SPITR.BALANCE)>0) AND ((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & flxParts.TextMatrix(MyRow, 8) & "));"

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

Private Sub txtWayBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

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
