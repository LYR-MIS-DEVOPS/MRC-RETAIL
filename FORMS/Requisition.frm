VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Requisition 
   Caption         =   "Spare Parts Requisition"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAmt 
      Height          =   285
      Left            =   8640
      TabIndex        =   44
      Top             =   8760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtJOid 
      Height          =   285
      Left            =   9000
      TabIndex        =   42
      Top             =   8040
      Width           =   615
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   8520
      TabIndex        =   41
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtORno 
      Height          =   285
      Left            =   8760
      TabIndex        =   40
      Top             =   8400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtInvNo 
      Height          =   285
      Left            =   2160
      TabIndex        =   31
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   315
      Left            =   9000
      TabIndex        =   30
      Tag             =   "XYF"
      Top             =   1800
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
      Image           =   "Requisition.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   6960
      TabIndex        =   29
      Tag             =   "XYF"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtReqID 
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Top             =   8880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPrintCopy 
      Height          =   285
      Left            =   1320
      TabIndex        =   27
      Top             =   8520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   7920
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   8400
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   7920
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   8400
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   1200
      TabIndex        =   24
      Top             =   8880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCust 
      Height          =   1815
      Left            =   2640
      TabIndex        =   23
      Top             =   8040
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
      TabIndex        =   22
      Top             =   8040
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
   Begin VB.TextBox txtWCRid 
      Height          =   285
      Left            =   960
      TabIndex        =   21
      Top             =   9480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAltA 
      Caption         =   "&A"
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Tag             =   "XY"
      Top             =   10920
      Width           =   135
   End
   Begin VB.Timer tmrQtyTip 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   8040
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   240
      TabIndex        =   19
      Text            =   "Qty. On Hand: 0"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtFlx 
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   8520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAcctNo 
      Height          =   285
      Left            =   8640
      TabIndex        =   17
      Top             =   10560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   5055
      Left            =   240
      TabIndex        =   7
      Tag             =   "XY"
      Top             =   2160
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   0
      FormatString    =   "^     |^ QTY. |^ UNIT |<     PART NUMBER         |<    DESCRIPTION            |>   UNIT PRICE  |>TOTAL  PRICE "
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
      Height          =   315
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   10
      Tag             =   "XYF"
      Top             =   7245
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   240
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   720
      Width           =   9735
      Begin VB.CheckBox chkOnAcct 
         Caption         =   "ON ACCOUNT"
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
         Height          =   255
         Left            =   3960
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   210
         Width           =   1575
      End
      Begin VB.TextBox txtReqNo 
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
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   180
         Width           =   1695
      End
      Begin VB.TextBox txtDate 
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
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   495
         Width           =   1695
      End
      Begin VB.TextBox txtCustomer 
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
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   495
         Width           =   4215
      End
      Begin VB.TextBox txtJobNo 
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
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "REQUISITION DATE"
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
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   525
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "CUSTOMER NAME"
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
         Left            =   120
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   525
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "JOB ORDER NO."
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
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "REQUISITION NO."
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
         TabIndex        =   9
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
            Picture         =   "Requisition.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Requisition.frx":1474
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Requisition.frx":234E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Requisition.frx":3228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Requisition.frx":4102
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "Add &New"
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
      Image           =   "Requisition.frx":4FDC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   2
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "E&xit"
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
      Image           =   "Requisition.frx":5576
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7200
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "Requisition.frx":5B10
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   8640
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Cancel "
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
      Image           =   "Requisition.frx":60AA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "Requisition.frx":6644
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   33
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Summary"
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
      Image           =   "Requisition.frx":6BDE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   34
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Print"
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
      Image           =   "Requisition.frx":7178
      cBack           =   -2147483633
   End
   Begin VB.Label lblRemarks 
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
      TabIndex        =   39
      Tag             =   "XYF"
      Top             =   7320
      Width           =   6615
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "REQUISITION FORM"
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
      Height          =   270
      Index           =   2
      Left            =   7320
      TabIndex        =   37
      Top             =   480
      Width           =   2610
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "REQUISITION FORM"
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
      Height          =   270
      Index           =   3
      Left            =   7350
      TabIndex        =   38
      Top             =   480
      Width           =   2610
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
      TabIndex        =   32
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
      m_MDI_Width     =   10290
      m_MDI_Height    =   8310
      OriginalBorder  =   2
      OriginalCaption =   27
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
      Left            =   7680
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   7290
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
      TabIndex        =   14
      Top             =   1905
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
      Height          =   275
      Index           =   0
      Left            =   240
      TabIndex        =   12
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
      Height          =   275
      Index           =   1
      Left            =   270
      TabIndex        =   13
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
      TabIndex        =   15
      Top             =   1935
      Width           =   2415
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
Attribute VB_Name = "Requisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
                                                                                                           
Dim PrtSrch As Boolean, ActSrch As Boolean, PrtBln As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim tmrCount As Integer, BtnIndx As Integer
Dim AddNew As Boolean
Dim fRow As Integer, fCol As Integer
Public MySource As String, JustScroll As Boolean

Private Sub ClearEntries()
    txtJobNo.Text = ""
    txtCustomer.Text = ""
    txtReqNo.Text = ""
    txtDate.Text = ""
    txtTotal.Text = ""
    txtReqID.Text = ""
    txtFind.Text = ""
    txtInvNo.Text = ""
    txtCustID.Text = ""
    chkOnAcct.Visible = False
    chkOnAcct.Value = 0
    flxOrd.Rows = 1
    flxOrd.Rows = 2
    txtPrintCopy = ""
    txtAmt.Text = ""
End Sub
                                                                
Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    JustScroll = bln
End Sub

Private Sub cmdAltA_Click()
    Call mnuAdd_Click
End Sub

Private Sub cmdFind_Click()
    SPfind.MySource = "REQUISITION"
    SPfind.Show vbModal
End Sub

Private Sub GetCustomerTerm()
    Dim rsCust As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CUSTOMERID From CUSTOMER " & _
          "WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
    'mySQl
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CUSTOMERID From CUSTOMER " & _
          "WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
    
    With rsCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtCustomer.Text = !CustomerName
            txtCustID.Text = !CUSTOMERID
        End If
        .Close
    End With
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 5) As String, Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=39));"
            
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=39;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            Call ClearEntries
            Call LockData(False)
            txtDate.Text = Date
            AddNew = True
            Call cmdTransDisp(1)
            If MySource = "Job Order" Then
                txtJobNo.Text = JobOrder.txtJobNo.Text
                If Val(JobOrder.txtCustID) > 0 And JobOrder.chkOnAcct.Value = 1 Then
                    txtCustID.Text = JobOrder.txtCustID.Text
                    txtJOid.Text = JobOrder.txtJOid.Text
                    lblRemarks.Caption = JobOrder.txtDesignation & " " & JobOrder.txtCustomer.Text & "/ Plate No. " & JobOrder.flxMC.TextMatrix(1, 3)
                    Call GetCustomerTerm
                    chkOnAcct.Visible = True
                    chkOnAcct.Value = 1
                Else
                    txtCustomer.Text = JobOrder.txtCustomer.Text
                End If
                'txtReqNo.SetFocus
            Else
                txtJobNo.SetFocus
            End If
        
        Case 1  'Edit
            If txtReqID.Text = "" Then
                MsgBox "No available Requisition to edit", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If Val(txtPrintCopy.Text) > 0 Then
                MsgBox "Unable to edit Req. No. " & txtReqNo.Text & " because a copy was already printed.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If InStr(lblRemarks.Caption, "CANCELLED") > 0 Then
                MsgBox "Unable to edit REQ.# " & txtReqNo.Text & " because it has already been Cancelled.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            '***Access
            'Qry = "SELECT SPSALESORDER.* From SPSALESORDER WHERE (((SPSALESORDER.REQID)=" & txtReqID.Text & "));"
            'old Qry = "SELECT SPSALESORDER.* From SPSALESORDER WHERE (((SPSALESORDER.REQID)=" & txtReqID.Text & ") AND ((SPSALESORDER.INVOICENO)<>'CANCELLED'));"
            'mySQL
            Qry = "SELECT SPSALESORDER.* From SPSALESORDER WHERE SPSALESORDER.REQID=" & txtReqID.Text & ";"
            
            If CtrlNumExist(Qry) Then
                MsgBox "Unable to edit REQ.#" & txtReqNo.Text & " because Sales Order has been served.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
                
            Call cmdTransDisp(1)
            Call LockData(False)
            cmdTrans(1).Tag = txtReqNo.Text
            AddNew = False
        Case 2  'Summary
            Unload Me
            FrmSummary.mySrc = "Requisition"
            FrmSummary.Show
            'ReqSummary.Show
        Case 3  'Print
            If SysBranch <> "LYR MOTOR SHOP AND SPAREPARTS DAVAO" Then
                Ans = MsgBox("Do you want to Print Requisition?", vbQuestion + vbYesNo, "Print")
                If Ans = vbNo Then Exit Sub
                PrnVal(0) = txtCustomer.Text
                PrnVal(1) = txtJobNo.Text
                PrnVal(2) = txtDate.Text
                PrnVal(3) = "TOTAL:"
                PrnVal(4) = txtTotal.Text
                PrnVal(5) = lblRemarks.Caption
                Call PrintMyForm(PrnVal, "REQUISITION", flxOrd)
            End If
            '***Access
            'LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.PRINTCOPY = [REQUISITION].[PRINTCOPY]+1 " & _
                           " WHERE (((REQUISITION.REQID)=" & txtReqID.Text & "));"
            'mySQL
            LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.PRINTCOPY = REQUISITION.PRINTCOPY+1" & _
                           " WHERE REQUISITION.REQID=" & txtReqID.Text & ";"
        Case 4  'Exit
            Unload Me
        Case 5  'Save
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            myMsg = ValidateRequisition
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Requisition?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'Save Spare Parts Invoice
            Call SaveRequisition
            MsgBox "Requisition successfully Saved!", vbInformation + vbOKOnly, "Save"
            Call cmdTransDisp(2)
            If MySource = "Job Order" Then
                cmdTrans(2).Enabled = False
                If AddNew Then
                    With JobOrder.flxJob
                        Ans = -1
                        For i = .Rows - 1 To 1 Step -1
                            If InStr(UCase$(.TextMatrix(i, 1)), "REQUISITION") > 0 Then
                                .TextMatrix(i, 1) = .TextMatrix(i, 1) & ", " & txtReqNo.Text
                                JobOrder.cboReq.AddItem txtReqNo.Text
                                JobOrder.cboReq.Text = txtReqNo.Text
                                Ans = 99
                                Exit For
                            End If
                        Next i
                        If Ans = -1 Then
                            .AddItem vbTab & "Requisition No. " & txtReqNo.Text
                            JobOrder.cboReq.AddItem txtReqNo.Text
                            JobOrder.cboReq.Text = txtReqNo.Text
                        End If
                    End With
                End If
            End If
            cmdTrans(1).Tag = ""
            Call LockData(True)
        Case 6  'Cancel
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            Ans = MsgBox("Do you want to cancel Requisition?", vbQuestion + vbYesNo)
            If Ans = vbNo Then Exit Sub
            Call cmdTransDisp(2)
            Call ClearEntries
            Call LockData(True)
            If Not AddNew Then
                Call DisplayRequisition(cmdTrans(1).Tag)
                cmdTrans(1).Tag = ""
            End If
            If MySource = "Job Order" Then cmdTrans(2).Enabled = False
    End Select
            
End Sub


Private Sub flxOrd_Click()
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
    
    If ActCol = 0 Or ActRow = 0 Or JustScroll Then Exit Sub
    Select Case ActCol
        Case 1    'qty
            With txtFlx
                If flxOrd.TextMatrix(ActRow, 1) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 1)
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
            Call QtyTip(True)
        Case 3 'PartNumber
            With txtPartNo
                If .Text <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 1) & "|" & flxOrd.TextMatrix(ActRow, 2) & "|" & flxOrd.TextMatrix(ActRow, 3) & "|" & flxOrd.TextMatrix(ActRow, 4) & "|" & flxOrd.TextMatrix(ActRow, 5) & "|" & flxOrd.TextMatrix(ActRow, 6) & "|" & flxOrd.TextMatrix(ActRow, 7) & "|" & flxOrd.TextMatrix(ActRow, 8)
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
    End Select
End Sub
                             
Private Sub flxOrd_DblClick()
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
                       
    If ActCol = 0 Then
        If flxOrd.TextMatrix(ActRow, 0) = "" Then
            '***Acces
            'flxOrd.Row = ActRow: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
            'LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'FOUND' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtReqID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(ActRow, 10) & "));"
            'mySQL
            flxOrd.Row = ActRow: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
            LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'FOUND' " & _
                           "WHERE SPSORDERDETAIL.SPSORDERID=" & txtReqID.Text & " AND SPSORDERDETAIL.SPAREPARTID=" & flxOrd.TextMatrix(ActRow, 10) & ";"
            
        Else
            flxOrd.TextMatrix(ActRow, 0) = ""
            '***Access
            'LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'BLANK' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtReqID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(ActRow, 10) & "));"
            'mySQl
            LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'BLANK' " & _
                           "WHERE SPSORDERDETAIL.SPSORDERID=" & txtReqID.Text & " AND SPSORDERDETAIL.SPAREPARTID=" & flxOrd.TextMatrix(ActRow, 10) & ";"
            
        End If
    End If
End Sub

Private Sub flxOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxOrd_Click
End Sub
                                                  
Private Sub flxOrd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If JustScroll Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuStatus)
        End If
    Else
        If Button = vbRightButton Then
            txtFlx.Visible = False
            Call PopupMenu(mnuRows)
        End If
    End If
End Sub

Private Sub flxParts_Click()
    Dim TtlPrice As Single
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 1)
        'Insert Additional retrieval here
        flxOrd.TextMatrix(ActRow, 8) = .TextMatrix(.Row, 5) 'QtyOnHand
        flxOrd.TextMatrix(ActRow, 4) = flxParts.TextMatrix(flxParts.Row, 2) 'Description
        flxOrd.TextMatrix(ActRow, 5) = Format$(flxParts.TextMatrix(flxParts.Row, 4), "###,##0.00") 'Price
        flxOrd.TextMatrix(ActRow, 7) = flxParts.TextMatrix(flxParts.Row, 3)    'Sparepart ID
        If Val(flxOrd.TextMatrix(ActRow, 1)) > 0 And Val(flxParts.TextMatrix(flxParts.Row, 5)) > 0 Then 'kung naa Qty and Price compute total
            TtlPrice = Val(flxOrd.TextMatrix(ActRow, 1)) * Val(flxParts.TextMatrix(flxParts.Row, 5))
            flxOrd.TextMatrix(ActRow, 6) = Format(TtlPrice, "#,###,##0.00")
            Call ComputeTotal
        Else
            flxOrd.TextMatrix(ActRow, 6) = "0.00"
            Call ComputeTotal
        End If
    End With
End Sub

Private Sub flxParts_EnterCell()
    Dim TtlPrice As Single
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 1)
        'Insert Additional retrieval here
        flxOrd.TextMatrix(ActRow, 8) = flxParts.TextMatrix(flxParts.Row, 5) 'QtyOnHand
        flxOrd.TextMatrix(ActRow, 4) = flxParts.TextMatrix(flxParts.Row, 2) 'Description
        flxOrd.TextMatrix(ActRow, 5) = Format$(flxParts.TextMatrix(flxParts.Row, 4), "###,##0.00") 'Price
        flxOrd.TextMatrix(ActRow, 7) = flxParts.TextMatrix(flxParts.Row, 3)    'Sparepart ID
        If Val(flxOrd.TextMatrix(ActRow, 1)) > 0 And Val(flxParts.TextMatrix(flxParts.Row, 5)) > 0 Then 'kung naa Qty and Price compute total
            TtlPrice = Val(flxOrd.TextMatrix(ActRow, 1)) * Val(flxParts.TextMatrix(flxParts.Row, 5))
            flxOrd.TextMatrix(ActRow, 6) = Format(TtlPrice, "#,###,##0.00")
            Call ComputeTotal
        Else
            flxOrd.TextMatrix(ActRow, 6) = "0.00"
            Call ComputeTotal
        End If
    End With
End Sub

Private Sub flxParts_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        If txtPartNo.Tag <> "" Then
            Tmp = Split(txtPartNo.Tag, "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxOrd.TextMatrix(ActRow, i + 1) = Tmp(i)
            Next i
        End If
        txtPartNo.Visible = False
        flxParts.Visible = False
        flxOrd.SetFocus
        flxOrd.Col = 3
    End If

End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        'If all fields are filled then addrow focus on next entry
        If Val(flxOrd.TextMatrix(ActRow, 1)) > Val(flxOrd.TextMatrix(ActRow, 8)) And flxOrd.TextMatrix(ActRow, 1) <> "" Then
            MsgBox "Qty Ordered must not be greater than Qty Available!", vbExclamation + vbOKOnly, "Validation"
            flxOrd.Row = ActRow
            flxOrd.Col = 2
            Call flxOrd_Click
            Exit Sub
        End If
        If flxOrd.TextMatrix(ActRow, 4) <> "" And flxOrd.TextMatrix(ActRow, 1) <> "" Then
            If ActRow = flxOrd.Rows - 1 Then flxOrd.AddItem ""
            flxOrd.SetFocus
            flxOrd.Row = flxOrd.Rows - 1
            flxOrd.Col = 3
        ElseIf flxOrd.TextMatrix(ActRow, 4) = "" Then   'No Description focus on Partnumber of current row
            flxOrd.SetFocus
            flxOrd.Col = 3
        ElseIf flxOrd.TextMatrix(ActRow, 1) = "" Then   'No Qty Focus on same column
            flxOrd.SetFocus
            flxOrd.Col = 1
        End If
        flxParts.Visible = False
    End If
End Sub


Private Sub Form_Load()

    If rsParts.State = adStateClosed Then Call LoadSpareParts
    Call ClearEntries
    Call LockData(True)
    If MySource = "Job Order" Then
        With JobOrder
            If .cboReq.Text <> "" Then
                Call DisplayRequisition(.cboReq.Text)
                If .txtOR.Text <> "" Then
                    If Val(txtPrintCopy.Text) > 0 Then
                        Call cmdTransDisp(3)    'Print and Exit
                    Else
                        Call cmdTransDisp(2)
                        cmdTrans(2).Enabled = False
                    End If
                Else
                        Call cmdTransDisp(2)
                        cmdTrans(2).Enabled = False
                End If
            Else
                Call cmdTrans_Click(0)
            End If
        End With
    Else
        Call cmdTransDisp(2)
    End If
    MotoRep.sBar.Panels(6).Text = "Spare Parts Requisition"
End Sub

Private Sub cmdTransDisp(MyCode As Integer)
    Dim i As Integer
    
    If MyCode = 1 Then  'Save and Cancel
        For i = 0 To 4
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
        cmdTrans(5).Left = 0
        cmdTrans(6).Left = cmdTrans(5).Width
    ElseIf MyCode = 2 Then  'Addnew to Exit
        For i = 0 To 4
            cmdTrans(i).Left = cmdTrans(i).Width * i
            cmdTrans(i).Enabled = True
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
    ElseIf MyCode = 3 Then  'Print and Exit only
        For i = 0 To 6
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(3).Visible = True
        cmdTrans(4).Visible = True
        cmdTrans(3).Left = 0
        cmdTrans(4).Left = cmdTrans(5).Width
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsParts.State = adStateOpen Then rsParts.Close
    MotoRep.sBar.Panels(6).Text = ""
    MySource = ""
End Sub


Private Sub mnuAdd_Click()
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 7) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    End If
    If flxOrd.Rows >= 13 Then
        MsgBox "You have reached the maximum Quantity per Requisition" & vbCrLf & "Please use another Requisition for additional Spare Parts", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    End If
    flxOrd.AddItem ""
    flxOrd.SetFocus
    flxOrd.Row = flxOrd.Rows - 1
    flxOrd.Col = 3
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        'If Not AddNew Then lstDelete.AddItem flxOrd.TextMatrix(ActRow, 7)
        If flxOrd.Rows = 2 Then
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        Else
            flxOrd.RemoveItem ActRow
        End If
        Call ComputeTotal
    End If
End Sub


Public Sub DisplayRequisition(MyTrans As String)
    Dim rsReq As New ADODB.Recordset
    Dim rsRDtls As New ADODB.Recordset
    Dim Qry As String, mySrc As String, MyUnit As String, myRem As String
    Dim TtlPrice As Single, OwnCode As Integer
    
    '***Access
    'Call ClearEntries
    'Qry = "SELECT REQUISITION.* From Requisition WHERE (((REQUISITION.REQNO)='" & MyTrans & "'));"
    'mySQl
    Qry = "SELECT REQUISITION.* From Requisition WHERE REQUISITION.REQNO='" & MyTrans & "';"
    
    rsReq.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsReq.RecordCount > 0 Then
        rsReq.MoveFirst
        txtJobNo.Text = rsReq!JobNo
        txtCustomer.Text = rsReq!CUSTOMER
        txtReqNo.Text = rsReq!REQNO
        txtDate.Text = rsReq!REQDATE
        txtPrintCopy.Text = IIf(IsNull(rsReq!PRINTCOPY), 0, rsReq!PRINTCOPY)
        txtORNo.Text = IIf(IsNull(rsReq!ORNO), "", rsReq!ORNO)
        txtInvNo.Text = IIf(IsNull(rsReq!InvoiceNO), "", rsReq!InvoiceNO)
        lblRemarks.Caption = IIf(IsNull(rsReq!Remarks), "", rsReq!Remarks)
        If txtORNo.Text <> "" Then lblRemarks.Caption = "O.R.#: " & txtORNo.Text & "  " & lblRemarks.Caption
        If txtInvNo.Text <> "" Then lblRemarks.Caption = "Invoice#: " & txtInvNo.Text & "  " & lblRemarks.Caption
        txtCustID.Text = IIf(IsNull(rsReq!CUSTOMERID), "", rsReq!CUSTOMERID)
        txtAmt.Text = IIf(IsNull(rsReq!Amount), 0, rsReq!Amount)
        If Val(txtCustID.Text) > 0 Then
            chkOnAcct.Value = 1
            chkOnAcct.Visible = True
        Else
            chkOnAcct.Value = 0
            chkOnAcct.Visible = False
        End If
        txtReqID.Text = rsReq!ReqID
    End If
    rsReq.Close
    
    With rsRDtls
        '***Access
        'Qry = "SELECT SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, [PRICE]*[QTY] AS TTLAMT, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.QTYONHAND " & _
              "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE (((SPSORDERDETAIL.REQID)=" & txtReqID.Text & "));"
        'mySQL
         Qry = "SELECT SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, SPSORDERDETAIL.PRICE*SPSORDERDETAIL.QTY AS TTLAMT, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.QTYONHAND " & _
              "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE SPSORDERDETAIL.REQID=" & txtReqID.Text & ";"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = 1
            Do While Not .EOF
                flxOrd.AddItem vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlAmt, "###,##0.00") & vbTab & !SparepartID & vbTab & !QtyOnHand
                .MoveNext
            Loop
        End If
        .Close
    End With
    Call ComputeTotal
    
    If MySource = "Job Order" Then
        If Val(JobOrder.txtCustID.Text) <> Val(txtCustID.Text) Then
            LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.CUSTOMER = '" & JobOrder.txtCustomer.Text & "', REQUISITION.CUSTOMERID = " & JobOrder.txtCustID.Text & " WHERE (((REQUISITION.REQID)=" & txtReqID.Text & "));"
            txtCustomer.Text = JobOrder.txtCustomer.Text
            txtCustID.Text = JobOrder.txtCustID.Text
        End If
    End If
        
        
End Sub
Private Sub SaveRequisition()
    Dim rsReq As New ADODB.Recordset
    Dim rsRDtls As New ADODB.Recordset
    Dim rsOnAccount As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String, nSkm As String, myTtl As Single, Ans As Integer
    Dim nPay As Integer, nTrm As Integer
    
    'Saving Table Spare Parts Invoice
    '***Access
    'myTtl = Val(Replace(txtTotal.Text, ",", ""))
    'If AddNew Then
    '    rsReq.Open "Requisition", LYRcon, adOpenKeyset, adLockOptimistic
    '    rsReq.AddNew
    'Else
    '    Qry = "SELECT REQUISITION.* From Requisition " & _
    '          "WHERE (((REQUISITION.REQID)=" & txtReqID.Text & "));"
    '    rsReq.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    rsReq.MoveFirst
    'End If
    'rsReq!JobNo = txtJobNo.Text
    'rsReq!CUSTOMER = txtCustomer.Text
    'rsReq!REQNO = txtReqNo.Text
    'rsReq!REQDATE = txtDate.Text
    'rsReq!Remarks = Left$(lblRemarks.Caption, 100)
    'rsReq!Amount = myTtl
    'rsReq!JOBORDERID = Val(txtJOid.Text)
    'If Val(txtCustID.Text) > 0 Then
    '    rsReq!CUSTOMERID = txtCustID.Text
    'End If
    'rsReq.Update
    'txtReqID.Text = rsReq!ReqID
    'rsReq.Close
    
    
    'mySQL
    myTtl = Val(Replace(txtTotal.Text, ",", ""))
    If AddNew Then
        LYRcon.Execute "INSERT INTO requisition( JOBNO, CUSTOMER, REQNO, REQDATE, REMARKS, AMOUNT, JOBORDERID )" & _
            "VALUES('" & txtJobNo.Text & "', '" & txtCustomer.Text & "', '" & txtReqNo.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
            "'" & Left$(lblRemarks.Caption, 100) & "', " & myTtl & ", " & Val(txtJOid.Text) & ");"
            
      Qry = "SELECT requisition.ReqID From Requisition Where requisition.JobNo= '" & txtJobNo.Text & "' AND requisition.REQNO = '" & txtReqNo.Text & "';"
      txtReqID.Text = GetFieldValue(Qry)
  
    
    End If
    If Not AddNew Then
        LYRcon.Execute "UPDATE Requisition SET JOBNO = '" & txtJobNo.Text & "', CUSTOMER = '" & txtCustomer.Text & "', " & _
            "REQNO = '" & txtReqNo.Text & "', REQDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
            "REMARKS = '" & Left$(lblRemarks.Caption, 100) & "', AMOUNT = " & myTtl & ", JOBORDERID = " & Val(txtJOid.Text) & " " & _
            "Where Requisition.ReqID = " & txtReqID.Text & ";"
    End If
    
    If Val(txtCustID.Text) > 0 Then
             MyCusID = txtCustID.Text
            LYRcon.Execute "UPDATE Requisition SET CUSTOMERID=" & txtCustID.Text & " Where requisition.ReqID = " & txtReqID.Text & ";"
    End If
        
    
    'Saving Table Spare Parts Order Details
    If Not AddNew Then
        '***Access
        'LYRcon.Execute "DELETE SPSORDERDETAIL.* From SPSORDERDETAIL " & _
                       "WHERE (((SPSORDERDETAIL.REQID)=" & txtReqID.Text & "));"
        LYRcon.Execute "DELETE From SPSORDERDETAIL " & _
                       "WHERE SPSORDERDETAIL.REQID=" & txtReqID.Text & ";"
    End If
    For i = 1 To flxOrd.Rows - 1
        With flxOrd
            LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( REQID, SPAREPARTID, QTY, UNIT, PRICE) " & _
                       "VALUES (" & txtReqID.Text & ", " & .TextMatrix(i, 7) & ", " & .TextMatrix(i, 1) & ", '" & .TextMatrix(i, 2) & "', " & Val(Replace(.TextMatrix(i, 5), ",", "")) & ");"
        End With
    Next i
    'Saving On Account Requisition to On Account Terms
'    If chkOnAcct.Value = 1 And Val(txtCustID.Text) > 0 Then
'        'Get Terms
'        Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
'              "WHERE (((ONACCTERMS.STATUS)='Default') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
'        With rsTerm
'            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'            If .RecordCount > 0 Then
'                .MoveFirst
'                nTrm = !Terms
'                nSkm = !PayScheme
'                nPay = !NPays
'            Else
'                nTrm = 30
'                nSkm = "MONTHLY"
'                nPay = 1
'            End If
'            .Close
'        End With
'
'        With rsOnAccount
'            If AddNew Then
'                .Open "ONACCTERMS", LYRcon, adOpenKeyset, adLockOptimistic
'                .AddNew
'            Else
'                Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
'                      "WHERE (((ONACCTERMS.REQUISITIONID)=" & txtReqID.Text & ") AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
'                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'                .MoveFirst
'            End If
'            !CUSTOMERID = txtCustID.Text
'            !PayScheme = nSkm
'            !Terms = nTrm
'            !NPays = nPay
'            !PayStart = StartPay(nSkm)
'            !BillingAmt = Format(myTtl / nPay, "####0.00")
'            !JOBORDERNO = txtJobNo.Text
'            !BALANCE = myTtl
'            !AcctAmount = myTtl
'            !REQUISITIONID = txtReqID.Text
'            !Status = "Temp"
'            .Update
'            .Close
'        End With
'    End If
End Sub

Function StartPay(S As String) As Date
    Dim iPS As Integer, sPay As Date
    
    If S = "WEEKLY" Then iPS = 7
    If S = "BI-MONTHLY" Then iPS = 15
    If S = "MONTHLY" Then iPS = 30
    If iPS = 7 Then
        If Day(txtDate.Text) < 5 Then
            sPay = DateSerial(Year(txtDate.Text), Month(txtDate.Text), 7)
        ElseIf Day(txtDate.Text) < 12 Then
            sPay = DateSerial(Year(txtDate.Text), Month(txtDate.Text), 14)
        ElseIf Day(txtDate.Text) < 19 Then
            sPay = DateSerial(Year(txtDate.Text), Month(txtDate.Text), 21)
        ElseIf Day(txtDate.Text) < 28 Then
            sPay = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, txtDate.Text)), Month(DateAdd("m", 1, txtDate.Text)), 1))
        Else
            sPay = DateAdd("m", 1, DateSerial(Year(txtDate.Text), Month(txtDate.Text), 7))
        End If
        StartPay = sPay
    ElseIf iPS = 15 Then
        StartPay = DateAdd("d", iPS, txtDate.Text)
    ElseIf iPS = 30 Then
        StartPay = DateAdd("m", 1, txtDate.Text)
    End If
End Function


Function ValidateRequisition() As String
    Dim MyErr As String, Qry As String
    Dim i As Integer
    
    MyErr = ""
    If txtJobNo.Text = "" Then
        MyErr = "Incomplete entry on field Job Order No."
    Else
        If AddNew Then
            '***Access
            'Qry = "SELECT SRVCJOBORDER.JOBORDERNO " & _
                  "From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtJobNo.Text & "'));"
            'mySQL
            Qry = "SELECT SRVCJOBORDER.JOBORDERNO " & _
                  "From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtJobNo.Text & "';"
            
            If Not CtrlNumExist(Qry) Then
                MyErr = "Job Order No. " & txtJobNo.Text & " does not exist!"
            End If
        End If
    End If
    If Not IsDate(txtDate.Text) Then MyErr = MyErr & vbCrLf & "Invalid entry on field Date"
    If txtReqNo.Text = "" Then
        MyErr = MyErr & vbCrLf & "Incomplete entry on field Requisition No."
    Else
        If (Not AddNew And cmdTrans(1).Tag <> txtReqNo.Text) Or AddNew Then
            '***Access
            'Qry = "SELECT REQUISITION.REQNO From Requisition " & _
                  "WHERE (((REQUISITION.REQNO)='" & txtReqNo.Text & "'));"
            'mySQL
            Qry = "SELECT REQUISITION.REQNO From Requisition " & _
                  "WHERE REQUISITION.REQNO='" & txtReqNo.Text & "';"
            
            If CtrlNumExist(Qry) Then
                MyErr = "Requisition No. " & txtReqNo.Text & " already exist!"
            End If
        End If
    End If
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 1) = "" And flxOrd.TextMatrix(flxOrd.Rows - 1, 7) = "" Then
        If flxOrd.Rows = 2 Then
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        Else
            flxOrd.RemoveItem flxOrd.Rows - 1
        End If
    End If
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 1) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 7) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row of Ordered Spare Parts"
    If chkOnAcct.Value = 1 Then
        If Val(txtCustID.Text) = 0 Then MyErr = MyErr & vbCrLf & "Customer ID not Updated! Pls. Inform your Programmer"
    End If
    For i = 1 To flxOrd.Rows - 1
        If flxOrd.TextMatrix(i, 1) = "" Then MyErr = MyErr & vbCrLf & "Ordered Spare Parts Sales @ Row " & i & " no Qty.!"
    Next i
    ValidateRequisition = MyErr
End Function

Private Sub LoadSpareParts()
    Dim rsSet As New ADODB.Recordset
    Dim Qry As String
    Dim pPct As Single
        
    
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.SPAREPARTID, SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    With flxParts
        .Width = flxOrd.ColWidth(4)
        .ColWidth(0) = 0
        .ColWidth(1) = flxOrd.ColWidth(4)
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
    End With
End Sub


Private Sub RXResizer1_ResizeComplete()
    With flxOrd
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.09
        .ColWidth(2) = (.Width - 300) * 0.09
        .ColWidth(3) = (.Width - 300) * 0.23
        .ColWidth(4) = (.Width - 300) * 0.27
        .ColWidth(5) = (.Width - 300) * 0.15
        .ColWidth(6) = (.Width - 300) * 0.15
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .Col = 6: .Row = 1
        txtTotal.Left = flxOrd.Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5)
        txtTotal.Width = .ColWidth(6)
        flxParts.ColWidth(1) = .ColWidth(4)
        Label3.Left = txtTotal.Left - (Label3.Width + 50)
        Label3.Top = txtTotal.Top
    End With
    lblHeading(0).Left = Frame1.Left
    lblHeading(1).Left = Frame1.Left + 30
    lblHeading(0).Top = Frame1.Top - (lblHeading(0).Height + 30)
    lblHeading(1).Top = lblHeading(0).Top + 30
    lblHeading(2).Top = lblHeading(0).Top
    lblHeading(3).Top = lblHeading(1).Top
    lblHeading(2).Left = Frame1.Left + Frame1.Width - lblHeading(2).Width
    lblHeading(3).Left = Frame1.Left + Frame1.Width - lblHeading(3).Width - 30
    lblHeading(4).Top = flxOrd.Top - lblHeading(4).Height + 30
    lblHeading(5).Top = flxOrd.Top - lblHeading(5).Height
    lblHeading(4).Left = flxOrd.Left
    lblHeading(5).Left = flxOrd.Left + 30
End Sub



Private Sub QtyTip(Disp As Boolean)
    If Disp Then
        With txtQtyTip
            .Top = flxOrd.CellTop + flxOrd.Top + 290
            .Left = flxOrd.Left + flxOrd.ColWidth(0) + 300
            .Text = "Qty. on Hand: " & flxOrd.TextMatrix(ActRow, 8)
            .Visible = True
            .ZOrder
        End With
    Else
        txtQtyTip.Visible = False
    End If
End Sub


Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFlx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtFlx.Text = Empty
        flxOrd.TextMatrix(ActRow, 1) = IIf(txtFlx.Tag = "XcdQty", "", txtFlx.Tag)
        txtFlx.Tag = "skiplostfocus"
        flxOrd.SetFocus
        flxOrd.Col = 1
        txtFlx.Visible = False
        Call QtyTip(False)
    End If
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtReqNo.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        flxOrd.SetFocus
        flxOrd.Col = 3  'partnumber
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxOrd.SetFocus
        flxOrd.Col = 3  'partnumber
    End If
End Sub

Private Sub txtFlx_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If ActCol = 1 Then
        If KeyAscii > 26 Then
            If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then
        Call txtFlx_LostFocus
        If txtFlx.Tag = "XcdQty" Then txtFlx.Tag = "": Exit Sub
        'If all fields are filled then addrow focus on next entry
        If flxOrd.TextMatrix(ActRow, 4) <> "" And flxOrd.TextMatrix(ActRow, 1) <> "" Then
            If ActRow = flxOrd.Rows - 1 Then flxOrd.AddItem ""
            flxOrd.SetFocus
            flxOrd.Row = flxOrd.Rows - 1
            flxOrd.Col = 3
        ElseIf flxOrd.TextMatrix(ActRow, 4) = "" Then   'No Description focus on Partnumber of current row
            flxOrd.SetFocus
            flxOrd.Col = 3
        ElseIf flxOrd.TextMatrix(ActRow, 1) = "" Then   'No Qty Focus on same column
            flxOrd.SetFocus
            flxOrd.Col = 1
        End If
    End If
End Sub

Private Sub txtFlx_LostFocus()
    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    With flxOrd
        ' Qty > Available and SPID<>"" then
        If Val(txtFlx.Text) > Val(.TextMatrix(ActRow, 8)) And .TextMatrix(ActRow, 7) <> "" Then
            MsgBox "Order Qty. must not exceed Qty. On-Hand (" & .TextMatrix(ActRow, 8) & ")!", vbCritical + vbOKOnly, "Validation"
            txtFlx.Tag = "XcdQty"
            txtFlx.SelStart = 0
            txtFlx.SelLength = Len(txtFlx.Text)
            txtFlx.SetFocus
            Exit Sub
        End If
    End With
    If InStr(flxOrd.TextMatrix(ActRow, 4), "ENGINE OIL") > 0 Or InStr(flxOrd.TextMatrix(ActRow, 4), "OIL 4T") > 0 Then
        flxOrd.TextMatrix(ActRow, 2) = IIf(Val(flxOrd.TextMatrix(ActRow, 1)) > 1, "Qrts.", "Qrt.")
    Else
        flxOrd.TextMatrix(ActRow, 2) = IIf(Val(flxOrd.TextMatrix(ActRow, 1)) > 1, "Pcs.", "Pc.")
    End If
    flxOrd.TextMatrix(ActRow, ActCol) = txtFlx.Text
    Call ComputeTotal
    txtFlx.Visible = False
    Call QtyTip(False)
End Sub

Private Sub ComputeTotal()
    Dim Tmp As Single
    Dim i As Integer, j As Integer
    
    With flxOrd
        Tmp = 0
        For j = 1 To .Rows - 1
            .TextMatrix(j, 6) = Format$(Val(.TextMatrix(j, 1)) * Val(Replace(.TextMatrix(j, 5), ",", "")), "###,##0.00")
            Tmp = Tmp + Val(Replace(.TextMatrix(j, 6), ",", ""))
        Next j
        txtTotal.Text = Format$(Tmp, "#,##0.00")
    End With
    If Val(txtAmt.Text) > 0 Then
        If Tmp <> Val(txtAmt.Text) Then
            txtAmt.Text = Tmp
            'old LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.AMOUNT = " & Tmp & " WHERE (((REQUISITION.REQID)=" & txtReqID.Text & "));"
        End If
    End If
            
End Sub

Private Sub txtJobNo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Call txtJobNo_LostFocus
        txtCustomer.SetFocus
    End If

End Sub

Private Sub txtJobNo_LostFocus()
    Dim Qry As String
        
    If txtJobNo.Text <> "" Then
        '***Access
        'Qry = "SELECT SRVCJOBORDER.JOBORDERNO " & _
              "From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtJobNo.Text & "'));"
        'mySQL
        Qry = "SELECT SRVCJOBORDER.JOBORDERNO " & _
              "From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtJobNo.Text & "';"
        
        If Not CtrlNumExist(Qry) Then
            MsgBox "Job Order No. " & txtJobNo.Text & " does not exist"
            txtJobNo.SelStart = 0
            txtJobNo.SelLength = Len(txtJobNo.Text)
        Else
            Call GetJOCustomer(txtJobNo.Text)
        End If
    End If
End Sub

Function GetJOCustomer(MyJo As String) As String
    Dim rsJO As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCJOBORDER.JOBORDERID, SRVCJOBORDER.CUSTOMERID, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.PLATENO FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO " & _
          "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtJobNo.Text & "'));"
    'mySQL
    Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCJOBORDER.JOBORDERID, SRVCJOBORDER.CUSTOMERID, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.PLATENO FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO " & _
          "WHERE SRVCJOBORDER.JOBORDERNO='" & txtJobNo.Text & "';"
    
    With rsJO
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            txtJOid.Text = !JOBORDERID
            If !CUSTOMERID > 0 Then
                txtCustID.Text = !CUSTOMERID
                txtJOid.Text = !JOBORDERID
                lblRemarks.Caption = IIf(IsNull(!DESIGNATION), "", !DESIGNATION & ": ") & !CUSTOMER & " Plate No. " & IIf(IsNull(!PLATENO), "", !PLATENO)
                chkOnAcct.Visible = True
                chkOnAcct.Value = 1
                Call GetCustomerTerm
            Else
                txtCustomer.Text = IIf(IsNull(!CUSTOMER), "", !CUSTOMER)
                chkOnAcct.Visible = False
                chkOnAcct.Value = 0
            End If
        Else
            GetJOCustomer = ""
        End If
        .Close
    End With
End Function

Private Sub txtPartNo_Change()
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
            .Col = 0: .ColSel = 3
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            flxOrd.TextMatrix(ActRow, 8) = flxParts.TextMatrix(flxParts.Row, 5) 'QtyOnHand
            flxOrd.TextMatrix(ActRow, 4) = flxParts.TextMatrix(flxParts.Row, 2) 'Description
            flxOrd.TextMatrix(ActRow, 5) = Format$(flxParts.TextMatrix(flxParts.Row, 4), "###,##0.00") 'Price
            flxOrd.TextMatrix(ActRow, 7) = flxParts.TextMatrix(flxParts.Row, 3)    'Sparepart ID
            If Val(flxOrd.TextMatrix(ActRow, 1)) > 0 And Val(flxParts.TextMatrix(flxParts.Row, 5)) > 0 Then 'kung naa Qty and Price compute total
                TtlPrice = Val(flxOrd.TextMatrix(ActRow, 1)) * Val(flxParts.TextMatrix(flxParts.Row, 5))
                flxOrd.TextMatrix(ActRow, 6) = Format(TtlPrice, "#,###,##0.00")
                Call ComputeTotal
            Else
                flxOrd.TextMatrix(ActRow, 6) = "0.00"
                Call ComputeTotal
            End If
            If InStr(flxOrd.TextMatrix(ActRow, 4), "ENGINE OIL") > 0 Or InStr(flxOrd.TextMatrix(ActRow, 4), "OIL 4T") > 0 Then
                flxOrd.TextMatrix(ActRow, 2) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Qrts.", "Qrt.")
            Else
                flxOrd.TextMatrix(ActRow, 2) = IIf(Val(flxOrd.TextMatrix(ActRow, 2)) > 1, "Pcs.", "Pc.")
            End If
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            flxOrd.TextMatrix(ActRow, 1) = ""
            flxOrd.TextMatrix(ActRow, 5) = ""
            flxOrd.TextMatrix(ActRow, 6) = ""
            flxOrd.TextMatrix(ActRow, 7) = ""
            flxOrd.TextMatrix(ActRow, 8) = ""
            flxOrd.TextMatrix(ActRow, 4) = ""
            flxOrd.TextMatrix(ActRow, 2) = ""
            .Visible = False
            Call ComputeTotal
        End If
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
            flxOrd.Col = 3
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
        Call txtPartNo_LostFocus
        'If all fields are filled then addrow focus on next entry
        If Val(flxOrd.TextMatrix(ActRow, 1)) > Val(flxOrd.TextMatrix(ActRow, 8)) And flxOrd.TextMatrix(ActRow, 1) <> "" Then
            MsgBox "Qty Ordered must not be greater than Qty Available!", vbExclamation + vbOKOnly, "Validation"
            flxOrd.Row = ActRow
            flxOrd.Col = 1
            Call flxOrd_Click
            Exit Sub
        End If
        If flxOrd.TextMatrix(ActRow, 4) <> "" And flxOrd.TextMatrix(ActRow, 1) <> "" Then
            If ActRow = flxOrd.Rows - 1 Then flxOrd.AddItem ""
            flxOrd.SetFocus
            flxOrd.Row = flxOrd.Rows - 1
            flxOrd.Col = 3
        ElseIf flxOrd.TextMatrix(ActRow, 4) = "" Then   'No Description focus on Partnumber of current row
            flxOrd.SetFocus
            flxOrd.Col = 3
        ElseIf flxOrd.TextMatrix(ActRow, 1) = "" Then   'No Qty Focus on same column
            flxOrd.SetFocus
            flxOrd.Col = 1
        End If
        flxParts.Visible = False
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    If flxOrd.TextMatrix(ActRow, 7) = "" Then
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
        txtPartNo.Visible = False
        flxParts.Visible = False
        Exit Sub
    End If
    flxOrd.TextMatrix(ActRow, 3) = txtPartNo.Text
    txtPartNo.Visible = False
    flxParts.Visible = False
End Sub

Private Sub txtReqNo_GotFocus()
    If txtReqNo.Text <> "" Then
        txtReqNo.SelStart = 0
        txtReqNo.SelLength = Len(txtReqNo.Text)
    End If
End Sub

Private Sub txtReqNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtReqNo_LostFocus()
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT REQUISITION.REQNO From Requisition " & _
          "WHERE (((REQUISITION.REQNO)='" & txtReqNo.Text & "'));"
    'mySQl
    Qry = "SELECT REQUISITION.REQNO From Requisition " & _
          "WHERE REQUISITION.REQNO='" & txtReqNo.Text & "';"
    
    If CtrlNumExist(Qry) Then
            MsgBox "Requisition No. " & txtReqNo.Text & " already exist!"
    End If
End Sub
