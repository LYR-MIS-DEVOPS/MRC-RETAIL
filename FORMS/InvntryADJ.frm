VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form InvntryADJ 
   Caption         =   "Inventory Adjustments"
   ClientHeight    =   8070
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   11520
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtChrgAmt 
      Height          =   285
      Left            =   7200
      TabIndex        =   44
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboEmp 
      Height          =   315
      Left            =   5760
      TabIndex        =   43
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid flxEmp 
      Height          =   855
      Left            =   6480
      TabIndex        =   42
      Tag             =   "XYF"
      Top             =   7080
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1508
      _Version        =   393216
      Cols            =   4
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      ScrollBars      =   2
      FormatString    =   "       |^NAME OF EMPLOYEE                   |^ CHARGE AMT."
   End
   Begin VB.CheckBox chkCharges 
      Caption         =   "APPLY CHARGES TO EMPLOYEES:"
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
      Left            =   7800
      TabIndex        =   41
      Tag             =   "XYF"
      Top             =   6840
      Width           =   3135
   End
   Begin VB.TextBox txtRef 
      Height          =   285
      Left            =   7080
      MaxLength       =   20
      TabIndex        =   40
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid flxPRN 
      Height          =   975
      Left            =   3120
      TabIndex        =   39
      Top             =   9360
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   5
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   375
      Left            =   6480
      TabIndex        =   38
      Top             =   5760
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1800
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   3240
      TabIndex        =   37
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtQtyA 
      Height          =   285
      Left            =   4440
      TabIndex        =   36
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   3240
      TabIndex        =   35
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1815
      Left            =   480
      TabIndex        =   34
      Top             =   8640
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
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   480
      TabIndex        =   31
      Tag             =   "XYF"
      Top             =   5685
      Width           =   5895
      Begin VB.TextBox txtRemarks 
         Height          =   645
         Left            =   240
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   1440
         Width           =   5535
      End
      Begin VB.TextBox txtReason 
         Height          =   645
         Left            =   240
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label lbl2 
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
         Left            =   240
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "REASON FOR ADJUSTMENT:"
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
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   360
      TabIndex        =   30
      Text            =   "Qty. Purchased: 0"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   8040
      TabIndex        =   29
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   27
      Tag             =   "XYF"
      Top             =   5685
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
      Left            =   8760
      TabIndex        =   17
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxAdj 
      Height          =   3375
      Left            =   480
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   2280
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   "^    |<PART NUMBER      |<DESCRIPTION          |^TYPE         |^REFERENCE  |^ QTY. ADJ.      |>PRICE |>TOTAL |^             |"
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
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   480
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   720
      Width           =   10695
      Begin VB.ComboBox cboAcctTitle 
         Height          =   315
         Left            =   1800
         TabIndex        =   47
         Tag             =   "XYF"
         Top             =   890
         Width           =   3255
      End
      Begin VB.TextBox txtAdjType 
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
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   200
         Width           =   3255
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   538
         Width           =   3255
      End
      Begin VB.TextBox txtAdjDate 
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
         Left            =   8520
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   538
         Width           =   1815
      End
      Begin VB.TextBox txtAdjNo 
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
         Left            =   8520
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   200
         Width           =   1815
      End
      Begin lvButton.lvButtons_H cmdCardexRed 
         Height          =   375
         Left            =   9120
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Cardex Red"
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
         Image           =   "InvntryADJ.frx":0000
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCardex 
         Height          =   375
         Left            =   7680
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Cardex"
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
         Image           =   "InvntryADJ.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "ACCOUNT TITLE:"
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
         TabIndex        =   49
         Tag             =   "XYF"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "REFERENCE NO."
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
         TabIndex        =   48
         Tag             =   "XYF"
         Top             =   568
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ADJUSTMENT DATE"
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
         Left            =   6720
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   570
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ADJUSTMENT NO."
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
         Left            =   6720
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   225
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "ADJUSTMENT TYPE:"
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
         Width           =   1815
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Add New"
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
      Image           =   "InvntryADJ.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   19
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "Su&mmary"
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
      Image           =   "InvntryADJ.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4500
      TabIndex        =   20
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "InvntryADJ.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   21
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "InvntryADJ.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7500
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "InvntryADJ.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9000
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "InvntryADJ.frx":2736
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   24
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "InvntryADJ.frx":2CD0
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPrepBy 
      Height          =   285
      Left            =   7800
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   6240
      Width           =   3255
   End
   Begin VB.TextBox txtAprvBy 
      Height          =   285
      Left            =   7800
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   6525
      Width           =   3255
   End
   Begin VB.Label lblTotal 
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
      Left            =   7920
      TabIndex        =   28
      Tag             =   "XYF"
      Top             =   5700
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   6480
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "PREPARED BY:"
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
      Left            =   6480
      TabIndex        =   25
      Tag             =   "XYF"
      Top             =   6240
      Width           =   1335
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11760
      m_MDI_Height    =   8640
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADJUSTMENTS"
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
      Left            =   9495
      TabIndex        =   15
      Tag             =   "TRF"
      Top             =   450
      Width           =   1440
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
      Caption         =   "ADJUSTMENTS"
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
      Left            =   9525
      TabIndex        =   16
      Tag             =   "TRF"
      Top             =   480
      Width           =   1440
   End
   Begin VB.Menu mnuADJ 
      Caption         =   "SpareParts"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row         Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCardex 
         Caption         =   "Show Cardex"
      End
      Begin VB.Menu mnuCardexRed 
         Caption         =   "Show Cardex Red"
      End
   End
   Begin VB.Menu mnuEmp 
      Caption         =   "Employee"
      Visible         =   0   'False
      Begin VB.Menu mnuAddEmp 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuDelEmp 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "InvntryADJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim rsEmployee As New ADODB.Recordset
Dim ActRow As Integer, ActCol As Integer, BtnIndx As Integer
Dim AddNew As Boolean, LockFlx As Boolean, DispOnly As Boolean, PrtSrch As Boolean, PrtBln As Boolean, CustSrch As Boolean

'New Variable Declaration
'By: ARL
'Date: 2020-08-28
Dim rsAcctTitle As New ADODB.Recordset
Dim arrAcctTitle() As String
Dim Qry As String
Dim myAcctCode As String
Dim AcctSrch As Boolean

Public Sub CboDisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Qry = "call accounting_corp.sp_chart_of_accts_list()"
    If rsAcctTitle.State = adStateOpen Then rsAcctTitle.Close
    rsAcctTitle.Open Qry, LYRcon, adOpenStatic, adLockPessimistic
    If rsAcctTitle.State = 1 Then
       Call modObjTools.addToListObjectWithArrayGivenField(cboAcctTitle, rsAcctTitle, arrAcctTitle, "Acct_Code", "Acct_Title")
    End If
    If rsAcctTitle.State <> 0 Then rsAcctTitle.Close
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Private Sub ClearEntries()
    txtAdjType.Text = ""
    txtRefNo.Text = ""
    txtAdjNo.Text = ""
    txtAdjDate.Text = ""
    txtReason.Text = ""
    txtTotal.Text = ""
    txtRemarks.Text = ""
    txtPrepBy.Text = ""
    txtAprvBy.Text = ""
    flxAdj.Rows = 1
    flxAdj.Rows = 2
    txtPrnCopy.Text = ""
    cboAcctTitle.Text = ""
    chkCharges.Value = 0
    flxEmp.Rows = 1
    flxEmp.Rows = 2
End Sub

Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    Frame2.Enabled = Not bln
    txtPrepBy.Locked = bln
    txtAprvBy.Locked = bln
    LockFlx = bln
    chkCharges.Enabled = Not bln
End Sub

Public Sub ComputeTotal()
    Dim GrdTtl As Single, i As Integer
    Dim TtlPrice As Currency
    
    GrdTtl = 0
    For i = 1 To flxAdj.Rows - 1
        TtlPrice = Val(flxAdj.TextMatrix(i, 5)) * PureNum(flxAdj.TextMatrix(i, 6))
        GrdTtl = GrdTtl + TtlPrice
        flxAdj.TextMatrix(i, 7) = Format(TtlPrice, "#,##0.00")
    Next i
    txtTotal.Text = Format$(GrdTtl, "#,##0.00")
    
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

Private Sub cboAcctTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then AcctSrch = True
End Sub

Private Sub cboAcctTitle_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    
        If cboAcctTitle.Text = "" Then
            Exit Sub
        End If
        
        cbolistIndex = modObjTools.getListIndex(cboAcctTitle, cboAcctTitle.Text)
        
        If cbolistIndex < 0 Then
            cboAcctTitle.SetFocus
        Else
            cboAcctTitle.ListIndex = cbolistIndex
        End If
        txtAdjNo.SetFocus
    End If
End Sub

Private Sub cboAcctTitle_LostFocus()
    myAcctCode = ""
    If cboAcctTitle.Text = "" Then Exit Sub
    myAcctCode = arrAcctTitle(cboAcctTitle.ListIndex, 0)
End Sub

Private Sub cboEmp_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If CustSrch Then
        CustSrch = False
        Exit Sub
    End If
    With cboEmp
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsEmployee.MoveFirst
            rsEmployee.Move i
            flxEmp.TextMatrix(ActRow, 3) = rsEmployee!EMPID
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CustSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            flxEmp.TextMatrix(ActRow, 3) = ""
        End If
    End With
End Sub

Private Sub cboEmp_Click()
    Call cboEmp_Change
End Sub

Private Sub cboEmp_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboEmp.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboEmp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then CustSrch = True
    If KeyCode = vbKeyEscape Then
        cboEmp.Text = Empty
        flxEmp.TextMatrix(ActRow, 1) = cboEmp.Tag
        cboEmp.Tag = "skiplostfocus"
        flxEmp.SetFocus
        flxEmp.Col = 1
        cboEmp.Visible = False
    End If
End Sub

Private Sub cboEmp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxEmp.SetFocus
        flxEmp.Col = 2
    End If
End Sub

Private Sub cboEmp_LostFocus()
    If cboEmp.Tag = "skiplostfocus" Then
        cboEmp.Tag = ""
        Exit Sub
    End If
    If Val(flxEmp.TextMatrix(ActRow, 3)) = 0 Then
        MsgBox "Please select a name from the list!", vbExclamation + vbOKOnly, "Validation Error"
        cboEmp.Visible = False
        Exit Sub
    End If
    flxEmp.TextMatrix(ActRow, 1) = cboEmp.Text
    cboEmp.Visible = False
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cboType.Text = Empty
        flxAdj.TextMatrix(ActRow, 3) = cboType.Tag
        cboType.Tag = "skiplostfocus"
        flxAdj.SetFocus
        flxAdj.Col = 2
        cboType.Visible = False
    End If
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxAdj.SetFocus
        flxAdj.Row = ActRow
        flxAdj.Col = 4
    End If
End Sub

Private Sub cboType_LostFocus()
    If cboType.Tag = "skiplostfocus" Then
        cboType.Tag = ""
        Exit Sub
    End If
    If cboType.ListIndex = -1 Then
        flxAdj.TextMatrix(ActRow, 3) = ""
        cboType.Visible = False
        Exit Sub
    End If
    flxAdj.TextMatrix(ActRow, 3) = cboType.Text
    cboType.Visible = False
End Sub

Private Sub chkCharges_Click()
    Dim Qry As String
    
    If chkCharges.Value = 1 Then
        flxEmp.Visible = True
        If rsEmployee.State = adStateOpen Then Exit Sub
        With rsEmployee
            '***Access
            'Qry = "SELECT [LASTNAME] & IIf(IsNull([FIRSTNAME]),'',', ' & [FIRSTNAME]) & IIf(IsNull([MIDDLENAME]),'',' ' & Left$([MIDDLENAME],1) & '.') AS EMPNAME, EMPLOYEES.EMPID " & _
                  "From Employees Where (((Employees.INACTIVE) = False)) ORDER BY [LASTNAME] & IIf(IsNull([FIRSTNAME]),'',', ' & [FIRSTNAME]) & IIf(IsNull([MIDDLENAME]),'',' ' & Left$([MIDDLENAME],1) & '.');"
            'mySQL
            Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPNAME, EMPLOYEES.EMPID " & _
                  "From Employees Where Employees.INACTIVE = 0 ORDER BY concat(Employees.LASTNAME,' ', left(Employees.FIRSTNAME,1),'.');"
            
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                cboEmp.Clear
                Do While Not .EOF
                    cboEmp.AddItem !EmpName
                    .MoveNext
                Loop
            End If
        End With
    Else
        flxEmp.Visible = False
    End If
End Sub

Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub

Private Sub cmdCardex_Click()
    SPCardex.PNadj = "INVADJ"
    SPCardex.Show vbModal
End Sub

Private Sub cmdCardexRed_Click()
    SPCardexRed.PNadj = "INVADJ"
    SPCardexRed.Show vbModal
End Sub

Private Sub LoadSpareParts()
    Dim Qry As String
        
    
    Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    With flxParts
        .Width = flxAdj.ColWidth(4)
        .ColWidth(0) = 0
        .ColWidth(1) = 0    'SPid
        .ColWidth(2) = flxAdj.ColWidth(4)   'PartNo
        .ColWidth(3) = 0    'Description
        .ColWidth(4) = 0    'Store Price
        .ColWidth(5) = 0    'Qty On Hand
    End With
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 12) As String
    Dim Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=52));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=52;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            BtnIndx = 4
            tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            txtAdjType.Text = "INVENTORY"
            txtAdjDate.Text = Date
            AddNew = True
            txtAdjNo.Text = GetAdjustmentNo
            txtPrepBy.Text = GetSettings("OIC")
            txtAprvBy.Text = GetSettings("MANAGER")
            
            txtRefNo.SetFocus
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=52));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=52;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If Val(txtPrnCopy.Text) > 0 Then
                MsgBox "Editing is not allowed because Adjustment was already printed.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            ElseIf Val(txtAdjNo.Text) = 0 Then
                MsgBox "No available Adjustments to Edit", vbInformation + vbOKOnly, "Confirm P.O."
                Exit Sub
            End If
            BtnIndx = 4
            tmrFloat.Enabled = True
            AddNew = False
            cmdTrans(1).Tag = txtAdjNo.Text
            Call LockData(False)
        Case 2  'Summary
            'InvAdjSumm.Show
            FrmSummary.mySrc = "Inventory Adjustment"
            FrmSummary.Show
            Unload Me
            Exit Sub
        Case 3  'Print
            Ans = MsgBox("Do you want to Print Inventory Adjustment?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            flxPRN.Rows = 1
            For i = 1 To flxAdj.Rows - 1
                flxPRN.AddItem vbTab & flxAdj.TextMatrix(i, 1) & vbTab & flxAdj.TextMatrix(i, 3) & vbTab & flxAdj.TextMatrix(i, 4) & vbTab & flxAdj.TextMatrix(i, 5)
            Next i
            PrnVal(0) = "TYPE: INVENTORY"
            PrnVal(1) = "DATE: " & txtAdjDate.Text
            PrnVal(2) = "REF#: " & txtRefNo.Text
            PrnVal(3) = "PART NUMBER"
            PrnVal(4) = "TYPE"
            PrnVal(5) = "REFERENCE"
            PrnVal(6) = "ADJ. QTY."
            PrnVal(7) = txtPrepBy.Text
            PrnVal(8) = txtAprvBy.Text
            PrnVal(9) = IIf(chkCharges.Value = 1, "NOTE: CHARGES APPLIED TO EMPLOYEES", "")
            PrnVal(10) = "O.I.C."
            PrnVal(11) = "MANAGER"
            PrnVal(12) = IIf(txtReason.Text <> "", "REASON: " & txtReason.Text, "")
            Call PrintMyForm(PrnVal, "INVENTORY ADJ", flxPRN)
            If Val(txtPrnCopy.Text) = 0 Then
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
                Call UpdateSPITR
            Else
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
            End If
            '***Access
            'LYRcon.Execute "UPDATE INVNTRYADJ SET INVNTRYADJ.PRINTCOPY = [INVNTRYADJ].[PRINTCOPY]+1 WHERE (((INVNTRYADJ.ADJNO)='" & txtAdjNo.Text & "'));"
            'mySQL
            LYRcon.Execute "UPDATE INVNTRYADJ SET INVNTRYADJ.PRINTCOPY = INVNTRYADJ.PRINTCOPY+1 WHERE INVNTRYADJ.ADJNO='" & txtAdjNo.Text & "';"
            
            'Print process follows
            
        Case 4  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.ToolbarButton.Visible = True
            Unload Me
            Exit Sub
        Case 5  'Save
            myMsg = ValidateInvAdj
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Adjustment?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SaveInvAdj
            MsgBox "Adjustment Successfully Saved!", vbInformation + vbOKOnly, "Save"
                
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            Call LockData(True)
            
        Case 6  'Cancel
            Ans = MsgBox("Do you want to Cancel Adjustment?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            If AddNew Then
                Call ClearEntries
            Else
                Call DisplayInvAdj(cmdTrans(1).Tag)
            End If
            Call LockData(True)
    End Select

End Sub

Private Sub flxAdj_Click()
    ActRow = flxAdj.Row
    ActCol = flxAdj.Col
    
    If ActCol = 0 Or ActRow = 0 Or LockFlx Then Exit Sub
    Select Case ActCol
        Case 1 'PartNumber
            With txtPartNo
                If flxAdj.TextMatrix(ActRow, 1) <> "" Then flxAdj.TextMatrix(ActRow, 9) = flxAdj.TextMatrix(ActRow, 1) & "|" & flxAdj.TextMatrix(ActRow, 2) & "|" & flxAdj.TextMatrix(ActRow, 3) & "|" & flxAdj.TextMatrix(ActRow, 4) & "|" & flxAdj.TextMatrix(ActRow, 5) & "|" & flxAdj.TextMatrix(ActRow, 6) & "|" & flxAdj.TextMatrix(ActRow, 7) & "|" & flxAdj.TextMatrix(ActRow, 8)
                .Top = flxAdj.CellTop + flxAdj.Top
                .Left = flxAdj.CellLeft + flxAdj.Left
                .Width = flxAdj.CellWidth
                .Text = flxAdj.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 3    'Type
            With cboType
                If flxAdj.TextMatrix(ActRow, 3) <> "" Then .Tag = flxAdj.TextMatrix(ActRow, 3)
                .Top = flxAdj.CellTop + flxAdj.Top
                .Left = flxAdj.CellLeft + flxAdj.Left
                .Width = flxAdj.CellWidth
                .Text = flxAdj.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 4  'Reference
            With txtRef
                If flxAdj.TextMatrix(ActRow, 4) <> "" Then .Tag = flxAdj.TextMatrix(ActRow, 4)
                .Top = flxAdj.CellTop + flxAdj.Top
                .Left = flxAdj.CellLeft + flxAdj.Left
                .Width = flxAdj.CellWidth
                .Text = flxAdj.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 5    'qty
            With txtQtyA
                If flxAdj.TextMatrix(ActRow, 5) <> "" Then .Tag = flxAdj.TextMatrix(ActRow, 5)
                .Top = flxAdj.CellTop + flxAdj.Top
                .Left = flxAdj.CellLeft + flxAdj.Left
                .Width = flxAdj.CellWidth
                .Text = flxAdj.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
    End Select
End Sub

Private Sub flxAdj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxAdj_Click
End Sub

Private Sub flxAdj_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If LockFlx Then Exit Sub
    If Button = vbRightButton And ActCol = 0 Then
        If flxAdj.TextMatrix(ActRow, 8) = "" Then
            mnuCardex.Enabled = False
            mnuCardexRed.Enabled = False
        Else
            mnuCardex.Enabled = True
            mnuCardexRed.Enabled = True
        End If
        Call PopupMenu(mnuADJ)
    End If
End Sub

Private Sub flxAdj_Scroll()
    txtPartNo.Visible = False
    cboType.Visible = False
    txtQtyA.Visible = False
    txtRef.Visible = False
End Sub

Private Sub flxEmp_Click()
    ActRow = flxEmp.Row
    ActCol = flxEmp.Col
    
    If ActCol = 0 Or ActRow = 0 Or LockFlx Then Exit Sub
    Select Case ActCol
        Case 1    'Employee Name
            With cboEmp
                If flxEmp.TextMatrix(ActRow, 1) <> "" Then .Tag = flxEmp.TextMatrix(ActRow, 1)
                .Top = flxEmp.CellTop + flxEmp.Top
                .Left = flxEmp.CellLeft + flxEmp.Left
                .Width = flxEmp.CellWidth
                .Text = flxEmp.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 2  'Amount Charge
            With txtChrgAmt
                If flxEmp.TextMatrix(ActRow, 2) <> "" Then .Tag = flxEmp.TextMatrix(ActRow, 2)
                .Top = flxEmp.CellTop + flxEmp.Top
                .Left = flxEmp.CellLeft + flxEmp.Left
                .Width = flxEmp.CellWidth
                .Text = flxEmp.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
    End Select
End Sub

Private Sub flxEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxEmp_Click
End Sub

Private Sub flxEmp_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If LockFlx Then Exit Sub
    If Button = vbRightButton And ActCol = 0 Then
        Call PopupMenu(mnuEmp)
    End If
End Sub

Private Sub flxEmp_Scroll()
    cboEmp.Visible = False
    txtChrgAmt.Visible = False
End Sub

Private Sub flxParts_Click()
    Dim TtlPrice As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxAdj.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 3) 'Description
        flxAdj.TextMatrix(ActRow, 6) = Format$(Val(.TextMatrix(.Row, 4)), "#,##0.00") 'Price
        flxAdj.TextMatrix(ActRow, 8) = .TextMatrix(.Row, 1) 'SPid
        flxAdj.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 8)
        'Qty Adj Available
        If Val(flxAdj.TextMatrix(ActRow, 5)) > 0 Then
            TtlPrice = Val(flxAdj.TextMatrix(ActRow, 5)) * Val(.TextMatrix(.Row, 4))
            flxAdj.TextMatrix(ActRow, 7) = Format(TtlPrice, "#,###,##0.00")
        Else
            flxAdj.TextMatrix(ActRow, 7) = "0.00"
        End If
        Call ComputeTotal
    End With
End Sub

Private Sub flxParts_EnterCell()
    Dim TtlPrice As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxAdj.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 3) 'Description
        flxAdj.TextMatrix(ActRow, 6) = Format$(Val(.TextMatrix(.Row, 4)), "#,##0.00") 'Price
        flxAdj.TextMatrix(ActRow, 8) = .TextMatrix(.Row, 1) 'SPid
        'Qty and SRP >0
        If Val(flxAdj.TextMatrix(ActRow, 5)) > 0 Then
            TtlPrice = Val(flxAdj.TextMatrix(ActRow, 5)) * Val(.TextMatrix(.Row, 4))
            flxAdj.TextMatrix(ActRow, 7) = Format(TtlPrice, "#,###,##0.00")
        Else
            flxAdj.TextMatrix(ActRow, 7) = "0.00"
        End If
        Call ComputeTotal
    End With
End Sub

Private Sub flxParts_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        Tmp = Split(flxParts.Tag, "|")
        For i = LBound(Tmp) To UBound(Tmp)
            flxAdj.TextMatrix(ActRow, i + 1) = Tmp(i)
        Next i
        txtPartNo.Visible = False
        flxParts.Visible = False
        flxAdj.SetFocus
        flxAdj.Col = 1
    End If

End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        If flxAdj.TextMatrix(ActRow, 8) = "" Then   'No SPID
            flxAdj.SetFocus
            flxAdj.Row = ActRow
            flxAdj.Col = 1
        Else
            flxAdj.SetFocus
            flxAdj.Row = ActRow
            flxAdj.Col = 3
        End If
    End If
End Sub

Private Sub flxParts_LostFocus()
Dim i As Long, j As Long
    'Validate existing PartNo
    For i = 1 To flxAdj.Rows - 1
        If i <> ActRow Then
            If flxAdj.TextMatrix(i, 1) = txtPartNo.Text Then
                flxAdj.TextMatrix(ActRow, 1) = txtPartNo.Text
                txtPartNo.Visible = False
                flxParts.Visible = False
                Ans = MsgBox("Part Number " & txtPartNo.Text & " already exist in row " & i, vbInformation + vbOKOnly, "Duplicate Part Number")
                For j = 1 To flxAdj.Cols - 1
                    flxAdj.TextMatrix(ActRow, j) = ""
                Next j
                DoEvents
                flxAdj.Row = i
                flxAdj.Col = 1
                flxAdj.SetFocus
                txtPartNo.Tag = "PN exist"
                Exit Sub
            End If
        End If
    Next i
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Call LockData(True)
    Call CboDisplay
    cboType.AddItem "DEBIT(+)"
    cboType.AddItem "CREDIT(-)"
    cmdA.Top = 15000
    Timer1.Enabled = True
    MotoRep.sBar.Panels(6).Text = "Spare Parts Inventory Adjustment Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
    If rsEmployee.State = adStateOpen Then rsEmployee.Close
    Unload SPCardex
    Unload SPCardexRed
End Sub

Private Sub Label15_Click()

End Sub

Private Sub mnuAdd_Click()
    If flxAdj.TextMatrix(ActRow, 8) = "" Or flxAdj.TextMatrix(ActRow, 3) = "" Or Val(flxAdj.TextMatrix(ActRow, 5)) = 0 Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        flxAdj.SetFocus
        flxAdj.Col = ActCol
        Exit Sub
    Else
        If flxAdj.Rows = 6 Then
            MsgBox "You can only add upto 5 Part Numbers per Adjustment!", vbExclamation + vbOKOnly, "Limit Exceed"
            Exit Sub
        End If
        If ActRow = flxAdj.Rows - 1 Then flxAdj.AddItem ""
        flxAdj.SetFocus
        flxAdj.Col = 1
        flxAdj.Row = flxAdj.Rows - 1
    End If
End Sub

Private Sub mnuAddEmp_Click()
    If flxEmp.TextMatrix(ActRow, 3) = "" Or Val(flxEmp.TextMatrix(ActRow, 2)) = 0 Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        flxEmp.SetFocus
        flxEmp.Col = ActCol
        Exit Sub
    Else
        If ActRow = flxEmp.Rows - 1 Then flxEmp.AddItem ""
        flxEmp.SetFocus
        flxEmp.Col = 1
        flxEmp.Row = flxEmp.Rows - 1
    End If
End Sub

Private Sub mnuCardex_Click()
    SPCardex.PNadj = flxAdj.TextMatrix(ActRow, 1)
    SPCardex.Show vbModal
End Sub

Private Sub mnuCardexRed_Click()
    SPCardexRed.PNadj = flxAdj.TextMatrix(ActRow, 1)
    SPCardexRed.Show vbModal
End Sub

Private Sub mnuDelEmp_Click()
    Dim Ans As Integer, i As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If ActRow = 1 And flxEmp.Rows = 2 Then
            flxEmp.Rows = 1
            flxEmp.Rows = 2
        Else
            flxEmp.RemoveItem ActRow
        End If
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer, i As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If ActRow = 1 And flxAdj.Rows = 2 Then
            flxAdj.Rows = 1
            flxAdj.Rows = 2
        Else
            flxAdj.RemoveItem ActRow
        End If
        Call ComputeTotal
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxAdj
        .ColWidth(0) = 350
        .ColWidth(1) = (.Width - 380) * 0.17
        .ColWidth(2) = (.Width - 380) * 0.28
        .ColWidth(3) = (.Width - 380) * 0.12
        .ColWidth(4) = (.Width - 380) * 0.12
        .ColWidth(5) = (.Width - 380) * 0.1
        .ColWidth(6) = (.Width - 380) * 0.1
        .ColWidth(7) = (.Width - 380) * 0.1
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6)
    End With
    txtTotal.Top = flxAdj.Height + flxAdj.Top
    txtTotal.Left = flxAdj.Left + (flxAdj.Width - (flxAdj.ColWidth(6) + 400))
    txtTotal.Width = flxAdj.ColWidth(5)
    lblTotal.Top = txtTotal.Top
    lblTotal.Left = txtTotal.Left - lblTotal.Width 'flxadj.Left + (flxadj.Width - (flxadj.ColWidth(4) + flxadj.ColWidth(5)))
    'lbl2.Top = txtRemarks.Top
    'lbl2.Left = txtRemarks.Left - lbl2.Width
    'lbl2.FontSize = txtRemarks.FontSize
    With flxEmp
        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - 400) * 0.7
        .ColWidth(2) = (.Width - 400) * 0.28
        .ColWidth(3) = 0
    End With
End Sub


Private Sub SaveInvAdj()
    Dim rsInvAdj As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    'Saving Inventory Adjustment
    '***Access
    'With rsInvAdj
    '    If AddNew Then
    '        .Open "InvntryAdj", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '    Else
    '        Qry = "SELECT INVNTRYADJ.* From InvntryADJ WHERE INVNTRYADJ.ADJNO='" & cmdTrans(1).Tag & "';"
    
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '        .MoveFirst
    '    End If
    '    !AdjNo = txtAdjNo.Text
    '    !DocRef = txtRefNo.Text
    '    !AdjDate = txtAdjDate.Text
    '    !REASON = txtReason.Text
    '    !Remarks = txtRemarks.Text
    '    !PrepBy = txtPrepBy.Text
    '    !AprvBy = txtAprvBy.Text
         'Access
    '    '!ApplyCharges = IIf(chkCharges.Value = 1, True, False)
    '    'mySQL
    '    !ApplyCharges = IIf(chkCharges.Value = 1, 1, 0)
    '    !UPDATEDBY = UserID
    '    !LASTUPDATE = Now
    '    .Update
    '    .Close
    'End With
    
    'mySQL
    AppCharges = IIf(chkCharges.Value = 1, 1, 0)
    If AddNew Then
        LYRcon.Execute "Insert Into invntryadj( ADJNO, DOCREF, ADJDATE, REASON, REMARKS, PREPBY, " & _
            "APRVBY, APPLYCHARGES, UPDATEDBY, LASTUPDATE, ACCT_CODE)VALUES('" & txtAdjNo.Text & "', '" & txtRefNo.Text & "', " & _
            "'" & Format(txtAdjDate.Text, "yyyy-mm-dd") & "', '" & txtReason.Text & "', '" & txtRemarks.Text & "', '" & txtPrepBy.Text & "', " & _
            "'" & txtAprvBy.Text & "', " & AppCharges & ", " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & myAcctCode & "');"
    Else
        LYRcon.Execute "Update InvntryADJ SET ADJNO = '" & txtAdjNo.Text & "', DOCREF = '" & txtRefNo.Text & "', " & _
            "ADJDATE = '" & Format(txtAdjDate.Text, "yyyy-mm-dd") & "', REASON = '" & txtReason.Text & "', REMARKS = '" & txtRemarks.Text & "', " & _
            "PREPBY = '" & txtPrepBy.Text & "', APRVBY = '" & txtAprvBy.Text & "', APPLYCHARGES = " & AppCharges & ", " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', ACCT_CODE = '" & myAcctCode & "' " & _
            "Where invntryadj.ADJNO = '" & cmdTrans(1).Tag & "';"
    End If
    
    If chkCharges.Value = 1 Then
        If Not AddNew Then
            '***Access
            'LYRcon.Execute "DELETE INVADJCHARGES.*, INVADJCHARGES.ADJNO From INVADJCHARGES WHERE (((INVADJCHARGES.ADJNO)='" & cmdTrans(1).Tag & "'));"
            'mySQL
            LYRcon.Execute "DELETE From INVADJCHARGES WHERE INVADJCHARGES.ADJNO='" & cmdTrans(1).Tag & "';"
            
        End If
        With flxEmp
            For i = 1 To .Rows - 1
                '***Access
                'LYRcon.Execute "INSERT INTO INVADJCHARGES ( ADJNO, EMPID, AMTCHARGE) " & _
                            "VALUES ('" & txtAdjNo.Text & "', " & .TextMatrix(i, 3) & ", " & PureNum(.TextMatrix(i, 2)) & ");"
                'mySQL
                LYRcon.Execute "INSERT INTO INVADJCHARGES ( ADJNO, EMPID, AMTCHARGE) " & _
                            "VALUES ('" & txtAdjNo.Text & "', " & .TextMatrix(i, 3) & ", " & Val(Replace(.TextMatrix(i, 2), ",", "")) & ");"
            Next i
        End With
    End If
    'Sales Return Details
    If Not AddNew Then
        '***Access
        'LYRcon.Execute "DELETE INVADJDETAILS.* From INVADJDETAILS WHERE (((INVADJDETAILS.ADJNO)='" & cmdTrans(1).Tag & "'));"
        'mySQL
        LYRcon.Execute "DELETE From INVADJDETAILS WHERE INVADJDETAILS.ADJNO='" & cmdTrans(1).Tag & "';"
    End If
    With flxAdj
        For i = 1 To .Rows - 1
            '***Access
            'LYRcon.Execute "INSERT INTO INVADJDETAILS ( ADJNO, SPAREPARTID, ADJQTY, ADJTYPE, REFERENCE, PRICE) " & _
                            "VALUES ('" & txtAdjNo.Text & "', " & .TextMatrix(i, 8) & ", " & .TextMatrix(i, 5) & ", '" & .TextMatrix(i, 3) & "', '" & .TextMatrix(i, 4) & "', " & PureNum(.TextMatrix(i, 6)) & ");"
            'mySQL
            LYRcon.Execute "INSERT INTO INVADJDETAILS ( ADJNO, SPAREPARTID, ADJQTY, ADJTYPE, REFERENCE, PRICE) " & _
                            "VALUES ('" & txtAdjNo.Text & "', " & .TextMatrix(i, 8) & ", " & .TextMatrix(i, 5) & ", '" & .TextMatrix(i, 3) & "', '" & .TextMatrix(i, 4) & "', " & Val(Replace(.TextMatrix(i, 6), ",", "")) & ");"
        
        Next i
    End With
    'Add Control Number
    If AddNew Then
        eMsg = "Error Saving Adjustment Counter"
        '***Access
        'LYRcon.Execute "INSERT INTO ADJCNTR ( ADJNO, ADJTYPE) VALUES (" & Val(txtAdjNo.Text) & ", 'INVENTORY');"
        'mySQL
        LYRcon.Execute "INSERT INTO ADJCNTR ( ADJNO, ADJTYPE) VALUES ('" & txtAdjNo.Text & "', 'INVENTORY');"
        
    Else
        eMsg = "Error Update Adjustment Counter"
        '***Access
        'LYRcon.Execute "UPDATE ADJCNTR SET ADJCNTR.ADJNO = " & Val(txtAdjNo.Text) & _
                       " WHERE (((ADJCNTR.ADJNO)=" & Val(cmdTrans(1).Tag) & ") AND ((ADJCNTR.ADJTYPE)='INVENTORY'));"
        'mySQL
        LYRcon.Execute "UPDATE ADJCNTR SET ADJCNTR.ADJNO = '" & txtAdjNo.Text & "' WHERE ADJCNTR.ADJNO='" & cmdTrans(1).Tag & "' AND ADJCNTR.ADJTYPE='INVENTORY';"
        
    End If
    
End Sub
Private Sub UpdateSPITR()
    Dim rsITR As New ADODB.Recordset
    Dim i As Integer, myTtl As Currency, OldBal As Currency, CdtOut As Integer, DbtReceive As Integer
    'mySQL
    'Dim DbtReceive As Double, CdtOut As Double
    
    With flxAdj
        For i = 1 To .Rows - 1
            'Update SpareParts Qty on hand
            If InStr(.TextMatrix(i, 3), "DEBIT") > 0 Then
                '***Access
                'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+" & Val(.TextMatrix(i, 5)) & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
                'mySQL
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+" & Val(.TextMatrix(i, 5)) & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
                
            Else
                '***Access
                'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]-" & Val(.TextMatrix(i, 5)) & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
                'mySQl
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND-" & Val(.TextMatrix(i, 5)) & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
                
                
            End If
            'Saving Inventory Transaction Register
            '***Access
            'rsITR.Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
            'rsITR.AddNew
            'rsITR!REFERENCE = "ADJ#" & txtAdjNo.Text
            'rsITR!SparepartID = .TextMatrix(i, 8)
            'If InStr(.TextMatrix(i, 3), "DEBIT") > 0 Then
            '    rsITR!RECEIVE = Val(.TextMatrix(i, 5))
            'Else
            '    rsITR!OUT = Val(.TextMatrix(i, 5))
            'End If
            'rsITR!TRANSDATE = txtAdjDate.Text
            'rsITR!TRANSTYPE = "INVENTORY ADJ."
            'rsITR!Price = PureNum(.TextMatrix(i, 6))
            'rsITR!REMARKS = .TextMatrix(i, 3)
            'rsITR!PACKINGLIST = Left$(txtRefNo.Text, 20)
            'rsITR!Source = .TextMatrix(i, 4)
            'rsITR!UPDATEDBY = UserID
            'rsITR!LASTUPDATE = Now
            'rsITR.Update
            'rsITR.Close
            
            'mySQL
            LYRcon.Execute "Insert Into spitr( REFERENCE, SPAREPARTID, TRANSDATE, " & _
                "TRANSTYPE, PRICE, REMARKS, PACKINGLIST, SOURCE, UPDATEDBY, LASTUPDATE)" & _
                "VALUES('ADJ#" & txtAdjNo.Text & "', " & .TextMatrix(i, 8) & ", " & _
                "'" & Format(txtAdjDate.Text, "yyyy-mm-dd") & "', 'INVENTORY ADJ.', " & _
                "" & Val(Replace(.TextMatrix(i, 6), ",", "")) & ", '" & .TextMatrix(i, 3) & "', " & _
                "'" & Left(txtRefNo.Text, 20) & "', '" & .TextMatrix(i, 4) & "', " & _
                "" & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
            
            If InStr(.TextMatrix(i, 3), "DEBIT") > 0 Then
                LYRcon.Execute "UPDATE spitr SET RECEIVE = " & Val(.TextMatrix(i, 5)) & " Where spitr.REFERENCE = '" & "ADJ#" & txtAdjNo.Text & "' AND spitr.SPAREPARTID = " & .TextMatrix(i, 8) & ";"
            Else
                LYRcon.Execute "UPDATE spitr SET `OUT` = " & Val(.TextMatrix(i, 5)) & " Where spitr.REFERENCE = '" & "ADJ#" & txtAdjNo.Text & "' AND spitr.SPAREPARTID = " & .TextMatrix(i, 8) & ";"
            End If
        Next i
    End With
End Sub


Function GetAdjustmentNo() As String
    Dim rsAdjCntr As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    'Access
    'Qry = "SELECT Max(ADJCNTR.ADJNO) AS MaxOfADJNO FROM ADJCNTR;"
    'mySQL
    Qry = "SELECT Max(round(ADJCNTR.ADJNO)) AS MaxOfADJNO FROM ADJCNTR;"
    With rsAdjCntr
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, .Fields(0) + 1)
        Else
            MyNo = 1
        End If
        .Close
    End With
    GetAdjustmentNo = Format(MyNo, "0000")
End Function


Function ValidateInvAdj() As String
    Dim vMsg As String, Qry As String
    Dim Cntr As Integer
    Dim i As Long, j As Long
    
    vMsg = ""

    If Not IsDate(txtAdjDate.Text) Then vMsg = "Invalid Adjustment Date"
    If txtAdjNo.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field Adjustment No."
    Else
        If AddNew Then
            '***Access
            'Qry = "SELECT ADJCNTR.ADJNO From ADJCNTR WHERE (((ADJCNTR.ADJNO)=" & Val(txtAdjNo.Text) & "));"
            'mySQL
            Qry = "SELECT ADJCNTR.ADJNO From ADJCNTR WHERE ADJCNTR.ADJNO='" & txtAdjNo.Text & "';"
            
            If CtrlNumExist(Qry) Then
                vMsg = vMsg & vbCrLf & "Adjustment No. " & txtAdjNo.Text & " already exist!"
            End If
        End If
    End If
    If txtAprvBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Approved By"
    If txtPrepBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Prepared By"
    If chkCharges.Value = 1 Then
        If Val(flxEmp.TextMatrix(flxEmp.Rows - 1, 3)) = 0 Or Val(flxEmp.TextMatrix(flxEmp.Rows - 1, 2)) = 0 Then
            If flxEmp.Row <= 2 Then
                vMsg = vMsg & vbCrLf & "Incomplete entry on Names of employees with charges"
            Else
                flxEmp.RemoveItem flxEmp.Rows - 1
            End If
        End If
    End If
    If flxAdj.TextMatrix(flxAdj.Rows - 1, 3) = "" Or flxAdj.TextMatrix(flxAdj.Rows - 1, 8) = "" Or Val(flxAdj.TextMatrix(flxAdj.Rows - 1, 5)) = 0 Then
        If flxAdj.Rows <= 2 Then
            vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Adjusted Spare Parts"
        Else
            flxAdj.RemoveItem flxAdj.Rows - 1
        End If
    End If
    
    With flxAdj
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 5)) = 0 Then
                vMsg = vMsg & vbCrLf & "No Quantity adjusted at row " & i
            End If
        Next i
    End With
    
    For i = 1 To flxAdj.Rows - 1
        For j = 1 To flxAdj.Rows - 1
            If i <> j Then
                If flxAdj.TextMatrix(i, 1) = flxAdj.TextMatrix(j, 1) Then
                    'MsgBox "Part Number " & txtPartNo.Text & " already exist in row " & i, vbInformation + vbOKOnly, "Duplicate Part Number"
                    vMsg = vMsg & vbCrLf & "Part Number " & flxOrd.TextMatrix(i, 1) & " already exist in row " & i
                End If
            End If
        Next j
    Next i
    If cboAcctTitle.Text = "" Then vMsg = vMsg & vbCrLf & "No Account Title Found!"
    
    ValidateInvAdj = vMsg
    
End Function

Public Sub DisplayInvAdj(MyRef As String)
    Dim rsAdj As New ADODB.Recordset
    Dim rsChrg As New ADODB.Recordset
    Dim rsDtls As New ADODB.Recordset
    Dim Qry As String, AdjType As String
    Dim AdjQ As Integer
    
    Call ClearEntries
    '***Access
    'Qry = "SELECT INVNTRYADJ.* From InvntryADJ WHERE (((INVNTRYADJ.ADJNO)='" & MyRef & "'));"
    'mySQL
    Qry = "SELECT INVNTRYADJ.* From InvntryADJ WHERE INVNTRYADJ.ADJNO='" & MyRef & "';"
    
    With rsAdj
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtAdjType.Text = "INVENTORY"
            txtRefNo.Text = IIf(IsNull(!DocRef), "", !DocRef)
            txtAdjNo.Text = !ADJNO
            txtAdjDate.Text = !ADJDATE
            txtReason.Text = IIf(IsNull(!REASON), "", !REASON)
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtPrepBy.Text = !PREPBY
            txtAprvBy.Text = !APRVBY
            txtPrnCopy.Text = IIf(IsNull(!PRINTCOPY), 0, !PRINTCOPY)
            'Access
            'If !ApplyCharges = True Then
            'mySQL
            If !APPLYCHARGES = 1 Then
                chkCharges.Value = 1
            Else
                chkCharges.Value = 0
            End If
        End If
        .Close
    End With
    
    If chkCharges.Value = 1 Then    'Display Employees applied charges
        '***Access
        'Qry = "SELECT [LASTNAME] & IIf(IsNull([FIRSTNAME]),'',', ' & [FIRSTNAME]) & IIf(IsNull([MIDDLENAME]),'',' ' & Left$([MIDDLENAME],1) & '.') AS EMPNAME, INVADJCHARGES.AMTCHARGE, Employees.EMPID " & _
              "FROM INVADJCHARGES INNER JOIN Employees ON INVADJCHARGES.EMPID = Employees.EMPID WHERE (((INVADJCHARGES.ADJNO)='" & MyRef & "'));"
        'mySQL
        Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPNAME, INVADJCHARGES.AMTCHARGE, Employees.EMPID " & _
              "FROM INVADJCHARGES INNER JOIN Employees ON INVADJCHARGES.EMPID = Employees.EMPID WHERE INVADJCHARGES.ADJNO='" & MyRef & "';"
        
        With rsChrg
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                flxEmp.Rows = 1
                .MoveFirst
                Do While Not .EOF
                    flxEmp.AddItem vbTab & !EmpName & vbTab & !AmtCharge & vbTab & !EMPID
                    .MoveNext
                Loop
            End If
            .Close
        End With
    End If
    
    If Val(txtPrnCopy.Text) = 0 Or InStr(txtRemarks.Text, "CANCELLED") > 0 Then
        '***Access
        'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, INVADJDETAILS.ADJTYPE, INVADJDETAILS.REFERENCE, INVADJDETAILS.ADJQTY, INVADJDETAILS.PRICE, [ADJQTY]*[PRICE] AS TTLAMT, INVADJDETAILS.SPAREPARTID " & _
              "FROM INVADJDETAILS INNER JOIN SPAREPARTS ON INVADJDETAILS.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((INVADJDETAILS.ADJNO)='" & MyRef & "'));"
        'mySQL
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, INVADJDETAILS.ADJTYPE, INVADJDETAILS.REFERENCE, INVADJDETAILS.ADJQTY, INVADJDETAILS.PRICE, INVADJDETAILS.ADJQTY*INVADJDETAILS.PRICE AS TTLAMT, INVADJDETAILS.SPAREPARTID " & _
              "FROM INVADJDETAILS INNER JOIN SPAREPARTS ON INVADJDETAILS.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE INVADJDETAILS.ADJNO='" & MyRef & "';"
        
        With rsDtls
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxAdj.Rows = 1
                Do While Not .EOF
                    flxAdj.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !AdjType & vbTab & IIf(IsNull(!REFERENCE), "", !REFERENCE) & vbTab & !AdjQty & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!TtlAmt, "#,##0.00") & vbTab & !SparepartID
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
        '***Access
        'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.RECEIVE, SPITR.OUT, SPITR.SOURCE, SPITR.PRICE, SPITR.SPAREPARTID " & _
              "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='ADJ#" & MyRef & "') AND ((SPITR.TRANSTYPE)='INVENTORY ADJ.'));"
        'mySQl
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.RECEIVE, SPITR.`OUT`, SPITR.SOURCE, SPITR.PRICE, SPITR.SPAREPARTID " & _
              "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='ADJ#" & MyRef & "' AND SPITR.TRANSTYPE='INVENTORY ADJ.';"
     
        With rsDtls
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxAdj.Rows = 1
                Do While Not .EOF
                    If !RECEIVE > 0 Then
                        AdjType = "DEBIT(+)"
                        AdjQ = !RECEIVE
                    Else
                        AdjType = "CREDIT(-)"
                        AdjQ = !OUT
                    End If
                    flxAdj.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & AdjType & vbTab & IIf(IsNull(!Source), "", !Source) & vbTab & AdjQ & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!Price * AdjQ, "#,##0.00") & vbTab & !SparepartID
                    .MoveNext
                Loop
            End If
            .Close
        End With
    End If
    Call LockData(True)
    Call ComputeTotal
    
End Sub

Private Sub Timer1_Timer()
    Call LoadSpareParts
    Load SPCardex
    Load SPCardexRed
    Timer1.Enabled = False
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 7 Then
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 5 Then
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
            BtnIndx = 5 'Index of Button Save
            cmdTrans(5).Top = -375
            cmdTrans(5).Left = 0
            cmdTrans(5).Visible = True
            cmdTrans(6).Top = -375
            cmdTrans(6).Left = cmdTrans(4).Width
            cmdTrans(6).Visible = True
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
        If BtnIndx = 4 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 4
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub


Private Sub txtAdjDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890/-"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxAdj.SetFocus
        flxAdj.Row = 1
        flxAdj.Col = 2
    End If
End Sub

Private Sub txtAdjNo_GotFocus()
    If Val(txtAdjNo.Text) > 0 Then
        txtAdjNo.SelStart = 0
        txtAdjNo.SelLength = Len(txtAdjNo.Text)
    End If
End Sub

Private Sub txtAdjNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAdjDate.SetFocus
End Sub

Private Sub txtAdjNo_LostFocus()
    Dim Qry As String
    
    'Verify AdjNo if existing
    If Val(txtAdjNo.Text) > 0 And AddNew Then
        '***Access
        'Qry = "SELECT ADJCNTR.ADJNO From ADJCNTR WHERE (((ADJCNTR.ADJNO)=" & Val(txtAdjNo.Text) & "));"
        'mySQL
        Qry = "SELECT INVNTRYADJ.ADJNO From INVNTRYADJ WHERE INVNTRYADJ.ADJNO='" & txtAdjNo.Text & "';"
        
        If CtrlNumExist(Qry) Then
            MsgBox "Adjustment No. " & txtAdjNo.Text & " already exist!", vbExclamation + vbOKOnly, "Validation Error"
        End If
    End If
    
End Sub

Private Sub txtaprvby_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If LockFlx Then
            cmdTrans(0).SetFocus
        Else
            cmdTrans(5).SetFocus
        End If
    End If
End Sub
                                            
Private Sub txtChrgAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtChrgAmt.Text = Empty
        flxEmp.TextMatrix(ActRow, 4) = txtChrgAmt.Tag
        txtChrgAmt.Tag = "skiplostfocus"
        flxEmp.SetFocus
        flxEmp.Col = 2
        txtChrgAmt.Visible = False
    End If
End Sub

Private Sub txtChrgAmt_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        With flxEmp
            .SetFocus
            If .TextMatrix(ActRow, 2) <> "" And Val(.TextMatrix(ActRow, 3)) > 0 Then
                If ActRow = .Rows - 1 Then .AddItem ""
                .Row = .Rows - 1
                .Col = 1
            ElseIf .TextMatrix(ActRow, 3) = "" Then
                .Row = ActRow
                .Col = 1
            ElseIf .TextMatrix(ActRow, 2) = "" Then
                .Row = ActRow
                .Col = 1
            End If
        End With
    End If

End Sub

Private Sub txtChrgAmt_LostFocus()
    If txtChrgAmt.Tag = "skiplostfocus" Then
        txtChrgAmt.Tag = ""
        Exit Sub
    End If
    flxEmp.TextMatrix(ActRow, 2) = Format(txtChrgAmt.Text, "#,##0.00")
    txtChrgAmt.Visible = False
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
            .Col = 0: .ColSel = 5
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            flxAdj.TextMatrix(ActRow, 8) = flxParts.TextMatrix(flxParts.Row, 1) 'SP id
            flxAdj.TextMatrix(ActRow, 2) = flxParts.TextMatrix(flxParts.Row, 3) 'Description
            flxAdj.TextMatrix(ActRow, 6) = Format$(Val(flxParts.TextMatrix(flxParts.Row, 4)), "###,##0.00") 'Price
            If Val(flxAdj.TextMatrix(ActRow, 5)) > 0 Then   'Qty Adj.
                TtlPrice = Val(flxAdj.TextMatrix(ActRow, 5)) * Val(flxParts.TextMatrix(flxParts.Row, 4))
                flxAdj.TextMatrix(ActRow, 7) = Format(TtlPrice, "#,###,##0.00")
                Call ComputeTotal
            Else
                flxAdj.TextMatrix(ActRow, 7) = "0.00"
            End If
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
'            rsParts.Requery
            flxAdj.TextMatrix(ActRow, 2) = ""
            flxAdj.TextMatrix(ActRow, 5) = ""
            flxAdj.TextMatrix(ActRow, 6) = ""
            flxAdj.TextMatrix(ActRow, 7) = ""
            flxAdj.TextMatrix(ActRow, 4) = ""
            .Visible = False
        End If

    End With
End Sub
Private Sub validateQty()
    AvlQty = 0
    If txtPartNo.Tag = "" Then
        If txtQtyA.Tag = "" Then
            If cboType.Tag = "" Then
                With flxOrd
                    Call qtyCheck(txtPartNo)    'VALIDATE
                    AvlQty = (avQty - Val(txtQtyA.Text))
                    If cboType.Text = "CREDIT(-)" Then
                        If AvlQty < 0 Then
                            MsgBox "Available Qty. On-Hand (" & avQty & ") ", vbCritical + vbOKOnly, "Validation"
                            flxAdj.TextMatrix(ActRow, 5) = ""
                            Exit Sub
                        End If
                    End If
                End With
            End If
        End If
    End If
    Exit Sub

End Sub
Private Sub qtyCheck(PartNo As String)
    Dim mainRs As New ADODB.Recordset
    
    Qry = ""
    avQty = 0
    Qry = "SELECT qtyonhand FROM spareparts  WHERE partnumber = '" & PartNo & "';"
        
With mainRs
    If .State = adStateOpen Then .Close
         .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
               .MoveFirst
                DoEvents
                avQty = !QtyOnHand
             End If
                 .Close
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
            Tmp = Split(flxAdj.TextMatrix(ActRow, 8), "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxAdj.TextMatrix(ActRow, i + 1) = Tmp(i)
            Next i
            flxParts.Tag = "skiplostfocus"
            flxAdj.SetFocus
            flxAdj.Col = 1
            txtPartNo.Visible = False
            flxParts.Visible = False
        Case vbKeyDelete
            PrtSrch = True
        Case vbKeyBack
            PrtSrch = True
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If flxAdj.TextMatrix(ActRow, 7) = "" Then   'No SPID
            flxAdj.SetFocus
            flxAdj.Row = ActRow
            flxAdj.Col = 1
        Else
            flxAdj.SetFocus
            flxAdj.Row = ActRow
            flxAdj.Col = 3
        End If
        flxParts.Visible = False
    End If
End Sub

Private Sub txtPartNo_LostFocus()
Dim i As Long, j As Long

    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    
    'Validate existing PartNo
    For i = 1 To flxAdj.Rows - 1
        If i <> ActRow Then
            If flxAdj.TextMatrix(i, 1) = txtPartNo.Text Then
                flxAdj.TextMatrix(ActRow, 1) = txtPartNo.Text
                txtPartNo.Visible = False
                flxParts.Visible = False
                Ans = MsgBox("Part Number " & txtPartNo.Text & " already exist in row " & i, vbInformation + vbOKOnly, "Duplicate Part Number")
                For j = 1 To flxAdj.Cols - 1
                    flxAdj.TextMatrix(ActRow, j) = ""
                Next j
                DoEvents
                flxAdj.Row = i
                flxAdj.Col = 1
                flxAdj.SetFocus
                txtPartNo.Tag = "PN exist"
                Exit Sub
            End If
        End If
    Next i
    
    If flxAdj.TextMatrix(ActRow, 7) = "" Then
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
        txtPartNo.Visible = False
        flxParts.Visible = False
        Exit Sub
    End If
    flxAdj.TextMatrix(ActRow, 1) = txtPartNo.Text
    txtPartNo.Visible = False
    flxParts.Visible = False
End Sub

Private Sub txtPrepBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAprvBy.SetFocus
End Sub


Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtPrice.Text = Empty
        flxAdj.TextMatrix(ActRow, 6) = txtPrice.Tag
        txtPrice.Tag = "skiplostfocus"
        flxAdj.SetFocus
        flxAdj.Col = 6
        txtPrice.Visible = False
    End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        With flxAdj
            .SetFocus
            If .TextMatrix(ActRow, 8) <> "" And .TextMatrix(ActRow, 3) <> "" And Val(.TextMatrix(ActRow, 5)) > 0 Then
                If ActRow = .Rows - 1 Then .AddItem ""
                .Row = .Rows - 1
                .Col = 1
            ElseIf .TextMatrix(ActRow, 8) = "" Then
                .Row = ActRow
                .Col = 1
            ElseIf .TextMatrix(ActRow, 3) = "" Then
                .Row = ActRow
                .Col = 3
            ElseIf .TextMatrix(ActRow, 5) = "" Then
                .Row = ActRow
                .Col = 5
            End If
        End With
    End If
End Sub

Private Sub txtPrice_LostFocus()
    If txtPrice.Tag = "skiplostfocus" Then
        txtPrice.Tag = ""
        Exit Sub
    End If
    flxAdj.TextMatrix(ActRow, 6) = Format(txtPrice.Text, "#,##0.00")
    Call ComputeTotal
    txtPrice.Visible = False
End Sub

Private Sub txtQtyA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtQtyA.Text = Empty
        flxAdj.TextMatrix(ActRow, 4) = txtQtyA.Tag
        txtQtyA.Tag = "skiplostfocus"
        flxAdj.SetFocus
        flxAdj.Col = 4
        txtQtyA.Visible = False
    End If
        txtQtyA.Tag = ""
        Call validateQty
End Sub

Private Sub txtQtyA_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        With flxAdj
            .SetFocus
            If .TextMatrix(ActRow, 8) <> "" And .TextMatrix(ActRow, 3) <> "" And Val(.TextMatrix(ActRow, 5)) > 0 Then
                If ActRow = .Rows - 1 Then .AddItem ""
                .Row = .Rows - 1
                .Col = 1
            ElseIf .TextMatrix(ActRow, 8) = "" Then
                .Row = ActRow
                .Col = 1
            ElseIf .TextMatrix(ActRow, 3) = "" Then
                .Row = ActRow
                .Col = 3
            ElseIf .TextMatrix(ActRow, 5) = "" Then
                .Row = ActRow
                .Col = 5
            End If
        End With
    End If
End Sub

Private Sub txtQtyA_LostFocus()
    If txtQtyA.Tag = "skiplostfocus" Then
        txtQtyA.Tag = ""
        Exit Sub
    End If
'    With flxAdj
'        'If Val(txtQtyA.Text) > Val(.TextMatrix(ActRow, 1)) And .TextMatrix(ActRow, 10) <> "" Then
'        If Val(txtQtyA.Text) > Val(.TextMatrix(ActRow, 1)) And .TextMatrix(ActRow, 11) <> "" Then
'            MsgBox "Order Qty. must not exceed Qty. On-Hand (" & .TextMatrix(ActRow, 1) & ")!", vbCritical + vbOKOnly, "Validation"
'            txtQtyA.Tag = "XcdQty"
'            txtQtyA.SelStart = 0
'            txtQtyA.SelLength = Len(txtQtyA.Text)
'            txtQtyA.SetFocus
'            Exit Sub
'        End If
    flxAdj.TextMatrix(ActRow, 5) = txtQtyA.Text
    Call ComputeTotal
    txtQtyA.Visible = False
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub txtRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtRef.Text = Empty
        flxAdj.TextMatrix(ActRow, 4) = txtRef.Tag
        txtRef.Tag = "skiplostfocus"
        flxAdj.SetFocus
        flxAdj.Col = 4
        txtRef.Visible = False
    End If
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxAdj.SetFocus
        flxAdj.Row = ActRow
        flxAdj.Col = 5
    End If
End Sub

Private Sub txtRef_LostFocus()
    If txtRef.Tag = "skiplostfocus" Then
        txtRef.Tag = ""
        Exit Sub
    End If
    flxAdj.TextMatrix(ActRow, 4) = txtRef.Text
    txtRef.Visible = False
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAdjNo.SetFocus
End Sub

Function GetPartsMan(MyID As Integer) As String
    Dim rsUser As New ADODB.Recordset
    Dim Qry As String
    '***Access
    'Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER WHERE (((PCUSER.EMPID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER WHERE PCUSER.EMPID=" & MyID & ";"
    
    With rsUser
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetPartsMan = IIf(IsNull(!UserName), "", UCase$(!UserName))
        Else
            GetPartsMan = ""
        End If
        .Close
    End With
End Function


Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAprvBy.SetFocus
End Sub
