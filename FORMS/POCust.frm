VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form POCust 
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11535
   ControlBox      =   0   'False
   Icon            =   "POCust.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   11535
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtAcct 
      Height          =   285
      Left            =   10440
      TabIndex        =   58
      Top             =   9480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   8160
   End
   Begin MSFlexGridLib.MSFlexGrid flxPRN 
      Height          =   1575
      Left            =   2760
      TabIndex        =   54
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   1
      Cols            =   6
   End
   Begin VB.TextBox txtPOid 
      Height          =   285
      Left            =   5160
      TabIndex        =   53
      Top             =   10080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox ChkStatus 
      Caption         =   "APPROVED"
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
      TabIndex        =   50
      Tag             =   "XYF"
      Top             =   6360
      Width           =   2415
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
      TabIndex        =   39
      Top             =   8760
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CheckBox chkDP 
      Caption         =   "DP"
      Height          =   255
      Left            =   360
      TabIndex        =   38
      Top             =   9600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox lstDelete 
      Height          =   840
      ItemData        =   "POCust.frx":17D2A
      Left            =   8160
      List            =   "POCust.frx":17D2C
      TabIndex        =   37
      Top             =   9360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtJOid 
      Height          =   285
      Left            =   1800
      TabIndex        =   36
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEmpID 
      Height          =   285
      Left            =   10680
      TabIndex        =   34
      Top             =   10080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   360
      TabIndex        =   32
      Top             =   9360
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
      TabIndex        =   31
      Top             =   9000
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
      TabIndex        =   30
      Text            =   "No Discount"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1455
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
      TabIndex        =   27
      Tag             =   "XYF"
      Top             =   6285
      Width           =   1620
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8520
      TabIndex        =   26
      Top             =   10080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   7800
      TabIndex        =   25
      Top             =   10080
      Visible         =   0   'False
      Width           =   615
   End
   Begin lvButton.lvButtons_H cmdAdvFind 
      Height          =   375
      Left            =   9960
      TabIndex        =   23
      Tag             =   "XYF"
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Find [F12]"
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
      Image           =   "POCust.frx":17D2E
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   240
      TabIndex        =   21
      Tag             =   "XYF"
      Top             =   6720
      Width           =   11055
      Begin VB.ComboBox cboCustomer 
         Height          =   315
         Left            =   1320
         TabIndex        =   57
         Tag             =   "XYF"
         Top             =   195
         Width           =   2655
      End
      Begin VB.TextBox txtEngine 
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
         Left            =   4320
         MaxLength       =   30
         TabIndex        =   55
         Tag             =   "XYF"
         Top             =   195
         Width           =   2055
      End
      Begin VB.TextBox txtAprvBy 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   52
         Tag             =   "XYF"
         Top             =   555
         Width           =   3015
      End
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1320
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Tag             =   "XYF"
         Top             =   555
         Width           =   5055
      End
      Begin VB.TextBox txtPrepBy 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   190
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "E#:"
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
         Left            =   3960
         TabIndex        =   56
         Tag             =   "XYF"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   1
         Left            =   6240
         TabIndex        =   51
         Tag             =   "XYF"
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lblRef 
         Caption         =   "CUSTOMER:"
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
         TabIndex        =   49
         Tag             =   "XYF"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCustomer 
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
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Tag             =   "XYF"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   0
         Left            =   6285
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.TextBox txtPrintCopy 
      Height          =   285
      Left            =   5160
      TabIndex        =   20
      Top             =   9720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   8760
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   9240
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   8760
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   9240
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   9840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1815
      Left            =   6000
      TabIndex        =   16
      Top             =   8880
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
      Top             =   8880
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   600
      TabIndex        =   14
      Text            =   "Qty. On Hand: 0"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtFlx 
      Height          =   315
      Left            =   3480
      TabIndex        =   13
      Top             =   10080
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
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   2280
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   0
      FormatString    =   $"POCust.frx":182C8
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
      Height          =   1140
      Left            =   240
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   720
      Width           =   11055
      Begin VB.TextBox txtTerm 
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
         Left            =   960
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox txtTO 
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   44
         Tag             =   "XYF"
         Text            =   "MOTO-R REPUBLIC CORPORATION"
         Top             =   300
         Width           =   4215
      End
      Begin VB.TextBox txtPOno 
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
         Left            =   9120
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   300
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   315
         Left            =   9120
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   660
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
         Caption         =   "TERM:"
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
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   690
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "TO:"
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
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O. No."
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
         Left            =   7440
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   330
         Width           =   1575
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
         Left            =   7560
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   690
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
            Picture         =   "POCust.frx":18351
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POCust.frx":1922B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POCust.frx":1A105
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POCust.frx":1AFDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POCust.frx":1BEB9
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
      Caption         =   "&New P.O."
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
      Image           =   "POCust.frx":1CD93
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   2
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit P.O."
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
      Image           =   "POCust.frx":1D32D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Save P.O."
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
      Image           =   "POCust.frx":1D8C7
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10080
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Cancel P.O."
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
      Image           =   "POCust.frx":1DE61
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   22
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Edit P.O."
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
      Image           =   "POCust.frx":1E3FB
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxBO 
      Height          =   1095
      Left            =   7560
      TabIndex        =   40
      Top             =   10320
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
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   41
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print P.O."
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
      Image           =   "POCust.frx":1E995
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   42
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "POCust.frx":1EF2F
      cBack           =   -2147483633
   End
   Begin VB.Label lblInvoice 
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
      Left            =   2760
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   6360
      Width           =   5205
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
      TabIndex        =   24
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
      m_MDI_Height    =   8820
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
      Top             =   2025
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
      Top             =   2055
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
End
Attribute VB_Name = "POCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim rsCustomer As New ADODB.Recordset

Dim CustSrch As Boolean, ActSrch As Boolean, PrtBln As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim tmrCount As Integer, BtnIndx As Integer
Dim AddNew As Boolean, PmanSrch As Boolean
Public JustScroll As Boolean
Dim fRow As Integer, fCol As Integer
Public MySource As String

Private Sub ClearEntries()
    txtPONO.Text = ""
    lblInvoice.Caption = ""
    txtTotal.Text = ""
    txtCustID.Text = ""
    txtTerm.Text = ""
    flxOrd.Rows = 1
    flxOrd.Rows = 2
    txtPrintCopy.Text = ""
    txtPrepBy.Text = ""
    txtRemarks.Text = ""
    cboCustomer.Text = ""
    txtAprvBy.Text = ""
    
End Sub
                                                                
Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    Frame1.Enabled = Not bln
    ChkStatus.Enabled = Not bln
    JustScroll = bln
End Sub
Private Sub cboCustomer_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String, myMsg As String
    
    If CustSrch Then
        CustSrch = False
        Exit Sub
    End If
    With cboCustomer
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsCustomer.MoveFirst
            rsCustomer.Move i
            txtCustID.Text = rsCustomer!CUSTOMERID
            '***Access
            'myMsg = "SELECT ONACCTERMS.TERMS From ONACCTERMS WHERE (((ONACCTERMS.STATUS)='Default') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
            'mySQl
            myMsg = "SELECT ONACCTERMS.TERMS From ONACCTERMS WHERE ONACCTERMS.STATUS='Default' AND ONACCTERMS.CUSTOMERID=" & txtCustID.Text & ";"
            
            txtTerm.Text = Val(GetFieldValue(myMsg)) & " " & "DAYS"
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CustSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustID.Text = ""
            txtTerm.Text = ""
        End If
    End With
End Sub

Private Sub cboCustomer_Click()
    Call cboCustomer_Change
End Sub

Private Sub cboCustomer_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboCustomer.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then CustSrch = True
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtEngine.SetFocus
End Sub

Private Sub ChkStatus_Click()
    If ChkStatus.Value = 1 Then
        txtAprvBy.Text = GetUserName
    Else
        txtAprvBy.Text = ""
    End If
End Sub

Private Sub cmdAdvFind_Click()
    SPfind.MySource = "POCUST"
    SPfind.Show vbModal
    'MotoRep.sBar.Panels(6).Text = "Spare Parts Sales Order"
End Sub

Private Sub cmdAltA_Click()
    Call mnuAdd_Click
End Sub


Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 7) As String, Qry As String, MyAcc As String
    
    'On Error GoTo LocalErr
    Select Case Index
        Case 0  'New
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=40));"
            'mySQl
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=40;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            Call ClearEntries
            Call LockData(False)
            AddNew = True
            Call BtnSaveCancel(True)
            txtPrepBy.Text = GetUserName
            If UserType = "Administrator" Then
                ChkStatus.Enabled = True
            Else
                ChkStatus.Enabled = False
            End If
            txtPONO.SetFocus
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=40));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=40;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If txtPOid.Text = "" Then
                MsgBox "No available P.O. to edit", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If Val(txtPrintCopy.Text) > 0 Then
                MsgBox "Unable to edit P.O. " & txtPONO.Text & " because its already printed.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If InStr(txtRemarks.Text, "CANCELLED") > 0 Then
                MsgBox "Unable to edit P.O. " & txtPONO.Text & " because it has already been Cancelled.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If UserType = "Administrator" Then
                ChkStatus.Enabled = True
            Else
                ChkStatus.Enabled = False
            End If
            txtPOid.Tag = txtPONO.Text
            Call BtnSaveCancel(True)
            Call LockData(False)
            AddNew = False
        Case 2  'PRINT
            If ChkStatus.Value = 0 Then
                Ans = MsgBox("Do you want to Print TEMPORARY P.O.?", vbQuestion + vbYesNo, "Print")
                If Ans = vbNo Then Exit Sub
                Call PrintTemp
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Print P.O.?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            FlxPrn.Rows = 1
            For i = 1 To flxOrd.Rows - 1
                FlxPrn.AddItem vbTab & flxOrd.TextMatrix(i, 2) & vbTab & flxOrd.TextMatrix(i, 3) & vbTab & flxOrd.TextMatrix(i, 4) & vbTab & flxOrd.TextMatrix(i, 7) & vbTab & flxOrd.TextMatrix(i, 8)
            Next i
            PrnVal(0) = txtDate.Text
            PrnVal(1) = txtTO.Text
            PrnVal(2) = "LYR BLDG. II, SOBRECARY ST. TAGUM CITY"
            PrnVal(3) = txtTerm.Text
            PrnVal(4) = txtPrepBy.Text
            PrnVal(5) = txtAprvBy.Text
            PrnVal(6) = txtTotal.Text
            myMsg = ""
            If txtEngine.Text <> "" Then myMsg = "Note: Engine No. " & txtEngine.Text
            If txtRemarks.Text <> "" Then myMsg = IIf(myMsg <> "", myMsg & vbCrLf & txtRemarks.Text, "Note: " & txtRemarks.Text)
            PrnVal(7) = myMsg
            Call PrintMyForm(PrnVal, "REBUILT PO", FlxPrn)
            txtPrintCopy.Text = Val(txtPrintCopy.Text) + 1
            '***Access
            'LYRcon.Execute "UPDATE CUSTPO SET CUSTPO.PRINTCOPY = [CUSTPO].[PRINTCOPY]+1 WHERE (((CUSTPO.CPOID)=" & txtPOid.Text & "));"
            
            'mySQl
            LYRcon.Execute "UPDATE CUSTPO SET CUSTPO.PRINTCOPY = CUSTPO.PRINTCOPY+1 WHERE CUSTPO.CPOID=" & txtPOid.Text & ";"
            

        Case 3  'SUMMARY
            FrmSummary.mySrc = "Customer P.O."
            FrmSummary.Show
            'POSummary.Show
            Unload Me
        Case 4  'Exit
            If MySource = "" Then
                'MotoRep.imgBG.Visible = True
                MotoRep.toolBarButton.Visible = True
                'MotoRep.tbMain.Visible = True
            End If
            Unload Me
        Case 5  'Save
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            myMsg = ValidatePO
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save P.O.?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'Save Spare Parts Invoice
            
            Call SavePurchaseOrder
            MsgBox "P.O. Successfully Saved!", vbInformation + vbOKOnly, "P.O."
            Call LockData(True)
            Call BtnSaveCancel(False)
        Case 6  'Cancel
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            Ans = MsgBox("Do you want to cancel P.O.?", vbQuestion + vbYesNo)
            If Ans = vbNo Then Exit Sub
            Call BtnSaveCancel(False)
            If AddNew Then
                Call ClearEntries
            Else
                Call ShowRecord(txtPOid.Text)
            End If
            Call LockData(True)
    End Select
    
End Sub

Private Sub PrintTemp()
    Dim rsPrn As New ADODB.Recordset
    Dim crPrn As New Rpt2HPO
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPSORDERDETAIL.PRICE, [PRICE]*[QTY] AS TTLPRICE " & _
          "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE (((SPSORDERDETAIL.CPOID)=" & txtPOid.Text & "));"
    
    'mySQL
    Qry = "SELECT SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPSORDERDETAIL.PRICE, SPSORDERDETAIL.PRICE*SPSORDERDETAIL.QTY AS TTLPRICE " & _
          "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE SPSORDERDETAIL.CPOID=" & txtPOid.Text & ";"
    
    
    rsPrn.Open Qry, LYRcon, adOpenDynamic, adLockOptimistic
    crPrn.DiscardSavedData
    crPrn.Database.SetDataSource rsPrn
    crPrn.ParameterFields(1).ClearCurrentValueAndRange
    crPrn.ParameterFields(1).AddCurrentValue txtPONO.Text
    crPrn.ParameterFields(2).ClearCurrentValueAndRange
    crPrn.ParameterFields(2).AddCurrentValue txtTerm.Text
    crPrn.ParameterFields(3).ClearCurrentValueAndRange
    crPrn.ParameterFields(3).AddCurrentValue cboCustomer.Text
    crPrn.ParameterFields(4).ClearCurrentValueAndRange
    crPrn.ParameterFields(4).AddCurrentValue txtEngine.Text
    crPrn.ParameterFields(5).ClearCurrentValueAndRange
    crPrn.ParameterFields(5).AddCurrentValue txtPrepBy.Text
    crPrn.PrintOut , 1
    rsPrn.Close

End Sub

Private Sub cmdTrans_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "POCUST"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
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
            '***Access
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
    '***Access
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & MyId & "));"
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
    '***Access
    'Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyId & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>='" & Format(StartDate, "yyyy-mm-dd") & "' And (SPITR.TRANSDATE)<='" & Format(EndDate, "yyyy-mm-dd") & "'));"
    
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
    End Select
End Sub

Private Sub flxOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxOrd_Click
End Sub
                                                  
Private Sub flxOrd_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "POCUST"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
End Sub

Private Sub flxOrd_LeaveCell()
    txtNoDiscount.Visible = False
    txtAppModel.Visible = False
    txtSQtyBck.Visible = False
End Sub

Private Sub flxOrd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Not JustScroll Then
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
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" And flxOrd.Rows <= 12 Then
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


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "POCUST"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
End Sub


Private Sub Form_Load()
    Dim Qry As String
    
    ActRow = 0
    ActCol = 0
    JustScroll = True
    txtDate.Text = Format(Date, "mm/dd/yyyy")

    Call BtnSaveCancel(False)
    Call LockData(True)
    Timer1.Enabled = True
    
    MotoRep.sBar.Panels(6).Text = "Customer's Purchase Order"
End Sub

Private Sub BtnSaveCancel(bln As Boolean)
    Dim i As Integer
    
    If bln Then
        For i = 0 To 4
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(5).Visible = True
        cmdTrans(5).Left = 0
        cmdTrans(6).Visible = True
        cmdTrans(6).Left = cmdTrans(5).Width
    Else
        For i = 0 To 4
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsParts.State = adStateOpen Then rsParts.Close
    If rsCustomer.State = adStateOpen Then rsCustomer.Close
    MySource = ""
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub mnuAdd_Click()
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    Else
        If flxOrd.Rows = 13 Then
            MsgBox "You can only process upto 12 Part Numbers only", vbExclamation + vbOKOnly, "Exceed limit"
            Exit Sub
        End If
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


Public Sub ShowRecord(MyID As Integer)
    Dim rsOrder As New ADODB.Recordset
    Dim rsPOdtls As New ADODB.Recordset
    Dim Qry As String, mySrc As String, MyUnit As String, myRem As String
    Dim TtlPrice As Single, OwnCode As Integer
    
    Call ClearEntries
        
    '***Access
    'Qry = "SELECT CUSTPO.* From CUSTPO WHERE (((CUSTPO.CPOID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT CUSTPO.* From CUSTPO WHERE CUSTPO.CPOID=" & MyID & ";"
    
    Call ClearEntries
    rsOrder.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsOrder.RecordCount > 0 Then
        rsOrder.MoveFirst
        txtPONO.Text = rsOrder!CPONO
        txtDate.Text = Format(rsOrder!CPODATE, "mm/dd/yyyy")
        txtCustID.Text = rsOrder!CUSTOMERID
        CustSrch = True
        cboCustomer.Text = rsOrder!CUSTOMER
        txtEngine.Text = IIf(IsNull(rsOrder!ENGINE), "", rsOrder!ENGINE)
        txtRemarks.Text = IIf(IsNull(rsOrder!Remarks), "", rsOrder!Remarks)
        txtPrepBy.Text = rsOrder!PREPBY
        txtAprvBy.Text = IIf(IsNull(rsOrder!APRVBY), "", rsOrder!APRVBY)
        lblInvoice.Caption = IIf(IsNull(rsOrder!InvoiceNO), "", "INVOICE NO. " & rsOrder!InvoiceNO)
        txtJOid.Text = IIf(IsNull(rsOrder!JOBORDERID), "", rsOrder!JOBORDERID)
        'access
        'If rsOrder!Approved = True Then
        
        If rsOrder!Approved = 1 Then
            ChkStatus.Value = 1
        Else
            ChkStatus.Value = 0
        End If
        txtPOid.Text = MyID
    End If
    rsOrder.Close
    '***Access
    'Qry = "SELECT ONACCTERMS.TERMS From ONACCTERMS WHERE (((ONACCTERMS.STATUS)='Default') AND ((ONACCTERMS.CUSTOMERID)=" & Val(txtCustID.Text) & "));"
    'mySQL
    Qry = "SELECT ONACCTERMS.TERMS From ONACCTERMS WHERE ONACCTERMS.STATUS='Default' AND ONACCTERMS.CUSTOMERID=" & Val(txtCustID.Text) & ";"
    txtTerm.Text = Val(GetFieldValue(Qry)) & " " & "DAYS"
    With rsPOdtls
        '***Access
        'Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, [PRICE]*[QTY] AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID " & _
              "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE (((SPSORDERDETAIL.CPOID)=" & MyId & "));"
        'mySQL
        Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, SPSORDERDETAIL.PRICE*SPSORDERDETAIL.QTY AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID " & _
              "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE SPSORDERDETAIL.CPOID=" & MyID & ";"
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = 1
            Do While Not .EOF
                flxOrd.AddItem vbTab & !QtyOnHand & vbTab & !Qty & vbTab & IIf(IsNull(!Unit), "", !Unit) & vbTab & !PartNumber & vbTab & !Description & vbTab & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlPrice, "###,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SparepartID   '& vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel)
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    If Val(txtJOid.Text) > 0 Then
        '***Access
        'Qry = "SELECT SRVCJOBORDER.JOBORDERNO From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERID)=" & txtJOid.Text & "));"
        'mySQL
        Qry = "SELECT SRVCJOBORDER.JOBORDERNO From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERID=" & txtJOid.Text & ";"
        lblInvoice.Caption = IIf(lblInvoice.Caption <> "", lblInvoice.Caption & ", ", "") & "JOB ORDER# " & GetFieldValue(Qry)
    End If
    Call ComputeTotal
End Sub

Private Sub SavePurchaseOrder()
    Dim rsPO As New ADODB.Recordset
    Dim rsPOdtls As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String, myTtl As Single, Ans As Integer, NewOnAcct As Boolean
    Dim eMsg As String
                                  
    'Saving Table Spare Parts Invoice
    On Error GoTo LocalErr
    eMsg = "Saving Purchase Order"
    
    '***Access
    'If AddNew Then
    '    rsPO.Open "CUSTPO", LYRcon, adOpenKeyset, adLockOptimistic
    '    rsPO.AddNew
    'Else
        '***Access
        'Qry = "SELECT CUSTPO.* From CUSTPO WHERE (((CUSTPO.CPOID)=" & txtPOid.Text & "));"
        'mySQL
    '    Qry = "SELECT CUSTPO.* From CUSTPO WHERE CUSTPO.CPOID=" & txtPOid.Text & ";"
        
    '    rsPO.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    rsPO.MoveFirst
    'End If
    'rsPO!CPONO = txtPONO.Text
    'rsPO!CPODATE = txtDate.Text
    'rsPO!PrepBy = txtPrepBy.Text
    'rsPO!AprvBy = txtAprvBy.Text
    'rsPO!CUSTOMER = cboCustomer.Text
    'rsPO!CUSTOMERID = Val(txtCustID.Text)
    'rsPO!JOBORDERID = Val(txtJOid.Text)
    'rsPO!ENGINE = txtEngine.Text
    'rsPO!REMARKS = txtRemarks.Text
    'rsPO!POAmount = Val(Replace(txtTotal.Text, ",", ""))
    'If ChkStatus.Value = 1 Then
    '    rsPO!Approved = 1
    'Else
    '    rsPO!Approved = 0
    'End If
    'rsPO.Update
    'rsPO.Close
    
    'mySQL
    'Saving Table Spare Parts Order Details
    If ChkStatus.Value = 1 Then
        POApp = 1
    Else
        POApp = 0
    End If
    
    If AddNew Then
        LYRcon.Execute "Insert Into custpo(CPONO, CPODATE, PREPBY, APRVBY, CUSTOMER, " & _
            "CUSTOMERID, JOBORDERID, ENGINE, REMARKS, POAMOUNT, APPROVED) " & _
            "VALUES('" & txtPONO.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
            "'" & txtPrepBy.Text & "', '" & txtAprvBy.Text & "', '" & cboCustomer.Text & "', " & _
            "" & Val(txtCustID.Text) & ", " & Val(txtJOid.Text) & ", '" & txtEngine.Text & "', " & _
            "'" & txtRemarks.Text & "', " & Val(Replace(txtTotal.Text, ",", "")) & ", " & POApp & ");"
    Else
        LYRcon.Execute "UPDATE custpo SET CPONO = '" & txtPONO.Text & "', CPODATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
            "PREPBY = '" & txtPrepBy.Text & "', APRVBY = '" & txtAprvBy.Text & "', CUSTOMER = '" & cboCustomer.Text & "', " & _
            "CUSTOMERID = " & Val(txtCustID.Text) & ", JOBORDERID = " & Val(txtJOid.Text) & ", ENGINE = '" & txtEngine.Text & "', " & _
            "REMARKS = '" & txtRemarks.Text & "', POAMOUNT = " & Val(Replace(txtTotal.Text, ",", "")) & ", APPROVED = " & POApp & ";"
    End If
    If AddNew Then
    Qry = "SELECT custpo.CPOID, custpo.CPONO From custpo Where custpo.CPONO = '" & txtPONO.Text & "';"
    txtPOid.Text = GetFieldValue(Qry)
    End If
    
    If Not AddNew Then
        '***Access
        'LYRcon.Execute "DELETE SPSORDERDETAIL.*, SPSORDERDETAIL.CPOID From SPSORDERDETAIL WHERE (((SPSORDERDETAIL.CPOID)=" & txtPOid.Text & "));"
        'mySQl
        LYRcon.Execute "DELETE From SPSORDERDETAIL WHERE SPSORDERDETAIL.CPOID=" & txtPOid.Text & ";"
    
    End If
    With flxOrd
        For i = 1 To flxOrd.Rows - 1
            eMsg = "Line 1136: Adding New Sales Order Details"
            LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( CPOID, SPAREPARTID, QTY, UNIT, PRICE) " & _
                           "VALUES (" & txtPOid.Text & ", " & .TextMatrix(i, 10) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ");"
        Next i
    End With
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbcrlr & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    
End Sub

Function ValidatePO() As String
    Dim MyErr As String, Qry As String
    Dim i As Integer
    
    MyErr = ""
    If txtPONO.Text = "" Then
        MyErr = "Incomplete entry on field P.O. Number"
    Else
        If AddNew Or (Not AddNew And txtPONO.Text <> txtPOid.Tag) Then
            '***Access
            'Qry = "SELECT CUSTPO.CPONO From CUSTPO WHERE (((CUSTPO.CPONO)='" & txtPONO.Text & "'));"
            'mySQL
            Qry = "SELECT CUSTPO.CPONO From CUSTPO WHERE CUSTPO.CPONO='" & txtPONO.Text & "';"
            If CtrlNumExist(Qry) Then
                MyErr = "P.O. Number " & txtPONO.Text & " already exist!"
            End If
        End If
    End If
    If Not IsDate(txtDate.Text) Then MyErr = MyErr & vbCrLf & "Invalid entry on field Date"
    If txtPrepBy.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field Prepare By"
    If cboCustomer.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field Customer"
    If Val(txtCustID.Text) = 0 Then MyErr = MyErr & vbCrLf & "Invalid Customer ID"
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
    ValidatePO = MyErr
End Function



Private Sub RXResizer1_ResizeComplete()
    
    On Error GoTo LocalErr
    With flxOrd
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.05
        .ColWidth(2) = (.Width - 300) * 0.09
        .ColWidth(3) = (.Width - 300) * 0.09
        .ColWidth(4) = (.Width - 300) * 0.22
        .ColWidth(5) = (.Width - 300) * 0.3
        .ColWidth(6) = 0
        .ColWidth(7) = (.Width - 300) * 0.12
        .ColWidth(8) = (.Width - 300) * 0.12
        .ColWidth(9) = 0
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
    Exit Sub
LocalErr:
    'Unload Me
End Sub

Private Sub Timer1_Timer()
    Dim Qry As String, TmpN As String
    Dim MyID As Integer
        
    Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
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
    
    
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME " & _
          "From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"
    With rsCustomer
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If cboCustomer.Text <> "" Then TmpN = cboCustomer.Text
            cboCustomer.Clear
            Do While Not .EOF
                cboCustomer.AddItem !CustomerName
                .MoveNext
            Loop
        End If
    End With
    If TmpN <> "" Then
        CustSrch = True
        cboCustomer.Text = TmpN
    End If
    If MySource = "JOB ORDER" Then
        With JobOrder
            If .cboPO.Text <> "" Then
                '***Access
                'Qry = "SELECT CUSTPO.CPOID From CUSTPO WHERE (((CUSTPO.CPONO)='" & .cboPO.Text & "'));"
                'mySQl
                Qry = "SELECT CUSTPO.CPOID From CUSTPO WHERE CUSTPO.CPONO='" & .cboPO.Text & "';"
                
                MyID = GetFieldValue(Qry)
                Call ShowRecord(MyID)
            Else
                If .txtCustomer.Text <> "" Then
                    Call cmdTrans_Click(0)
                    If Val(JobOrder.txtCustID) > 0 And JobOrder.chkOnAcct.Value = 1 Then
                        '***Access
                        'Qry = "SELECT CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE (((CUSTOMER.CUSTOMERID)=" & .txtCustID.Text & "));"
                        'mySQl
                        Qry = "SELECT CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE CUSTOMER.CUSTOMERID=" & .txtCustID.Text & ";"
                        
                        cboCustomer.Text = GetFieldValue(Qry)
                        txtRemarks.Text = .txtDesignation & " " & .txtCustomer.Text & "/ Plate No. " & .flxMC.TextMatrix(1, 3)
                    Else 'mySQL STOP HERE
                        'No txt component found
                        txtCustomer.Text = JobOrder.txtCustomer.Text
                    End If
                    txtPONO.SetFocus
                    lblInvoice.Caption = IIf(lblInvoice.Caption <> "", lblInvoice.Caption & ", ", "") & "Job Order# " & .txtJobNo.Text
                    txtEngine.Text = .flxMC.TextMatrix(1, 1)
                    txtJOid.Text = .txtJOid.Text
                End If
            End If
        End With
    End If
    Timer1.Enabled = False
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

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "POCUST"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
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
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" And flxOrd.Rows <= 12 Then
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


Private Sub txtEngine_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then txtRemarks.SetFocus
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
        txtPONO.SetFocus
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
    If ActCol = 1 Then
        If KeyAscii > 26 Then
            If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then
        Call txtFlx_LostFocus
        If txtFlx.Tag = "XcdQty" Then txtFlx.Tag = "": Exit Sub
        'If all fields are filled then addrow focus on next entry
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" And flxOrd.Rows <= 12 Then
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

Private Sub ComputeTotal()
    Dim i As Integer
    Dim TtlPrice As Currency, GrndTtl As Currency, DiscAmt As Currency, TtlDiscAmt As Currency
    Dim Pctd As Single, myPrice As Currency
    
    With flxOrd
        GrndTtl = 0
        TtlDiscAmt = 0
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 2)) > 0 And .TextMatrix(i, 7) <> "" Then
                TtlPrice = Val(.TextMatrix(i, 2)) * PureNum(.TextMatrix(i, 7))
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
        txtTotal.Text = Format$(GrndTtl, "#,##0.00")
    End With
            
End Sub

Private Sub txtPOno_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "POCUST"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
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
        Case vbKeyBack
            PrtSrch = True
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        'Call txtPartNo_LostFocus
        flxOrd.SetFocus
        'If all fields are filled then addrow focus on next entry
        If Val(flxOrd.TextMatrix(ActRow, 2)) > Val(flxOrd.TextMatrix(ActRow, 1)) And flxOrd.TextMatrix(ActRow, 2) <> "" Then
            MsgBox "Qty Ordered must not be greater than Qty Available!", vbExclamation + vbOKOnly, "Validation"
            flxOrd.Row = ActRow
            flxOrd.Col = 2
            Call flxOrd_Click
            Exit Sub
        End If
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" And flxOrd.Rows <= 12 Then
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

Private Sub txtPOno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtAddress.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        txtDate.SetFocus
    End If
End Sub

Private Sub txtPOno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDate.SetFocus
End Sub

Private Sub txtPOno_LostFocus()
    Dim Qry As String
    
    If txtPONO.Text <> "" Then
        If AddNew Or (Not AddNew And txtPONO.Text <> txtPOid.Tag) Then
            '***Access
            'Qry = "SELECT CUSTPO.CPONO From CUSTPO WHERE (((CUSTPO.CPONO)='" & txtPONO.Text & "'));"
            'mySQl
            Qry = "SELECT CUSTPO.CPONO From CUSTPO WHERE CUSTPO.CPONO='" & txtPONO.Text & "';"
            If CtrlNumExist(Qry) Then
                MsgBox "P.O. Number " & txtPONO.Text & " already exist!", vbExclamation + vbOKOnly, "Validation Error"
                txtPONO.Text = ""
                Exit Sub
            End If
        End If
    End If
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
    
    strValid = "0123456789.X"
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
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" And flxOrd.Rows <= 12 Then
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

Function GetCost(SPid As Long) As Single
    Dim rsCost As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPAREPARTS.COST From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & SPid & "));"
    'mySQL
    Qry = "SELECT SPAREPARTS.COST From SPAREPARTS WHERE SPAREPARTS.SPAREPARTID=" & SPid & ";"
    
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

Private Sub txtPrice_LostFocus()
    Dim AmtDsc As Currency, LessAmt As Currency, Pctd As Double
    
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
            If Val(txtPrice.Text) < AmtDsc Then
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
            If Val(txtPrice.Text) < AmtDsc Then
                MsgBox "Price must not be less than maximum discounted Price: " & Format(AmtDsc, "#,##0.00") & "!", vbCritical + vbOKOnly, "Validation"
                txtPrice.Tag = "XcdQty"
                txtPrice.SelStart = 0
                txtPrice.SelLength = Len(txtPrice.Text)
                txtPrice.SetFocus
                Exit Sub
            End If
        End If
        If Val(txtPrice.Text) < Val(.TextMatrix(ActRow, 13)) Then
            Pctd = (Val(.TextMatrix(ActRow, 13)) - Val(txtPrice.Text)) / Val(.TextMatrix(ActRow, 13))
            If Val(txtPrice.Text) > 0 Then .TextMatrix(ActRow, 6) = Format(Pctd, "#0.000%")
        End If
        If Val(txtPrice.Text) < AmtDsc Then 'Price is less than allowable discount
            .TextMatrix(ActRow, 3) = IIf(InStr(.TextMatrix(ActRow, 3), "*") > 0, "", .TextMatrix(ActRow, 3) & "*")
        Else
            .TextMatrix(ActRow, 3) = Replace(.TextMatrix(ActRow, 3), "*", "")
        End If
    End With
    flxOrd.TextMatrix(ActRow, 7) = Format(txtPrice.Text, "#,##0.00")
    Call ComputeTotal
    txtPrice.Visible = False

End Sub
