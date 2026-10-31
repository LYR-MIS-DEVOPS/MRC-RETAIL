VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form arPR 
   Caption         =   "Provisional Receipt"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   13080
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   57
      Tag             =   "XY"
      Text            =   "HoldGrayBackColor"
      Top             =   2940
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   13560
      TabIndex        =   56
      Tag             =   "XY"
      Text            =   "HoldBackColor"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrEnabledtmrShowRecord 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10800
      Top             =   0
   End
   Begin VB.Timer tmrShowRecord 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11280
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar PbaR 
      Height          =   135
      Left            =   120
      TabIndex        =   55
      Tag             =   "xy"
      Top             =   3480
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtTtlAlloc 
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
      Left            =   6480
      TabIndex        =   54
      Tag             =   "XYF"
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox txtAllocAmt 
      Height          =   285
      Left            =   11400
      TabIndex        =   53
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtInvNo 
      Height          =   285
      Left            =   9840
      TabIndex        =   52
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin lvButton.lvButtons_H cmdAutoAlloc 
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Tag             =   "xyf"
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
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
      Image           =   "ProvReceipt.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtTtlInvAmt 
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
      Left            =   5040
      TabIndex        =   49
      Tag             =   "XYF"
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox txtChkBal 
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
      Left            =   10320
      TabIndex        =   47
      Tag             =   "XYF"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   7080
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   9360
      Width           =   375
   End
   Begin VB.TextBox txtChksPRid 
      Height          =   285
      Left            =   6960
      TabIndex        =   39
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   7920
      TabIndex        =   38
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Height          =   315
      Left            =   7920
      TabIndex        =   37
      Top             =   9000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboBranch 
      Height          =   315
      Left            =   2640
      TabIndex        =   36
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   1920
      TabIndex        =   35
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid flxAOP 
      Height          =   4095
      Left            =   120
      TabIndex        =   11
      Tag             =   "XY"
      Top             =   3960
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   16777215
      ScrollBars      =   2
      FormatString    =   $"ProvReceipt.frx":059A
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCust 
      Height          =   1695
      Left            =   13560
      TabIndex        =   27
      Top             =   495
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2990
      _Version        =   393216
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Height          =   1800
      Left            =   120
      TabIndex        =   26
      Tag             =   "XY"
      Top             =   1680
      Width           =   12855
      Begin VB.ComboBox cboChkStat 
         Height          =   315
         Left            =   1800
         TabIndex        =   44
         Tag             =   "XYF"
         Top             =   1320
         Width           =   4335
      End
      Begin VB.ComboBox cboChkType 
         Height          =   315
         Left            =   10200
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtPaidTo 
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
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   947
         Width           =   4335
      End
      Begin MSMask.MaskEdBox txtChkDate 
         Height          =   315
         Left            =   10200
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   570
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtChkAmt 
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
         Left            =   10200
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   947
         Width           =   2535
      End
      Begin VB.TextBox txtChkNo 
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
         Left            =   10200
         MaxLength       =   20
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   203
         Width           =   2535
      End
      Begin VB.TextBox txtBank 
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
         MaxLength       =   70
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   575
         Width           =   4335
      End
      Begin VB.TextBox txtAcctName 
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
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   203
         Width           =   4335
      End
      Begin VB.ComboBox cbocashier 
         Height          =   315
         Left            =   4080
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Check Type "
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
         Index           =   11
         Left            =   8880
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   1365
         Width           =   900
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Check Status   "
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
         Index           =   10
         Left            =   360
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   1372
         Width           =   1095
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Check Amount     "
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
         Index           =   9
         Left            =   8880
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   1005
         Width           =   1275
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Check Date     "
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
         Index           =   8
         Left            =   8880
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   630
         Width           =   1050
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Check Number     "
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
         Index           =   7
         Left            =   8880
         TabIndex        =   31
         Tag             =   "XYF"
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Paid To   "
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
         Index           =   6
         Left            =   360
         TabIndex        =   30
         Tag             =   "XYF"
         Top             =   1005
         Width           =   660
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Name of Bank   "
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
         Index           =   5
         Left            =   360
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   627
         Width           =   1140
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Account Name   "
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
         Index           =   4
         Left            =   360
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Cashier"
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
         Index           =   12
         Left            =   3480
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   1395
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   720
      Width           =   12855
      Begin MSMask.MaskEdBox txtPRDate 
         Height          =   315
         Left            =   11040
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPRno 
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
         Left            =   11040
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   210
         Width           =   1695
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
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   600
         Width           =   4335
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
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   210
         Width           =   4335
      End
      Begin VB.Label lblH 
         Caption         =   "P.R. Date   "
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
         Left            =   10080
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "P.R. Number   "
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
         Index           =   2
         Left            =   9840
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Address   "
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
         Left            =   360
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name   "
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
         Left            =   360
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   240
         Width           =   1275
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "New PR"
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
      Image           =   "ProvReceipt.frx":0653
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   12
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Edit PR"
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
      Image           =   "ProvReceipt.frx":0BED
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   13
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "View Summary"
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
      Image           =   "ProvReceipt.frx":1187
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   14
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Exit PR"
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
      Image           =   "ProvReceipt.frx":1721
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Save PR"
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
      Image           =   "ProvReceipt.frx":1CBB
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Cancel PR"
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
      Image           =   "ProvReceipt.frx":2255
      cBack           =   -2147483633
   End
   Begin VB.Label lblTtlInvAmt 
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
      Left            =   2640
      TabIndex        =   50
      Tag             =   "XYF"
      Top             =   8100
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Check Balance:  "
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
      TabIndex        =   48
      Tag             =   "XYF"
      Top             =   3660
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "ALLOCATION OF CHECK"
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
      TabIndex        =   34
      Tag             =   "XYF"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Provisional Receipt"
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
      Left            =   9600
      TabIndex        =   20
      Top             =   480
      Width           =   3360
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Provisional Receipt"
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
      Left            =   9630
      TabIndex        =   21
      Top             =   510
      Width           =   3360
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO-R REPUBLIC CORP."
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
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   4000
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO-R REPUBLIC CORP."
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
      Left            =   150
      TabIndex        =   19
      Top             =   510
      Width           =   4000
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   1320
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13320
      m_MDI_Height    =   9165
   End
   Begin VB.Menu mnuAlloc 
      Caption         =   "Allocation"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Row             Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "arPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCustomer As New ADODB.Recordset
Dim rsAcctCust As New ADODB.Recordset
Dim CustSrch As Boolean, CustBln As Boolean, AddNew As Boolean, JustScroll As Boolean, ShowRec As Boolean
Dim ActRow As Integer, ActCol As Integer, ThisBranch As String, CashierID As Integer

Public OwnerForm As String
Public CHKSPRID As String

Private Sub ClearEntries()
    txtChkBal.Locked = False
    txtCustomer.Text = ""
    txtAddress.Text = ""
    txtPRno.Text = ""
    txtPRDate.Text = "__/__/__"
    txtAcctName.Text = ""
    txtBank.Text = ""
    cboChkStat.Text = ""
    txtChkNo.Text = ""
    txtChkDate.Text = "__/__/__"
    txtChkAmt.Text = ""
    flxAOP.Rows = 1
    flxAOP.Rows = 2
    txtCustID.Text = ""
    cboChkType.Text = ""
    txtChksPRid.Text = ""
    txtChksPRid.Tag = ""
    txtChkBal.Text = ""
    txtTtlAlloc.Text = ""
    txtTtlInvAmt.Text = ""
End Sub

Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    Frame2.Enabled = Not bln
'   flxAOP.Enabled = Not Bln
    JustScroll = bln
End Sub

Private Sub tbSaveCancel(bln As Boolean)
    Dim i As Integer
    
    If Not bln Then
        For i = 0 To 3
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
    Else
        For i = 0 To 3
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(4).Left = 0
        cmdTrans(5).Left = cmdTrans(4).Width
        cmdTrans(4).Visible = True
        cmdTrans(5).Visible = True
    End If
End Sub

Private Sub cboAccount_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If CustSrch Then
        CustSrch = False
        Exit Sub
    End If
    With cboAccount
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsAcctCust.MoveFirst
            rsAcctCust.Move i
            flxAOP.TextMatrix(ActRow, 5) = rsAcctCust!CUSTOMERID
            flxAOP.TextMatrix(ActRow, 3) = rsAcctCust!address
            flxAOP.TextMatrix(ActRow, 2) = IIf(IsNull(rsAcctCust!ContactPerson), "", rsAcctCust!ContactPerson)
            flxAOP.TextMatrix(ActRow, 4) = Format(rsAcctCust!OSbal, "#,##0.00")
            
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CustSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            flxAOP.TextMatrix(ActRow, 5) = ""
            flxAOP.TextMatrix(ActRow, 3) = ""
            flxAOP.TextMatrix(ActRow, 2) = ""
            flxAOP.TextMatrix(ActRow, 4) = ""
        End If
    End With
End Sub

Private Sub cboAccount_Click()
    Call cboAccount_Change
End Sub

Private Sub cboAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cboAccount.Text = ""
        flxAOP.TextMatrix(ActRow, 1) = cboAccount.Tag
        cboAccount.Tag = "skiplostfocus"
        flxAOP.SetFocus
        flxAOP.Col = 0
        cboAccount.Visible = False
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        CustSrch = True
        flxAOP.TextMatrix(ActRow, 5) = ""
        flxAOP.TextMatrix(ActRow, 3) = ""
        flxAOP.TextMatrix(ActRow, 2) = ""
        flxAOP.TextMatrix(ActRow, 4) = ""
    End If
End Sub

Private Sub cboAccount_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call cboAccount_LostFocus
        flxAOP.SetFocus
        If ActRow = flxAOP.Rows - 1 Then flxAOP.AddItem ""
        flxAOP.Col = 1
    End If
End Sub

Private Sub cboAccount_LostFocus()
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String, eMsg As String
    
    If cboAccount.Tag = "skiplostfocus" Then
        cboAccount.Tag = ""
        Exit Sub
    End If
    If Val(flxAOP.TextMatrix(ActRow, 5)) = 0 Then
        MsgBox "No record of customer, Please select from the list.", vbExclamation + vbOKOnly, "Validation Error"
        cboAccount.Visible = False
        Exit Sub
    End If
    flxAOP.TextMatrix(ActRow, 1) = cboAccount.Text
    cboAccount.Visible = False
End Sub

Private Sub HideObj()
    flxCust.Visible = False
   ' cboAccount.Visible = False
    txtRemarks.Visible = False
    txtAmount.Visible = False
End Sub

Private Sub cboBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cboBranch.Text = Empty
        flxAOP.TextMatrix(ActRow, 1) = cboBranch.Tag
        cboBranch.Tag = "skiplostfocus"
        flxAOP.SetFocus
        flxAOP.Col = 1
        cboBranch.Visible = False
    End If
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        Call cboBranch_LostFocus
        flxAOP.SetFocus
        flxAOP.Col = 2
    End If
End Sub

Private Sub cboBranch_LostFocus()
    If cboBranch.Tag = "skiplostfocus" Then
        cboBranch.Tag = ""
        Exit Sub
    End If
    If cboBranch.ListIndex = -1 Then
        MsgBox "Please select a valid Branch from the list!", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    
    flxAOP.TextMatrix(ActRow, 1) = cboBranch.Text
    cboBranch.Visible = False

End Sub

Private Sub cbocashier_Click()
    If cbocashier.ListIndex = -1 Then Exit Sub
    CashierID = cbocashier.ItemData(cbocashier.ListIndex)
End Sub
Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub
Function GetPRNo() As String
    Dim rsPRno As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    Qry = "SELECT IFNULL(Max(CHECKS.PRNO),0) AS MaxOfPRNO FROM CHECKS;"

    With rsPRno
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(!MaxOfPRNO), 1, Val(!MaxOfPRNO) + 1)
        Else
            MyNo = 1
        End If
        .Close
    End With
    GetPRNo = Format(MyNo, "0000")
End Function

Private Sub cmdAutoAlloc_Click()
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String, MyBal As Currency, TtlInvAmt As Currency, TtlAlloc As Currency
    Dim i As Integer, j As Integer
    
    If Val(txtCustID.Text) = 0 Or Val(Replace(txtChkBal.Text, ",", "")) = 0 Then Exit Sub
    
    With flxAOP
        If .TextMatrix(.Rows - 1, 1) = "" And .TextMatrix(.Rows - 1, 2) = "" Then
            If .Rows = 2 Then .Rows = 1 Else .RemoveItem .Rows - 1
        End If
    End With
'FLXAOP:1-BRANCH,2-INV.NO,3-INV.DATE,4-TERM,5-DUEDATE,6-INV.AMT,7-ALLOC.AMT,8-OR.NO,9-CUSTOMER,10-SPSORDERID,11-CustomerID
    Qry = "SELECT onaccterms.SPINVNO, onaccterms.TRANSDATE, onaccterms.TERMS, onaccterms.ACCTAMOUNT, onaccterms.SPSORDERID, customer.CUSTOMERNAME, customer.CUSTOMERID, onaccterms.`STATUS`, onaccterms.BALANCE " & _
          "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) " & _
          "Where customer.CUSTOMERID = " & Val(txtCustID.Text) & " And onaccterms.`STATUS` = 'OPEN' Order By onaccterms.TRANSDATE;"
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyBal = Val(Replace(txtChkBal.Text, ",", ""))
            TtlInvAmt = 0
            TtlAlloc = 0
            Do While Not .EOF
                flxAOP.AddItem vbTab & ThisBranch & vbTab & !SPINVNO & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !Terms & vbTab & Format(GetDueDate(!TRANSDATE, !Terms)) & vbTab & Format(!AcctAmount, "#,##0.00") & vbTab & _
                IIf(MyBal >= !AcctAmount, Format(!AcctAmount, "#,##0.00"), Format(MyBal, "#,##0.00")) & vbTab & vbTab & !CustomerName & vbTab & !SPSORDERID & vbTab & !CUSTOMERID
                TtlInvAmt = TtlInvAmt + !AcctAmount
                If MyBal >= !AcctAmount Then
                    MyBal = MyBal - !AcctAmount
                    TtlAlloc = TtlAlloc + !AcctAmount
                Else
                    MyBal = 0
                    TtlAlloc = TtlAlloc + MyBal
                End If
                .MoveNext
            Loop
        Else
            flxAOP.AddItem ""
        End If
        .Close
    End With
    txtTtlInvAmt.Text = Format(TtlInvAmt, "#,##0.00")
    txtTtlAlloc.Text = Format(TtlAlloc, "#,##0.00")
    If flxAOP.Rows > 2 Then
        With flxAOP
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 7)) = 0 Then
                    j = i
                    Exit For
                End If
            Next i
            .Col = 1: .Row = j: .ColSel = 10: .RowSel = .Rows - 1: .FillStyle = flexFillRepeat: .CellForeColor = QBColor(8)  ' .CellBackColor = RGB(255, 0, 0)
            .Col = 2

        End With
        
    End If
    cmdAutoAlloc.Visible = False

End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
    Dim Ans As Integer, Desig As String
    Dim eMsg As String, Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=46));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=46;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            cboChkStat.Width = txtPaidTo.Width
            Call ClearEntries
            Call LockData(False)
            Call tbSaveCancel(True)
            txtCustomer.SetFocus
            txtPRDate.Text = Format$(Date, "mm/dd/yy")
            AddNew = True
            ShowRec = False
            flxAOP.Enabled = False
            txtPRno.Text = GetPRNo

            
        Case 1  'Edit
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=46));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=46;"
            MyAcc = GetFieldValue(Qry)
            
            Qry = "SELECT PCUSER.DESIGNATION From PCUSER " & _
                "WHERE PCUSER.EMPID=" & UserID & ";"
            Desig = UCase$(GetFieldValue(Qry))
            
            'If MyAcc <> "FULL ACCESS" Then
            If MyAcc <> "FULL ACCESS" Or Desig <> "SYSTEM ADMIN." Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            
            If Val(txtChksPRid.Text) = 0 Then
                MsgBox "No available P.R. to edit!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            
            
            If Val(Replace(txtChkBal.Text, ",", "")) >= 1 Then
                cmdAutoAlloc.Enabled = True
            Else
                cmdAutoAlloc.Visible = False
                Qry = "SELECT checks.MGRCFID From checks Where checks.CHKSPRID = " & Val(txtChksPRid.Text) & ";"
                If Val(GetFieldValue(Qry)) > 0 Then
                    MsgBox "EDITING is not Allowed because a Cashflow was already processed using this Check!", vbExclamation + vbOKOnly, "Edit"
                    Exit Sub
                End If
            End If
            Ans = MsgBox("Do you want to Edit PR?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            cboChkStat.Width = cboChkStat.Width - (cbocashier.Width + lblH(12).Width)
            Call LockData(False)
            Call tbSaveCancel(True)
            txtCustomer.SetFocus
            txtChksPRid.Tag = txtPRno.Text
            If Val(txtChkBal.Text) > 1 Then JustScroll = False
            AddNew = False
            flxAOP.Enabled = False
            
            txtChkBal.Locked = True
            
        Case 2  'Summary
            arCheckSumm.Show
            Unload Me
        Case 3  'Quit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
        Case 4  'Save
            eMsg = ValidatePR
            If eMsg <> "" Then
                MsgBox eMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save PR?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'Call Save Check
            cboChkStat.Width = txtPaidTo.Width
            Call SaveCheck
            'MsgBox "P.R. successfully saved!", vbOKOnly + vbInformation, "Save"
        Case 5  'Cancel
            Call HideObj
            Ans = MsgBox("Do you want to Cancel PR?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            cboChkStat.Width = txtPaidTo.Width
            Call ClearEntries
            Call LockData(True)
            Call tbSaveCancel(False)
    End Select
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cmdTrans_Click()"

End Sub

Public Sub ShowRecord(MyID As Long)
    Dim rsCheck As New ADODB.Recordset
    Dim rsAlloc As New ADODB.Recordset
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String, TtlInvAmt As Currency, TtlAlloc As Currency, i As Integer
    
    ShowRec = True
    'mySQL
    Qry = "SELECT CHECKS.* From CHECKS WHERE CHECKS.CHKSPRID=" & MyID & ";"
    With cboChkStat
        .Clear
        .AddItem "CURRENT-CLEARED"
        .AddItem "PDC DUE"
        .AddItem "POST-DATED"
    End With
            
    With rsCheck
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            CustSrch = True
            txtCustomer.Text = IIf(IsNull(!CUSTOMER), "", !CUSTOMER)
            txtAddress.Text = IIf(IsNull(!address), "", !address)
            txtPRno.Text = IIf(IsNull(!PRNO), "", !PRNO)
            If Not IsNull(!PRDATE) Then txtPRDate.Text = Format(!PRDATE, "mm/dd/yy")
            txtCustID.Text = !CUSTOMERID
            txtAcctName.Text = !CHKACCTNAME
            txtBank.Text = !BANK
            txtChkNo.Text = !CHECKNO
            txtChkAmt.Text = Format(!Amount, "#,##0.00")
            If !STATUS = "PDC DUE" Then
                cboChkStat.ListIndex = 1
                'cboChkStat.Text = "CURRENT-FOR DEPOSIT"
            ElseIf !STATUS = "CURRENT-CLEARED" Then
                cboChkStat.ListIndex = 0
                cboChkStat.Text = IIf(IsNull(!STATUS), "", !STATUS)
            ElseIf !STATUS = "POST-DATED" Then
                cboChkStat.ListIndex = 2
                cboChkStat.Text = IIf(IsNull(!STATUS), "", !STATUS)
            Else
                cboChkStat.Text = IIf(IsNull(!STATUS), "", !STATUS)
            End If
            txtPaidTo.Text = !PaidTo
            txtChkDate.Text = Format(!CHECKDATE, "mm/dd/yy")
            cboChkType.Text = IIf(IsNull(!CHECKTYPE), "", !CHECKTYPE)
            txtChksPRid.Text = !CHKSPRID
            txtChkBal.Text = Format(!CHECKBAL, "#,##0.00")
            If !CHECKBAL > 0 Then
                cmdAutoAlloc.Visible = True
                cmdAutoAlloc.Enabled = False
            Else
                cmdAutoAlloc.Visible = False
            End If
        End If
        .Close
    End With
    'Display Allocated Invoices
    
    'Qry = "SELECT * From checkalloc Where checkalloc.CHKSPRID = " & Val(txtChksPRid.Text) & ";"
    
    DoEvents
    Dim REC_TYPE As String
    REC_TYPE = ""
    Call modAllocation.sp_check_allocation_list(Val(txtChksPRid.Text), rsAlloc)
    With rsAlloc
        '.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .State = 1 Then
            If .RecordCount > 0 Then
                .MoveFirst
                flxAOP.Rows = 1
                TtlInvAmt = 0
                TtlAlloc = 0
                Pbar.Value = 0
                Pbar.Max = .RecordCount * 2
                Pbar.Visible = True
                REC_TYPE = !REC_TYPE
                Do While Not .EOF
                   '                  0         1                                         2
                    flxAOP.AddItem vbTab & IIf(IsNull(!Branch), "", !Branch) & vbTab & IIf(IsNull(!RefNo), "", !RefNo)
                    flxAOP.TextMatrix(flxAOP.Rows - 1, 7) = Format(!AllocAmt, "#,##0.00")
                    flxAOP.TextMatrix(flxAOP.Rows - 1, 10) = !SPSORDERID
                    flxAOP.TextMatrix(flxAOP.Rows - 1, 11) = !CUSTOMERID
                    flxAOP.TextMatrix(flxAOP.Rows - 1, 12) = IIf(IsNull(!JOBORDERID), "", !JOBORDERID)
                    flxAOP.TextMatrix(flxAOP.Rows - 1, 13) = !VOUCHERID
                    
                    ' 0    1           2          3              4         5             6                7                8             9
                    '    |^BRANCH   |^INV. NO.  |^INV. DATE   |^TERM    |^DUE DATE    |>INV. AMOUNT    |>ALLOC. AMOUNT  |^RECEIPT NO. |<CUSTOMER NAME
                    If REC_TYPE = "NEW" Then
                        flxAOP.TextMatrix(flxAOP.Rows - 1, 3) = !TRANSDATE
                        flxAOP.TextMatrix(flxAOP.Rows - 1, 4) = !Terms
                        flxAOP.TextMatrix(flxAOP.Rows - 1, 5) = !DUEDATE
                        flxAOP.TextMatrix(flxAOP.Rows - 1, 6) = !AcctAmount
                        flxAOP.TextMatrix(flxAOP.Rows - 1, 8) = !RECEIPT_NO
                        flxAOP.TextMatrix(flxAOP.Rows - 1, 9) = !CustomerName
                    End If
                    
                    .MoveNext
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                Loop
            Else
                txtTtlInvAmt.Text = ""
                txtTtlAlloc.Text = ""
            End If
        Else
            txtTtlInvAmt.Text = ""
            txtTtlAlloc.Text = ""
        End If
        If .State <> 0 Then .Close
    End With
    
    If REC_TYPE = "OLD" Then
        With flxAOP
            Dim joNum As String
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" And (Val(.TextMatrix(i, 10)) > 0 Or Val(.TextMatrix(i, 12)) > 0 Or Val(.TextMatrix(i, 13))) Then
                    
                    If .TextMatrix(i, 1) = ThisBranch Then
                        If .TextMatrix(i, 12) > 0 Then 'ADDITIONAL
                                Qry = "SELECT onaccterms.SPINVNO, onaccterms.JOBORDERNO, onaccterms.JOBORDERID, onaccterms.TRANSDATE, onaccterms.TERMS, onaccterms.ACCTAMOUNT, onaccterms.SPSORDERID, customer.CUSTOMERNAME, customer.CUSTOMERID, onaccterms.`STATUS`, onaccterms.BALANCE " & _
                                        "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) " & _
                                        "Where onaccterms.JOBORDERID = '" & Val(.TextMatrix(i, 12)) & "' Order By onaccterms.TRANSDATE Desc;"
                        
                        ElseIf .TextMatrix(i, 13) > 0 Then
                            
                                Qry = "SELECT disbursement.CVNUMBER, disbursement.CVDATE, disbursement.AMOUNT, customer.CUSTOMERID, customer.CUSTOMERNAME, Disbursement.VOUCHERID " & _
                                        "From Disbursement INNER JOIN customer ON (disbursement.CUSTOMERID = customer.CUSTOMERID) " & _
                                        "Where Disbursement.VOUCHERID = '" & Val(.TextMatrix(i, 13)) & "' AND disbursement.CUSTOMERID > 0 order by disbursement.cvdate desc"
                        
                        Else
                            Qry = "SELECT onaccterms.SPINVNO, onaccterms.TRANSDATE, onaccterms.TERMS, onaccterms.ACCTAMOUNT, onaccterms.SPSORDERID, customer.CUSTOMERNAME, customer.CUSTOMERID, onaccterms.`STATUS`, onaccterms.BALANCE " & _
                                  "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) " & _
                                  "Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 10)) & " Order By onaccterms.TRANSDATE Desc;"
                        End If
                        
                        rsInv.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                        If rsInv.RecordCount > 0 Then
                            rsInv.MoveFirst
                            'FLXAOP:1-BRANCH,2-INV.NO,3-INV.DATE,4-TERM,5-DUEDATE,6-INV.AMT,7-ALLOC.AMT,8-OR.NO,9-CUSTOMER,10-SPSORDERID,11-CustomerID
                            If .TextMatrix(i, 13) > 0 Then
                                .TextMatrix(i, 3) = Format(rsInv!CVDate, "mm/dd/yyyy")
                                .TextMatrix(i, 4) = ""
                                .TextMatrix(i, 5) = ""
                                .TextMatrix(i, 6) = Format(rsInv!Amount, "#,##0.00")
                                .TextMatrix(i, 9) = rsInv!CustomerName
                            Else
                                .TextMatrix(i, 3) = Format(rsInv!TRANSDATE, "mm/dd/yyyy")
                                .TextMatrix(i, 4) = rsInv!Terms
                                .TextMatrix(i, 5) = DateAdd("d", rsInv!Terms, rsInv!TRANSDATE) 'Format(GetDueDate(rsInv!TransDate, rsInv!Terms), "mm/dd/yyyy")
                                .TextMatrix(i, 6) = Format(rsInv!AcctAmount, "#,##0.00")
                                .TextMatrix(i, 9) = rsInv!CustomerName
                                    If .TextMatrix(i, 12) > 0 Then
                                        .TextMatrix(i, 8) = GetFieldValue("SELECT Ledger.RefNo From Ledger Where ledger.JOBID = '" & .TextMatrix(i, 12) & "' AND ledger.CUSTOMERID = " & .TextMatrix(i, 11) & " AND ledger.REFTYPE = 'Payment' Order By Ledger.REFDATE Desc;")
                                    Else
                                        .TextMatrix(i, 8) = GetFieldValue("SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & .TextMatrix(i, 2) & "' AND ledger.CUSTOMERID = " & .TextMatrix(i, 11) & " AND ledger.REFTYPE = 'Payment' Order By Ledger.REFDATE Desc;")
                                    End If
                            End If
                        End If
                        rsInv.Close
                   
                    Else
                        If ThisBranch = "MRC2" Then
                            ActRow = i
                            If flxAOP.TextMatrix(i, 1) = "VAN1" Then
                                If ConnectVAN1 Then Call ShowBranchInv(.TextMatrix(i, 1), .TextMatrix(i, 2))
                            ElseIf flxAOP.TextMatrix(i, 1) = "VAN2" Then
                                If ConnectVAN2 Then Call ShowBranchInv(.TextMatrix(i, 1), .TextMatrix(i, 2))
                            ElseIf flxAOP.TextMatrix(i, 1) = "VAN3" Then
                                If ConnectVan3 Then Call ShowBranchInv(.TextMatrix(i, 1), .TextMatrix(i, 2))
                            ElseIf flxAOP.TextMatrix(i, 1) = "MRC1" Then
                                If ConnectMRC1 Then Call ShowBranchInv(.TextMatrix(i, 1), .TextMatrix(i, 2))
                            End If
                        End If
                    End If
                End If
                Pbar.Value = Pbar.Value + 1
                DoEvents
            Next i
        End With
    End If
    Call ComputeTotal
    ShowRec = False
    Pbar.Visible = False
    
    'Insert Allocated Payments
    'Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CONTACTPERSON, CUSTOMER.ADDRESS, CUSTOMER.OSBAL, CHECKALLOC.CUSTOMERID, CHECKALLOC.CHKSPRID " & _
          "FROM CHECKALLOC INNER JOIN CUSTOMER ON CHECKALLOC.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((CHECKALLOC.CHKSPRID)=" & MyID & "));"
    'mySQL
'    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CONTACTPERSON, CUSTOMER.ADDRESS, CUSTOMER.OSBAL, CHECKALLOC.CUSTOMERID, CHECKALLOC.CHKSPRID " & _
          "FROM CHECKALLOC INNER JOIN CUSTOMER ON CHECKALLOC.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE CHECKALLOC.CHKSPRID=" & MyID & ";"

    
'    With rsAlloc
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            flxAOP.Rows = 1
'            Do While Not .EOF
'                flxAOP.AddItem vbTab & !CustomerName & vbTab & !ContactPerson & vbTab & !Address & vbTab & Format(!OSbal, "#,##0.00") & vbTab & !CustomerID
'                .MoveNext
'            Loop
'        End If
'        .Close
'    End With
End Sub

Private Sub SaveCheck()
On Error GoTo er
    Dim ChkAmt As Double, i As Integer, Qry As String, myDate4BrkDwn As String, myDate4Dpsit As String
    Dim myStatus As String
    Dim tmpRs As New ADODB.Recordset
    If cboChkStat.Text = "PDC DUE" Then
        myDate4Dpsit = Trim(txtPRDate.Text)
        myStatus = "PDC DUE"
        myDate4BrkDwn = ""
    ElseIf cboChkStat.Text = "CURRENT-CLEARED" Then
        myDate4BrkDwn = Trim(txtPRDate.Text)
        myDate4Dpsit = ""
        myStatus = cboChkStat.Text
    Else
        myStatus = cboChkStat.Text
    End If
    ChkAmt = Val(Replace(txtChkAmt.Text, ",", ""))
    
    If AddNew Then
        'LYRcon.Execute "Insert Into checks(CUSTOMER, ADDRESS, PRNO, PRDATE, CUSTOMERID, CHKACCTNAME, BANK, " & _
        '    "CHECKNO, AMOUNT, `STATUS`, PAIDTO, CHECKDATE, CHECKTYPE, CHECKBAL, UPDATEDBY, LASTUPDATE)" & _
        '    "VALUES('" & txtCustomer.Text & "', '" & txtAddress.Text & "', '" & txtPRno.Text & "', " & _
        '    "'" & Format(txtPRDate.Text, "yyyy-mm-dd") & "', " & Val(txtCustID.Text) & ", '" & txtAcctName.Text & "', " & _
        '    "'" & txtBank.Text & "', '" & txtChkNo.Text & "', " & ChkAmt & ", '" & myStatus & "', '" & txtPaidTo.Text & "', " & _
        '    "'" & Format(txtChkDate.Text, "yyyy-mm-dd") & "', '" & cboChkType.Text & "', " & ChkAmt & ", " & UserID & ", " & _
        '    "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        'Qry = "select checks.CHKSPRID from checks where checks.CHECKNO='" & txtChkNo.Text & "' AND checks.CUSTOMER = '" & txtCustomer.Text & "' AND checks.CHECKDATE = '" & Format(txtChkDate.Text, "yyyy-mm-dd") & "';"
        'txtChksPRid.Text = GetFieldValue(Qry)
        'If IsDate(myDate4BrkDwn) Then
        '    LYRcon.Execute "UPDATE checks SET DATE4BRKDWN = '" & Format(DateValue(myDate4BrkDwn), "YYYY-MM-DD") & "' Where checks.CHKSPRID = " & txtChksPRid.Text & ";"
        'End If
        'If IsDate(myDate4Dpsit) Then
        '    LYRcon.Execute "UPDATE checks SET DATE4DPSIT = '" & Format(DateValue(myDate4Dpsit), "YYYY-MM-DD") & "' Where checks.CHKSPRID = " & txtChksPRid.Text & ";"
        'End If
        
        Call modAllocation.sp_checks_create(Val(txtCustID.Text), txtCustomer.Text, txtAddress.Text, txtPRno.Text, Format(txtPRDate.Text, "yyyy-mm-dd"), txtAcctName.Text _
                    , txtBank.Text, txtChkNo.Text, ChkAmt, myStatus, txtPaidTo.Text, Format(txtChkDate.Text, "yyyy-mm-dd"), cboChkType.Text, myDate4BrkDwn, myDate4Dpsit, tmpRs)
        
        
    Else
        'LYRcon.Execute "Update checks SET CUSTOMER = '" & txtCustomer.Text & "', ADDRESS = '" & txtAddress.Text & "', " & _
        '     "PRNO = '" & txtPRno.Text & "', PRDATE = '" & Format(txtPRDate.Text, "yyyy-mm-dd") & "', CUSTOMERID = " & Val(txtCustID.Text) & ", " & _
        '     "CHKACCTNAME = '" & txtAcctName.Text & "', BANK = '" & txtBank.Text & "', CHECKNO = '" & txtChkNo.Text & "', " & _
        '     "AMOUNT = " & ChkAmt & ", `STATUS` = '" & myStatus & "', PAIDTO = '" & txtPaidTo.Text & "', " & _
        '     "CHECKDATE = '" & Format(txtChkDate.Text, "yyyy-mm-dd") & "', CHECKTYPE = '" & cboChkType.Text & "', " & _
        '     "UPDATEDBY = " & CashierID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
        '     "Where checks.CHKSPRID = " & txtChksPRid.Text & ";"
        
        Call modAllocation.sp_checks_update(Val(txtChksPRid.Text), Val(txtCustID.Text), txtCustomer.Text, txtAddress.Text, txtPRno.Text, Format(txtPRDate.Text, "yyyy-mm-dd"), txtAcctName.Text _
                    , txtBank.Text, txtChkNo.Text, ChkAmt, myStatus, txtPaidTo.Text, Format(txtChkDate.Text, "yyyy-mm-dd"), cboChkType.Text, tmpRs)
        
    End If
    
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            If tmpRs!level = "Granted" Then
                If AddNew Then
                    txtChksPRid.Text = tmpRs!CheckID
                End If
                MsgBox tmpRs!message, vbInformation, "SaveCheck()"
                
                Call LockData(True)
                Call tbSaveCancel(False)
            Else
                If AddNew Then
                    txtChksPRid.Text = ""
                End If
                MsgBox tmpRs!message, vbCritical, "Error->SaveCheck()"
            End If
        Else
            GoTo UNKNOWN_PARAMETER
        End If
        
    Else 'tmpRS.State = 0
UNKNOWN_PARAMETER:
        MsgBox "UNKNOWN EXPECTED RETURN PARAMETER ON THE CURRENT EXECUTION!", vbCritical, "Error->SaveCheck()"
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    
    'Comment By SMC
    '2016-08-11
    
'    'Update Check Allocation
'    If cboChkType.Text = "CUSTOMER CHECK" And Val(txtChkBal.Text) > 0 Then
'        If Not AddNew Then
'            LYRcon.Execute "DELETE From checkalloc Where checkalloc.CHKSPRID = " & Val(txtChksPRid.Text) & " AND checkalloc.ORDATE IS NULL;"
'        End If
'        With flxAOP
'            For i = 1 To .Rows - 1
'                If .TextMatrix(i, 9) <> "" And Val(.TextMatrix(i, 10)) > 0 And Val(.TextMatrix(i, 7)) > 0 And Val(.TextMatrix(i, 6)) > 0 And .TextMatrix(i, 8) = "" Then
'                    If UCase$(Left$(.TextMatrix(i, 2), 3)) = "JO#" Then
'                        MyJo = Replace$(.TextMatrix(i, 2), "JO#", "")
'                    ElseIf UCase$(Left$(.TextMatrix(i, 2), 4)) = "PCV#" Then
'                        MyPCV = Replace$(.TextMatrix(i, 2), "PCV#", "")
'                    Else
'                        MyPCV = ""
'                        MyJo = ""
'                    End If
'                    If MyJo <> "" Then
'                        LYRcon.Execute "INSERT INTO checkalloc(CUSTOMERID, CHKSPRID, JOBORDERID, ALLOCAMT, BRANCH, REFNO) VALUES(" & Val(.TextMatrix(i, 11)) & ", " & Val(txtChksPRid.Text) & ", " & Val(.TextMatrix(i, 10)) & ", " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "');"
'
'                    ElseIf MyPCV <> "" Then
'                        LYRcon.Execute "INSERT INTO checkalloc(CUSTOMERID, CHKSPRID, VOUCHERID, ALLOCAMT, BRANCH, REFNO) VALUES(" & Val(.TextMatrix(i, 11)) & ", " & Val(txtChksPRid.Text) & ", " & Val(.TextMatrix(i, 10)) & ", " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "');"
'                    Else
'                        LYRcon.Execute "INSERT INTO checkalloc(CUSTOMERID, CHKSPRID, SPSORDERID, ALLOCAMT, BRANCH, REFNO) " & _
'                        "VALUES(" & Val(.TextMatrix(i, 11)) & ", " & Val(txtChksPRid.Text) & ", " & Val(.TextMatrix(i, 10)) & ", " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "');"
'                    End If
'                    'update table based of Branch
'                    If flxAOP.TextMatrix(i, 1) = ThisBranch Then
''                        If MyPCV <> "" Then
''                            Exit Sub
'                        If MyJo <> "" Then
'                            LYRcon.Execute "Update onaccterms SET CHKSPRID = " & Val(txtChksPRid.Text) & ", PAYSTART = '" & Format(DateValue(.TextMatrix(i, 5)), "yyyy-mm-dd") & "' " & _
'                            "Where onaccterms.JOBORDERID = " & Val(.TextMatrix(i, 10)) & " AND onaccterms.CUSTOMERID = " & Val(.TextMatrix(i, 11)) & ";"
'                        Else
'                            LYRcon.Execute "Update onaccterms SET CHKSPRID = " & Val(txtChksPRid.Text) & ", PAYSTART = '" & Format(DateValue(.TextMatrix(i, 5)), "yyyy-mm-dd") & "' Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 10)) & " AND onaccterms.CUSTOMERID = " & Val(.TextMatrix(i, 11)) & ";"
'
'                        End If
'                    ElseIf flxAOP.TextMatrix(i, 1) = "MRC1" Then
'                        conMRC1.Execute "Update onaccterms SET CHKSPRID = " & Val(txtChksPRid.Text) & ", PAYSTART = '" & Format(DateValue(.TextMatrix(i, 5)), "yyyy-mm-dd") & "' " & _
'                        "Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 10)) & " AND onaccterms.CUSTOMERID = " & Val(.TextMatrix(i, 11)) & ";"
'                    ElseIf flxAOP.TextMatrix(i, 1) = "VAN1" Then
'                        conVAN1.Execute "Update onaccterms SET CHKSPRID = " & Val(txtChksPRid.Text) & ", PAYSTART = '" & Format(DateValue(.TextMatrix(i, 5)), "yyyy-mm-dd") & "' " & _
'                        "Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 10)) & " AND onaccterms.CUSTOMERID = " & Val(.TextMatrix(i, 11)) & ";"
'                    ElseIf flxAOP.TextMatrix(i, 1) = "VAN2" Then
'                        conVAN2.Execute "Update onaccterms SET CHKSPRID = " & Val(txtChksPRid.Text) & ", PAYSTART = '" & Format(DateValue(.TextMatrix(i, 5)), "yyyy-mm-dd") & "' " & _
'                        "Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 10)) & " AND onaccterms.CUSTOMERID = " & Val(.TextMatrix(i, 11)) & ";"
'                    ElseIf flxAOP.TextMatrix(i, 1) = "VAN3" Then
'                        conVAN3.Execute "Update onaccterms SET CHKSPRID = " & Val(txtChksPRid.Text) & ", PAYSTART = '" & Format(DateValue(.TextMatrix(i, 5)), "yyyy-mm-dd") & "' " & _
'                        "Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 10)) & " AND onaccterms.CUSTOMERID = " & Val(.TextMatrix(i, 11)) & ";"
'                    End If
'                ElseIf i = 1 And .TextMatrix(i, 2) = "" Then
'                    LYRcon.Execute "INSERT INTO checkalloc(CUSTOMERID, CHKSPRID) " & _
'                    "VALUES(" & Val(txtCustID.Text) & ", " & Val(txtChksPRid.Text) & ");"
'                    'To allow viewing of checks in Receipt Transaction
'                End If
'            Next i
'        End With
'    End If

Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->SaveCheck()"
End Sub

Function ValidatePR() As String
    Dim vMsg As String
    
    vMsg = ""
    If cboChkType.Text <> "RE-DISCOUNT" Then
        If txtCustID.Text = "" And (flxAOP.Rows = 2 And Val(flxAOP.TextMatrix(1, 5)) = 0) Then vMsg = "Invalid Customer Name"
    End If
    If txtPRno.Text = "" Then
        If cboChkStat.Text <> "CURRENT-CLEARED" Then
            vMsg = vMsg & vbCrLf & "Incomplete entry on field P.R. Number"
        End If
    Else
        If AddNew Or (Not AddNew And txtPRno.Text <> txtChksPRid.Tag) Then
            'If CtrlNumExist("SELECT CHECKS.* From CHECKS WHERE (((CHECKS.PRNO)='" & txtPRno.Text & "'));") Then vMsg = vMsg & vbCrLf & "P.R. NUMBER already exist"
            'mySQL
            If CtrlNumExist("SELECT CHECKS.* From CHECKS WHERE CHECKS.PRNO='" & txtPRno.Text & "';") Then vMsg = vMsg & vbCrLf & "P.R. NUMBER already exist"
        End If
    End If
    If Not IsDate(txtPRDate.Text) Then vMsg = vMsg & vbCrLf & "Invalid entry on field P.R. Date"
    If txtBank.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Bank Name"
    If txtChkNo.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Check Number"
    If Not IsDate(txtChkDate.Text) Then vMsg = vMsg & vbCrLf & "Invalid entry on field Check Date"
    If txtChkAmt.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Check Amount"
    If cboChkStat.Text = "" Or cboChkStat.ListIndex = -1 Then vMsg = vMsg & vbCrLf & "Invalid entry on field Check Status"
    If cboChkType.Text = "" Or cboChkType.ListIndex = -1 Then vMsg = vMsg & vbCrLf & "Invalid entry on field Check Type"
    'Remove Last Blank Row of flxAOP
    With flxAOP
        If .TextMatrix(.Rows - 1, 10) = "" And .Rows > 2 Then .RemoveItem .Rows - 1
        'If .TextMatrix(.Rows - 1, 5) = "" And cboChkType.Text = "CUSTOMER CHECK" Then vMsg = vMsg & vbCrLf & "No Customer allocated for this Check"
        'If cboChkType.Text <> "RE-DISCOUNT" Then
        '    If .TextMatrix(.Rows - 1, 5) = "" And Val(txtCustID.Text) = 0 Then vMsg = vMsg & vbCrLf & "You are required to allocate this Check to a Customer"
        'End If
    End With

    ValidatePR = vMsg

End Function

Private Sub flxAOP_Click()
    Exit Sub
    
    ActRow = flxAOP.Row
    ActCol = flxAOP.Col
    
    'FLXAOP:1-BRANCH,2-INV.NO,3-INV.DATE,4-TERM,5-DUEDATE,6-INV.AMT,7-ALLOC.AMT,8-OR.NO,9-CUSTOMER,10-SPSORDERID,11-CustomerID
    
    If ActCol = 0 Or ActRow = 0 Or JustScroll Then Exit Sub   'Or Val(txtCustID.Text) > 0 Then Exit Sub
    'If flxAOP.TextMatrix(ActRow, ActCol) <> "" Then Exit Sub
    If Not AddNew And Val(Replace(txtChkBal.Text, ",", "")) = 0 Then
        MsgBox "Editing is not allowed because check balance is already zero.", vbExclamation + vbOKOnly, "Check Allocation"
        Exit Sub
    End If
    Select Case ActCol
        Case 1
            With cboBranch
                If flxAOP.TextMatrix(ActRow, 1) <> "" Then .Tag = flxAOP.TextMatrix(ActRow, 1)
                .Top = flxAOP.CellTop + flxAOP.Top
                .Left = flxAOP.CellLeft + flxAOP.Left
                .Width = flxAOP.CellWidth
                .Text = flxAOP.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        
        Case 2
            With txtInvNo
                If flxAOP.TextMatrix(ActRow, 2) <> "" Then .Tag = flxAOP.TextMatrix(ActRow, 2)
                .Top = flxAOP.CellTop + flxAOP.Top
                .Left = flxAOP.CellLeft + flxAOP.Left
                .Width = flxAOP.CellWidth
                .Text = flxAOP.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 7
            If Not AddNew And flxAOP.TextMatrix(ActRow, 8) <> "" Then
                MsgBox "Editing is not allowed because a Receipt was already issued", vbExclamation + vbOKOnly, "Check Allocation"
                Exit Sub
            End If
            With txtAllocAmt
                If flxAOP.TextMatrix(ActRow, 7) <> "" Then .Tag = flxAOP.TextMatrix(ActRow, 7)
                .Top = flxAOP.CellTop + flxAOP.Top
                .Left = flxAOP.CellLeft + flxAOP.Left
                .Width = flxAOP.CellWidth
                .Text = flxAOP.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        
    End Select
End Sub

Private Sub flxAOP_DblClick()
    Exit Sub
    
    ActRow = flxAOP.Row
    ActCol = flxAOP.Col
    
    'FLXAOP:1-BRANCH,2-INV.NO,3-INV.DATE,4-TERM,5-DUEDATE,6-INV.AMT,7-ALLOC.AMT,8-OR.NO,9-CUSTOMER,10-SPSORDERID,11-CustomerID
    
    If ActCol = 0 Or ActRow = 0 Or JustScroll Then Exit Sub   'Or Val(txtCustID.Text) > 0 Then Exit Sub
    If Not AddNew And Val(Replace(txtChkBal.Text, ",", "")) = 0 Then
        MsgBox "Editing is not allowed because check balance is already zero.", vbExclamation + vbOKOnly, "Check Allocation"
        Exit Sub
    End If
    
    Select Case ActCol
        Case 1    'BRANCH
            With cboBranch
                If flxAOP.TextMatrix(ActRow, 1) <> "" Then .Tag = flxAOP.TextMatrix(ActRow, 1)
                .Top = flxAOP.CellTop + flxAOP.Top
                .Left = flxAOP.CellLeft + flxAOP.Left
                .Width = flxAOP.CellWidth
                .Text = flxAOP.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 2
            With txtInvNo
                If flxAOP.TextMatrix(ActRow, 2) <> "" Then .Tag = flxAOP.TextMatrix(ActRow, 2)
                .Top = flxAOP.CellTop + flxAOP.Top
                .Left = flxAOP.CellLeft + flxAOP.Left
                .Width = flxAOP.CellWidth
                .Text = flxAOP.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 7
            With txtAllocAmt
                If flxAOP.TextMatrix(ActRow, 7) <> "" Then .Tag = flxAOP.TextMatrix(ActRow, 7)
                .Top = flxAOP.CellTop + flxAOP.Top
                .Left = flxAOP.CellLeft + flxAOP.Left
                .Width = flxAOP.CellWidth
                .Text = flxAOP.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        
    End Select

End Sub

Private Sub flxAOP_KeyPress(KeyAscii As Integer)
    Exit Sub
    If KeyAscii = 13 Then Call flxAOP_DblClick
End Sub

Private Sub flxAOP_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Exit Sub
    If Button = vbRightButton And ActCol = 0 Then
        Call PopupMenu(mnuAlloc)
    End If
End Sub

Private Sub flxCust_Click()
    txtCustomer.Text = flxCust.TextMatrix(flxCust.Row, 2)
    txtAddress.Text = flxCust.TextMatrix(flxCust.Row, 3)
    txtAcctName.Text = flxCust.TextMatrix(flxCust.Row, 2)
    txtCustID.Text = flxCust.TextMatrix(flxCust.Row, 1)
End Sub

Private Sub flxCust_EnterCell()
    CustSrch = True
    txtCustomer.Text = flxCust.TextMatrix(flxCust.Row, 2)
    txtAddress.Text = flxCust.TextMatrix(flxCust.Row, 3)
    txtAcctName.Text = flxCust.TextMatrix(flxCust.Row, 2)
    txtCustID.Text = flxCust.TextMatrix(flxCust.Row, 1)
End Sub

Private Sub flxCust_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxCust.Visible = False
        txtCustomer.SetFocus
    End If
End Sub

Private Sub flxCust_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxCust.Visible = False
        Call txtCustomer_LostFocus
        txtAddress.SetFocus
    End If
End Sub

Private Sub Form_Load()
On Error GoTo er
    Dim rsCashier As New ADODB.Recordset
    Dim MyID As Integer, Qry As String
    
    If SysBranch = "LYR MOTORPARTS CORP. - SOBRECARY" Then
        ThisBranch = "MRC1"
        cboBranch.AddItem "MRC1"
        'cboBranch.AddItem "MRC2"
        'cboBranch.AddItem "VAN1"
        'cboBranch.AddItem "VAN2"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - WHOLESALE" Then
        ThisBranch = "MRC2"
        cboBranch.AddItem "MRC1"
        cboBranch.AddItem "MRC2"
        cboBranch.AddItem "VAN1"
        cboBranch.AddItem "VAN2"
        cboBranch.AddItem "VAN3"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - GANTE" Then
        ThisBranch = "MRC2 RETAIL"
        cboBranch.AddItem "MRC2 RETAIL"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN-2 [LHC403]" Then
        ThisBranch = "VAN1"
        cboBranch.AddItem "VAN1"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN 2" Then
        ThisBranch = "VAN2"
        cboBranch.AddItem "VAN2"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - ULAS" Then
        ThisBranch = "ULAS"
        cboBranch.AddItem "ULAS"
    ElseIf SysBranch = "MC PARTS DEPOT TORIL" Then
        ThisBranch = "TORIL"
        cboBranch.AddItem "TORIL"
    ElseIf SysBranch = "LYR MOTORPARTS PANABO" Then
        ThisBranch = "PANABO"
        cboBranch.AddItem "PANABO"
    ElseIf SysBranch = "LYR MOTOR SHOP AND SPAREPARTS DAVAO" Then
        ThisBranch = "LYR SHOP DAVAO"
        cboBranch.AddItem "LYR SHOP DAVAO"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - LAGAO" Then
        ThisBranch = "GENSAN"
        cboBranch.AddItem "GENSAN"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN-55 [AAH1250]" Then
        ThisBranch = "VAN3"
        cboBranch.AddItem "VAN3"
    End If
    Timer1.Enabled = True
    Call tbSaveCancel(False)
    Call LockData(True)
    
    MyID = -1
    Qry = "SELECT concat(employees.FIRSTNAME,' ', IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.')),' ', employees.LASTNAME) AS CASHIERS, employees.EmpId " & _
              "From employees Where employees.`POSITION` = 'CASHIER' Order By  employees.LASTNAME;"
    With rsCashier
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cbocashier.AddItem !Cashiers
                cbocashier.ItemData(cbocashier.ListCount - 1) = !EMPID
                If !EMPID = UserID Then MyID = cbocashier.ListCount - 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    If cbocashier.ListCount > 0 And MyID >= 0 Then
        cbocashier.ListIndex = MyID
    End If
    
    DoEvents
    
    Call RXResizer1_ResizeComplete
    DoEvents
    If OwnerForm = "arCheckSumm" And Trim(CHKSPRID) <> "" Then
        DoEvents
        DoEvents
        tmrEnabledtmrShowRecord.Enabled = True
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->Form_Load()"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsCustomer.State = adStateOpen Then rsCustomer.Close
    If rsAcctCust.State = adStateOpen Then rsAcctCust.Close
    If conVAN1.State = adStateOpen Then conVAN1.Close
    If conVAN2.State = adStateOpen Then conVAN2.Close
    If conMRC1.State = adStateOpen Then conMRC1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OwnerForm = ""
    CHKSPRID = ""
End Sub

 

Private Sub mnuAdd_Click()
    'validate Last Row
    With flxAOP
        If .TextMatrix(.Rows - 1, 9) = "" And (Val(.TextMatrix(.Rows - 1, 10)) = 0 Or Val(.TextMatrix(.Rows - 1, 11)) = 0) Then
            MsgBox "Incomplete entry of Last Row!", vbExclamation + vbOKOnly, "Validation Error"
            Exit Sub
        End If
        .AddItem ""
        .SetFocus
        .Row = .Rows - 1
        .Col = 1
    End With
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    If Not AddNew And flxAOP.TextMatrix(ActRow, 8) <> "" Then
        MsgBox "Cannot delete row because an Receipt has been issued!", vbInformation + vbOKOnly, "Warning"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to Remove Invoice No. " & flxAOP.TextMatrix(ActRow, 2) & " from " & flxAOP.TextMatrix(ActRow, 1) & "?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbNo Then Exit Sub
    If ActRow = 1 And flxAOP.Rows = 2 Then
        flxAOP.Rows = 1
        flxAOP.Rows = 2
    Else
        flxAOP.RemoveItem ActRow
    End If
    'Call ComputeTotal
End Sub

Private Sub RXResizer1_ResizeComplete()
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    With flxAOP
        .Col = 0: .Row = 0
        .ColSel = .Cols - 1: .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .Row = 1
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.1
        .ColWidth(2) = (.Width - 300) * 0.09
        .ColWidth(3) = (.Width - 300) * 0.09
        .ColWidth(4) = (.Width - 300) * 0.07
        .ColWidth(5) = (.Width - 300) * 0.09
        .ColWidth(6) = (.Width - 300) * 0.11
        .ColWidth(7) = (.Width - 300) * 0.11
        .ColWidth(8) = (.Width - 300) * 0.09
        .ColWidth(9) = (.Width - 300) * 0.22
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        txtTtlInvAmt.Left = .Left + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(0)
        txtTtlAlloc.Left = txtTtlInvAmt.Left + .ColWidth(6)
        txtTtlInvAmt.Width = .ColWidth(6) - 100
        txtTtlAlloc.Width = .ColWidth(7)
        lblTtlInvAmt.Left = txtTtlInvAmt.Left - lblTtlInvAmt.Width
    End With
    With flxCust
        .ColWidth(0) = 0
        .ColWidth(1) = 1
        .ColWidth(2) = txtCustomer.Width - 300
        .ColWidth(3) = 0    'Address
        .ColWidth(4) = 0    'count Acct no
        .Width = txtCustomer.Width
    End With
    cmdA.Top = 15000
End Sub

Private Sub Timer1_Timer()
On Error GoTo er
    Dim Qry As String
    
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"
    
    rsCustomer.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxCust.DataSource = rsCustomer
    With flxCust
        .Width = txtCustomer.Width
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = txtCustomer.Width
        .ColWidth(3) = 0    'address
    End With
    'Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CONTACTPERSON, CUSTOMER.OSBAL, CUSTOMER.CTYPE " & _
    '      "From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"
    'With rsAcctCust
    '    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    If .RecordCount > 0 Then
    '        .MoveFirst
    '        'cboAccount.Clear
    '        Do While Not .EOF
    '            cboAccount.AddItem !CustomerName
    '            .MoveNext
    '        Loop
    '    End If
    'End With
    With cboChkType
        .AddItem "CUSTOMER CHECK"
        .AddItem "PETTY CASH"
        .AddItem "RE-DISCOUNT"
        .AddItem "OTHER CHECKS"
    End With
    With cboChkStat
        .AddItem "CURRENT-CLEARED"
        .AddItem "PDC DUE"
        .AddItem "POST-DATED"
    End With
    Timer1.Enabled = False

Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->Timer1_Timer()"
End Sub

Private Sub tmrEnabledtmrShowRecord_Timer()
On Error GoTo er
    If tmrEnabledtmrShowRecord.Enabled = False Then Exit Sub
    
    tmrEnabledtmrShowRecord.Enabled = False
    tmrShowRecord.Enabled = True
    Call Unload(frm_OnProcess)
    frm_OnProcess.OwnerForm = "arPR"
    frm_OnProcess.Show vbModal
Exit Sub
er:
    tmrShowRecord.Enabled = False
    tmrEnabledtmrShowRecord.Enabled = False
    MotoRep.Enabled = True
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
End Sub

Private Sub tmrShowRecord_Timer()
On Error GoTo er
    Dim Qry As String
    Dim CntCHKSPRID As Integer
    If Trim(CHKSPRID) = "" Then Exit Sub
    If tmrShowRecord.Enabled = False Then Exit Sub
    tmrShowRecord.Enabled = False
    MotoRep.Enabled = False
    Screen.MousePointer = vbHourglass
    
    Call ShowRecord(CLng(CHKSPRID))
    DoEvents
    
    Qry = "SELECT Cast(ifnull(Count(CHKSPRID),0) as Char) as CHKSPRID FROM checksledger Where CHKSPRID=" & CHKSPRID & ""
    CntCHKSPRID = GetFieldValue(Qry)
    If CntCHKSPRID > 0 Then
        txtChkAmt.Locked = True
        txtChkAmt.BackColor = txtHoldGrayBackColor.BackColor
        
        txtChkBal.Locked = True
        txtChkBal.BackColor = txtHoldGrayBackColor.BackColor
    Else
        txtChkAmt.Locked = False
        txtChkAmt.BackColor = txtHoldWhiteBackColor.BackColor
        
        txtChkBal.Locked = False
        txtChkBal.BackColor = txtHoldWhiteBackColor.BackColor
    End If
    
    
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    MotoRep.Enabled = True
    Screen.MousePointer = vbDefault
Exit Sub
er:
    Call Unload(frm_OnProcess)
    Set frm_OnProcess = Nothing
    MotoRep.Enabled = True
    Screen.MousePointer = vbDefault
    tmrShowRecord.Enabled = False
    MsgBox Err.Description, vbCritical, "Error->tmrShowRecord_Timer()"
End Sub

Private Sub txtAcctName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtBank.SetFocus
    If KeyCode = vbKeyUp Then txtPRDate.SetFocus
End Sub

Private Sub txtAcctName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBank.SetFocus
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtPRno.SetFocus
    If KeyCode = vbKeyUp Then txtCustomer.SetFocus
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPRno.SetFocus
End Sub

Private Sub txtAllocAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtAllocAmt.Text = Empty
        flxAOP.TextMatrix(ActRow, 7) = txtAllocAmt.Tag
        txtAllocAmt.Tag = "skiplostfocus"
        flxAOP.SetFocus
        flxAOP.Col = 7
        txtAllocAmt.Visible = False
    End If

End Sub

Private Sub txtAllocAmt_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxAOP.SetFocus
        flxAOP.Col = 8
    End If

End Sub

Private Sub txtAllocAmt_LostFocus()
    If txtAllocAmt.Tag = "skiplostfocus" Then
        txtAllocAmt.Tag = ""
        Exit Sub
    End If
    flxAOP.TextMatrix(ActRow, 7) = ""
    Call ComputeTotal
    If Val(Replace(txtTtlAlloc.Text, ",", "")) + Val(txtAllocAmt.Text) > Val(Replace(txtChkBal.Text, ",", "")) Then
        MsgBox "Total allocation must no be more than CHECK BALANCE!", vbExclamation + vbOKOnly
        Exit Sub
    End If
    If Val(Replace(txtAllocAmt.Text, ",", "")) > Val(Replace(flxAOP.TextMatrix(ActRow, 6), ",", "")) Then
        MsgBox "Total allocation must no be more than Invoice Amount BALANCE!", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    
    flxAOP.TextMatrix(ActRow, 7) = Format(txtAllocAmt.Text, "#,##0.00")
    txtTtlAlloc.Text = Format(Val(Replace(txtTtlAlloc.Text, ",", "")) + Val(Replace(txtAllocAmt.Text, ",", "")), "#,##0.00")
    txtAllocAmt.Visible = False
    

End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtAmount.Text = ""
        flxAOP.TextMatrix(ActRow, 4) = txtAmount.Tag
        txtAmount.Tag = "skiplostfocus"
        flxAOP.SetFocus
        flxAOP.Col = 4
        txtAmount.Visible = False
    End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Call txtAmount_LostFocus
        flxAOP.SetFocus
        flxAOP.Col = 0
    End If
End Sub

Private Sub txtAmount_LostFocus()
    If txtAmount.Tag = "skiplostfocus" Then
        txtAmount.Tag = ""
        Exit Sub
    End If
    flxAOP.TextMatrix(ActRow, 4) = IIf(txtAmount.Text <> "", Format$(txtAmount.Text, "###,##0.00"), "")
    txtAmount.Visible = False
    Call ComputeTotal
End Sub

Private Sub ComputeTotal()
    Dim InvTtl As Variant, AllocTtl As Variant
    Dim i As Integer
    
    InvTtl = 0
    AllocTtl = 0
    For i = 1 To flxAOP.Rows - 1
        InvTtl = InvTtl + Val(Replace(flxAOP.TextMatrix(i, 6), ",", ""))
        If Not AddNew Then
            If flxAOP.TextMatrix(i, 8) = "" Then
                AllocTtl = AllocTtl + Val(Replace(flxAOP.TextMatrix(i, 7), ",", ""))
                'AllocTtl = AllocTtl + flxAOP.TextMatrix(i, 7)
            Else
                'If Val(Replace(flxAOP.TextMatrix(i, 7), ",", "")) < Val(Replace(flxAOP.TextMatrix(i, 6), ",", "")) Then
                '    AllocTtl = AllocTtl + Val(Replace(flxAOP.TextMatrix(i, 7), ",", ""))
                'End If
            End If
        Else
            AllocTtl = AllocTtl + Val(Replace(flxAOP.TextMatrix(i, 7), ",", ""))
        End If
    Next i
    
    txtTtlInvAmt.Text = Format(InvTtl, "#,##0.00")
    txtTtlAlloc.Text = Format(AllocTtl, "#,##0.00")
    
End Sub

Private Sub ComputeTotalAlloc()
    Dim AllocTtl As Single
    Dim i As Integer
    
    AllocTtl = 0
    For i = 1 To flxAOP.Rows - 1
        If flxAOP.TextMatrix(i, 8) = "" Then
            AllocTtl = AllocTtl + Val(Replace(flxAOP.TextMatrix(i, 7), ",", ""))
        End If
    Next i
    
    txtTtlAlloc.Text = Format(AllocTtl, "#,##0.00")
    
End Sub


Private Sub txtBank_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtPaidTo.SetFocus
    If KeyCode = vbKeyUp Then txtAcctName.SetFocus
End Sub

Private Sub txtBank_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPaidTo.SetFocus
End Sub

Private Sub txtChkAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then cboChkType.SetFocus
    If KeyCode = vbKeyUp Then txtChkDate.SetFocus
End Sub

Private Sub txtChkAmt_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        cboChkType.SetFocus
    End If
End Sub

Private Sub txtChkAmt_LostFocus()
    If txtChkAmt.Text <> "" Then
        txtChkAmt.Text = FormatNumber(txtChkAmt.Text, 2, vbTrue)
        If AddNew Then txtChkBal.Text = FormatNumber(txtChkAmt.Text, 2, vbTrue)
    End If
End Sub

Private Sub txtChkDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtChkAmt.SetFocus
    If KeyCode = vbKeyUp Then txtChkNo.SetFocus
End Sub

Private Sub txtChkDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtChkAmt.SetFocus
End Sub

Private Sub txtChkDate_LostFocus()
    If IsDate(txtPRDate.Text) And IsDate(txtChkDate.Text) Then
        If cboChkStat.ListIndex = -1 Or cboChkStat.Text = "" Then
            If txtPRDate.Text <= txtChkDate.Text Then
                cboChkStat.ListIndex = 1
            Else
                cboChkStat.ListIndex = 2
            End If
        End If
    End If
End Sub

Private Sub txtChkNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtChkDate.SetFocus
    If KeyCode = vbKeyUp Then cboChkStat.SetFocus
End Sub

Private Sub txtChkNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtChkDate.SetFocus
End Sub

Private Sub cboChkStat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtChkNo.SetFocus
End Sub

Private Sub txtCustomer_Change()
    With flxCust
        If CustSrch Then
            CustSrch = False
            Exit Sub
        End If
        If txtCustomer.Text = "" Then .Visible = False: Exit Sub
        rsCustomer.Find "[CUSTOMERNAME] LIKE '" + txtCustomer.Text + "%'", , adSearchForward, 1
        If Not rsCustomer.EOF Then
            .Top = Frame1.Top + txtCustomer.Top + txtCustomer.Height + 30
            .Left = Frame1.Left + txtCustomer.Left
            .Width = txtCustomer.Width
            .Visible = True
            If Not CustBln Then
                txtCustomer.Tag = Len(txtCustomer.Text)
                CustBln = True
            End If
            .TopRow = rsCustomer.AbsolutePosition - 1
            .Row = rsCustomer.AbsolutePosition - 1
            .Col = 0: .ColSel = 3
            txtCustomer.Text = .TextMatrix(.Row, 2)
            'Insert Additional retrieval here
            txtAddress.Text = .TextMatrix(.Row, 3)
            txtCustID.Text = .TextMatrix(.Row, 1)
            txtAcctName.Text = .TextMatrix(.Row, 2)
                 
            txtCustomer.SelStart = Val(txtCustomer.Tag)
            txtCustomer.SelLength = Len(txtCustomer.Text)
            CustBln = False
        Else
            'rsCustomer.Requery
            txtAddress.Text = ""
            txtCustID.Text = ""
            txtAcctName.Text = ""
            .Visible = False
        End If
    End With
End Sub

Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If flxCust.Visible Then
                flxCust.SetFocus
            Else
                txtAddress.SetFocus
            End If
        Case vbKeyEscape
            txtCustomer.Text = ""
            flxCust.Visible = False
        Case vbKeyDelete
            CustSrch = True
            flxCust.Visible = False
        Case vbKeyBack
            CustSrch = True
            flxCust.Visible = False
    End Select

End Sub

Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxCust.Visible = False
        txtAddress.SetFocus
    End If
End Sub

Private Sub txtCustomer_LostFocus()
    txtPaidTo.Text = StrConv(SysBranch, vbProperCase)
    cboChkType.ListIndex = 0
    If Val(txtCustID.Text) > 0 Then
        ActRow = 1
        'cboAccount.Text = txtCustomer.Text
        cmdAutoAlloc.Visible = True
    End If
End Sub


Private Sub txtInvNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtInvNo.Text = Empty
        flxAOP.TextMatrix(ActRow, 2) = txtInvNo.Tag
        txtInvNo.Tag = "skiplostfocus"
        flxAOP.SetFocus
        flxAOP.Col = 2
        txtInvNo.Visible = False
    End If
End Sub

Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call txtInvNo_LostFocus
        flxAOP.SetFocus
        flxAOP.Col = 2
    End If

End Sub

Private Sub txtInvNo_LostFocus()
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String, X As Integer
    Dim MyJo As String, MyPCV As String
    
    If txtInvNo.Tag = "skiplostfocus" Then
        txtInvNo.Tag = ""
        Exit Sub
    End If
    If txtInvNo.Text <> "" And flxAOP.TextMatrix(ActRow, 1) <> "" Then 'Inv & Branch name not blank
        If ActRow > 1 And flxAOP.Rows > 2 Then
            X = InvExistInFlx(txtInvNo.Text, ActRow)
            If X > 0 Then
                MsgBox "Invoice No. " & txtInvNo.Text & " already exist in  row " & X & "!", vbExclamation + vbOKOnly, "Invoice Exist"
                Exit Sub
            End If
        End If
        If flxAOP.TextMatrix(ActRow, 1) = ThisBranch Then
            If UCase$(Left$(txtInvNo.Text, 3)) = "JO#" Then
                MyJo = Replace$(txtInvNo.Text, "JO#", "")
                Qry = "SELECT onaccterms.SPINVNO, onaccterms.TRANSDATE, onaccterms.TERMS, onaccterms.ACCTAMOUNT, onaccterms.SPSORDERID, customer.CUSTOMERNAME, customer.CUSTOMERID, onaccterms.`STATUS`, onaccterms.BALANCE, onaccterms.JOBORDERNO, onaccterms.JOBORDERID " & _
                      "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) " & _
                      "Where onaccterms.JOBORDERNO = '" & MyJo & "' AND onaccterms.SPINVNO IS NULL AND onaccterms.JOBORDERID > 1 Order By onaccterms.TRANSDATE DESC"
            
            ElseIf UCase$(Left$(txtInvNo.Text, 4)) = "PCV#" Then
                MyJo = ""
                MyPCV = Replace$(txtInvNo.Text, "PCV#", "")
                Qry = "SELECT disbursement.CVNUMBER, disbursement.CVDATE, disbursement.AMOUNT, customer.CUSTOMERID, customer.CUSTOMERNAME, Disbursement.VOUCHERID " & _
                        "From Disbursement INNER JOIN customer ON (disbursement.CUSTOMERID = customer.CUSTOMERID) " & _
                            "Where disbursement.CVNUMBER = '" & MyPCV & "' AND disbursement.CUSTOMERID > 0"
            
            Else
                MyPCV = ""
                MyJo = ""
                Qry = "SELECT onaccterms.SPINVNO, onaccterms.TRANSDATE, onaccterms.TERMS, onaccterms.ACCTAMOUNT, onaccterms.SPSORDERID, customer.CUSTOMERNAME, customer.CUSTOMERID, onaccterms.`STATUS`, onaccterms.BALANCE, onaccterms.JOBORDERNO, onaccterms.JOBORDERID  " & _
                      "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) " & _
                      "Where onaccterms.SPINVNO = '" & txtInvNo.Text & "' Order By onaccterms.TRANSDATE Desc;"
            End If
            rsInv.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsInv.RecordCount > 0 Then
                rsInv.MoveFirst
'FLXAOP:1-BRANCH,2-INV.NO,3-INV.DATE,4-TERM,5-DUEDATE,6-INV.AMT,7-ALLOC.AMT,8-OR.NO,9-CUSTOMER,10-SPSORDERID,11-CustomerID
                With flxAOP
                    .TextMatrix(ActRow, 2) = txtInvNo.Text
                    '.TextMatrix(ActRow, 3) = IIf(MyPCV <> "", Format(rsInv!CVDate, "mm/dd/yyyy"), Format(rsInv!TRANSDATE, "mm/dd/yyyy"))
                        If MyPCV <> "" Then
                            .TextMatrix(ActRow, 3) = Format(rsInv!CVDate, "mm/dd/yyyy")
                        Else
                            .TextMatrix(ActRow, 3) = Format(rsInv!TRANSDATE, "mm/dd/yyyy")
                        End If
                    '.TextMatrix(ActRow, 4) = IIf(MyPCV <> "", "", rsInv!Terms)
                        If MyPCV <> "" Then
                            .TextMatrix(ActRow, 4) = ""
                        Else
                            .TextMatrix(ActRow, 4) = rsInv!Terms
                        End If
                    '.TextMatrix(ActRow, 5) = IIf(MyPCV <> "", "", DateAdd("d", rsInv!Terms, rsInv!TRANSDATE)) 'Format(GetDueDate(rsInv!TransDate, rsInv!Terms), "mm/dd/yyyy")
                        If MyPCV <> "" Then
                            .TextMatrix(ActRow, 5) = ""
                        Else
                            .TextMatrix(ActRow, 5) = DateAdd("d", rsInv!Terms, rsInv!TRANSDATE)
                        End If
                    '.TextMatrix(ActRow, 6) = IIf(MyPCV <> "", Format(rsInv!Amount, "#,##0.00"), Format(rsInv!Balance, "#,##0.00"))
                        If MyPCV <> "" Then
                            .TextMatrix(ActRow, 6) = Format(rsInv!Amount, "#,##0.00")
                        Else
                            .TextMatrix(ActRow, 6) = Format(rsInv!Balance, "#,##0.00")
                        End If
                    If Val(txtChksPRid.Text) > 0 Then
                        If MyPCV <> "" Then
                            Qry = "SELECT checkalloc.ALLOCAMT From checkalloc Where checkalloc.CUSTOMERID = " & rsInv!CUSTOMERID & " AND checkalloc.CHKSPRID = " & Val(txtChksPRid.Text) & " AND checkalloc.VOUCHERID = " & .TextMatrix(ActRow, 13) & ";"
                        Else
                            Qry = "SELECT checkalloc.ALLOCAMT From checkalloc Where checkalloc.CUSTOMERID = " & rsInv!CUSTOMERID & " AND checkalloc.CHKSPRID = " & Val(txtChksPRid.Text) & " AND " & IIf(MyJo <> "", "checkalloc.JOBORDERID = " & rsInv!JOBORDERID, "checkalloc.SPSORDERID = " & rsInv!SPSORDERID) & ";"
                        End If
                        .TextMatrix(ActRow, 7) = Format(Val(GetFieldValue(Qry)), "#,##0.00")
                    Else
                        .TextMatrix(ActRow, 7) = ""
                    End If
                    .TextMatrix(ActRow, 9) = rsInv!CustomerName
                        
                        'If MyJo <> 0 Then
                        '    .TextMatrix(ActRow, 10) = rsInv!JOBORDERID
                        If MyPCV <> "" Then
                            .TextMatrix(ActRow, 10) = rsInv!VOUCHERID
                        ElseIf MyJo <> "" Then
                            .TextMatrix(ActRow, 10) = rsInv!JOBORDERID
                        Else
                            .TextMatrix(ActRow, 10) = rsInv!SPSORDERID
                        End If
                    .TextMatrix(ActRow, 11) = rsInv!CUSTOMERID
                        
                    If ShowRec Then
                        
                        If MyPCV <> "" Then
                            .TextMatrix(ActRow, 6) = Format(rsInv!Amount, "#,##0.00")
                        Else
                            .TextMatrix(ActRow, 6) = Format(rsInv!AcctAmount, "#,##0.00")
                        End If
                        
                        If MyJo <> "" Then
                            Qry = "SELECT Ledger.RefNo From Ledger Where ledger.JOBID = " & rsInv!JOBORDERID & " AND ledger.SPINVNO = '" & MyJo & "' AND Ledger.CUSTOMERID = " & rsInv!CUSTOMERID & " Order By Ledger.REFDATE Desc;"
                        Else
                            Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & RefNo & "' AND ledger.REFTYPE = 'Payment' AND Ledger.CUSTOMERID = " & rsInv!CUSTOMERID & " Order By Ledger.REFDATE Desc;"
                        End If
                        .TextMatrix(ActRow, 8) = GetFieldValue(Qry)
                    Else
                        If MyPCV = "" Then
                            If rsInv!STATUS = "CLOSED" Or rsInv!Balance = 0 Then
                                MsgBox IIf(MyJo <> "", "", "Invoice No. ") & txtInvNo.Text & " already CLOSED!", vbExclamation + vbOKOnly, "Account Closed"
                                .TextMatrix(ActRow, 7) = ""
                                If MyJo <> "" Then
                                    Qry = "SELECT Ledger.RefNo From Ledger Where ledger.JOBID = " & rsInv!JOBORDERID & " AND ledger.SPINVNO = '" & MyJo & "' AND Ledger.CUSTOMERID = " & rsInv!CUSTOMERID & " Order By Ledger.REFDATE Desc;"
                                Else
                                    Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & RefNo & "' AND ledger.REFTYPE = 'Payment' AND Ledger.CUSTOMERID = " & rsInv!CUSTOMERID & " Order By Ledger.REFDATE Desc;"
                                End If
                                'Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & RefNo & "' AND ledger.REFTYPE = 'Payment' Order By Ledger.REFDATE Desc;"
                                .TextMatrix(ActRow, 8) = GetFieldValue(Qry)
                            End If
                        End If
                    End If
                End With
            Else
                MsgBox "No record found for " & IIf(MyJo <> "", "", "Invoice No. ") & txtInvNo.Text & "!", vbExclamation + vbOKOnly, "Record not found"
                rsInv.Close
                Exit Sub
            End If
            rsInv.Close
        Else
            If ThisBranch = "MRC2" Then
                If flxAOP.TextMatrix(ActRow, 1) = "VAN1" Then
                    If ConnectVAN1 Then Call ShowBranchInv(flxAOP.TextMatrix(ActRow, 1), txtInvNo.Text)
                ElseIf flxAOP.TextMatrix(ActRow, 1) = "VAN2" Then
                    If ConnectVAN2 Then Call ShowBranchInv(flxAOP.TextMatrix(ActRow, 1), txtInvNo.Text)
                ElseIf flxAOP.TextMatrix(ActRow, 1) = "VAN3" Then
                    If ConnectVan3 Then Call ShowBranchInv(flxAOP.TextMatrix(ActRow, 1), txtInvNo.Text)
                ElseIf flxAOP.TextMatrix(ActRow, 1) = "MRC1" Then
                    If ConnectMRC1 Then Call ShowBranchInv(flxAOP.TextMatrix(ActRow, 1), txtInvNo.Text)
                End If
            End If
        End If
    End If
    txtInvNo.Visible = False
    Call ComputeTotal
End Sub

Sub ShowBranchInv(MyBrn As String, RefNo As String)
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String, X As Integer


    Qry = "SELECT onaccterms.SPINVNO, onaccterms.TRANSDATE, onaccterms.TERMS, onaccterms.ACCTAMOUNT, onaccterms.SPSORDERID, customer.CUSTOMERNAME, customer.CUSTOMERID, onaccterms.`STATUS`, onaccterms.BALANCE " & _
          "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) " & _
          "Where onaccterms.SPINVNO = '" & RefNo & "' Order By onaccterms.TRANSDATE;"
    
    If MyBrn = "VAN1" Then
        rsInv.Open Qry, conVAN1, adOpenKeyset, adLockOptimistic
    ElseIf MyBrn = "VAN2" Then
        rsInv.Open Qry, conVAN2, adOpenKeyset, adLockOptimistic
    ElseIf MyBrn = "VAN3" Then
        rsInv.Open Qry, conVAN3, adOpenKeyset, adLockOptimistic
    ElseIf MyBrn = "MRC2" Then
        rsInv.Open Qry, conMRC2, adOpenKeyset, adLockOptimistic
    ElseIf MyBrn = "MRC1" Then
        rsInv.Open Qry, conMRC1, adOpenKeyset, adLockOptimistic
    End If
    
    If rsInv.RecordCount > 0 Then
        rsInv.MoveFirst
'FLXAOP:1-BRANCH,2-INV.NO,3-INV.DATE,4-TERM,5-DUEDATE,6-INV.AMT,7-ALLOC.AMT,8-OR.NO,9-CUSTOMER,10-SPSORDERID,11-CustomerID
        With flxAOP
            If txtInvNo.Text <> "" Then .TextMatrix(ActRow, 2) = txtInvNo.Text
            .TextMatrix(ActRow, 3) = Format(rsInv!TRANSDATE, "mm/dd/yyyy")
            .TextMatrix(ActRow, 4) = rsInv!Terms
            .TextMatrix(ActRow, 5) = DateAdd("d", rsInv!Terms, rsInv!TRANSDATE) 'Format(GetDueDate(rsInv!TRANSDATE, rsInv!Terms), "mm/dd/yyyy")
            .TextMatrix(ActRow, 6) = Format(rsInv!Balance, "#,##0.00")
            If Val(txtChksPRid.Text) > 0 Then
                Qry = "SELECT checkalloc.ALLOCAMT From checkalloc Where checkalloc.CUSTOMERID = " & rsInv!CUSTOMERID & " AND checkalloc.SPSORDERID = " & rsInv!SPSORDERID & " AND checkalloc.CHKSPRID = " & Val(txtChksPRid.Text) & ";"
                .TextMatrix(ActRow, 7) = Format(Val(GetFieldValue(Qry)), "#,##0.00")
            Else
                .TextMatrix(ActRow, 7) = ""
            End If
            .TextMatrix(ActRow, 9) = rsInv!CustomerName
            .TextMatrix(ActRow, 10) = rsInv!SPSORDERID
            .TextMatrix(ActRow, 11) = rsInv!CUSTOMERID
            If ShowRec Then
                .TextMatrix(ActRow, 6) = Format(rsInv!AcctAmount, "#,##0.00")
                Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & RefNo & "' AND ledger.REFTYPE = 'Payment' Order By Ledger.REFDATE Desc;"
                .TextMatrix(ActRow, 8) = GetFieldValueBranch(Qry, MyBrn)
            Else
                If rsInv!STATUS = "CLOSED" Or rsInv!Balance = 0 Then
                    MsgBox "Invoice No. " & txtInvNo.Text & " already CLOSED!", vbExclamation + vbOKOnly, "Account Closed"
                    .TextMatrix(ActRow, 7) = ""
                    Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & RefNo & "' AND ledger.REFTYPE = 'Payment' Order By Ledger.REFDATE Desc;"
                    .TextMatrix(ActRow, 8) = GetFieldValueBranch(Qry, MyBrn)
                End If
            End If
        End With
    Else
        MsgBox "No record found for Invoice No. " & RefNo & "!", vbExclamation + vbOKOnly, "Record not found"
        rsInv.Close
        Exit Sub
    End If
    rsInv.Close

End Sub

Function InvExistInFlx(MyInv As String, cRow As Integer) As Integer
    Dim i As Integer
    
    With flxAOP
        For i = 1 To .Rows - 1
            If i <> cRow Then
                If .TextMatrix(i, 1) = MyInv Then
                    InvExistInFlx = i
                    Exit Function
                End If
            End If
        Next i
    End With
    InvExistInFlx = 0
End Function

Private Sub txtPaidTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then cboChkStat.SetFocus
    If KeyCode = vbKeyUp Then txtBank.SetFocus
End Sub

Private Sub txtPaidTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboChkStat.SetFocus
End Sub

Private Sub txtPRDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtAcctName.SetFocus
    If KeyCode = vbKeyUp Then txtPRno.SetFocus
End Sub

Private Sub txtPRDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtAcctName.SetFocus
End Sub

Private Sub txtPRDate_LostFocus()
    If IsDate(txtPRDate.Text) And IsDate(txtChkDate.Text) Then
        If txtPRDate.Text <= txtChkDate.Text Then
            cboChkStat.ListIndex = 1
        Else
            cboChkStat.ListIndex = 2
        End If
    End If
End Sub

Private Sub txtPRno_GotFocus()
    With txtPRno
        If Len(.Text) > 0 Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub

Private Sub txtPRno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtPRDate.SetFocus
    If KeyCode = vbKeyUp Then txtAddress.SetFocus
End Sub

Private Sub txtPRno_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890_"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtPRDate.SetFocus
End Sub

Private Sub txtPRno_LostFocus()
    If txtPRno.Text <> "" Then
        'If CtrlNumExist("SELECT CHECKS.* From CHECKS WHERE (((CHECKS.PRNO)='" & txtPRno.Text & "'));") Then
        'mySQL
        If CtrlNumExist("SELECT CHECKS.* From CHECKS WHERE CHECKS.PRNO='" & txtPRno.Text & "';") Then
           MsgBox "P.R. NUMBER already exist!", vbExclamation + vbOKOnly, "Validation Error"
        End If
    End If
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtRemarks.Text = ""
        flxAOP.TextMatrix(ActRow, 3) = txtRemarks.Tag
        txtRemarks.Tag = "skiplostfocus"
        flxAOP.SetFocus
        flxAOP.Col = 3
        txtRemarks.Visible = False
    End If
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtRemarks_LostFocus
        flxAOP.SetFocus
        flxAOP.Col = 4
    End If
End Sub

Private Sub txtRemarks_LostFocus()
    If txtRemarks.Tag = "skiplostfocus" Then
        txtRemarks.Tag = ""
        Exit Sub
    End If
    flxAOP.TextMatrix(ActRow, 3) = txtRemarks.Text
    txtRemarks.Visible = False
End Sub
